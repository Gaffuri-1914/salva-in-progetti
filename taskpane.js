"use strict";

const CONFIG = {
  clientId: "8f6460ec-3e33-4b3a-ae30-99b1b15c7d4d",
  authority: "https://login.microsoftonline.com/39ec9c38-5f9f-4fc5-97e9-315e5b0ca844",
  redirectUri: "https://gaffuri-1914.github.io/salva-in-progetti/redirect.html",
  driveId: "b!QwxxLPRUY0W2n6m_gXiZ1__Bq8JA22tCggLffX4ovNQmQLwSt0JzS5JfzAnsczno",
  scopes: ["Files.ReadWrite.All"],
};
const GRAPH = "https://graph.microsoft.com/v1.0";
const SIMPLE_UPLOAD_MAX = 4 * 1024 * 1024;
const CHUNK = 320 * 1024 * 12; // upload session chunks must be multiples of 320 KiB
const ROOT_OPTION = "__root__";

const $ = (id) => document.getElementById(id);
let pca;
let projects = [];
let subfolders = [];

Office.onReady(async () => {
  renderAttachments();
  if (!Office.context.requirements.isSetSupported("Mailbox", "1.14")) {
    $("saveMail").disabled = true;
    $("saveMail").parentElement.title = "Non supportato da questa versione di Outlook";
  }
  $("search").addEventListener("input", filterProjects);
  $("project").addEventListener("change", selectProject);
  $("folder").addEventListener("change", updateSaveButton);
  $("saveMail").addEventListener("change", updateSaveButton);
  $("attachments").addEventListener("change", updateSaveButton);
  $("save").addEventListener("click", save);
  try {
    await loadProjects();
    $("search").disabled = false;
    $("projectInfo").textContent = `${projects.length} progetti. Scrivi il numero o parte del nome.`;
    $("search").focus();
  } catch (e) {
    showError("Impossibile caricare i progetti", e);
  }
});

// ---------- autenticazione ----------
async function getToken() {
  if (!pca) {
    pca = await msal.createNestablePublicClientApplication({
      auth: { clientId: CONFIG.clientId, authority: CONFIG.authority, redirectUri: CONFIG.redirectUri },
      cache: { cacheLocation: "localStorage" },
    });
  }
  const account = pca.getActiveAccount() || pca.getAllAccounts()[0];
  const request = { scopes: CONFIG.scopes, account };
  try {
    return (await pca.acquireTokenSilent(request)).accessToken;
  } catch {
    const result = await pca.acquireTokenPopup(request);
    pca.setActiveAccount(result.account);
    return result.accessToken;
  }
}

async function graph(path, options = {}) {
  const token = await getToken();
  const res = await fetch(path.startsWith("https://") ? path : GRAPH + path, {
    ...options,
    headers: { Authorization: `Bearer ${token}`, ...(options.headers || {}) },
  });
  if (!res.ok) throw new Error(`Graph ${res.status}: ${await res.text()}`);
  return res.status === 204 ? null : res.json();
}

async function listChildren(itemPath) {
  let url = `/drives/${CONFIG.driveId}/${itemPath}/children?$select=id,name,folder,webUrl&$top=999`;
  const items = [];
  while (url) {
    const page = await graph(url);
    items.push(...page.value);
    url = page["@odata.nextLink"];
  }
  return items.filter((i) => i.folder).sort((a, b) => a.name.localeCompare(b.name, "it", { numeric: true }));
}

// ---------- progetti / sottocartelle ----------
async function loadProjects() {
  projects = (await listChildren("root")).reverse();
}

function filterProjects() {
  const q = $("search").value.trim().toLowerCase();
  const list = $("project");
  list.replaceChildren();
  resetFolders();
  if (!q) { list.disabled = true; return; }
  const matches = projects.filter((p) => /^\d+$/.test(q) ? p.name.startsWith(q) : p.name.toLowerCase().includes(q)).slice(0, 50);
  for (const p of matches) list.add(new Option(p.name, p.id));
  list.disabled = matches.length === 0;
  $("projectInfo").textContent = matches.length ? `${matches.length} risultati` : "Nessun progetto trovato";
  if (matches.length === 1) { list.selectedIndex = 0; selectProject(); }
}

function resetFolders() {
  subfolders = [];
  $("folder").replaceChildren();
  $("folder").disabled = true;
  updateSaveButton();
}

async function selectProject() {
  const projectId = $("project").value;
  resetFolders();
  if (!projectId) return;
  $("projectInfo").textContent = "Caricamento sottocartelle…";
  try {
    subfolders = await listChildren(`items/${projectId}`);
    const sel = $("folder");
    sel.add(new Option("(cartella principale del progetto)", ROOT_OPTION));
    for (const f of subfolders) sel.add(new Option(f.name, f.id));
    const last = localStorage.getItem("sip-lastFolder");
    const match = subfolders.find((f) => f.name === last);
    sel.value = match ? match.id : (subfolders[0]?.id ?? ROOT_OPTION);
    sel.disabled = false;
    $("projectInfo").textContent = $("project").selectedOptions[0].text;
  } catch (e) {
    showError("Impossibile leggere le sottocartelle", e);
  }
  updateSaveButton();
}

// ---------- allegati ----------
function renderAttachments() {
  const box = $("attachments");
  const atts = Office.context.mailbox.item.attachments || [];
  if (!atts.length) { box.textContent = "Nessun allegato"; return; }
  for (const a of atts) {
    const label = document.createElement("label");
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.value = a.id;
    cb.dataset.name = a.name;
    cb.dataset.type = a.attachmentType;
    const isCloud = a.attachmentType === Office.MailboxEnums.AttachmentType.Cloud;
    cb.checked = !a.isInline && !isCloud;
    cb.disabled = isCloud;
    label.append(cb, ` ${a.name} (${formatSize(a.size)})${a.isInline ? " – nel testo" : ""}${isCloud ? " – link cloud, non salvabile" : ""}`);
    box.append(label);
  }
}

function formatSize(bytes) {
  return bytes > 1048576 ? `${(bytes / 1048576).toFixed(1)} MB` : `${Math.max(1, Math.round(bytes / 1024))} KB`;
}

function selectedAttachments() {
  return [...$("attachments").querySelectorAll("input:checked")];
}

function updateSaveButton() {
  const hasTarget = !$("folder").disabled && $("folder").value;
  const hasContent = selectedAttachments().length > 0 || $("saveMail").checked;
  $("save").disabled = !(hasTarget && hasContent);
}

function getAttachmentContent(id) {
  return new Promise((resolve, reject) =>
    Office.context.mailbox.item.getAttachmentContentAsync(id, (r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error)));
}

function getMailAsEml() {
  return new Promise((resolve, reject) =>
    Office.context.mailbox.item.getAsFileAsync((r) =>
      r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject(r.error)));
}

function base64ToBytes(b64) {
  const bin = atob(b64);
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return bytes;
}

function contentToBytes(content) {
  const F = Office.MailboxEnums.AttachmentContentFormat;
  return content.format === F.Base64 ? base64ToBytes(content.content) : new TextEncoder().encode(content.content);
}

function safeFileName(name) {
  const cleaned = name.replace(/[\\/:*?"<>|\x00-\x1f]/g, "_").replace(/\s+/g, " ").trim().replace(/^[.\s]+|[.\s]+$/g, "");
  return (cleaned || "file").slice(0, 150);
}

// ---------- upload ----------
async function upload(folderPath, fileName, bytes) {
  const target = `/drives/${CONFIG.driveId}/${folderPath}:/${encodeURIComponent(fileName)}:`;
  if (bytes.length <= SIMPLE_UPLOAD_MAX) {
    return graph(`${target}/content?@microsoft.graph.conflictBehavior=rename`, { method: "PUT", body: bytes });
  }
  const session = await graph(`${target}/createUploadSession`, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "rename" } }),
  });
  let result;
  for (let start = 0; start < bytes.length; start += CHUNK) {
    const end = Math.min(start + CHUNK, bytes.length);
    // uploadUrl is pre-authenticated: no Authorization header allowed
    const res = await fetch(session.uploadUrl, {
      method: "PUT",
      headers: { "Content-Range": `bytes ${start}-${end - 1}/${bytes.length}` },
      body: bytes.slice(start, end),
    });
    if (!res.ok) throw new Error(`Upload ${res.status}: ${await res.text()}`);
    result = await res.json();
  }
  return result;
}

async function save() {
  const folderId = $("folder").value === ROOT_OPTION ? $("project").value : $("folder").value;
  const folderPath = `items/${folderId}`;
  const folderName = $("folder").selectedOptions[0].text;
  localStorage.setItem("sip-lastFolder", folderName);
  $("save").disabled = true;
  setStatus("Salvataggio in corso…");
  const saved = [];
  const failed = [];

  for (const cb of selectedAttachments()) {
    try {
      const content = await getAttachmentContent(cb.value);
      let name = cb.dataset.name;
      if (cb.dataset.type === Office.MailboxEnums.AttachmentType.Item && !/\.\w{2,4}$/.test(name)) name += ".eml";
      const item = await upload(folderPath, safeFileName(name), contentToBytes(content));
      saved.push(item.name);
    } catch (e) {
      failed.push(`${cb.dataset.name}: ${e.message || e}`);
    }
  }

  if ($("saveMail").checked) {
    try {
      const item = Office.context.mailbox.item;
      const date = (item.dateTimeCreated || new Date()).toISOString().slice(0, 10);
      const fileName = safeFileName(`${date} - ${item.subject || "senza oggetto"}`).slice(0, 140) + ".eml";
      const uploaded = await upload(folderPath, fileName, base64ToBytes(await getMailAsEml()));
      saved.push(uploaded.name);
    } catch (e) {
      failed.push(`Mail: ${e.message || e}`);
    }
  }

  const project = $("project").selectedOptions[0].text;
  let msg = saved.length ? `Salvati in "${project}" › ${folderName}:\n• ${saved.join("\n• ")}` : "";
  if (failed.length) msg += `${msg ? "\n\n" : ""}Errori:\n• ${failed.join("\n• ")}`;
  setStatus(msg, failed.length ? "err" : "ok");
  updateSaveButton();
}

// ---------- UI ----------
function setStatus(text, cls = "") {
  $("status").textContent = text;
  $("status").className = cls;
}

function showError(prefix, e) {
  console.error(e);
  setStatus(`${prefix}: ${e.message || e}`, "err");
  $("projectInfo").textContent = "";
}
