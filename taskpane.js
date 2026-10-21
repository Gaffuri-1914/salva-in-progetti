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
const ROTATABLE = { pdf: "application/pdf", jpg: "image/jpeg", jpeg: "image/jpeg", png: "image/png" };

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
  $("attachments").addEventListener("click", onRotateClick);
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
    const isCloud = a.attachmentType === Office.MailboxEnums.AttachmentType.Cloud;
    let { base, ext } = splitName(a.name);
    if (a.attachmentType === Office.MailboxEnums.AttachmentType.Item && !ext) ext = ".eml";

    const row = document.createElement("div");
    row.className = "att";
    row.dataset.id = a.id;
    row.dataset.original = a.name;
    row.dataset.ext = ext;
    row.dataset.rotation = "0";

    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = !a.isInline && !isCloud;
    cb.disabled = isCloud;
    cb.title = "Salva questo allegato";

    const name = document.createElement("input");
    name.type = "text";
    name.className = "name";
    name.value = base;
    name.disabled = isCloud;
    name.title = "Nome del file (modificabile)";

    const extSpan = document.createElement("span");
    extSpan.className = "ext";
    extSpan.textContent = ext;

    const line = document.createElement("div");
    line.className = "line";
    line.append(cb, name, extSpan);

    const info = document.createElement("div");
    info.className = "hint";
    info.textContent = `${formatSize(a.size)}${a.isInline ? " – nel testo" : ""}${isCloud ? " – link cloud, non salvabile" : ""}`;

    if (!isCloud && ROTATABLE[ext.slice(1).toLowerCase()]) {
      const left = rotateButton("↺", -90, "Ruota 90° in senso antiorario");
      const right = rotateButton("↻", 90, "Ruota 90° in senso orario");
      const preview = rotateButton("Anteprima", 0, "Mostra l'anteprima (per i PDF: prima pagina)");
      const angle = document.createElement("span");
      angle.className = "angle";
      info.append(" ", left, right, preview, angle);
    }
    row.append(line, info);
    box.append(row);
  }
}

function splitName(name) {
  const m = name.match(/^(.+?)(\.[A-Za-z0-9]{1,5})?$/);
  return m ? { base: m[1], ext: m[2] || "" } : { base: name, ext: "" };
}

function rotateButton(text, delta, title) {
  const b = document.createElement("button");
  b.type = "button";
  b.className = "rot";
  b.textContent = text;
  b.title = title;
  b.dataset.delta = delta;
  return b;
}

function onRotateClick(e) {
  const btn = e.target.closest("button.rot");
  if (!btn) return;
  const row = btn.closest(".att");
  const rotation = (Number(row.dataset.rotation) + Number(btn.dataset.delta) + 360) % 360;
  row.dataset.rotation = rotation;
  row.querySelector(".angle").textContent = rotation ? ` ruotato ${rotation}° orario` : "";
  showPreview(row);
}

async function showPreview(row) {
  let box = row.querySelector(".preview");
  if (!box) {
    box = document.createElement("div");
    box.className = "preview";
    box.textContent = "Caricamento anteprima…";
    row.append(box);
    try {
      const bytes = await attachmentBytes(row.dataset.id);
      const ext = row.dataset.ext.slice(1).toLowerCase();
      const img = document.createElement("img");
      img.alt = "Anteprima";
      img.src = ext === "pdf" ? await pdfThumbnail(bytes) : URL.createObjectURL(new Blob([bytes], { type: ROTATABLE[ext] }));
      box.replaceChildren(img);
    } catch (e) {
      box.textContent = `Anteprima non disponibile: ${e.message || e}`;
      return;
    }
  }
  const img = box.querySelector("img");
  if (img) img.style.transform = `rotate(${row.dataset.rotation}deg)`;
}

let pdfjsModule;
async function pdfThumbnail(bytes) {
  const pdfjs = await (pdfjsModule ??= import("./pdf.min.mjs"));
  pdfjs.GlobalWorkerOptions.workerSrc = new URL("pdf.worker.min.mjs", location.href).href;
  // pdf.js detaches the buffer it receives, so pass a copy
  const task = pdfjs.getDocument({ data: bytes.slice() });
  const doc = await task.promise;
  try {
    const page = await doc.getPage(1);
    const base = page.getViewport({ scale: 1 });
    const viewport = page.getViewport({ scale: 340 / Math.max(base.width, base.height) });
    const canvas = document.createElement("canvas");
    canvas.width = Math.ceil(viewport.width);
    canvas.height = Math.ceil(viewport.height);
    await page.render({ canvasContext: canvas.getContext("2d"), viewport }).promise;
    return canvas.toDataURL("image/png");
  } finally {
    task.destroy();
  }
}

function formatSize(bytes) {
  return bytes > 1048576 ? `${(bytes / 1048576).toFixed(1)} MB` : `${Math.max(1, Math.round(bytes / 1024))} KB`;
}

function selectedAttachments() {
  return [...$("attachments").querySelectorAll(".att")].filter((row) => row.querySelector("input[type=checkbox]").checked);
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

const contentCache = new Map();
function attachmentBytes(id) {
  if (!contentCache.has(id)) {
    contentCache.set(id, getAttachmentContent(id).then(contentToBytes).catch((e) => {
      contentCache.delete(id);
      throw e;
    }));
  }
  return contentCache.get(id);
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

async function rotate(bytes, ext, angle) {
  if (ext === "pdf") {
    await loadScript("pdf-lib.min.js");
    const doc = await PDFLib.PDFDocument.load(bytes, { ignoreEncryption: true });
    for (const page of doc.getPages()) {
      page.setRotation(PDFLib.degrees((page.getRotation().angle + angle) % 360));
    }
    return doc.save();
  }
  const type = ROTATABLE[ext];
  const img = await createImageBitmap(new Blob([bytes], { type }));
  const canvas = document.createElement("canvas");
  const sideways = angle % 180 !== 0;
  canvas.width = sideways ? img.height : img.width;
  canvas.height = sideways ? img.width : img.height;
  const ctx = canvas.getContext("2d");
  ctx.translate(canvas.width / 2, canvas.height / 2);
  ctx.rotate((angle * Math.PI) / 180);
  ctx.drawImage(img, -img.width / 2, -img.height / 2);
  const blob = await new Promise((resolve) => canvas.toBlob(resolve, type, 0.92));
  return new Uint8Array(await blob.arrayBuffer());
}

const scripts = {};
function loadScript(src) {
  return (scripts[src] ??= new Promise((resolve, reject) => {
    const s = document.createElement("script");
    s.src = src;
    s.onload = resolve;
    s.onerror = () => reject(new Error(`Impossibile caricare ${src}`));
    document.head.append(s);
  }));
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

  for (const row of selectedAttachments()) {
    try {
      let bytes = await attachmentBytes(row.dataset.id);
      const ext = row.dataset.ext;
      const rotation = Number(row.dataset.rotation);
      if (rotation) bytes = await rotate(bytes, ext.slice(1).toLowerCase(), rotation);
      const base = row.querySelector(".name").value.trim() || splitName(row.dataset.original).base;
      const item = await upload(folderPath, safeFileName(base + ext), bytes);
      saved.push(item.name);
    } catch (e) {
      failed.push(`${row.dataset.original}: ${e.message || e}`);
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
