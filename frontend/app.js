const state = {
  preview: [],
  showName: "",
  saveDirHandle: null,
  saveDirName: "Not selected",
};

const $ = (id) => document.getElementById(id);

function setStatus(message, isError = false) {
  const el = $("opsStatus");
  el.textContent = message;
  el.style.color = isError ? "#b91c1c" : "#0f766e";
}

function updateSaveLocationUi() {
  const el = $("saveLocationValue");
  if (!el) return;
  el.value = state.saveDirName || "Not selected";
}

function setDownloadButtonVisibility(isVisible) {
  const btn = $("downloadAllBtn");
  if (!btn) return;
  btn.hidden = !isVisible;
  if (!isVisible) {
    btn.disabled = false;
    btn.textContent = "Download All PDFs (.zip)";
  }
}

function resetDownloadAvailability() {
  setDownloadButtonVisibility(false);
}

function getFiles() {
  const csvFile = $("allocationCsvFile").files[0];
  const pdfFile = $("ticketsPdfFile").files[0];
  if (!csvFile || !pdfFile) {
    throw new Error("Please select both allocation CSV and tickets PDF.");
  }
  return { csvFile, pdfFile };
}

function inferShowNameFromFileName(name) {
  if (!name) return "";
  let base = name.replace(/\.[^.]+$/, "").trim();
  base = base.replace(/\s+\(\d+\)\s*[A-Za-z]?$/, "").trim();
  const parts = base.split(/\s+-\s+/);
  return (parts[0] || base).trim();
}

function detectShowName() {
  const manual = $("showNameInput").value.trim();
  if (manual) return manual;

  const csvName = $("allocationCsvFile").files[0]?.name || "";
  const pdfName = $("ticketsPdfFile").files[0]?.name || "";
  return inferShowNameFromFileName(csvName) || inferShowNameFromFileName(pdfName) || "Show";
}

function maybePrefillShowName() {
  const input = $("showNameInput");
  if (input.value.trim()) return;
  input.value = detectShowName();
}

function toAbsoluteUrl(pathOrUrl) {
  try {
    return new URL(pathOrUrl, window.location.origin).toString();
  } catch (_) {
    return pathOrUrl || "";
  }
}

function buildMailtoHref(row) {
  const email = (row.email || "").trim();
  if (!email) return "";
  const showName = state.showName || "Show";
  const subject = `Your ${showName} tickets are here!`;
  const body = [
    `Hi - here are your tickets for ${showName}. Do let me know if you have any questions, but otherwise please check all the information including the date to make sure everything is correct and please keep them somewhere safe on your phone so that the bar code can be scanned on arrival.`,
    "",
    "Please do shout if you have any questions, but otherwise, have a brilliant time!",
    "",
    "Thanks",
    "",
    "Annabelle",
  ].join("\n");
  return `mailto:${encodeURIComponent(email)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
}

function getPdfOpenHref(row) {
  const dragAssets = buildDragAssets(row);
  if (dragAssets?.localUrl) return dragAssets.localUrl;
  if (row.pdf_data_url) return row.pdf_data_url;
  return "";
}

function getPdfDragHref(row) {
  const dragAssets = buildDragAssets(row);
  if (dragAssets?.localUrl) return dragAssets.localUrl;
  if (row.pdf_data_url) return row.pdf_data_url;
  return "";
}

function getPdfDownloadHref(row) {
  if (row.pdf_download_url) return toAbsoluteUrl(row.pdf_download_url);
  if (row.pdf_url) return toAbsoluteUrl(row.pdf_url);
  return "";
}

const dragAssetCache = new Map();

function isSafariBrowser() {
  const ua = navigator.userAgent || "";
  return /Safari/i.test(ua) && !/Chrome|CriOS|Chromium|Edg|OPR|Android/i.test(ua);
}

function setDragData(dt, type, value) {
  if (!value) return;
  try {
    dt.setData(type, value);
  } catch (_) {}
}

function buildDragAssets(row) {
  const dataUrl = row.pdf_data_url || "";
  const fileName = row.pdf_file || "ticket.pdf";
  if (!dataUrl.startsWith("data:application/pdf;base64,")) return null;
  const cacheKey = `${fileName}|${dataUrl.length}`;
  if (dragAssetCache.has(cacheKey)) return dragAssetCache.get(cacheKey);

  try {
    const base64 = dataUrl.slice("data:application/pdf;base64,".length);
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) {
      bytes[i] = binary.charCodeAt(i);
    }
    const file = new File([bytes], fileName, { type: "application/pdf" });
    const localUrl = URL.createObjectURL(file);
    const assets = { file, localUrl };
    dragAssetCache.set(cacheKey, assets);
    return assets;
  } catch (_) {
    return null;
  }
}

function clearDragAssetCache() {
  for (const assets of dragAssetCache.values()) {
    if (assets?.localUrl) {
      URL.revokeObjectURL(assets.localUrl);
    }
  }
  dragAssetCache.clear();
}

function wireDragPdf(linkEl, row) {
  const safari = isSafariBrowser();
  const href = safari ? getPdfDownloadHref(row) : getPdfDragHref(row);
  const fileName = row.pdf_file || "ticket.pdf";
  const dragAssets = buildDragAssets(row);
  const dragFile = dragAssets?.file || null;
  if (!href && !dragFile) return;

  linkEl.draggable = true;
  linkEl.setAttribute("draggable", "true");
  linkEl.style.webkitUserDrag = "element";
  if (safari) {
    // On Safari/macOS, preserve native link drag behavior for Mail.
    return;
  }
  linkEl.addEventListener("dragstart", (event) => {
    const dt = event.dataTransfer;
    if (!dt) return;
    dt.effectAllowed = "copy";

    try {
      dt.clearData();
    } catch (_) {}

    let hasNativeFile = false;
    if (dragFile && dt.items && typeof dt.items.add === "function") {
      try {
        const added = dt.items.add(dragFile);
        hasNativeFile = Boolean(added && added.kind === "file");
      } catch (_) {}
    }

    if (href) {
      setDragData(dt, "DownloadURL", `application/pdf:${fileName}:${href}`);
      if (!hasNativeFile) {
        setDragData(dt, "text/uri-list", href);
      }
    }

    if (!hasNativeFile && !href) {
      setDragData(dt, "text/plain", fileName);
    }
  });
}

async function uploadAndFetch(path) {
  const { csvFile, pdfFile } = getFiles();
  const form = new FormData();
  form.append("allocation_csv", csvFile);
  form.append("tickets_pdf", pdfFile);

  const res = await fetch(path, {
    method: "POST",
    body: form,
  });

  if (!res.ok) {
    let detail = `HTTP ${res.status}`;
    try {
      const json = await res.json();
      if (json.detail) detail = json.detail;
    } catch (_) {}
    throw new Error(detail);
  }

  return res;
}

function renderPreview() {
  const tbody = $("previewTable").querySelector("tbody");
  tbody.innerHTML = "";

  for (const row of state.preview) {
    const tr = document.createElement("tr");

    const emailTd = document.createElement("td");
    if (row.email) {
      const mailLink = document.createElement("a");
      mailLink.href = buildMailtoHref(row);
      mailLink.textContent = row.email;
      emailTd.appendChild(mailLink);
    }

    const pdfTd = document.createElement("td");
    const openHref = getPdfOpenHref(row);
    const dragHref = getPdfDragHref(row);
    const downloadHref = getPdfDownloadHref(row);
    const safari = isSafariBrowser();
    if (openHref && row.pdf_file) {
      const fileLink = document.createElement("a");
      fileLink.href = openHref;
      fileLink.textContent = row.pdf_file;
      fileLink.target = "_blank";
      fileLink.rel = "noopener noreferrer";
      fileLink.className = "pdf-open-link";

      const dragLink = document.createElement("a");
      dragLink.href = safari ? (downloadHref || dragHref || openHref) : (dragHref || openHref);
      dragLink.download = row.pdf_file || "ticket.pdf";
      dragLink.textContent = "Drag PDF";
      dragLink.className = "pdf-drag-link";
      dragLink.setAttribute("role", "button");
      dragLink.title = "Drag into Mail to attach the PDF";
      if (dragHref || openHref) {
        wireDragPdf(dragLink, row);
      } else {
        dragLink.classList.add("is-disabled");
      }

      pdfTd.appendChild(fileLink);
      pdfTd.appendChild(document.createTextNode(" "));
      pdfTd.appendChild(dragLink);
    } else {
      pdfTd.textContent = row.pdf_file || "";
    }

    tr.appendChild(emailTd);
    tr.appendChild(pdfTd);
    tbody.appendChild(tr);
  }
}

function renderStats(stats) {
  const el = $("previewStats");
  if (!stats) {
    el.textContent = "";
    return;
  }
  const requested = Number(stats.requested_seat_count || 0);
  const matched = Number(stats.matched_seat_count || 0);
  const missing = Number(stats.missing_seat_count || 0);
  const outputs = Number(stats.output_pdf_count || 0);
  const status = missing === 0 && requested === matched ? "All matched" : "Check missing seats";
  el.textContent = `Seats requested: ${requested} | Tickets found in PDF: ${matched} | Missing: ${missing} | Output PDFs: ${outputs} (${status})`;
}

function renderFailures(failures) {
  const summary = $("previewFailures");
  const wrap = $("previewFailureDetails");
  const tbody = $("previewFailureTable").querySelector("tbody");

  summary.textContent = failures.length
    ? `${failures.length} booking(s) have missing seats in the PDF and may produce incomplete bundles.`
    : "";

  tbody.innerHTML = "";
  if (!failures.length) {
    wrap.hidden = true;
    return;
  }

  for (const failure of failures) {
    const tr = document.createElement("tr");
    const emailTd = document.createElement("td");
    emailTd.textContent = failure.email || "";
    const bookingTd = document.createElement("td");
    bookingTd.textContent = failure.booking_reference || "";
    const seatsTd = document.createElement("td");
    seatsTd.textContent = (failure.missing_seats || []).join(", ");
    tr.appendChild(emailTd);
    tr.appendChild(bookingTd);
    tr.appendChild(seatsTd);
    tbody.appendChild(tr);
  }
  wrap.hidden = false;
}

function setBuildLoading(isLoading) {
  const btn = $("buildBtn");
  const downloadBtn = $("downloadAllBtn");
  const chooseFolderBtn = $("chooseFolderBtn");
  const progress = $("buildProgress");
  btn.disabled = isLoading;
  if (downloadBtn && !downloadBtn.hidden) {
    downloadBtn.disabled = isLoading;
  }
  if (chooseFolderBtn) {
    chooseFolderBtn.disabled = isLoading;
  }
  btn.textContent = isLoading ? "Building..." : "Build Email PDF List";
  progress.hidden = !isLoading;
}

function setDownloadLoading(isLoading) {
  const btn = $("downloadAllBtn");
  const buildBtn = $("buildBtn");
  const chooseFolderBtn = $("chooseFolderBtn");
  if (!btn) return;
  btn.disabled = isLoading;
  if (buildBtn) {
    buildBtn.disabled = isLoading;
  }
  if (chooseFolderBtn) {
    chooseFolderBtn.disabled = isLoading;
  }
  btn.textContent = isLoading ? "Preparing ZIP..." : "Download All PDFs (.zip)";
}

function triggerBlobDownload(blob, fileName) {
  const nav = window.navigator;
  if (nav && typeof nav.msSaveOrOpenBlob === "function") {
    nav.msSaveOrOpenBlob(blob, fileName);
    return;
  }

  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = fileName;
  document.body.appendChild(link);
  link.click();
  link.remove();
  setTimeout(() => URL.revokeObjectURL(url), 1000);
}

function supportsDirectoryPicker() {
  return typeof window.showDirectoryPicker === "function";
}

async function ensureDirectoryWritePermission(dirHandle) {
  if (!dirHandle || typeof dirHandle.queryPermission !== "function") return true;
  const opts = { mode: "readwrite" };
  let permission = await dirHandle.queryPermission(opts);
  if (permission === "granted") return true;
  if (typeof dirHandle.requestPermission !== "function") return false;
  permission = await dirHandle.requestPermission(opts);
  return permission === "granted";
}

function clearSaveDirectorySelection() {
  state.saveDirHandle = null;
  state.saveDirName = "Not selected";
  updateSaveLocationUi();
}

async function writeBlobToSelectedDirectory(blob, fileName) {
  if (!state.saveDirHandle) return false;
  const hasPermission = await ensureDirectoryWritePermission(state.saveDirHandle);
  if (!hasPermission) {
    throw new Error("Folder permission was not granted. Please choose the folder again.");
  }

  const fileHandle = await state.saveDirHandle.getFileHandle(fileName, { create: true });
  const writable = await fileHandle.createWritable();
  try {
    await writable.write(blob);
  } finally {
    await writable.close();
  }
  return true;
}

$("buildBtn").addEventListener("click", async () => {
  setBuildLoading(true);
  try {
    maybePrefillShowName();
    state.showName = detectShowName();
    const res = await uploadAndFetch("/ticket-bundles/preview");
    const data = await res.json();
    clearDragAssetCache();
    state.preview = data.rows || [];
    renderPreview();
    renderStats(data.stats || null);
    renderFailures(data.failures || []);
    const hasRows = state.preview.length > 0;
    setDownloadButtonVisibility(hasRows);
    setStatus(
      hasRows
        ? `Built list: ${state.preview.length} email/PDF rows.`
        : "Build complete, but no rows were produced.",
      !hasRows
    );
  } catch (err) {
    setDownloadButtonVisibility(false);
    renderStats(null);
    renderFailures([]);
    setStatus(`Build failed: ${err.message}`, true);
  } finally {
    setBuildLoading(false);
  }
});

$("chooseFolderBtn").addEventListener("click", async () => {
  if (!supportsDirectoryPicker()) {
    setStatus("Folder picker is not available in this browser. Downloads will use the normal save dialog.");
    return;
  }

  try {
    const dirHandle = await window.showDirectoryPicker({ mode: "readwrite" });
    state.saveDirHandle = dirHandle;
    state.saveDirName = dirHandle?.name || "Selected folder";
    updateSaveLocationUi();
    setStatus(`Save location selected: ${state.saveDirName}`);
  } catch (err) {
    if (err && err.name === "AbortError") return;
    setStatus(`Could not select folder: ${err.message || "Unknown error"}`, true);
  }
});

$("downloadAllBtn").addEventListener("click", async () => {
  setDownloadLoading(true);
  try {
    maybePrefillShowName();
    const res = await uploadAndFetch("/ticket-bundles/generate");
    const blob = await res.blob();
    const showName = detectShowName().replace(/[^A-Za-z0-9_-]+/g, "_").replace(/^_+|_+$/g, "") || "Show";
    const fileName = `${showName}_ticket_bundle.zip`;
    if (state.saveDirHandle) {
      await writeBlobToSelectedDirectory(blob, fileName);
      setStatus(`ZIP saved to ${state.saveDirName}.`);
    } else {
      triggerBlobDownload(blob, fileName);
      setStatus("ZIP downloaded. Choose your save location in the browser download dialog.");
    }
  } catch (err) {
    if (state.saveDirHandle) {
      clearSaveDirectorySelection();
      setStatus(
        `Could not save to selected folder (${err.message}). Folder selection was cleared; please choose it again.`,
        true
      );
    } else {
      setStatus(`Download failed: ${err.message}`, true);
    }
  } finally {
    setDownloadLoading(false);
  }
});

$("allocationCsvFile").addEventListener("change", () => {
  maybePrefillShowName();
  resetDownloadAvailability();
});
$("ticketsPdfFile").addEventListener("change", () => {
  maybePrefillShowName();
  resetDownloadAvailability();
});
window.addEventListener("beforeunload", clearDragAssetCache);
updateSaveLocationUi();
resetDownloadAvailability();
