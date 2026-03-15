const DEFAULT_SAVE_LOCATION_HINT =
  "Annabelle's New Ticket Folder / ***SEATING & E-TICKETS / *TICKETS WAITING TO BE SENT OUT";
const SAVE_LOCATION_DB_NAME = "ticket-allocator-save-location";
const SAVE_LOCATION_STORE_NAME = "handles";
const SAVE_LOCATION_HANDLE_KEY = "default-save-dir";
const MONTH_ALIASES = {
  jan: "Jan",
  january: "Jan",
  feb: "Feb",
  february: "Feb",
  mar: "Mar",
  march: "Mar",
  apr: "Apr",
  april: "Apr",
  may: "May",
  jun: "Jun",
  june: "Jun",
  jul: "Jul",
  july: "Jul",
  aug: "Aug",
  august: "Aug",
  sep: "Sep",
  sept: "Sep",
  september: "Sep",
  oct: "Oct",
  october: "Oct",
  nov: "Nov",
  november: "Nov",
  dec: "Dec",
  december: "Dec",
};
const MONTH_PATTERN =
  "jan(?:uary)?|feb(?:ruary)?|mar(?:ch)?|apr(?:il)?|may|jun(?:e)?|jul(?:y)?|aug(?:ust)?|sep(?:t(?:ember)?)?|oct(?:ober)?|nov(?:ember)?|dec(?:ember)?";

const state = {
  preview: [],
  showName: "",
  saveDirHandle: null,
  saveDirName: DEFAULT_SAVE_LOCATION_HINT,
  detectedPerformanceDate: "",
  detectedPerformanceTime: "",
  manualPerformanceDate: "",
  manualPerformanceTime: "",
  needsPerformanceDetails: false,
  hasBuiltPreview: false,
};

const $ = (id) => document.getElementById(id);

function setStatus(message, isError = false) {
  const el = $("opsStatus");
  el.textContent = message;
  el.style.color = isError ? "#dc2626" : "#0f766e";
}

function updatePerformanceDetailsUi() {
  const wrap = $("performanceDetailsFields");
  const help = $("performanceDetailsHelp");
  const dateInput = $("performanceDateInput");
  const timeInput = $("performanceTimeInput");
  if (!wrap || !help || !dateInput || !timeInput) return;

  const shouldShow = state.hasBuiltPreview;
  if (shouldShow) {
    wrap.removeAttribute("hidden");
    help.removeAttribute("hidden");
  } else {
    wrap.setAttribute("hidden", "");
    help.setAttribute("hidden", "");
  }
  dateInput.value = state.manualPerformanceDate;
  timeInput.value = state.manualPerformanceTime;
}

function updateSaveLocationUi() {
  const el = $("saveLocationValue");
  if (!el) return;
  el.value = state.saveDirName || DEFAULT_SAVE_LOCATION_HINT;
}

function setDownloadButtonVisibility(isVisible) {
  const btn = $("downloadAllBtn");
  if (!btn) return;
  btn.hidden = !isVisible;
  if (!isVisible) {
    btn.disabled = false;
    btn.textContent = "Download All PDFs";
  }
}

function resetDownloadAvailability() {
  setDownloadButtonVisibility(false);
}

function resetPerformanceDetails() {
  state.detectedPerformanceDate = "";
  state.detectedPerformanceTime = "";
  state.manualPerformanceDate = "";
  state.manualPerformanceTime = "";
  state.needsPerformanceDetails = false;
  state.hasBuiltPreview = false;
  updatePerformanceDetailsUi();
}

function getFiles() {
  const csvFile = $("allocationCsvFile").files[0];
  const pdfFile = $("ticketsPdfFile").files[0];
  if (!csvFile || !pdfFile) {
    throw new Error("Please select both an allocation file and a ticket PDF.");
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

function normalizeMonth(value) {
  const key = (value || "").toLowerCase().replace(/[^a-z]/g, "");
  return MONTH_ALIASES[key] || "";
}

function normalizeDay(value) {
  const digits = String(value || "").replace(/\D/g, "");
  if (!digits) return "";
  const day = Number(digits);
  if (!Number.isInteger(day) || day <= 0 || day > 31) return "";
  return String(day);
}

function format12hTime(hourValue, minuteValue, meridiemValue) {
  const hour = Number(hourValue);
  const minute = Number(minuteValue);
  const meridiem = String(meridiemValue || "").trim().toLowerCase();
  if (!Number.isInteger(hour) || hour <= 0 || hour > 12) return "";
  if (!Number.isInteger(minute) || minute < 0 || minute > 59) return "";
  if (!["a", "p"].includes(meridiem)) return "";
  return `${hour}.${String(minute).padStart(2, "0")}`;
}

function format24hTime(hourValue, minuteValue) {
  const hour = Number(hourValue);
  const minute = Number(minuteValue);
  if (!Number.isInteger(hour) || hour < 0 || hour > 23) return "";
  if (!Number.isInteger(minute) || minute < 0 || minute > 59) return "";
  const displayHour = hour % 12 || 12;
  return `${displayHour}.${String(minute).padStart(2, "0")}`;
}

function dedupeStrings(values) {
  const out = [];
  const seen = new Set();
  for (const value of values) {
    const clean = String(value || "").trim();
    const key = clean.toLowerCase();
    if (!clean || seen.has(key)) continue;
    seen.add(key);
    out.push(clean);
  }
  return out;
}

function normalizePerformanceTimeValue(value) {
  const clean = String(value || "").trim();
  if (!clean) return "";

  const twelveHour = clean.match(/^(\d{1,2})(?::|\.)(\d{2})\s*([ap])\.?\s*m?\.?$/i);
  if (twelveHour) {
    return format12hTime(twelveHour[1], twelveHour[2], twelveHour[3]);
  }

  const twelveHourCompact = clean.match(/^(\d{1,2})\s*([ap])\.?\s*m?\.?$/i);
  if (twelveHourCompact) {
    return format12hTime(twelveHourCompact[1], "00", twelveHourCompact[2]);
  }

  const twentyFourHour = clean.match(/^([01]?\d|2[0-3])[:.](\d{2})$/);
  if (twentyFourHour) {
    return format24hTime(twentyFourHour[1], twentyFourHour[2]);
  }

  const bare = clean.match(/^(\d{1,2})\.(\d{2})$/);
  if (bare) {
    return `${Number(bare[1])}.${bare[2]}`;
  }

  return clean;
}

function extractDateCandidates(text) {
  if (!text) return [];
  const values = [];
  const monthDayRe = new RegExp(`\\b(${MONTH_PATTERN})[.\\s,_-]+(\\d{1,2})(?:st|nd|rd|th)?(?:[.,\\s_-]+\\d{2,4})?\\b`, "ig");
  const dayMonthRe = new RegExp(`\\b(\\d{1,2})(?:st|nd|rd|th)?[.\\s,_-]+(${MONTH_PATTERN})(?:[.,\\s_-]+\\d{2,4})?\\b`, "ig");

  for (const match of text.matchAll(monthDayRe)) {
    const month = normalizeMonth(match[1]);
    const day = normalizeDay(match[2]);
    if (month && day) values.push(`${month} ${day}`);
  }
  for (const match of text.matchAll(dayMonthRe)) {
    const month = normalizeMonth(match[2]);
    const day = normalizeDay(match[1]);
    if (month && day) values.push(`${month} ${day}`);
  }
  return dedupeStrings(values);
}

function extractTimeCandidates(text) {
  if (!text) return [];
  const values = [];
  const twelveHourRe = /\b(\d{1,2})(?::|\.)(\d{2})\s*([ap])\.?\s*m\.?\b/ig;
  const twelveHourCompactRe = /\b(\d{1,2})\s*([ap])\.?\s*m\.?\b/ig;
  const twentyFourHourRe = /\b([01]?\d|2[0-3])[:.](\d{2})\b/g;

  for (const match of text.matchAll(twelveHourRe)) {
    const formatted = format12hTime(match[1], match[2], match[3]);
    if (formatted) values.push(formatted);
  }
  for (const match of text.matchAll(twelveHourCompactRe)) {
    const formatted = format12hTime(match[1], "00", match[2]);
    if (formatted) values.push(formatted);
  }
  for (const match of text.matchAll(twentyFourHourRe)) {
    const formatted = format24hTime(match[1], match[2]);
    if (formatted) values.push(formatted);
  }
  return dedupeStrings(values);
}

function extractPerformanceMetadataFromFileNames() {
  const candidates = [
    $("allocationCsvFile").files[0]?.name || "",
    $("ticketsPdfFile").files[0]?.name || "",
  ].join(" \n ");
  const dates = extractDateCandidates(candidates);
  const times = extractTimeCandidates(candidates);
  return {
    performanceDate: dates.length === 1 ? dates[0] : "",
    performanceTime: times.length === 1 ? times[0] : "",
  };
}

function applyDetectedPerformanceDetails(fileNameMetadata, previewMetadata) {
  const pdfDate = previewMetadata?.performance_date || "";
  const pdfTime = previewMetadata?.performance_time || "";
  state.detectedPerformanceDate = fileNameMetadata.performanceDate || pdfDate;
  state.detectedPerformanceTime = normalizePerformanceTimeValue(fileNameMetadata.performanceTime || pdfTime);
  state.manualPerformanceDate = state.detectedPerformanceDate;
  state.manualPerformanceTime = state.detectedPerformanceTime;
  state.needsPerformanceDetails = !(state.detectedPerformanceDate && state.detectedPerformanceTime);
  updatePerformanceDetailsUi();
}

function getResolvedPerformanceDetails() {
  const dateValue = (state.manualPerformanceDate || state.detectedPerformanceDate || "").trim();
  const timeValue = (state.manualPerformanceTime || state.detectedPerformanceTime || "").trim();
  return {
    performanceDate: dateValue,
    performanceTime: timeValue,
    isComplete: Boolean(dateValue && timeValue),
  };
}

function sanitizeFolderName(value) {
  return String(value || "")
    .replace(/[\\/:*?"<>|]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function buildBundleFolderName() {
  const showName = detectShowName().trim() || "Show";
  const details = getResolvedPerformanceDetails();
  if (!details.performanceDate || !details.performanceTime) {
    throw new Error("Please fill in the performance date and time before saving.");
  }
  return sanitizeFolderName(`${showName} - ${details.performanceDate} - ${details.performanceTime}`);
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
  const fileName = row.pdf_file || "ticket.pdf";
  const dragAssets = buildDragAssets(row);
  const dragFile = dragAssets?.file || null;
  const blobUrl = dragAssets?.localUrl || "";

  if (!dragFile && !blobUrl) return;

  // Always set href to blob URL + download attribute.
  // This is local-only (never exposes the website) and gives the browser
  // a native file drag source for Apple Mail on macOS.
  if (blobUrl) {
    linkEl.href = blobUrl;
    linkEl.download = fileName;
  }

  linkEl.draggable = true;
  linkEl.setAttribute("draggable", "true");
  linkEl.style.webkitUserDrag = "element";

  // On Safari/macOS, the native <a href download> drag is the only thing
  // that works with Apple Mail. Do NOT add a custom dragstart.
  if (isSafariBrowser()) return;

  // On Chrome/other browsers, enhance with dt.items.add(file) for richer
  // drop targets. But keep the <a href> blob fallback intact by NOT calling
  // event.preventDefault() — if dt.items.add fails, the native drag still works.
  linkEl.addEventListener("dragstart", (event) => {
    const dt = event.dataTransfer;
    if (!dt) return;
    dt.effectAllowed = "copy";

    // Try to add the real File object. This gives the best result when
    // dropping into apps that support DataTransfer files.
    if (dragFile && dt.items && typeof dt.items.add === "function") {
      try {
        dt.items.add(dragFile);
      } catch (_) {}
    }

    // NO server URLs, NO DownloadURL, NO text/uri-list.
    // The <a href=blob: download=name> handles everything else natively.
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

  const emptyHint = $("emptyStateHint");
  if (emptyHint) emptyHint.hidden = state.preview.length > 0;

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
    // For the clickable filename, prefer the server URL (works in new tabs).
    // Blob URLs break when opened in a separate tab.
    const previewHref = (row.pdf_url ? toAbsoluteUrl(row.pdf_url) : "") || openHref;
    if (previewHref && row.pdf_file) {
      const fileLink = document.createElement("a");
      fileLink.href = previewHref;
      fileLink.textContent = row.pdf_file;
      fileLink.target = "_blank";
      fileLink.rel = "noopener noreferrer";
      fileLink.className = "pdf-open-link";

      const dragLink = document.createElement("a");
      dragLink.href = dragHref || openHref;
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
  el.textContent = `${requested} seats requested · ${matched} matched · ${missing} missing · ${outputs} PDFs (${status})`;
}

function renderFailures(failures) {
  const summary = $("previewFailures");
  const wrap = $("previewFailureDetails");
  const tbody = $("previewFailureTable").querySelector("tbody");

  summary.textContent = failures.length
    ? `${failures.length} booking(s) have missing seats and may produce incomplete bundles.`
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
  btn.textContent = isLoading ? "Preparing..." : "Download All PDFs";
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

function openSaveLocationDb() {
  return new Promise((resolve, reject) => {
    if (!window.indexedDB) {
      reject(new Error("IndexedDB is not available in this browser."));
      return;
    }

    const request = window.indexedDB.open(SAVE_LOCATION_DB_NAME, 1);
    request.onerror = () => reject(request.error || new Error("Could not open save-location storage."));
    request.onupgradeneeded = () => {
      const db = request.result;
      if (!db.objectStoreNames.contains(SAVE_LOCATION_STORE_NAME)) {
        db.createObjectStore(SAVE_LOCATION_STORE_NAME);
      }
    };
    request.onsuccess = () => resolve(request.result);
  });
}

async function getPersistedDirectoryHandle() {
  const db = await openSaveLocationDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(SAVE_LOCATION_STORE_NAME, "readonly");
    const store = tx.objectStore(SAVE_LOCATION_STORE_NAME);
    const request = store.get(SAVE_LOCATION_HANDLE_KEY);
    request.onerror = () => reject(request.error || new Error("Could not read saved folder."));
    request.onsuccess = () => resolve(request.result || null);
    tx.oncomplete = () => db.close();
    tx.onerror = () => reject(tx.error || new Error("Could not finish reading saved folder."));
  });
}

async function persistDirectoryHandle(dirHandle) {
  const db = await openSaveLocationDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(SAVE_LOCATION_STORE_NAME, "readwrite");
    const store = tx.objectStore(SAVE_LOCATION_STORE_NAME);
    const request = store.put(dirHandle, SAVE_LOCATION_HANDLE_KEY);
    request.onerror = () => reject(request.error || new Error("Could not save folder selection."));
    tx.oncomplete = () => {
      db.close();
      resolve();
    };
    tx.onerror = () => reject(tx.error || new Error("Could not finish saving folder selection."));
  });
}

async function clearPersistedDirectoryHandle() {
  const db = await openSaveLocationDb();
  return new Promise((resolve, reject) => {
    const tx = db.transaction(SAVE_LOCATION_STORE_NAME, "readwrite");
    const store = tx.objectStore(SAVE_LOCATION_STORE_NAME);
    const request = store.delete(SAVE_LOCATION_HANDLE_KEY);
    request.onerror = () => reject(request.error || new Error("Could not clear saved folder."));
    tx.oncomplete = () => {
      db.close();
      resolve();
    };
    tx.onerror = () => reject(tx.error || new Error("Could not finish clearing saved folder."));
  });
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
  state.saveDirName = DEFAULT_SAVE_LOCATION_HINT;
  updateSaveLocationUi();
}

async function clearSaveDirectorySelectionAndPersistedHandle() {
  clearSaveDirectorySelection();
  if (!window.indexedDB) return;
  try {
    await clearPersistedDirectoryHandle();
  } catch (_) {}
}

async function restoreSavedDirectorySelection() {
  if (!supportsDirectoryPicker() || !window.indexedDB) {
    clearSaveDirectorySelection();
    return;
  }

  try {
    const dirHandle = await getPersistedDirectoryHandle();
    if (!dirHandle) {
      clearSaveDirectorySelection();
      return;
    }

    const hasPermission = await ensureDirectoryWritePermission(dirHandle);
    if (!hasPermission) {
      await clearSaveDirectorySelectionAndPersistedHandle();
      setStatus(
        "Saved folder needs permission again. Please choose the save folder.",
        true
      );
      return;
    }

    state.saveDirHandle = dirHandle;
    state.saveDirName = dirHandle?.name || "Selected folder";
    updateSaveLocationUi();
  } catch (_) {
    await clearSaveDirectorySelectionAndPersistedHandle();
  }
}

function toDataUrlPdfBlob(dataUrl) {
  if (!dataUrl || !dataUrl.startsWith("data:application/pdf;base64,")) return null;
  const base64 = dataUrl.slice("data:application/pdf;base64,".length);
  try {
    const binary = atob(base64);
    const bytes = new Uint8Array(binary.length);
    for (let i = 0; i < binary.length; i += 1) {
      bytes[i] = binary.charCodeAt(i);
    }
    return new Blob([bytes], { type: "application/pdf" });
  } catch (_) {
    return null;
  }
}

function buildEmailFileContent(row) {
  const email = (row.email || "").trim();
  if (!email) return null;
  const showName = state.showName || "Show";
  const subject = `Your ${showName} tickets are here!`;
  const body = [
    `Hi - here are your tickets for ${showName}. Do let me know if you have any questions, but otherwise please check all the information including the date to make sure everything is correct and please keep them somewhere safe on your phone so that the bar code can be scanned on arrival.`,
    "",
    "Please do shout if you have any questions, but otherwise, have a brilliant time!",
  ].join("\n");

  const mailto = `mailto:${encodeURIComponent(email)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
  const mailtoXml = mailto.replace(/&/g, "&amp;");

  // .inetloc is the generic macOS internet-location format and handles mailto:
  // links more reliably than .webloc.
  const plist = [
    `<?xml version="1.0" encoding="UTF-8"?>`,
    `<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">`,
    `<plist version="1.0">`,
    `<dict>`,
    `\t<key>URL</key>`,
    `\t<string>${mailtoXml}</string>`,
    `</dict>`,
    `</plist>`,
  ].join("\n");

  return plist;
}

function emailFileName(pdfFileName) {
  if (!pdfFileName) return "email.inetloc";
  return pdfFileName.replace(/\.pdf$/i, "_email.inetloc");
}

async function writePreviewPdfsToSelectedDirectory(previewRows, folderName) {
  if (!state.saveDirHandle) return 0;
  const hasPermission = await ensureDirectoryWritePermission(state.saveDirHandle);
  if (!hasPermission) {
    throw new Error("Folder permission was not granted. Please choose the folder again.");
  }

  const outputDirHandle = await state.saveDirHandle.getDirectoryHandle(folderName, { create: true });
  let writtenCount = 0;

  for (const row of previewRows) {
    const fileName = row?.pdf_file || "";
    const blob = toDataUrlPdfBlob(row?.pdf_data_url || "");
    if (!fileName || !blob) continue;

    // Write PDF
    const fileHandle = await outputDirHandle.getFileHandle(fileName, { create: true });
    const writable = await fileHandle.createWritable();
    try {
      await writable.write(blob);
      writtenCount += 1;
    } finally {
      await writable.close();
    }

    // Write email compose shortcut
    const emailContent = buildEmailFileContent(row);
    if (emailContent) {
      const emailName = emailFileName(fileName);
      const emailHandle = await outputDirHandle.getFileHandle(emailName, { create: true });
      const emailWritable = await emailHandle.createWritable();
      try {
        await emailWritable.write(new Blob([emailContent], { type: "application/xml" }));
      } finally {
        await emailWritable.close();
      }
    }
  }

  return writtenCount;
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
    applyDetectedPerformanceDetails(extractPerformanceMetadataFromFileNames(), data.performance_metadata || null);
    renderPreview();
    renderStats(data.stats || null);
    renderFailures(data.failures || []);
    const hasRows = state.preview.length > 0;
    state.hasBuiltPreview = hasRows;
    updatePerformanceDetailsUi();
    setDownloadButtonVisibility(hasRows);
    setStatus(
      hasRows
        ? `Ready — ${state.preview.length} email/PDF rows built.`
        : "Build complete, but no rows were produced.",
      !hasRows
    );
  } catch (err) {
    resetPerformanceDetails();
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
    setStatus("Folder picker is not available in this browser.");
    return;
  }

  try {
    const dirHandle = await window.showDirectoryPicker({ mode: "readwrite" });
    state.saveDirHandle = dirHandle;
    state.saveDirName = dirHandle?.name || "Selected folder";
    try {
      await persistDirectoryHandle(dirHandle);
    } catch (_) {}
    updateSaveLocationUi();
    setStatus(`Save folder: ${state.saveDirName}`);
  } catch (err) {
    if (err && err.name === "AbortError") return;
    setStatus(`Could not select folder: ${err.message || "Unknown error"}`, true);
  }
});

// ---------------------------------------------------------------------------
// Minimal client-side ZIP builder (store method, no compression)
// ---------------------------------------------------------------------------

async function buildClientZip(previewRows, folderPrefix) {
  const entries = await _collectZipEntries(previewRows, folderPrefix);
  return _buildZipFromEntries(entries);
}

async function _collectZipEntries(previewRows, folderPrefix) {
  const entries = [];

  for (const row of previewRows) {
    const pdfName = row?.pdf_file || "";
    if (pdfName && row?.pdf_data_url) {
      const dataUrl = row.pdf_data_url;
      if (dataUrl.startsWith("data:application/pdf;base64,")) {
        const base64 = dataUrl.slice("data:application/pdf;base64,".length);
        const binary = atob(base64);
        const bytes = new Uint8Array(binary.length);
        for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
        entries.push({ name: `${folderPrefix}/${pdfName}`, data: bytes });
      }
    }

    const emailContent = buildEmailFileContent(row);
    if (emailContent && pdfName) {
      const emailName = emailFileName(pdfName);
      const emailBytes = new TextEncoder().encode(emailContent);
      entries.push({ name: `${folderPrefix}/${emailName}`, data: emailBytes });
    }
  }

  return entries;
}

function _buildZipFromEntries(entries) {
  const encoder = new TextEncoder();
  const centralRecords = [];
  const localParts = [];
  let offset = 0;

  for (const entry of entries) {
    const nameBytes = encoder.encode(entry.name);
    const data = entry.data;
    const crc = crc32(data);
    const size = data.length;

    // Local file header (30 bytes + name + data)
    const local = new ArrayBuffer(30 + nameBytes.length);
    const lv = new DataView(local);
    lv.setUint32(0, 0x04034b50, true);  // signature
    lv.setUint16(4, 20, true);           // version needed
    lv.setUint16(6, 0, true);            // flags
    lv.setUint16(8, 0, true);            // compression (store)
    lv.setUint16(10, 0, true);           // mod time
    lv.setUint16(12, 0, true);           // mod date
    lv.setUint32(14, crc, true);         // crc32
    lv.setUint32(18, size, true);        // compressed size
    lv.setUint32(22, size, true);        // uncompressed size
    lv.setUint16(26, nameBytes.length, true); // name length
    lv.setUint16(28, 0, true);           // extra length
    new Uint8Array(local).set(nameBytes, 30);

    localParts.push(new Uint8Array(local));
    localParts.push(data);

    // Central directory record
    const central = new ArrayBuffer(46 + nameBytes.length);
    const cv = new DataView(central);
    cv.setUint32(0, 0x02014b50, true);   // signature
    cv.setUint16(4, 20, true);            // version made by
    cv.setUint16(6, 20, true);            // version needed
    cv.setUint16(8, 0, true);             // flags
    cv.setUint16(10, 0, true);            // compression
    cv.setUint16(12, 0, true);            // mod time
    cv.setUint16(14, 0, true);            // mod date
    cv.setUint32(16, crc, true);          // crc32
    cv.setUint32(20, size, true);         // compressed size
    cv.setUint32(24, size, true);         // uncompressed size
    cv.setUint16(28, nameBytes.length, true); // name length
    cv.setUint16(30, 0, true);            // extra length
    cv.setUint16(32, 0, true);            // comment length
    cv.setUint16(34, 0, true);            // disk number
    cv.setUint16(36, 0, true);            // internal attrs
    cv.setUint32(38, 0, true);            // external attrs
    cv.setUint32(42, offset, true);       // local header offset
    new Uint8Array(central).set(nameBytes, 46);

    centralRecords.push(new Uint8Array(central));
    offset += 30 + nameBytes.length + size;
  }

  const centralStart = offset;
  let centralSize = 0;
  for (const rec of centralRecords) centralSize += rec.length;

  // End of central directory (22 bytes)
  const eocd = new ArrayBuffer(22);
  const ev = new DataView(eocd);
  ev.setUint32(0, 0x06054b50, true);     // signature
  ev.setUint16(4, 0, true);               // disk number
  ev.setUint16(6, 0, true);               // central dir disk
  ev.setUint16(8, entries.length, true);   // entries on disk
  ev.setUint16(10, entries.length, true);  // total entries
  ev.setUint32(12, centralSize, true);     // central dir size
  ev.setUint32(16, centralStart, true);    // central dir offset
  ev.setUint16(20, 0, true);              // comment length

  const allParts = [...localParts, ...centralRecords, new Uint8Array(eocd)];
  const totalSize = allParts.reduce((s, p) => s + p.length, 0);
  const result = new Uint8Array(totalSize);
  let pos = 0;
  for (const part of allParts) {
    result.set(part, pos);
    pos += part.length;
  }

  return new Blob([result], { type: "application/zip" });
}

function crc32(data) {
  let crc = 0xFFFFFFFF;
  for (let i = 0; i < data.length; i++) {
    crc ^= data[i];
    for (let j = 0; j < 8; j++) {
      crc = (crc >>> 1) ^ (crc & 1 ? 0xEDB88320 : 0);
    }
  }
  return (crc ^ 0xFFFFFFFF) >>> 0;
}

$("downloadAllBtn").addEventListener("click", async () => {
  setDownloadLoading(true);
  try {
    maybePrefillShowName();
    const folderName = buildBundleFolderName();
    if (state.saveDirHandle) {
      if (!state.preview.length) {
        throw new Error("Build the email list first.");
      }
      const writtenCount = await writePreviewPdfsToSelectedDirectory(state.preview, folderName);
      if (writtenCount === 0) {
        throw new Error("No PDF files were available to save.");
      }
      setStatus(`${writtenCount} PDF(s) and email drafts saved to ${state.saveDirName}/${folderName}.`);
    } else {
      if (!state.preview.length) {
        throw new Error("Build the email list first.");
      }
      const zipBlob = await buildClientZip(state.preview, folderName);
      const fileName = `${folderName}.zip`;
      triggerBlobDownload(zipBlob, fileName);
      setStatus("Download started.");
    }
  } catch (err) {
    if (state.saveDirHandle) {
      await clearSaveDirectorySelectionAndPersistedHandle();
      setStatus(
        `Could not save to folder (${err.message}). Please choose the folder again.`,
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
  resetPerformanceDetails();
  const file = $("allocationCsvFile").files[0];
  const nameEl = document.querySelector("#dropZoneAllocation .drop-file-name");
  if (nameEl) nameEl.textContent = file ? file.name : "";
});
$("ticketsPdfFile").addEventListener("change", () => {
  maybePrefillShowName();
  resetDownloadAvailability();
  resetPerformanceDetails();
  const file = $("ticketsPdfFile").files[0];
  const nameEl = document.querySelector("#dropZonePdf .drop-file-name");
  if (nameEl) nameEl.textContent = file ? file.name : "";
});
$("performanceDateInput").addEventListener("input", (event) => {
  state.manualPerformanceDate = event.target.value;
});
$("performanceTimeInput").addEventListener("input", (event) => {
  state.manualPerformanceTime = normalizePerformanceTimeValue(event.target.value);
});
window.addEventListener("beforeunload", clearDragAssetCache);
updateSaveLocationUi();
updatePerformanceDetailsUi();
resetDownloadAvailability();
restoreSavedDirectorySelection();

// ---------------------------------------------------------------------------
// Drag-and-drop file input zones
// ---------------------------------------------------------------------------

function setFileInputFiles(inputEl, file) {
  const dt = new DataTransfer();
  dt.items.add(file);
  inputEl.files = dt.files;
  inputEl.dispatchEvent(new Event("change", { bubbles: true }));
}

function setupDropZone(dropZoneId, inputId, acceptTest) {
  const zone = $(dropZoneId);
  const input = $(inputId);
  if (!zone || !input) return;

  let dragCounter = 0;

  zone.addEventListener("dragenter", (e) => {
    e.preventDefault();
    dragCounter++;
    zone.classList.add("drop-active");
  });

  zone.addEventListener("dragover", (e) => {
    e.preventDefault();
    e.dataTransfer.dropEffect = "copy";
  });

  zone.addEventListener("dragleave", (e) => {
    e.preventDefault();
    dragCounter--;
    if (dragCounter <= 0) {
      dragCounter = 0;
      zone.classList.remove("drop-active");
    }
  });

  zone.addEventListener("drop", (e) => {
    e.preventDefault();
    dragCounter = 0;
    zone.classList.remove("drop-active");

    const files = e.dataTransfer.files;
    if (!files || files.length === 0) return;

    const file = files[0];
    if (acceptTest && !acceptTest(file)) {
      setStatus(`"${file.name}" is not an accepted file type for this input.`, true);
      return;
    }

    setFileInputFiles(input, file);
    const nameEl = zone.querySelector(".drop-file-name");
    if (nameEl) nameEl.textContent = file.name;
  });
}

function isAllocationFile(file) {
  const name = (file.name || "").toLowerCase();
  return name.endsWith(".csv") || name.endsWith(".numbers");
}

function isPdfFile(file) {
  const name = (file.name || "").toLowerCase();
  return name.endsWith(".pdf") || file.type === "application/pdf";
}

setupDropZone("dropZoneAllocation", "allocationCsvFile", isAllocationFile);
setupDropZone("dropZonePdf", "ticketsPdfFile", isPdfFile);