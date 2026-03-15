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
  el.style.color = isError ? "#b91c1c" : "#0f766e";
}

function updatePerformanceDetailsUi() {
  const wrap = $("performanceDetailsFields");
  const help = $("performanceDetailsHelp");
  const dateInput = $("performanceDateInput");
  const timeInput = $("performanceTimeInput");
  if (!wrap || !help || !dateInput || !timeInput) return;

  const shouldShow = state.hasBuiltPreview && state.needsPerformanceDetails;
  wrap.hidden = !shouldShow;
  help.hidden = !shouldShow;
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
    btn.textContent = "Download All PDFs (.zip)";
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
  return `${hour}.${String(minute).padStart(2, "0")}${meridiem}m`;
}

function format24hTime(hourValue, minuteValue) {
  const hour = Number(hourValue);
  const minute = Number(minuteValue);
  if (!Number.isInteger(hour) || hour < 0 || hour > 23) return "";
  if (!Number.isInteger(minute) || minute < 0 || minute > 59) return "";
  const meridiem = hour < 12 ? "am" : "pm";
  const displayHour = hour % 12 || 12;
  return `${displayHour}.${String(minute).padStart(2, "0")}${meridiem}`;
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

function getPerformanceContextText() {
  return [
    $("allocationCsvFile").files[0]?.name || "",
    $("ticketsPdfFile").files[0]?.name || "",
    state.showName || "",
  ].join(" \n ");
}

function applyPerformanceContextHints(timeValue, contextText) {
  const cleanTime = String(timeValue || "").trim();
  const context = String(contextText || "");
  if (!cleanTime) return "";

  const matineeHint = /\bmat(?:inee)?\b/i.test(context);
  if (!matineeHint) return cleanTime;

  const match = cleanTime.match(/^(\d{1,2})\.(\d{2})(am|pm)$/i);
  if (!match) return cleanTime;

  const hour = Number(match[1]);
  const meridiem = match[3].toLowerCase();
  if (meridiem === "am" && hour >= 1 && hour <= 6) {
    return `${hour}.${match[2]}pm`;
  }
  return cleanTime;
}

function applyDetectedPerformanceDetails(fileNameMetadata, previewMetadata) {
  const pdfDate = previewMetadata?.performance_date || "";
  const pdfTime = previewMetadata?.performance_time || "";
  const contextText = getPerformanceContextText();
  state.detectedPerformanceDate = fileNameMetadata.performanceDate || pdfDate;
  state.detectedPerformanceTime = applyPerformanceContextHints(
    fileNameMetadata.performanceTime || pdfTime,
    contextText
  );
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
  if (!details.isComplete) {
    throw new Error("Please add the performance date and time before saving the folder.");
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
  const fileName = row.pdf_file || "ticket.pdf";
  const dragAssets = buildDragAssets(row);
  const dragFile = dragAssets?.file || null;
  const downloadHref = getPdfDownloadHref(row);

  if (!dragFile && !downloadHref) return;

  linkEl.draggable = true;
  linkEl.setAttribute("draggable", "true");
  linkEl.style.webkitUserDrag = "element";

  if (safari) {
    // Safari: use download link so native drag attaches via the link href.
    if (downloadHref) {
      linkEl.href = downloadHref;
      linkEl.download = fileName;
    }
    return;
  }

  linkEl.addEventListener("dragstart", (event) => {
    const dt = event.dataTransfer;
    if (!dt) return;
    dt.effectAllowed = "copyMove";

    try {
      dt.clearData();
    } catch (_) {}

    // Strategy 1: Add as a real File object via DataTransferItemList.
    // This is the most reliable way to get an actual PDF attachment
    // into email clients and file managers.
    let hasNativeFile = false;
    if (dragFile && dt.items && typeof dt.items.add === "function") {
      try {
        const added = dt.items.add(dragFile);
        hasNativeFile = Boolean(added && added.kind === "file");
      } catch (_) {}
    }

    // Strategy 2: DownloadURL triggers a Chrome download-on-drop.
    // Must use an absolute https:// URL (blob: won't work cross-origin).
    if (downloadHref && downloadHref.startsWith("http")) {
      setDragData(dt, "DownloadURL", `application/pdf:${fileName}:${downloadHref}`);
    }

    // Strategy 3: Fallback text/uri-list for apps that accept links.
    if (!hasNativeFile && downloadHref) {
      setDragData(dt, "text/uri-list", downloadHref);
      setDragData(dt, "text/plain", downloadHref);
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
      dragLink.href = downloadHref || dragHref || openHref;
      dragLink.download = row.pdf_file || "ticket.pdf";
      dragLink.textContent = "Drag PDF";
      dragLink.className = "pdf-drag-link";
      dragLink.setAttribute("role", "button");
      dragLink.title = "Drag into Mail or a folder to attach the PDF";
      if (downloadHref || dragHref || openHref) {
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
        "Saved Dropbox folder needs permission again. Please choose *TICKETS WAITING TO BE SENT OUT.",
        true
      );
      return;
    }

    state.saveDirHandle = dirHandle;
    state.saveDirName = dirHandle?.name || "*TICKETS WAITING TO BE SENT OUT";
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

    const fileHandle = await outputDirHandle.getFileHandle(fileName, { create: true });
    const writable = await fileHandle.createWritable();
    try {
      await writable.write(blob);
      writtenCount += 1;
    } finally {
      await writable.close();
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
        ? `Built list: ${state.preview.length} email/PDF rows.`
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
    setStatus("Folder picker is not available in this browser. Downloads will use the normal save dialog.");
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
    const folderName = buildBundleFolderName();
    if (state.saveDirHandle) {
      if (!state.preview.length) {
        throw new Error("Build the email list first.");
      }
      const writtenCount = await writePreviewPdfsToSelectedDirectory(state.preview, folderName);
      if (writtenCount === 0) {
        throw new Error("No PDF files were available to save.");
      }
      setStatus(`${writtenCount} PDF(s) saved to ${state.saveDirName}/${folderName}.`);
    } else {
      const res = await uploadAndFetch("/ticket-bundles/generate");
      const blob = await res.blob();
      const fileName = `${folderName}.zip`;
      triggerBlobDownload(blob, fileName);
      setStatus("ZIP downloaded. Choose your save location in the browser download dialog.");
    }
  } catch (err) {
    if (state.saveDirHandle) {
      await clearSaveDirectorySelectionAndPersistedHandle();
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
  state.manualPerformanceTime = event.target.value;
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
