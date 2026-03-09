const state = {
  preview: [],
  showName: "",
};

const $ = (id) => document.getElementById(id);

function setStatus(message, isError = false) {
  const el = $("opsStatus");
  el.textContent = message;
  el.style.color = isError ? "#b91c1c" : "#0f766e";
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
  if (row.pdf_url) return toAbsoluteUrl(row.pdf_url);
  if (row.pdf_data_url) return row.pdf_data_url;
  return "";
}

function getPdfDragHref(row) {
  if (row.pdf_download_url) return toAbsoluteUrl(row.pdf_download_url);
  if (row.pdf_data_url) return row.pdf_data_url;
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
    const assets = { file };
    dragAssetCache.set(cacheKey, assets);
    return assets;
  } catch (_) {
    return null;
  }
}

function wireDragPdf(linkEl, row) {
  const safari = isSafariBrowser();
  const href = safari
    ? (row.pdf_download_url ? toAbsoluteUrl(row.pdf_download_url) : getPdfDragHref(row))
    : getPdfDragHref(row);
  const fileName = row.pdf_file || "ticket.pdf";
  const dragAssets = buildDragAssets(row);
  const dragFile = dragAssets?.file || null;
  if (!href && !dragFile) return;

  linkEl.draggable = true;
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
      if (safari && !hasNativeFile) {
        // Safari Mail fallback: only URL file hints, no text fields (prevents body link insertion).
        setDragData(dt, "public.url", href);
        setDragData(dt, "public.url-name", fileName);
        setDragData(dt, "application/x-ticketallocator-pdf", fileName);
        return;
      }
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
    if (openHref && row.pdf_file) {
      const fileLink = document.createElement("a");
      fileLink.href = openHref;
      fileLink.textContent = row.pdf_file;
      fileLink.target = "_blank";
      fileLink.rel = "noopener noreferrer";
      fileLink.className = "pdf-open-link";

      const dragLink = document.createElement("button");
      dragLink.type = "button";
      dragLink.textContent = "Drag PDF";
      dragLink.className = "pdf-drag-link";
      dragLink.addEventListener("click", (event) => event.preventDefault());
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
  const progress = $("buildProgress");
  btn.disabled = isLoading;
  btn.textContent = isLoading ? "Building..." : "Build Email PDF List";
  progress.hidden = !isLoading;
}

$("buildBtn").addEventListener("click", async () => {
  setBuildLoading(true);
  try {
    maybePrefillShowName();
    state.showName = detectShowName();
    const res = await uploadAndFetch("/ticket-bundles/preview");
    const data = await res.json();
    state.preview = data.rows || [];
    renderPreview();
    renderStats(data.stats || null);
    renderFailures(data.failures || []);
    setStatus(`Built list: ${state.preview.length} email/PDF rows.`);
  } catch (err) {
    renderStats(null);
    renderFailures([]);
    setStatus(`Build failed: ${err.message}`, true);
  } finally {
    setBuildLoading(false);
  }
});

$("allocationCsvFile").addEventListener("change", maybePrefillShowName);
$("ticketsPdfFile").addEventListener("change", maybePrefillShowName);
