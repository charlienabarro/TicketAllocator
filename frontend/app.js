const state = {
  preview: [],
  showName: "",
  pdfBlobCache: {},
  pdfBlobLoads: {},
  pdfObjectUrlCache: {},
  pdfDataUrlCache: {},
  pdfStateByUrl: {},
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
  const subject = `${showName} tickets`;
  const body = [
    "Hi,",
    "",
    "Please find your tickets attached.",
    "",
    "Best regards,",
  ]
    .filter(Boolean)
    .join("\n");
  return `mailto:${encodeURIComponent(email)}?subject=${encodeURIComponent(subject)}&body=${encodeURIComponent(body)}`;
}

async function primePdfBlob(row) {
  const key = row.pdf_url || "";
  if (!key) return null;
  if (state.pdfBlobCache[key]) return state.pdfBlobCache[key];
  if (state.pdfBlobLoads[key]) return state.pdfBlobLoads[key];
  state.pdfStateByUrl[key] = "loading";

  state.pdfBlobLoads[key] = (async () => {
    try {
      let blob = null;
      for (let attempt = 0; attempt < 3; attempt += 1) {
        try {
          const res = await fetch(toAbsoluteUrl(key), { cache: "no-store" });
          if (!res.ok) continue;
          blob = await res.blob();
          if (blob && blob.size > 0) break;
        } catch (_) {}
      }
      if (!blob || blob.size === 0) {
        state.pdfStateByUrl[key] = "failed";
        return null;
      }
      state.pdfBlobCache[key] = blob;
      if (!state.pdfObjectUrlCache[key]) {
        state.pdfObjectUrlCache[key] = URL.createObjectURL(blob);
      }
      if (!state.pdfDataUrlCache[key]) {
        state.pdfDataUrlCache[key] = await blobToDataUrl(blob);
      }
      state.pdfStateByUrl[key] = "ready";
      return blob;
    } catch (_) {
      state.pdfStateByUrl[key] = "failed";
      return null;
    } finally {
      delete state.pdfBlobLoads[key];
    }
  })();

  return state.pdfBlobLoads[key];
}

async function downloadPdfToDevice(row) {
  const key = row.pdf_url || "";
  if (!key) return;
  await primePdfBlob(row);
  const blob = state.pdfBlobCache[key];
  if (!blob) {
    setStatus("PDF is not ready yet. Rebuild preview and try again.", true);
    return false;
  }

  const objectUrl = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = objectUrl;
  a.download = row.pdf_file || "ticket.pdf";
  document.body.appendChild(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(objectUrl), 4000);
  return true;
}

function blobToDataUrl(blob) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onloadend = () => resolve(String(reader.result || ""));
    reader.onerror = reject;
    reader.readAsDataURL(blob);
  });
}

async function primeAllPreviewPdfs(rows, onProgress) {
  const list = rows.filter((row) => row && row.pdf_url);
  for (let idx = 0; idx < list.length; idx += 1) {
    await primePdfBlob(list[idx]);
    if (onProgress) onProgress(idx + 1, list.length);
  }
}

function clearPdfCaches() {
  for (const key of Object.keys(state.pdfObjectUrlCache)) {
    try {
      URL.revokeObjectURL(state.pdfObjectUrlCache[key]);
    } catch (_) {}
  }
  state.pdfBlobCache = {};
  state.pdfBlobLoads = {};
  state.pdfObjectUrlCache = {};
  state.pdfDataUrlCache = {};
  state.pdfStateByUrl = {};
}

async function downloadAllPreviewPdfs() {
  if (!state.preview.length) {
    setStatus("Build a preview first.", true);
    return;
  }
  let done = 0;
  for (const row of state.preview) {
    if (!row.pdf_url) continue;
    await primePdfBlob(row);
    const ok = await downloadPdfToDevice(row);
    if (ok) done += 1;
    await new Promise((resolve) => setTimeout(resolve, 140));
  }
  setStatus(`Downloaded ${done} PDF file(s).`);
}

function wirePdfDrag(linkEl, row) {
  const key = row.pdf_url || "";
  if (!key) return;
  const absoluteUrl = toAbsoluteUrl(key);
  if (!absoluteUrl) return;
  linkEl.draggable = true;
  linkEl.title = "Drag to attach PDF";

  linkEl.addEventListener("pointerenter", () => {
    primePdfBlob(row);
  });
  linkEl.addEventListener("mousedown", () => {
    primePdfBlob(row);
  });

  linkEl.addEventListener("dragstart", (event) => {
    const dt = event.dataTransfer;
    if (!dt) return;
    dt.effectAllowed = "copy";
    const cached = state.pdfBlobCache[key];
    const fileName = row.pdf_file || "ticket.pdf";
    if (!cached) {
      event.preventDefault();
      setStatus("PDF still loading. Try drag again in a moment.", true);
      primePdfBlob(row);
      return;
    }
    let fileAdded = false;
    if (cached && dt.items && dt.items.add) {
      try {
        const file = new File([cached], fileName, { type: "application/pdf" });
        dt.items.add(file);
        fileAdded = true;
      } catch (_) {}
    }
    if (!fileAdded) {
      dt.setData("DownloadURL", `application/pdf:${fileName}:${absoluteUrl}`);
    }
    dt.setData("text/plain", "");
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
    } else {
      emailTd.textContent = "";
    }

    const pdfTd = document.createElement("td");
    if (row.pdf_url && row.pdf_file) {
      const key = row.pdf_url || "";
      const loadState = state.pdfStateByUrl[key] || "loading";
      const ready = loadState === "ready";
      const openLink = document.createElement("a");
      openLink.href = toAbsoluteUrl(key);
      openLink.textContent = row.pdf_file;
      openLink.className = "pdf-open-link";
      openLink.target = "_blank";
      openLink.rel = "noopener noreferrer";

      const dragChip = document.createElement("span");
      dragChip.textContent = "Drag PDF";
      dragChip.className = "pdf-drag-link";
      if (ready) {
        wirePdfDrag(dragChip, row);
      } else {
        dragChip.classList.add("is-disabled");
        dragChip.title = loadState === "failed" ? "Failed to load PDF" : "PDF still loading";
      }

      const readyChip = document.createElement("span");
      if (loadState === "ready") {
        readyChip.className = "pdf-ready is-ready";
        readyChip.textContent = "Ready";
      } else if (loadState === "failed") {
        readyChip.className = "pdf-ready is-failed";
        readyChip.textContent = "Failed";
      } else {
        readyChip.className = "pdf-ready is-loading";
        readyChip.textContent = "Loading";
      }

      pdfTd.appendChild(openLink);
      pdfTd.appendChild(document.createTextNode(" "));
      pdfTd.appendChild(dragChip);
      pdfTd.appendChild(document.createTextNode(" "));
      pdfTd.appendChild(readyChip);
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
    clearPdfCaches();
    state.preview = data.rows || [];
    renderPreview();
    await primeAllPreviewPdfs(state.preview, (done, total) => {
      renderPreview();
      setStatus(`Preparing PDFs for drag and download... ${done}/${total}`);
    });
    renderStats(data.stats || null);

    const failures = data.failures || [];
    renderFailures(failures);

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
$("downloadAllBtn").addEventListener("click", async () => {
  await downloadAllPreviewPdfs();
});
