const state = {
  preview: [],
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
    emailTd.textContent = row.email || "";

    const pdfTd = document.createElement("td");
    if (row.pdf_url && row.pdf_file) {
      const link = document.createElement("a");
      link.href = row.pdf_url;
      link.textContent = row.pdf_file;
      link.target = "_blank";
      link.rel = "noopener noreferrer";
      pdfTd.appendChild(link);
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
    const res = await uploadAndFetch("/ticket-bundles/preview");
    const data = await res.json();
    state.preview = data.rows || [];
    renderPreview();
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
