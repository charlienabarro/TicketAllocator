from __future__ import annotations

import base64
import csv
from collections import defaultdict
from datetime import datetime
from io import StringIO
from pathlib import Path
import re
import tempfile
from urllib.parse import quote
from uuid import uuid4

from fastapi import FastAPI, File, HTTPException, Query, Response, UploadFile
from fastapi.responses import FileResponse, StreamingResponse
from fastapi.staticfiles import StaticFiles

from allocator.allocator_engine import run_allocation
from allocator.booking_importer import parse_bookings_csv
from allocator.import_parsers import parse_seat_plan_csv, parse_ticket_stock_csv
from allocator.models import Allocation, AvailableSeat, Booking, MatchStatus, Performance, TheatreSeat
from allocator.seat_plan_ingestion import SeatPlanIngestor
from allocator.seat_range_expander import expand_ticket_stock_rows
from allocator.ticket_bundle import (
    ParsedTicketPageResult,
    TicketBundleError,
    build_pkpass_artifacts,
    build_output_filenames,
    build_booking_groups,
    build_group_pdf,
    build_bundle_zip,
    group_has_output_pdf,
    parse_ticket_pdf_page_results,
    parse_ticket_pdf_pages,
    parse_allocation_csv,
    parse_seat_list,
    split_groups_for_output,
)
from backend.schemas import (
    AllocationRowResponse,
    CreatePerformanceRequest,
    CreatePerformanceResponse,
    ImportBookingsRequest,
    ImportTicketStockRequest,
    LoadSeatPlanRequest,
    ManualAllocationRequest,
    RunAllocationRequest,
)
from backend.store import InMemoryStore

app = FastAPI(title="Seat Allocator API", version="0.1.0")
store = InMemoryStore()
ingestor = SeatPlanIngestor()
preview_download_cache: dict[str, dict[str, tuple[bytes, str]]] = {}
ROOT_DIR = Path(__file__).resolve().parent.parent
FRONTEND_DIR = ROOT_DIR / "frontend"
EMAIL_CELL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
app.mount("/static", StaticFiles(directory=str(FRONTEND_DIR)), name="static")


@app.get("/")
def frontend_home() -> FileResponse:
    return FileResponse(str(FRONTEND_DIR / "index.html"))


@app.post("/ticket-bundles/preview")
async def preview_ticket_bundle(
    allocation_csv: UploadFile = File(...),
    tickets_pdf: UploadFile = File(...),
) -> dict:
    allocation_rows = await _read_allocation_rows(allocation_csv)
    pdf_content = await tickets_pdf.read()
    expected_seats = _expected_seat_labels(allocation_rows)

    try:
        parsed_page_results = parse_ticket_pdf_page_results(pdf_content, expected_seats=expected_seats)
        seat_map = _seat_map_from_page_results(parsed_page_results)
        groups = build_booking_groups(allocation_rows, seat_map)
        parsed_pages = [result.to_parsed_ticket_page() for result in parsed_page_results if result.wallet_ready]
        pass_artifacts = build_pkpass_artifacts(parsed_pages)
    except TicketBundleError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    output_groups, excluded_groups = split_groups_for_output(groups)
    filenames = build_output_filenames(output_groups)
    preview_files = _build_preview_files(pdf_content, output_groups, filenames, pass_artifacts)
    preview_id = _store_preview_files(preview_files) if preview_files else None
    performance_metadata = _build_performance_metadata_from_page_results(parsed_page_results)
    missing_unique = {seat for group in groups for seat in group.missing_seats}
    matched_unique = {seat for group in groups for seat in group.seat_labels if seat in seat_map}
    wallet_failure_rows = _collect_wallet_failure_rows(output_groups, parsed_page_results)
    wallet_pass_count = len(pass_artifacts)

    return {
        "rows": [
            {
                "email": group.email,
                "pdf_file": filename,
                "pdf_url": f"/ticket-bundles/preview/{preview_id}/files/{quote(filename)}" if preview_id else "",
                "pdf_download_url": (
                    f"/ticket-bundles/preview/{preview_id}/files/{quote(filename)}?download=1" if preview_id else ""
                ),
                "pdf_data_url": (
                    f"data:application/pdf;base64,{base64.b64encode(preview_files[filename][0]).decode('ascii')}"
                ),
                "wallet_passes": _build_wallet_pass_rows(
                    group=group,
                    preview_id=preview_id,
                    pass_artifacts=pass_artifacts,
                    parsed_page_results=parsed_page_results,
                ),
                "wallet_failures": _build_wallet_failure_rows_for_group(group, parsed_page_results),
            }
            for group, filename in zip(output_groups, filenames)
        ],
        "failures": [
            {
                "booking_reference": group.booking_reference,
                "email": group.email,
                "missing_seats": group.missing_seats,
                "issue": _describe_group_output_issue(group),
            }
            for group in excluded_groups
        ],
        "wallet_failures": wallet_failure_rows,
        "stats": {
            "requested_seat_count": len(expected_seats),
            "matched_seat_count": len(matched_unique),
            "missing_seat_count": len(missing_unique),
            "output_pdf_count": len(output_groups),
            "wallet_pass_count": wallet_pass_count,
            "wallet_failure_count": len(wallet_failure_rows),
        },
        "performance_metadata": performance_metadata,
    }


@app.post("/ticket-bundles/generate")
async def generate_ticket_bundle(
    allocation_csv: UploadFile = File(...),
    tickets_pdf: UploadFile = File(...),
) -> StreamingResponse:
    allocation_rows = await _read_allocation_rows(allocation_csv)
    pdf_content = await tickets_pdf.read()
    expected_seats = _expected_seat_labels(allocation_rows)

    try:
        parsed_pages = parse_ticket_pdf_pages(pdf_content, expected_seats=expected_seats)
        seat_map = _seat_map_from_parsed_pages(parsed_pages)
        groups = build_booking_groups(allocation_rows, seat_map)
        zip_blob = build_bundle_zip(pdf_content, groups, parsed_pages=parsed_pages)
    except TicketBundleError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc

    headers = {"Content-Disposition": 'attachment; filename="ticket_bundle.zip"'}
    return StreamingResponse(iter([zip_blob]), media_type="application/zip", headers=headers)


@app.get("/ticket-bundles/preview/{preview_id}/files/{file_name}")
def download_preview_file(preview_id: str, file_name: str, download: bool = Query(default=False)) -> Response:
    preview = preview_download_cache.get(preview_id)
    if not preview:
        raise HTTPException(status_code=404, detail="Preview file set not found. Build preview again.")
    artifact = preview.get(file_name)
    if artifact is None:
        raise HTTPException(status_code=404, detail="Preview PDF not found.")
    blob, media_type = artifact
    disposition = "attachment" if download else "inline"
    headers = {"Content-Disposition": f'{disposition}; filename="{file_name}"'}
    return Response(content=blob, media_type=media_type, headers=headers)


@app.post("/performances", response_model=CreatePerformanceResponse)
def create_performance(payload: CreatePerformanceRequest) -> CreatePerformanceResponse:
    theatre_id = _find_or_create_theatre(payload.theatre_name, payload.theatre_city)
    performance_id = store.next_performance_id()

    store.performances[performance_id] = Performance(
        id=performance_id,
        theatre_id=theatre_id,
        show_name=payload.show_name,
        performance_date=payload.performance_date,
        performance_time=payload.performance_time,
        supplier_reference=payload.supplier_reference,
    )

    store.available_seats_by_performance.setdefault(performance_id, [])
    store.bookings_by_performance.setdefault(performance_id, [])
    store.allocations_by_performance.setdefault(performance_id, {})

    return CreatePerformanceResponse(performance_id=performance_id, theatre_id=theatre_id)


@app.post("/theatres/seat-plan")
def import_seat_plan(payload: LoadSeatPlanRequest) -> dict:
    theatre = store.theatres.get(payload.theatre_id)
    if not theatre:
        raise HTTPException(status_code=404, detail="Theatre not found")

    rows = parse_seat_plan_csv(payload.csv_content)
    result = ingestor.ingest_structured_rows(
        theatre_id=payload.theatre_id,
        theatre_name=theatre.name,
        city=theatre.city,
        rows=rows,
    )

    store.theatre_seats[payload.theatre_id] = result.theatre_seats
    return {
        "theatre_id": payload.theatre_id,
        "ingested_seat_count": len(result.theatre_seats),
        "requires_manual_review": result.requires_manual_review,
        "warnings": [w.__dict__ for w in result.warnings],
    }


@app.post("/imports/ticket-stock")
def import_ticket_stock(payload: ImportTicketStockRequest) -> dict:
    performance = _get_performance(payload.performance_id)

    parsed_rows = parse_ticket_stock_csv(payload.csv_content)
    expanded = expand_ticket_stock_rows(parsed_rows)

    by_key: dict[str, AvailableSeat] = {}
    for seat in store.available_seats_by_performance.setdefault(payload.performance_id, []):
        by_key[_available_key(seat)] = seat

    import_id = f"stock-{datetime.utcnow().isoformat()}"
    for seat in expanded:
        key = f"{seat.section}:{seat.row}:{seat.seat_number}"
        by_key[key] = AvailableSeat(
            performance_id=payload.performance_id,
            theatre_seat_id=key,
            section=seat.section,
            row=seat.row,
            seat_number=seat.seat_number,
            seat_label=seat.seat_label,
            source_import_id=import_id,
        )

    store.available_seats_by_performance[payload.performance_id] = sorted(
        by_key.values(), key=lambda s: (s.section, s.row, s.seat_number)
    )

    if performance.theatre_id not in store.theatre_seats:
        store.theatre_seats[performance.theatre_id] = _infer_theatre_seats_from_available(
            performance.theatre_id,
            store.available_seats_by_performance[payload.performance_id],
        )

    return {
        "performance_id": payload.performance_id,
        "imported_count": len(expanded),
        "total_available_seats": len(store.available_seats_by_performance[payload.performance_id]),
    }


@app.post("/imports/bookings")
def import_bookings(payload: ImportBookingsRequest) -> dict:
    _get_performance(payload.performance_id)

    known_sections = sorted(
        {seat.section for seat in store.available_seats_by_performance.get(payload.performance_id, [])}
    )

    preview_reader = csv.DictReader(StringIO(payload.csv_content))
    line_count = sum(1 for _ in preview_reader)
    start_id = store.next_booking_id_start(line_count)

    parsed = parse_bookings_csv(
        payload.csv_content,
        performance_id=payload.performance_id,
        starting_booking_id=start_id,
        known_sections=known_sections,
    )

    existing = {b.booking_reference: b for b in store.bookings_by_performance.setdefault(payload.performance_id, [])}
    for booking in parsed.bookings:
        existing[booking.booking_reference] = booking

    store.bookings_by_performance[payload.performance_id] = sorted(
        existing.values(), key=lambda b: b.booking_reference
    )

    for pref in parsed.preferences:
        store.preferences_by_booking_id[pref.booking_id] = pref

    return {
        "performance_id": payload.performance_id,
        "imported_bookings": len(parsed.bookings),
        "total_bookings": len(store.bookings_by_performance[payload.performance_id]),
    }


@app.post("/allocations/run")
def run_allocations(payload: RunAllocationRequest) -> dict:
    performance = _get_performance(payload.performance_id)
    bookings = store.bookings_by_performance.get(payload.performance_id, [])
    available = store.available_seats_by_performance.get(payload.performance_id, [])

    if not bookings:
        raise HTTPException(status_code=400, detail="No bookings loaded for this performance")
    if not available:
        raise HTTPException(status_code=400, detail="No available seats loaded for this performance")

    theatre_seats = store.theatre_seats.get(performance.theatre_id) or _infer_theatre_seats_from_available(
        performance.theatre_id, available
    )

    prefs = {
        booking.id: store.preferences_by_booking_id[booking.id]
        for booking in bookings
        if booking.id in store.preferences_by_booking_id
    }

    run = run_allocation(
        bookings=bookings,
        preferences=prefs,
        available_seats=available,
        theatre_seats=theatre_seats,
    )

    alloc_map: dict[int, Allocation] = {}
    for allocation in run.allocations:
        alloc_map[allocation.booking_id] = allocation
    store.allocations_by_performance[payload.performance_id] = alloc_map

    return {
        "performance_id": payload.performance_id,
        "allocated": len(run.allocations) - len(run.unallocated_booking_ids),
        "unallocated": len(run.unallocated_booking_ids),
    }


@app.get("/allocations/{performance_id}", response_model=list[AllocationRowResponse])
def get_allocations(performance_id: int) -> list[AllocationRowResponse]:
    _get_performance(performance_id)
    bookings = store.bookings_by_performance.get(performance_id, [])
    alloc_map = store.allocations_by_performance.get(performance_id, {})

    rows: list[AllocationRowResponse] = []
    for booking in bookings:
        allocation = alloc_map.get(
            booking.id,
            Allocation(
                booking_id=booking.id,
                assigned_seats=[],
                match_status=MatchStatus.UNALLOCATED,
                match_notes="Not allocated yet",
            ),
        )

        rows.append(
            AllocationRowResponse(
                booking_id=booking.id,
                booking_reference=booking.booking_reference,
                customer_name=booking.customer_name,
                quantity=booking.quantity,
                assigned_seats=allocation.assigned_seats,
                section=_extract_section_from_alloc(performance_id, allocation.assigned_seats),
                match_status=allocation.match_status.value,
                match_notes=allocation.match_notes,
                manually_edited=allocation.manually_edited,
            )
        )

    return rows


@app.post("/allocations/{booking_id}/manual")
def manual_allocation(booking_id: int, payload: ManualAllocationRequest) -> dict:
    performance_id, booking = _find_booking(booking_id)
    if booking is None:
        raise HTTPException(status_code=404, detail="Booking not found")

    available = {
        seat.seat_label: seat for seat in store.available_seats_by_performance.get(performance_id, [])
    }
    if any(label not in available for label in payload.assigned_seats):
        raise HTTPException(status_code=400, detail="One or more requested seats are not available")

    allocations = store.allocations_by_performance.setdefault(performance_id, {})
    for other_booking_id, alloc in allocations.items():
        if other_booking_id == booking_id:
            continue
        overlap = set(alloc.assigned_seats) & set(payload.assigned_seats)
        if overlap:
            raise HTTPException(status_code=400, detail=f"Seats already assigned: {sorted(overlap)}")

    if payload.assigned_seats and len(payload.assigned_seats) != booking.quantity:
        raise HTTPException(status_code=400, detail="Manual assignment must match booking quantity")

    status = MatchStatus.GOOD_MATCH if payload.assigned_seats else MatchStatus.UNALLOCATED
    allocations[booking_id] = Allocation(
        booking_id=booking_id,
        assigned_seats=payload.assigned_seats,
        match_status=status,
        match_notes="Manual override",
        manually_edited=True,
    )

    return {"performance_id": performance_id, "booking_id": booking_id, "assigned_seats": payload.assigned_seats}


@app.get("/exports/{performance_id}/csv")
def export_allocations_csv(performance_id: int) -> Response:
    _get_performance(performance_id)
    bookings = store.bookings_by_performance.get(performance_id, [])
    alloc_map = store.allocations_by_performance.get(performance_id, {})
    prefs = store.preferences_by_booking_id

    output = StringIO()
    writer = csv.writer(output)
    writer.writerow(
        [
            "Booking Ref",
            "Customer Name",
            "Quantity",
            "Preferences",
            "Notes",
            "Assigned Seats",
            "Section",
            "Match Status",
            "Match Notes",
        ]
    )

    for booking in bookings:
        allocation = alloc_map.get(
            booking.id,
            Allocation(booking_id=booking.id, match_status=MatchStatus.UNALLOCATED),
        )
        preference_text = prefs.get(booking.id).raw_request_text if booking.id in prefs else ""
        assigned = " ".join(allocation.assigned_seats)
        writer.writerow(
            [
                booking.booking_reference,
                booking.customer_name,
                booking.quantity,
                preference_text,
                booking.notes,
                assigned,
                _extract_section_from_alloc(performance_id, allocation.assigned_seats) or "",
                allocation.match_status.value,
                allocation.match_notes,
            ]
        )

    return Response(content=output.getvalue(), media_type="text/csv")


def _find_or_create_theatre(name: str, city: str) -> int:
    for theatre_id, theatre in store.theatres.items():
        if theatre.name.lower() == name.lower() and theatre.city.lower() == city.lower():
            return theatre_id

    theatre_id = store.next_theatre_id()
    from allocator.models import Theatre

    store.theatres[theatre_id] = Theatre(id=theatre_id, name=name, city=city)
    return theatre_id


def _get_performance(performance_id: int) -> Performance:
    performance = store.performances.get(performance_id)
    if not performance:
        raise HTTPException(status_code=404, detail="Performance not found")
    return performance


def _available_key(seat: AvailableSeat) -> str:
    return f"{seat.section}:{seat.row}:{seat.seat_number}"


def _infer_theatre_seats_from_available(theatre_id: int, available: list[AvailableSeat]) -> list[TheatreSeat]:
    by_row: dict[tuple[str, str], list[AvailableSeat]] = defaultdict(list)
    for seat in available:
        by_row[(seat.section, seat.row)].append(seat)

    inferred: list[TheatreSeat] = []
    for (section, row), row_seats in by_row.items():
        ordered = sorted(row_seats, key=lambda s: s.seat_number)
        center = (ordered[0].seat_number + ordered[-1].seat_number) / 2.0
        for i, seat in enumerate(ordered):
            inferred.append(
                TheatreSeat(
                    theatre_id=theatre_id,
                    section=section,
                    row=row,
                    seat_number=seat.seat_number,
                    seat_label=seat.seat_label,
                    is_aisle=(i == 0 or i == len(ordered) - 1),
                    is_accessible=False,
                    x_position=float(seat.seat_number - center),
                    y_position=float(_row_depth_value(row)),
                    adjacent_group_key=f"{section}:{row}:{seat.seat_number}",
                )
            )

    return sorted(inferred, key=lambda s: (s.section, s.row, s.seat_number))


def _row_depth_value(row: str) -> int:
    total = 0
    for ch in row.upper():
        if "A" <= ch <= "Z":
            total += ord(ch) - ord("A")
        elif ch.isdigit():
            total += int(ch)
    return total


def _extract_section_from_alloc(performance_id: int, assigned_seats: list[str]) -> str | None:
    if not assigned_seats:
        return None
    by_label = {
        seat.seat_label: seat.section
        for seat in store.available_seats_by_performance.get(performance_id, [])
    }
    sections = {by_label[label] for label in assigned_seats if label in by_label}
    if len(sections) == 1:
        return next(iter(sections))
    return None


def _find_booking(booking_id: int) -> tuple[int, Booking | None]:
    for performance_id, bookings in store.bookings_by_performance.items():
        for booking in bookings:
            if booking.id == booking_id:
                return performance_id, booking
    return -1, None


def _decode_csv(data: bytes) -> str:
    try:
        return data.decode("utf-8-sig")
    except UnicodeDecodeError:
        return data.decode("latin-1")


def _expected_seat_labels(allocation_rows: list[dict[str, str]]) -> set[str]:
    labels: set[str] = set()
    for row in allocation_rows:
        for seat in parse_seat_list(row.get("seats_raw", "")):
            labels.add(seat)
    return labels


def _seat_map_from_parsed_pages(parsed_pages) -> dict[str, int]:
    seat_map: dict[str, int] = {}
    for page in parsed_pages:
        if page.seat_label in seat_map:
            raise TicketBundleError(f"Duplicate ticket seat detected in PDF: {page.seat_label}")
        seat_map[page.seat_label] = page.page_index
    return seat_map


def _seat_map_from_page_results(parsed_page_results: list[ParsedTicketPageResult]) -> dict[str, int]:
    seat_map: dict[str, int] = {}
    for page in parsed_page_results:
        if page.seat_label in seat_map:
            raise TicketBundleError(f"Duplicate ticket seat detected in PDF: {page.seat_label}")
        seat_map[page.seat_label] = page.page_index
    return seat_map


def _build_preview_files(
    pdf_content: bytes,
    groups,
    filenames: list[str],
    pass_artifacts: dict[int, tuple[str, bytes]],
) -> dict[str, tuple[bytes, str]]:
    files: dict[str, tuple[bytes, str]] = {}
    for group, name in zip(groups, filenames):
        if not group_has_output_pdf(group):
            continue
        files[name] = (build_group_pdf(pdf_content, group), "application/pdf")
        for page_index in group.page_indexes:
            artifact = pass_artifacts.get(page_index)
            if artifact is None:
                continue
            pass_name, pass_blob = artifact
            files[pass_name] = (pass_blob, "application/vnd.apple.pkpass")
    return files


def _build_wallet_pass_rows(
    group,
    preview_id: str | None,
    pass_artifacts: dict[int, tuple[str, bytes]],
    parsed_page_results: list[ParsedTicketPageResult],
) -> list[dict[str, str]]:
    page_results_by_index = {page.page_index: page for page in parsed_page_results}
    wallet_rows: list[dict[str, str]] = []
    for page_index in group.page_indexes:
        artifact = pass_artifacts.get(page_index)
        if artifact is None:
            continue
        pass_name, pass_blob = artifact
        parsed_page = page_results_by_index.get(page_index)
        wallet_rows.append(
            {
                "seat_label": parsed_page.seat_label if parsed_page else "",
                "pass_file": pass_name,
                "pass_url": f"/ticket-bundles/preview/{preview_id}/files/{quote(pass_name)}" if preview_id else "",
                "pass_download_url": (
                    f"/ticket-bundles/preview/{preview_id}/files/{quote(pass_name)}?download=1" if preview_id else ""
                ),
                "pass_data_url": (
                    "data:application/vnd.apple.pkpass;base64,"
                    f"{base64.b64encode(pass_blob).decode('ascii')}"
                ),
            }
        )
    return wallet_rows


def _build_wallet_failure_rows_for_group(
    group,
    parsed_page_results: list[ParsedTicketPageResult],
) -> list[dict[str, str]]:
    page_results_by_index = {page.page_index: page for page in parsed_page_results}
    failures: list[dict[str, str]] = []
    for page_index in group.page_indexes:
        parsed_page = page_results_by_index.get(page_index)
        if parsed_page is None or parsed_page.wallet_ready:
            continue
        failures.append(
            {
                "seat_label": parsed_page.seat_label,
                "issue": parsed_page.wallet_error or "Wallet pass failed for this ticket.",
            }
        )
    return failures


def _collect_wallet_failure_rows(
    groups,
    parsed_page_results: list[ParsedTicketPageResult],
) -> list[dict[str, str]]:
    page_results_by_index = {page.page_index: page for page in parsed_page_results}
    failures: list[dict[str, str]] = []
    for group in groups:
        for page_index in group.page_indexes:
            parsed_page = page_results_by_index.get(page_index)
            if parsed_page is None or parsed_page.wallet_ready:
                continue
            failures.append(
                {
                    "email": group.email,
                    "booking_reference": group.booking_reference,
                    "seat_label": parsed_page.seat_label,
                    "issue": parsed_page.wallet_error or "Wallet pass failed for this ticket.",
                }
            )
    return failures


def _build_performance_metadata_from_page_results(parsed_pages) -> dict[str, str | bool | None]:
    if not parsed_pages:
        return {"performance_date": None, "performance_time": None, "confidence": False}
    first_page = next((page for page in parsed_pages if page.performance_date or page.performance_time), parsed_pages[0])
    return {
        "performance_date": first_page.performance_date,
        "performance_time": first_page.performance_time,
        "confidence": bool(first_page.performance_date and first_page.performance_time),
    }


def _store_preview_files(files: dict[str, tuple[bytes, str]]) -> str:
    preview_id = uuid4().hex
    preview_download_cache[preview_id] = files

    if len(preview_download_cache) > 20:
        oldest_key = next(iter(preview_download_cache))
        del preview_download_cache[oldest_key]

    return preview_id


def _describe_group_output_issue(group) -> str:
    if group.missing_seats:
        return f"No ticket found for: {', '.join(group.missing_seats)}"
    if not group.page_indexes:
        return "No matching ticket pages were found for this booking."
    return "This booking could not produce a PDF."


async def _read_allocation_rows(upload: UploadFile) -> list[dict[str, str]]:
    raw = await upload.read()
    name = (upload.filename or "").lower()
    try:
        if name.endswith(".numbers"):
            csv_text = _convert_numbers_bytes_to_csv(raw)
            return parse_allocation_csv(csv_text)
        return parse_allocation_csv(_decode_csv(raw))
    except TicketBundleError as exc:
        raise HTTPException(status_code=400, detail=str(exc)) from exc


def _convert_numbers_bytes_to_csv(numbers_blob: bytes) -> str:
    try:
        from numbers_parser import Document
    except Exception as exc:
        raise TicketBundleError(
            "Server cannot read .numbers files yet. Install dependency: numbers-parser."
        ) from exc

    with tempfile.TemporaryDirectory(prefix="ticket-bundle-") as tmp:
        in_path = Path(tmp) / "input.numbers"
        in_path.write_bytes(numbers_blob)

        try:
            doc = Document(str(in_path))
        except Exception as exc:
            raise TicketBundleError(
                "Failed to parse .numbers file. Ensure the file is valid and not password-protected."
            ) from exc

        tables: list[list[list[str]]] = []
        for sheet in doc.sheets:
            for table in sheet.tables:
                rows = table.rows(values_only=True)
                normalized_rows: list[list[str]] = [
                    [_numbers_cell_to_text(cell) for cell in row] for row in rows
                ]
                if any(any(cell for cell in row) for row in normalized_rows):
                    tables.append(normalized_rows)

        if not tables:
            raise TicketBundleError("Numbers file did not contain any readable table rows.")

        best_table = max(tables, key=_numbers_table_score)
        csv_text = _rows_to_csv(best_table)
        if not csv_text.strip():
            raise TicketBundleError("Numbers conversion produced empty CSV data.")
        return csv_text


def _numbers_cell_to_text(cell: object) -> str:
    if cell is None:
        return ""
    if isinstance(cell, bool):
        return "TRUE" if cell else "FALSE"
    return str(cell).strip()


def _numbers_table_score(rows: list[list[str]]) -> tuple[int, int, int]:
    non_empty_rows = [row for row in rows if any(cell.strip() for cell in row)]
    if not non_empty_rows:
        return (0, 0, 0)

    email_cells = sum(
        1
        for row in non_empty_rows
        for cell in row
        if EMAIL_CELL_RE.match(cell.strip())
    )
    seat_like_cells = sum(
        1
        for row in non_empty_rows
        for cell in row
        if re.search(r"[A-Za-z]{1,3}\s*[- ]?\s*\d{1,3}", cell) is not None
    )
    return (email_cells, seat_like_cells, len(non_empty_rows))


def _rows_to_csv(rows: list[list[str]]) -> str:
    if not rows:
        return ""
    width = max((len(row) for row in rows), default=0)
    output = StringIO()
    writer = csv.writer(output)
    for row in rows:
        padded = row + [""] * (width - len(row))
        writer.writerow(padded)
    return output.getvalue()
