from __future__ import annotations

import csv
import re
import zipfile
from dataclasses import dataclass
from io import BytesIO, StringIO


SEAT_TOKEN_RE = re.compile(r"\b([A-Za-z]{1,3})\s*[- ]?\s*(\d{1,3})\b")
SEAT_TOKEN_RE_REVERSED = re.compile(r"\b(\d{1,3})\s*[- ]?\s*([A-Za-z]{1,3})\b")
ROW_SEAT_LABEL_RE = re.compile(r"\bROW\b[\s:.-]*([A-Za-z]{1,3})[\s|/,-]*\bSEAT\b[\s:.-]*(\d{1,3})\b", re.IGNORECASE)
SEAT_ROW_LABEL_RE = re.compile(r"\bSEAT\b[\s:.-]*(\d{1,3})[\s|/,-]*\bROW\b[\s:.-]*([A-Za-z]{1,3})\b", re.IGNORECASE)
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
EMAIL_EXTRACT_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
SECTION_HINT_PATTERN = (
    r"(?:STALLS|CIRCLE|DRESS|GRAND|UPPER|LOWER|BALCONY|MEZZANINE|BOX|PIT|GALLERY|SECTION|PLATFORM\d*)"
)
SECTION_HINT_RE = re.compile(
    rf"\b{SECTION_HINT_PATTERN}\b",
    re.IGNORECASE,
)
SECTION_ROW_SEAT_RE = re.compile(
    rf"(?i)\b{SECTION_HINT_PATTERN}\s*[-:/]?\s*([A-Za-z]{{1,3}})\s*[- ]\s*(\d{{1,3}})"
)
ROW_SEAT_SECTION_RE = re.compile(
    rf"(?i)\b([A-Za-z]{{1,3}})\s*[- ]?\s*(\d{{1,3}})(?=(?:\s|[-/])*{SECTION_HINT_PATTERN})"
)
COMPACT_SEAT_AFTER_SECTION_RE = re.compile(
    rf"(?i)\b{SECTION_HINT_PATTERN}\s+([A-Za-z]{{1,3}})\s*(\d{{1,3}})(?=[A-Za-z]|$)"
)
SEAT_BEFORE_ORDER_RE = re.compile(
    r"(?i)\b([A-Za-z]{1,3})\s*[- ]?\s*(\d{1,3})(?=\s*ORDER(?:\b|[A-Z]))"
)
RESERVED_ROW_TOKENS = {
    "AM",
    "PM",
    "MON",
    "TUE",
    "TUES",
    "WED",
    "THU",
    "THUR",
    "THURS",
    "FRI",
    "SAT",
    "SUN",
    "JAN",
    "FEB",
    "MAR",
    "APR",
    "MAY",
    "JUN",
    "JUL",
    "AUG",
    "SEP",
    "SEPT",
    "OCT",
    "NOV",
    "DEC",
    "SW",
    "NW",
    "SE",
    "NE",
}
MONTH_NAME_TO_ABBR = {
    "JANUARY": "Jan",
    "JAN": "Jan",
    "FEBRUARY": "Feb",
    "FEB": "Feb",
    "MARCH": "Mar",
    "MAR": "Mar",
    "APRIL": "Apr",
    "APR": "Apr",
    "MAY": "May",
    "JUNE": "Jun",
    "JUN": "Jun",
    "JULY": "Jul",
    "JUL": "Jul",
    "AUGUST": "Aug",
    "AUG": "Aug",
    "SEPTEMBER": "Sep",
    "SEPT": "Sep",
    "SEP": "Sep",
    "OCTOBER": "Oct",
    "OCT": "Oct",
    "NOVEMBER": "Nov",
    "NOV": "Nov",
    "DECEMBER": "Dec",
    "DEC": "Dec",
}
MONTH_PATTERN = (
    r"JAN(?:UARY)?|FEB(?:RUARY)?|MAR(?:CH)?|APR(?:IL)?|MAY|JUN(?:E)?|JUL(?:Y)?|"
    r"AUG(?:UST)?|SEP(?:T(?:EMBER)?)?|OCT(?:OBER)?|NOV(?:EMBER)?|DEC(?:EMBER)?"
)
MONTH_DAY_RE = re.compile(
    rf"\b({MONTH_PATTERN})\.?\s+(\d{{1,2}})(?:ST|ND|RD|TH)?(?:,\s*\d{{2,4}})?\b",
    re.IGNORECASE,
)
DAY_MONTH_RE = re.compile(
    rf"\b(\d{{1,2}})(?:ST|ND|RD|TH)?\s+({MONTH_PATTERN})\.?(?:,\s*\d{{2,4}})?\b",
    re.IGNORECASE,
)
TIME_12H_RE = re.compile(r"\b(\d{1,2})(?::|\.)(\d{2})\s*([AP])\.?\s*M\.?\b", re.IGNORECASE)
TIME_12H_COMPACT_RE = re.compile(r"\b(\d{1,2})\s*([AP])\.?\s*M\.?\b", re.IGNORECASE)
TIME_24H_RE = re.compile(r"\b([01]?\d|2[0-3])[:.]([0-5]\d)\b")


@dataclass(slots=True)
class BookingTicketGroup:
    booking_reference: str
    customer_name: str
    email: str
    seat_labels: list[str]
    page_indexes: list[int]
    missing_seats: list[str]


def group_has_output_pdf(group: BookingTicketGroup) -> bool:
    return bool(group.page_indexes) and not group.missing_seats


def split_groups_for_output(
    groups: list[BookingTicketGroup],
) -> tuple[list[BookingTicketGroup], list[BookingTicketGroup]]:
    complete: list[BookingTicketGroup] = []
    excluded: list[BookingTicketGroup] = []

    for group in groups:
        if group_has_output_pdf(group):
            complete.append(group)
        else:
            excluded.append(group)

    return complete, excluded


class TicketBundleError(ValueError):
    pass


def parse_allocation_csv(content: str) -> list[dict[str, str]]:
    try:
        parsed_with_headers = _parse_allocation_csv_with_headers(content)
    except TicketBundleError:
        parsed_with_headers = []
    parsed_without_headers = _parse_allocation_csv_without_headers(content)

    if parsed_with_headers and parsed_without_headers:
        # Prefer the richer parse (usually with headers) if it did not drop rows.
        if len(parsed_with_headers) >= len(parsed_without_headers):
            return parsed_with_headers
        return parsed_without_headers
    if parsed_with_headers:
        return parsed_with_headers
    if parsed_without_headers:
        return parsed_without_headers

    raise TicketBundleError(
        "Could not parse allocation rows. Ensure the file contains email addresses and seat values."
    )


def _parse_allocation_csv_with_headers(content: str) -> list[dict[str, str]]:
    reader = csv.DictReader(StringIO(content))
    if not reader.fieldnames:
        return []

    normalized = {h.strip().lower(): h for h in reader.fieldnames if h}

    booking_col = _find_column(
        normalized,
        [
            "booking ref",
            "booking_reference",
            "booking reference",
            "ref",
            "order number",
            "order no",
            "order id",
        ],
        required=False,
    )
    email_col = _find_column(normalized, ["email", "customer email", "email address", "e-mail"])
    seats_col = _find_column(
        normalized,
        ["assigned seats", "assigned_seats", "seats", "seat labels", "allocated seats", "ticket seats"],
        required=False,
    )
    seat_row_col = _find_column(
        normalized,
        ["seat row", "seat_row", "row letter", "row"],
        required=False,
    )
    seat_number_col = _find_column(
        normalized,
        ["seat number", "seat_number", "seat no", "seat_no", "seat #", "seat"],
        required=False,
    )
    if not seats_col and not (seat_row_col and seat_number_col):
        raise TicketBundleError(
            "Missing required seat columns. Expected a seat label column or both row + seat number columns."
        )
    name_col = _find_column(
        normalized, ["customer name", "customer_name", "name", "lead name", "booker name"], required=False
    )

    rows: list[dict[str, str]] = []
    last_emails: list[str] = []
    last_customer_name = ""
    last_booking_ref = ""
    for raw in reader:
        booking_ref = (raw.get(booking_col, "") or "").strip() if booking_col else ""
        email_cell = (raw.get(email_col, "") or "").strip()
        emails = _extract_emails(email_cell)
        customer_name = (raw.get(name_col, "") or "").strip() if name_col else ""
        if _looks_like_unknown_name(customer_name):
            customer_name = ""
        seats_raw = _build_seats_raw(
            raw=raw,
            seats_col=seats_col,
            seat_row_col=seat_row_col,
            seat_number_col=seat_number_col,
        )
        if seats_raw:
            if not booking_ref and last_booking_ref:
                booking_ref = last_booking_ref
            if not emails and last_emails:
                emails = list(last_emails)
            if not customer_name and last_customer_name:
                customer_name = last_customer_name

        if not emails:
            continue

        for email in emails:
            row = {
                "booking_reference": booking_ref,
                "customer_name": customer_name,
                "email": email,
                "seats_raw": seats_raw,
            }
            rows.append(row)

        last_emails = list(emails)
        if customer_name:
            last_customer_name = customer_name
        if booking_ref:
            last_booking_ref = booking_ref

    if not rows:
        return []
    return rows


def _parse_allocation_csv_without_headers(content: str) -> list[dict[str, str]]:
    matrix: list[list[str]] = []
    for row in csv.reader(StringIO(content)):
        values = [cell.strip() for cell in row]
        if any(values):
            matrix.append(values)
    if not matrix:
        return []

    width = max(len(r) for r in matrix)
    if width == 0:
        return []

    def get_cell(r: list[str], idx: int) -> str:
        return r[idx].strip() if idx < len(r) else ""

    email_scores: list[int] = []
    seat_scores: list[int] = []
    for col in range(width):
        email_count = 0
        seat_count = 0
        for row in matrix:
            value = get_cell(row, col)
            if _extract_emails(value):
                email_count += 1
            if _looks_like_seat_cell(value):
                seat_count += 1
        email_scores.append(email_count)
        seat_scores.append(seat_count)

    email_col = max(range(width), key=lambda c: email_scores[c])
    if email_scores[email_col] == 0:
        return []

    seat_candidates = sorted(range(width), key=lambda c: seat_scores[c], reverse=True)
    seat_col = next((c for c in seat_candidates if c != email_col and seat_scores[c] > 0), None)
    row_col, number_col = _infer_split_seat_columns(matrix, email_col)
    if seat_col is None and (row_col is None or number_col is None):
        return []

    excluded = {email_col}
    if seat_col is not None:
        excluded.add(seat_col)
    if row_col is not None:
        excluded.add(row_col)
    if number_col is not None:
        excluded.add(number_col)
    booking_col = next((c for c in range(width) if c not in excluded), None)
    name_col = _infer_name_column(matrix, email_col=email_col, excluded=excluded)

    rows: list[dict[str, str]] = []
    last_emails: list[str] = []
    last_customer_name = ""
    last_booking_ref = ""
    for idx, row in enumerate(matrix):
        emails = _extract_emails(get_cell(row, email_col))
        customer_name = get_cell(row, name_col) if name_col is not None else ""
        if _looks_like_unknown_name(customer_name):
            customer_name = ""
        seats_raw = ""
        if seat_col is not None:
            seats_raw = get_cell(row, seat_col)
        if not seats_raw and row_col is not None and number_col is not None:
            seat_row = get_cell(row, row_col)
            seat_number = get_cell(row, number_col)
            seats_raw = _merge_split_seat_fields(seat_row, seat_number)

        booking_ref = get_cell(row, booking_col) if booking_col is not None else ""
        if seats_raw:
            if not booking_ref and last_booking_ref:
                booking_ref = last_booking_ref
            if not emails and last_emails:
                emails = list(last_emails)
            if not customer_name and last_customer_name:
                customer_name = last_customer_name

        if not emails:
            continue

        for email in emails:
            rows.append(
                {
                    "booking_reference": booking_ref or "",
                    "customer_name": customer_name,
                    "email": email,
                    "seats_raw": seats_raw,
                }
            )
        last_emails = list(emails)
        if customer_name:
            last_customer_name = customer_name
        if booking_ref:
            last_booking_ref = booking_ref

    return rows


def extract_pdf_page_seat_map(pdf_bytes: bytes, expected_seats: set[str] | None = None) -> dict[str, int]:
    PdfReader, _ = _load_pdf_backend()

    reader = PdfReader(BytesIO(pdf_bytes))
    seat_to_page: dict[str, int] = {}
    normalized_expected = _normalize_seat_labels(expected_seats or set())

    for page_index, page in enumerate(reader.pages):
        text = page.extract_text() or ""
        matches = _extract_seat_tokens(text)
        if not matches:
            matches = _extract_seat_tokens_from_page_content(page)

        if normalized_expected:
            filtered = [seat for seat in matches if seat in normalized_expected]
            if filtered:
                matches = filtered
            else:
                anchored = _extract_expected_seats_from_text(text, normalized_expected)
                if not anchored:
                    anchored = _extract_expected_seats_from_page_content(page, normalized_expected)
                if anchored:
                    matches = anchored

        for seat in matches:
            if normalized_expected and seat not in normalized_expected:
                continue
            seat_to_page.setdefault(seat, page_index)

    if not seat_to_page:
        raise TicketBundleError(
            "Could not detect seat labels in ticket PDF. Ensure each page contains labels like C1 or AA14."
        )

    return seat_to_page


def extract_ticket_performance_metadata(pdf_bytes: bytes) -> dict[str, str | bool | None]:
    PdfReader, _ = _load_pdf_backend()
    reader = PdfReader(BytesIO(pdf_bytes))

    text_parts: list[str] = []
    for page in reader.pages[:5]:
        try:
            page_text = page.extract_text() or ""
        except Exception:
            page_text = ""
        if page_text.strip():
            text_parts.append(page_text)

        content_tokens = _extract_string_literals_from_pdf_content(_extract_page_content_bytes(page))
        if content_tokens:
            text_parts.append(" ".join(content_tokens))

    combined_text = "\n".join(part for part in text_parts if part.strip())
    date_candidates = _extract_performance_date_candidates(combined_text)
    time_candidates = _extract_performance_time_candidates(combined_text)
    date_value = date_candidates[0] if len(date_candidates) == 1 else None
    time_value = time_candidates[0] if len(time_candidates) == 1 else None

    return {
        "performance_date": date_value,
        "performance_time": time_value,
        "confidence": bool(date_value and time_value),
    }


def _extract_performance_date_candidates(text: str) -> list[str]:
    if not text:
        return []

    matches: list[str] = []
    for match in MONTH_DAY_RE.finditer(text):
        month = _normalize_month(match.group(1))
        day = _normalize_day(match.group(2))
        if month and day:
            matches.append(f"{month} {day}")

    for match in DAY_MONTH_RE.finditer(text):
        month = _normalize_month(match.group(2))
        day = _normalize_day(match.group(1))
        if month and day:
            matches.append(f"{month} {day}")

    return _dedupe_strings(matches)


def _extract_performance_time_candidates(text: str) -> list[str]:
    if not text:
        return []

    matches: list[str] = []
    for match in TIME_12H_RE.finditer(text):
        formatted = _format_12h_time(match.group(1), match.group(2), match.group(3))
        if formatted:
            matches.append(formatted)

    for match in TIME_12H_COMPACT_RE.finditer(text):
        formatted = _format_12h_time(match.group(1), "00", match.group(2))
        if formatted:
            matches.append(formatted)

    for match in TIME_24H_RE.finditer(text):
        formatted = _format_24h_time(match.group(1), match.group(2))
        if formatted:
            matches.append(formatted)

    return _dedupe_strings(matches)


def _normalize_month(value: str) -> str | None:
    clean = re.sub(r"[^A-Za-z]", "", value or "").upper()
    return MONTH_NAME_TO_ABBR.get(clean)


def _normalize_day(value: str) -> str | None:
    try:
        day = int(re.sub(r"[^0-9]", "", value or ""))
    except ValueError:
        return None
    if day <= 0 or day > 31:
        return None
    return str(day)


def _format_12h_time(hour_value: str, minute_value: str, meridiem_value: str) -> str | None:
    try:
        hour = int(hour_value)
        minute = int(minute_value)
    except (TypeError, ValueError):
        return None
    if hour <= 0 or hour > 12 or minute < 0 or minute > 59:
        return None
    meridiem = (meridiem_value or "").strip().lower()
    if meridiem not in {"a", "p"}:
        return None
    return f"{hour}.{minute:02d}"


def _format_24h_time(hour_value: str, minute_value: str) -> str | None:
    try:
        hour = int(hour_value)
        minute = int(minute_value)
    except (TypeError, ValueError):
        return None
    if hour < 0 or hour > 23 or minute < 0 or minute > 59:
        return None

    display_hour = hour % 12 or 12
    return f"{display_hour}.{minute:02d}"


def _dedupe_strings(values: list[str]) -> list[str]:
    deduped: list[str] = []
    seen: set[str] = set()
    for value in values:
        key = value.strip().lower()
        if not key or key in seen:
            continue
        seen.add(key)
        deduped.append(value.strip())
    return deduped


def build_booking_groups(
    allocation_rows: list[dict[str, str]],
    seat_to_page: dict[str, int],
) -> list[BookingTicketGroup]:
    grouped_rows: dict[str, list[dict[str, str]]] = {}
    for idx, row in enumerate(allocation_rows):
        booking_ref = row["booking_reference"].strip()
        email_key = row["email"].strip().lower()
        normalized_ref = _normalize_booking_reference_for_grouping(booking_ref)
        key = f"{normalized_ref}|{email_key}" if normalized_ref else email_key
        if not key:
            key = f"row-{idx+1}"
        grouped_rows.setdefault(key, []).append(row)

    groups: list[BookingTicketGroup] = []
    for idx, group_rows in enumerate(grouped_rows.values(), start=1):
        first = group_rows[0]
        booking_ref = first["booking_reference"].strip() or f"ROW{idx}"
        email = first["email"].strip()
        customer_name = first["customer_name"].strip()

        seat_labels: list[str] = []
        pages: list[int] = []
        missing: list[str] = []
        for row in group_rows:
            for seat in parse_seat_list(row["seats_raw"]):
                if seat not in seat_labels:
                    seat_labels.append(seat)
                page = seat_to_page.get(seat)
                if page is None:
                    if seat not in missing:
                        missing.append(seat)
                    continue
                pages.append(page)

        groups.append(
            BookingTicketGroup(
                booking_reference=booking_ref,
                customer_name=customer_name,
                email=email,
                seat_labels=seat_labels,
                page_indexes=sorted(set(pages)),
                missing_seats=missing,
            )
        )

    return groups


def build_bundle_zip(pdf_bytes: bytes, groups: list[BookingTicketGroup]) -> bytes:
    complete_groups, _ = split_groups_for_output(groups)
    output = BytesIO()

    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        manifest_csv = _build_manifest_csv(complete_groups)
        archive.writestr("manifest.csv", manifest_csv)

        for group in complete_groups:
            pdf_blob = build_group_pdf(pdf_bytes, group)
            filename = output_pdf_filename(group)
            archive.writestr(filename, pdf_blob)

    output.seek(0)
    return output.getvalue()


def build_group_pdf(pdf_bytes: bytes, group: BookingTicketGroup) -> bytes:
    PdfReader, PdfWriter = _load_pdf_backend()
    reader = PdfReader(BytesIO(pdf_bytes))
    writer = PdfWriter()
    for idx in group.page_indexes:
        if idx < 0 or idx >= len(reader.pages):
            continue
        writer.add_page(reader.pages[idx])

    pdf_blob = BytesIO()
    writer.write(pdf_blob)
    return pdf_blob.getvalue()


def parse_seat_list(seats_raw: str) -> list[str]:
    normalized = re.sub(r"\bto\b", "-", seats_raw, flags=re.IGNORECASE)
    tokens = re.split(r"[;,\n]+", normalized)
    if len(tokens) == 1:
        tokens = re.split(r"\s{2,}", normalized)

    out: list[str] = []

    for token in tokens:
        clean = token.strip()
        if not clean:
            continue

        if "-" in clean or re.search(r"\bto\b", clean, flags=re.IGNORECASE):
            out.extend(_expand_range(clean))
            continue

        out.extend(_extract_seat_tokens(clean))

    # de-duplicate while preserving order
    deduped: list[str] = []
    seen: set[str] = set()
    for seat in out:
        if seat in seen:
            continue
        seen.add(seat)
        deduped.append(seat)

    return deduped


def _expand_range(token: str) -> list[str]:
    token = re.sub(r"\bto\b", "-", token, flags=re.IGNORECASE)
    parts = [p.strip() for p in token.split("-") if p.strip()]
    if len(parts) != 2:
        return _extract_seat_tokens(token)

    start = _parse_single_seat(parts[0])
    end = _parse_single_seat(parts[1], default_row=start[0] if start else None)
    if not start or not end:
        return _extract_seat_tokens(token)

    start_row, start_num = start
    end_row, end_num = end
    if start_row != end_row:
        return _extract_seat_tokens(token)

    lo, hi = sorted((start_num, end_num))
    return [f"{start_row}{i}" for i in range(lo, hi + 1)]


def _parse_single_seat(token: str, default_row: str | None = None) -> tuple[str, int] | None:
    number_only = re.search(r"\b(\d{1,3})\b", token)
    if default_row and number_only and not re.search(r"[A-Za-z]", token):
        return default_row.upper(), int(number_only.group(1))

    m = SEAT_TOKEN_RE.search(token)
    if not m:
        return None
    row = m.group(1).upper()
    number = int(m.group(2))
    return row, number


def _extract_seat_tokens(text: str) -> list[str]:
    out: list[str] = []

    for m in SEAT_TOKEN_RE.finditer(text):
        token = _seat_token(m.group(1), m.group(2))
        if token:
            out.append(token)

    # Tickets like "Stalls-D-30SECTIONROWSEAT" need section-aware extraction.
    for m in SECTION_ROW_SEAT_RE.finditer(text):
        token = _seat_token(m.group(1), m.group(2))
        if token:
            out.append(token)
    for m in ROW_SEAT_SECTION_RE.finditer(text):
        token = _seat_token(m.group(1), m.group(2))
        if token:
            out.append(token)
    for m in COMPACT_SEAT_AFTER_SECTION_RE.finditer(text):
        token = _seat_token(m.group(1), m.group(2))
        if token:
            out.append(token)
    for m in SEAT_BEFORE_ORDER_RE.finditer(text):
        token = _seat_token(m.group(1), m.group(2))
        if token:
            out.append(token)

    if out:
        return _dedupe_seat_tokens(out)

    # Some PDFs expose seat fields as labeled text blocks (ROW J / SEAT 11).
    for m in ROW_SEAT_LABEL_RE.finditer(text):
        token = _seat_token(m.group(1), m.group(2))
        if token:
            out.append(token)
    for m in SEAT_ROW_LABEL_RE.finditer(text):
        token = _seat_token(m.group(2), m.group(1))
        if token:
            out.append(token)
    if out:
        return _dedupe_seat_tokens(out)

    # Fallback for PDFs that render split boxes as "11 J STALLS" rather than "J 11".
    if "ROW" in text.upper() and "SEAT" in text.upper():
        for line in text.splitlines():
            line = line.strip()
            if not line:
                continue
            if not SECTION_HINT_RE.search(line):
                continue
            for m in SEAT_TOKEN_RE.finditer(line):
                token = _seat_token(m.group(1), m.group(2))
                if token:
                    out.append(token)
            for m in SEAT_TOKEN_RE_REVERSED.finditer(line):
                token = _seat_token(m.group(2), m.group(1))
                if token:
                    out.append(token)

    if out:
        return _dedupe_seat_tokens(out)

    # Last text-level fallback: infer split row/seat tokens across lines/blocks.
    return _extract_seats_from_token_sequence(_tokenize(text))


def _extract_seat_tokens_from_page_content(page) -> list[str]:
    data = _extract_page_content_bytes(page)
    if not data:
        return []
    return _extract_seat_tokens_from_pdf_content_data(data)


def _extract_expected_seats_from_page_content(page, expected_seats: set[str]) -> list[str]:
    data = _extract_page_content_bytes(page)
    if not data:
        return []
    return _extract_expected_seats_from_pdf_content_data(data, expected_seats)


def _extract_page_content_bytes(page) -> bytes:
    try:
        contents = page.get_contents()
    except Exception:
        return b""
    if not contents:
        return b""

    try:
        if isinstance(contents, list):
            data = b"".join(_pdf_content_data(c) for c in contents if c is not None)
        else:
            data = _pdf_content_data(contents)
    except Exception:
        return b""
    return data


def _pdf_content_data(content_obj) -> bytes:
    if hasattr(content_obj, "get_data"):
        return content_obj.get_data()
    if hasattr(content_obj, "getData"):  # PyPDF2 compatibility
        return content_obj.getData()
    return b""


def _extract_seat_tokens_from_pdf_content_data(data: bytes) -> list[str]:
    literals = _extract_string_literals_from_pdf_content(data)

    if not literals:
        return []

    combined_text = " ".join(literals)
    direct = _extract_seat_tokens(combined_text)
    if direct:
        return direct
    return _extract_seats_from_token_sequence(_tokenize(combined_text))


def _extract_expected_seats_from_pdf_content_data(data: bytes, expected_seats: set[str]) -> list[str]:
    literals = _extract_string_literals_from_pdf_content(data)
    if not literals:
        return []
    return _extract_expected_seats_from_text(" ".join(literals), expected_seats)


def _extract_string_literals_from_pdf_content(data: bytes) -> list[str]:
    stream_text = data.decode("latin-1", errors="ignore")
    literals: list[str] = []

    for m in re.finditer(r"\((?:\\.|[^\\()])*\)", stream_text):
        value = m.group(0)[1:-1]
        value = value.replace(r"\\", "\\")
        value = value.replace(r"\(", "(").replace(r"\)", ")")
        value = value.replace(r"\n", " ").replace(r"\r", " ").replace(r"\t", " ")
        cleaned = value.strip()
        if cleaned:
            literals.append(cleaned)
    return literals


def _extract_expected_seats_from_text(text: str, expected_seats: set[str]) -> list[str]:
    if not text or not expected_seats:
        return []
    tokens = _tokenize(text)
    return _extract_expected_seats_from_tokens(tokens, expected_seats)


def _extract_expected_seats_from_tokens(tokens: list[str], expected_seats: set[str]) -> list[str]:
    out: list[str] = []
    if not tokens or not expected_seats:
        return out

    normalized_tokens = [token.strip().upper() for token in tokens if token.strip()]
    for token in normalized_tokens:
        seat = _normalize_seat_label(token)
        if seat and seat in expected_seats:
            out.append(seat)
            continue

        reversed_match = re.fullmatch(r"(\d{1,3})\s*([A-Z]{1,3})", token)
        if not reversed_match:
            continue
        seat = _seat_token(reversed_match.group(2), reversed_match.group(1))
        if seat and seat in expected_seats:
            out.append(seat)

    for idx, first in enumerate(normalized_tokens):
        window_end = min(len(normalized_tokens), idx + 5)
        for next_idx in range(idx + 1, window_end):
            second = normalized_tokens[next_idx]

            if re.fullmatch(r"[A-Z]{1,3}", first) and re.fullmatch(r"\d{1,3}", second):
                seat = _seat_token(first, second)
                if seat and seat in expected_seats:
                    out.append(seat)

            if re.fullmatch(r"\d{1,3}", first) and re.fullmatch(r"[A-Z]{1,3}", second):
                seat = _seat_token(second, first)
                if seat and seat in expected_seats:
                    out.append(seat)

    return _dedupe_seat_tokens(out)


def _normalize_seat_labels(seat_labels: set[str]) -> set[str]:
    normalized: set[str] = set()
    for seat_label in seat_labels:
        seat = _normalize_seat_label(seat_label)
        if seat:
            normalized.add(seat)
    return normalized


def _normalize_seat_label(value: str) -> str | None:
    clean = value.strip().upper()
    if not clean:
        return None

    parsed_seats = parse_seat_list(clean)
    if parsed_seats:
        return parsed_seats[0]

    forward_match = re.fullmatch(r"([A-Z]{1,3})\s*[- ]?\s*(\d{1,3})", clean)
    if forward_match:
        return _seat_token(forward_match.group(1), forward_match.group(2))

    reversed_match = re.fullmatch(r"(\d{1,3})\s*[- ]?\s*([A-Z]{1,3})", clean)
    if reversed_match:
        return _seat_token(reversed_match.group(2), reversed_match.group(1))

    return None


def _extract_seats_from_token_sequence(tokens: list[str]) -> list[str]:
    if len(tokens) < 2:
        return []

    out: list[str] = []
    for idx in range(len(tokens) - 1):
        first = tokens[idx]
        second = tokens[idx + 1]

        if _is_row_token(first) and _is_seat_number_token(second):
            if _has_section_hint_nearby(tokens, idx):
                seat = _seat_token(first, second)
                if seat:
                    out.append(seat)
                continue

        if _is_seat_number_token(first) and _is_row_token(second):
            if _has_section_hint_nearby(tokens, idx):
                seat = _seat_token(second, first)
                if seat:
                    out.append(seat)

    return _dedupe_seat_tokens(out)


def _tokenize(text: str) -> list[str]:
    return re.findall(r"[A-Za-z]+|\d+", text)


def _is_row_token(token: str) -> bool:
    value = token.strip().upper()
    if not re.fullmatch(r"[A-Z]{1,3}", value):
        return False
    if value in {"ROW", "SEAT", "LEVEL", "ORDER", "ID", "AM", "PM"}:
        return False
    if SECTION_HINT_RE.fullmatch(value):
        return False
    return True


def _is_seat_number_token(token: str) -> bool:
    return bool(re.fullmatch(r"\d{1,3}", token.strip()))


def _has_section_hint_nearby(tokens: list[str], anchor_index: int, radius: int = 8) -> bool:
    start = max(0, anchor_index - radius)
    end = min(len(tokens), anchor_index + radius + 1)
    for idx in range(start, end):
        if SECTION_HINT_RE.fullmatch(tokens[idx].strip().upper()):
            return True
    return False


def _seat_token(row: str, seat: str) -> str | None:
    row_value = row.strip().upper()
    if not row_value or len(row_value) > 3 or not row_value.isalpha():
        return None
    if row_value in RESERVED_ROW_TOKENS:
        return None

    try:
        seat_number = int(seat)
    except (TypeError, ValueError):
        return None
    if seat_number <= 0 or seat_number > 999:
        return None
    return f"{row_value}{seat_number}"


def _dedupe_seat_tokens(tokens: list[str]) -> list[str]:
    deduped: list[str] = []
    seen: set[str] = set()
    for token in tokens:
        if token in seen:
            continue
        seen.add(token)
        deduped.append(token)
    return deduped


def _normalize_booking_reference_for_grouping(value: str) -> str:
    ref = value.strip().lower()
    if not ref:
        return ""

    # Notes-like values (e.g. "extra", "central") should not split bookings.
    if not any(ch.isdigit() for ch in ref):
        return ""

    letters = sum(1 for ch in ref if ch.isalpha())
    digits = sum(1 for ch in ref if ch.isdigit())
    if digits == 0 or (letters + digits) < 3:
        return ""

    return ref


def _infer_name_column(matrix: list[list[str]], email_col: int, excluded: set[int]) -> int | None:
    width = max(len(r) for r in matrix)

    def get_cell(r: list[str], idx: int) -> str:
        return r[idx].strip() if idx < len(r) else ""

    scores: list[int] = [0] * width
    for col in range(width):
        if col in excluded:
            continue
        score = 0
        for row in matrix:
            value = get_cell(row, col)
            if not value:
                continue
            if _extract_emails(value):
                continue
            if _looks_like_seat_cell(value):
                continue
            if re.search(r"\d", value):
                continue
            if re.search(r"[A-Za-z]{2,}", value):
                score += 1
        # Prefer nearby column (often immediately left of email).
        score -= abs(col - email_col)
        scores[col] = score

    candidate = max(range(width), key=lambda c: scores[c]) if width else None
    if candidate is None:
        return None
    if candidate in excluded or scores[candidate] <= 0:
        return None
    return candidate


def _looks_like_unknown_name(value: str) -> bool:
    clean = value.strip().lower()
    if not clean:
        return True
    normalized = re.sub(r"[^a-z]+", "", clean)
    return normalized in {"unknown", "psunknown", "nnameunknown", "noname"}


def _find_column(normalized_headers: dict[str, str], aliases: list[str], required: bool = True) -> str | None:
    for alias in aliases:
        if alias in normalized_headers:
            return normalized_headers[alias]

    if required:
        raise TicketBundleError(f"Missing required CSV column. Expected one of: {aliases}")
    return None


def _is_email(value: str) -> bool:
    return bool(EMAIL_RE.match(value.strip()))


def _extract_emails(value: str) -> list[str]:
    if not value:
        return []
    emails = [match.group(0).strip() for match in EMAIL_EXTRACT_RE.finditer(value)]
    if not emails and _is_email(value):
        emails = [value.strip()]
    deduped: list[str] = []
    seen: set[str] = set()
    for email in emails:
        key = email.lower()
        if key in seen:
            continue
        seen.add(key)
        deduped.append(email)
    return deduped


def _looks_like_seat_cell(value: str) -> bool:
    text = value.strip()
    if not text:
        return False
    if SEAT_TOKEN_RE.search(text):
        return True
    if re.search(r"\bto\b", text, flags=re.IGNORECASE) and re.search(r"\d", text):
        return True
    if "-" in text and re.search(r"\d", text):
        return True
    return False


def _build_seats_raw(
    raw: dict[str, str],
    seats_col: str | None,
    seat_row_col: str | None,
    seat_number_col: str | None,
) -> str:
    if seats_col:
        combined = (raw.get(seats_col, "") or "").strip()
        if combined:
            return combined
    if seat_row_col and seat_number_col:
        seat_row = (raw.get(seat_row_col, "") or "").strip()
        seat_number = (raw.get(seat_number_col, "") or "").strip()
        return _merge_split_seat_fields(seat_row, seat_number)
    return ""


def _merge_split_seat_fields(seat_row: str, seat_number: str) -> str:
    row_clean = seat_row.strip().upper()
    num_clean = seat_number.strip()
    if not row_clean or not num_clean:
        return ""
    if re.search(r"[A-Za-z]", num_clean):
        return num_clean
    return f"{row_clean}{num_clean}"


def _infer_split_seat_columns(matrix: list[list[str]], email_col: int) -> tuple[int | None, int | None]:
    width = max(len(r) for r in matrix)
    row_scores: list[int] = [0] * width
    number_scores: list[int] = [0] * width

    def get_cell(r: list[str], idx: int) -> str:
        return r[idx].strip() if idx < len(r) else ""

    for col in range(width):
        if col == email_col:
            continue
        for row in matrix:
            value = get_cell(row, col)
            if re.fullmatch(r"[A-Za-z]{1,3}", value):
                row_scores[col] += 1
            if re.fullmatch(r"\d{1,3}([-/]\d{1,3})?", value):
                number_scores[col] += 1

    row_col = max(range(width), key=lambda c: row_scores[c]) if width else None
    number_col = max(range(width), key=lambda c: number_scores[c]) if width else None
    if row_col is None or number_col is None:
        return None, None
    if row_col == email_col or number_col == email_col or row_col == number_col:
        return None, None
    if row_scores[row_col] == 0 or number_scores[number_col] == 0:
        return None, None
    return row_col, number_col


def _build_manifest_csv(groups: list[BookingTicketGroup]) -> str:
    filenames = build_output_filenames(groups)
    output = StringIO()
    writer = csv.writer(output)
    writer.writerow(["Email", "PDF File"])

    for group, filename in zip(groups, filenames):
        writer.writerow([group.email, filename])

    return output.getvalue()


def _safe_filename(value: str) -> str:
    cleaned = re.sub(r"[^A-Za-z0-9._-]+", "_", value.strip())
    return cleaned.strip("_") or "unknown"


def output_pdf_filename(group: BookingTicketGroup) -> str:
    customer_name = group.customer_name or ""
    if _looks_like_unknown_name(customer_name):
        customer_name = ""
    customer_name = customer_name.strip()

    if customer_name:
        safe_name = _safe_filename(customer_name)
        return f"{safe_name}_tickets.pdf"

    email = (group.email or "").strip()
    if email:
        safe_email = _safe_filename(email)
        return f"{safe_email}.pdf"

    safe_booking_ref = _safe_filename(group.booking_reference or "booking")
    return f"{safe_booking_ref}_tickets.pdf"


def build_output_filenames(groups: list[BookingTicketGroup]) -> list[str]:
    used: set[str] = set()
    output: list[str] = []
    duplicate_counts: dict[str, int] = {}

    for group in groups:
        base = output_pdf_filename(group)
        candidate = base

        if candidate in used:
            duplicate_counts[base] = duplicate_counts.get(base, 1) + 1
            suffix = _safe_filename((group.email or "").split("@")[0])
            if suffix:
                stem = base[:-4] if base.lower().endswith(".pdf") else base
                candidate = f"{stem}_{suffix}.pdf"
            if candidate in used:
                stem = base[:-4] if base.lower().endswith(".pdf") else base
                candidate = f"{stem}_{duplicate_counts[base]}.pdf"

        used.add(candidate)
        output.append(candidate)

    return output


def _load_pdf_backend():
    try:  # pragma: no branch
        from pypdf import PdfReader, PdfWriter

        return PdfReader, PdfWriter
    except Exception:
        try:
            from PyPDF2 import PdfReader, PdfWriter

            return PdfReader, PdfWriter
        except Exception as exc:  # pragma: no cover - environment dependent
            raise TicketBundleError(
                "PDF processing library missing. Install with: pip install pypdf"
            ) from exc
