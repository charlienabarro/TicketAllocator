from __future__ import annotations

import base64
import csv
import hashlib
import json
import os
import re
import zipfile
from dataclasses import dataclass
from datetime import datetime, timedelta
from io import BytesIO, StringIO


SEAT_TOKEN_RE = re.compile(r"\b([A-Za-z]{1,3})\s*[- ]?\s*(\d{1,3})\b")
SEAT_TOKEN_RE_REVERSED = re.compile(r"\b(\d{1,3})\s*[- ]?\s*([A-Za-z]{1,3})\b")
GENERIC_ADMISSION_TOKEN_RE = re.compile(
    r"\b(?:STANDING|DANCE\s*FLOOR(?:\s*GA)?|GA)\s*[- ]?\s*(\d{1,3})\b",
    re.IGNORECASE,
)
GENERIC_ADMISSION_RANGE_RE = re.compile(
    r"\b(?:STANDING|DANCE\s*FLOOR(?:\s*GA)?|GA)\s*(\d{1,3})\s*-\s*(\d{1,3})\b",
    re.IGNORECASE,
)
CANONICAL_GENERIC_ADMISSION_RE = re.compile(r"STANDING(\d{1,3})\b", re.IGNORECASE)
ROW_SEAT_LABEL_RE = re.compile(r"\bROW\b[\s:.-]*([A-Za-z]{1,3})[\s|/,-]*\bSEAT\b[\s:.-]*(\d{1,3})\b", re.IGNORECASE)
SEAT_ROW_LABEL_RE = re.compile(r"\bSEAT\b[\s:.-]*(\d{1,3})[\s|/,-]*\bROW\b[\s:.-]*([A-Za-z]{1,3})\b", re.IGNORECASE)
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
EMAIL_EXTRACT_RE = re.compile(r"[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}")
SECTION_HINT_PATTERN = (
    r"(?:STAGE|STALLS|CIRCLE|DRESS|GRAND|UPPER|LOWER|BALCONY|MEZZANINE|BOX|PIT|GALLERY|SECTION|PLATFORM\d*)"
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
PACKED_SEAT_AFTER_SECTION_RE = re.compile(
    rf"(?i)(?<![A-Za-z]){SECTION_HINT_PATTERN}\s*([A-Za-z]{{1,3}})\s*(\d{{1,3}})\b"
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
    "BY",
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
DAY_NAME_PATTERN = r"(?:Mon(?:day)?|Tue(?:sday)?|Wed(?:nesday)?|Thu(?:rsday)?|Fri(?:day)?|Sat(?:urday)?|Sun(?:day)?)"
FULL_DATE_WITH_YEAR_RE = re.compile(
    rf"(?:{DAY_NAME_PATTERN}\s+)?\d{{1,2}}\s+(?:{MONTH_PATTERN})\s+\d{{4}}",
    re.IGNORECASE,
)
VENUE_KEYWORD_RE = re.compile(
    r"\b("
    r"theatre|theater|arena|hall|playhouse|opera|centre|center|room|auditorium|vic|lyceum|apollo|"
    r"troubadour|criterion|hippodrome|coliseum|bridge|wyndham|garrick|aldwych|duke|royal|palace|"
    r"warehouse|studio|arts|festival"
    r")\b",
    re.IGNORECASE,
)
SHOW_NOISE_RE = re.compile(
    r"\b(order|booking|barcode|qr|seat|row|section|ticket|group|price|doors|performance|sponsor|production)\b",
    re.IGNORECASE,
)
SHOW_LABEL_BLACKLIST = {"event"}
LOWERCASE_TITLE_CONNECTORS = {
    "a",
    "an",
    "and",
    "as",
    "at",
    "by",
    "for",
    "from",
    "in",
    "of",
    "on",
    "or",
    "the",
    "to",
    "with",
}
DISCLAIMER_PHRASES = (
    "Tickets cannot be sold on for commercial gain",
    "Failure to comply will make this ticket",
    "Registered Charity",
    "Venue and Travel info",
    "No need to print your ticket",
)
KNOWN_SHOW_PATTERNS = [
    (re.compile(r"\babba voyage\b", re.IGNORECASE), "ABBA Voyage"),
    (re.compile(r"one flew over the cuckoo[’']?s nest", re.IGNORECASE), "One Flew Over the Cuckoo's Nest"),
    (re.compile(r"titan[íi]que", re.IGNORECASE), "Titanique"),
    (re.compile(r"paddington the musical", re.IGNORECASE), "Paddington The Musical"),
    (re.compile(r"\bhamilton\b", re.IGNORECASE), "Hamilton"),
    (re.compile(r"\bthe car man\b", re.IGNORECASE), "The Car Man"),
    (re.compile(r"jesus christ superstar", re.IGNORECASE), "Jesus Christ Superstar"),
    (re.compile(r"inter alia", re.IGNORECASE), "Inter Alia"),
    (re.compile(r"the devil wears prada", re.IGNORECASE), "The Devil Wears Prada"),
    (re.compile(r"the holy rosenbergs", re.IGNORECASE), "The Holy Rosenbergs"),
    (re.compile(r"back to the future: the musical", re.IGNORECASE), "Back To The Future: The Musical"),
    (re.compile(r"disney[’']?s the lion king", re.IGNORECASE), "Disney's The Lion King"),
    (re.compile(r"sinatra the musical", re.IGNORECASE), "Sinatra The Musical"),
    (re.compile(r"marie\s*&\s*rosetta", re.IGNORECASE), "Marie & Rosetta"),
]
KNOWN_VENUE_HINTS = [
    (re.compile(r"\babba arena\b", re.IGNORECASE), "ABBA Arena"),
    (re.compile(r"oldvictheatre\.com|the old vic", re.IGNORECASE), "The Old Vic"),
    (re.compile(r"criterion theatre|cri-terion-theatre\.co\.uk", re.IGNORECASE), "Criterion Theatre"),
    (re.compile(r"savoy theatre", re.IGNORECASE), "Savoy Theatre"),
    (re.compile(r"victoria palace theatre", re.IGNORECASE), "Victoria Palace Theatre"),
    (re.compile(r"wyndham'?s theatre", re.IGNORECASE), "Wyndham's Theatre"),
    (re.compile(r"london palladium|the london palladium", re.IGNORECASE), "The London Palladium"),
    (re.compile(r"menier chocolate factory", re.IGNORECASE), "Menier Chocolate Factory"),
    (re.compile(r"adelphi theatre", re.IGNORECASE), "Adelphi Theatre"),
    (re.compile(r"lyceum theatre", re.IGNORECASE), "Lyceum Theatre"),
    (re.compile(r"dominion theatre", re.IGNORECASE), "Dominion Theatre"),
    (re.compile(r"aldwych theatre", re.IGNORECASE), "Aldwych Theatre"),
    (re.compile(r"rosebery avenue,\s*london ec1r 4tn|sadler'?s wells", re.IGNORECASE), "Sadler's Wells Theatre"),
    (re.compile(r"sohoplace|4 soho place", re.IGNORECASE), "@sohoplace"),
]
PASS_ICON_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+jXioAAAAASUVORK5CYII="
)
WALLET_PASS_TYPE_IDENTIFIER = os.getenv("WALLET_PASS_TYPE_IDENTIFIER", "pass.com.ticketallocator.test")
WALLET_TEAM_IDENTIFIER = os.getenv("WALLET_TEAM_IDENTIFIER", "TKTTEST123")
WALLET_ORGANIZATION_NAME = os.getenv("WALLET_ORGANIZATION_NAME", "Ticket Allocator")
WALLET_PASS_DESCRIPTION = os.getenv("WALLET_PASS_DESCRIPTION", "Theatre ticket")


@dataclass(slots=True)
class BookingTicketGroup:
    booking_reference: str
    customer_name: str
    email: str
    seat_labels: list[str]
    page_indexes: list[int]
    missing_seats: list[str]


@dataclass(slots=True)
class ParsedTicketPage:
    page_index: int
    show_name: str
    performance_date: str
    performance_time: str
    venue_name: str
    row: str
    seat: str
    seat_label: str
    qr_payload: str


@dataclass(slots=True)
class ParsedTicketPageResult:
    page_index: int
    show_name: str
    performance_date: str
    performance_time: str
    venue_name: str
    row: str
    seat: str
    seat_label: str
    qr_payload: str | None
    wallet_error: str | None = None

    @property
    def wallet_ready(self) -> bool:
        return bool(self.qr_payload and not self.wallet_error)

    def to_parsed_ticket_page(self) -> ParsedTicketPage:
        if not self.wallet_ready:
            raise TicketBundleError(self.wallet_error or f"Wallet pass could not be built for seat {self.seat_label}.")
        return ParsedTicketPage(
            page_index=self.page_index,
            show_name=self.show_name,
            performance_date=self.performance_date,
            performance_time=self.performance_time,
            venue_name=self.venue_name,
            row=self.row,
            seat=self.seat,
            seat_label=self.seat_label,
            qr_payload=str(self.qr_payload),
        )


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
        if _should_ignore_time_match(text, match.start(), match.end()):
            continue
        formatted = _format_12h_time(match.group(1), match.group(2), match.group(3))
        if formatted:
            matches.append(formatted)

    for match in TIME_12H_COMPACT_RE.finditer(text):
        if _should_ignore_time_match(text, match.start(), match.end()):
            continue
        formatted = _format_12h_time(match.group(1), "00", match.group(2))
        if formatted:
            matches.append(formatted)

    for match in TIME_24H_RE.finditer(text):
        if _should_ignore_time_match(text, match.start(), match.end()):
            continue
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


def _should_ignore_time_match(text: str, start: int, end: int) -> bool:
    before = text[max(0, start - 24) : start].lower()
    after = text[end : min(len(text), end + 12)].lower()
    if "£" in before:
        return True
    if any(term in before for term in ("price", "commission", "levy", "restoration", "contact", "help@", "mon-fri", "9am")):
        return True
    if any(term in after for term in ("commission", "levy")) and "@" not in before:
        return False
    return False


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


def parse_ticket_pdf_page_results(
    pdf_bytes: bytes,
    expected_seats: set[str] | None = None,
    decode_wallet: bool = True,
) -> list[ParsedTicketPageResult]:
    fitz = _load_fitz_backend()
    document = fitz.open(stream=pdf_bytes, filetype="pdf")
    normalized_expected = _normalize_seat_labels(expected_seats or set())
    remaining_generic_admission_labels = _sorted_generic_admission_labels(normalized_expected)
    synthetic_generic_admission_count = 0
    raw_pages: list[dict[str, str | int | None]] = []
    aggregate_text_parts: list[str] = []

    for page_index in range(document.page_count):
        page = document.load_page(page_index)
        page_text = _extract_fitz_page_text(page)
        seat_label = _extract_page_seat_label_from_page(page, page_text, normalized_expected)
        if seat_label is None and _is_generic_admission_page(page_text):
            if remaining_generic_admission_labels:
                seat_label = remaining_generic_admission_labels.pop(0)
            elif not normalized_expected:
                synthetic_generic_admission_count += 1
                seat_label = _generic_admission_label(synthetic_generic_admission_count)
        if normalized_expected and seat_label is None:
            continue
        aggregate_text_parts.append(page_text)
        raw_pages.append(
            {
                "page_index": page_index,
                "text": page_text,
                "seat_label": seat_label,
            }
        )

    if not raw_pages:
        raise TicketBundleError(
            "Could not detect seat labels in ticket PDF. Ensure each page contains labels like C1 or AA14."
        )

    document_metadata = _extract_ticket_document_metadata(aggregate_text_parts)
    parsed_pages: list[ParsedTicketPageResult] = []
    for raw_page in raw_pages:
        page_index = int(raw_page["page_index"])
        page = document.load_page(page_index)
        page_text = str(raw_page["text"])
        seat_label = raw_page["seat_label"] or _extract_page_seat_label_from_page(page, page_text, normalized_expected)
        if not seat_label:
            raise TicketBundleError(f"Could not detect row and seat on ticket page {page_index + 1}.")
        row, seat = _split_seat_label(seat_label)
        page_metadata = _extract_ticket_page_metadata(page_text)
        show_name = page_metadata["show_name"] or document_metadata["show_name"]
        performance_date = page_metadata["performance_date"] or document_metadata["performance_date"]
        performance_time = page_metadata["performance_time"] or document_metadata["performance_time"]
        venue_name = page_metadata["venue_name"] or document_metadata["venue_name"]
        qr_payload: str | None = None
        wallet_error: str | None = None
        if decode_wallet:
            try:
                qr_payload = _decode_qr_payload_for_page(page)
            except TicketBundleError as exc:
                wallet_error = str(exc)

        missing_fields = [
            label
            for label, value in (
                ("show name", show_name),
                ("date", performance_date),
                ("time", performance_time),
                ("venue", venue_name),
            )
            if not value
        ]
        if missing_fields and not wallet_error:
            wallet_error = f"Could not detect {', '.join(missing_fields)} on ticket page {page_index + 1}."
        if decode_wallet and not qr_payload and not wallet_error:
            wallet_error = f"Could not decode the original QR code on ticket page {page_index + 1}."
        parsed_pages.append(
            ParsedTicketPageResult(
                page_index=page_index,
                show_name=str(show_name or ""),
                performance_date=str(performance_date or ""),
                performance_time=str(performance_time or ""),
                venue_name=str(venue_name or ""),
                row=row,
                seat=seat,
                seat_label=seat_label,
                qr_payload=qr_payload,
                wallet_error=wallet_error,
            )
        )

    return parsed_pages


def parse_ticket_pdf_pages(pdf_bytes: bytes, expected_seats: set[str] | None = None) -> list[ParsedTicketPage]:
    parsed_results = parse_ticket_pdf_page_results(pdf_bytes, expected_seats=expected_seats, decode_wallet=True)
    parsed_pages: list[ParsedTicketPage] = []
    for result in parsed_results:
        parsed_pages.append(result.to_parsed_ticket_page())
    return parsed_pages


def _extract_fitz_page_text(page) -> str:
    text_parts: list[str] = []
    try:
        page_text = page.get_text("text") or ""
    except Exception:
        page_text = ""
    if page_text.strip():
        text_parts.append(page_text)

    try:
        words = page.get_text("words") or []
    except Exception:
        words = []
    if words:
        ordered_words = [str(word[4]).strip() for word in sorted(words, key=lambda item: (item[5], item[6], item[7], item[0]))]
        word_text = " ".join(word for word in ordered_words if word)
        if word_text:
            text_parts.append(word_text)

    return "\n".join(part for part in text_parts if part.strip())


def _extract_page_seat_label(text: str, expected_seats: set[str]) -> str | None:
    matches = _extract_seat_tokens(text)
    if expected_seats:
        filtered = [seat for seat in matches if seat in expected_seats]
        if len(filtered) == 1:
            return filtered[0]
        if len(filtered) > 1:
            raise TicketBundleError(
                "Found multiple matching seats on a ticket page. Please ensure each page contains exactly one ticket."
            )
        anchored = _extract_expected_seats_from_text(text, expected_seats)
        if len(anchored) == 1:
            return anchored[0]
        if len(anchored) > 1:
            raise TicketBundleError(
                "Found multiple matching seats on a ticket page. Please ensure each page contains exactly one ticket."
            )
        return None

    if len(matches) == 1:
        return matches[0]
    if len(matches) > 1:
        raise TicketBundleError(
            "Found multiple seat labels on a ticket page. Please ensure each page contains exactly one ticket."
        )
    return None


def _extract_page_seat_label_from_page(page, text: str, expected_seats: set[str]) -> str | None:
    abba_label = _extract_abba_ticket_seat_label_from_page(page, text)
    if abba_label:
        if expected_seats and abba_label not in expected_seats:
            return None
        return abba_label
    if _is_generic_admission_page(text):
        return None
    return _extract_page_seat_label(text, expected_seats)


def _extract_abba_ticket_seat_label_from_page(page, text: str) -> str | None:
    if not _looks_like_abba_ticket_text(text):
        return None

    try:
        words = page.get_text("words") or []
    except Exception:
        words = []
    if not words:
        return None

    top_words = [word for word in words if float(word[1]) < 280]
    if not top_words:
        return None

    line_groups: list[dict[str, object]] = []
    for word in sorted(top_words, key=lambda item: (item[1], item[0])):
        token = str(word[4]).strip()
        if not token:
            continue
        y_pos = float(word[1])
        x_pos = float(word[0])
        if line_groups and abs(float(line_groups[-1]["y"]) - y_pos) <= 6:
            cast_words = line_groups[-1]["words"]
            assert isinstance(cast_words, list)
            cast_words.append((x_pos, token))
        else:
            line_groups.append({"y": y_pos, "words": [(x_pos, token)]})

    normalized_lines: list[list[str]] = []
    for line in line_groups:
        words_in_line = line["words"]
        assert isinstance(words_in_line, list)
        ordered = [token for _x, token in sorted(words_in_line, key=lambda item: item[0])]
        filtered = [token for token in ordered if token != "-"]
        if filtered:
            normalized_lines.append(filtered)

    for line in reversed(normalized_lines):
        if len(line) != 2:
            continue
        first, second = line
        if _is_row_token(first) and _is_seat_number_token(second):
            seat = _seat_token(first, second)
            if seat:
                return seat
        if _is_seat_number_token(first) and _is_row_token(second):
            seat = _seat_token(second, first)
            if seat:
                return seat

    return None


def _looks_like_abba_ticket_text(text: str) -> bool:
    lowered = (text or "").lower()
    return "abba voyage" in lowered and "abba arena" in lowered and "groups ticket" in lowered


def _is_generic_admission_page(text: str) -> bool:
    lowered = (text or "").lower()
    return _looks_like_abba_ticket_text(text) and "dance floor" in lowered


def _extract_ticket_document_metadata(texts: list[str]) -> dict[str, str | None]:
    combined_text = _normalize_metadata_text("\n".join(text for text in texts if text.strip()))
    return {
        "show_name": _first_candidate(_extract_show_name_candidates(combined_text)),
        "performance_date": _select_single_candidate(_extract_performance_date_candidates(combined_text)),
        "performance_time": _select_single_candidate(_extract_performance_time_candidates(combined_text)),
        "venue_name": _first_candidate(_extract_venue_candidates(combined_text)),
    }


def _extract_ticket_page_metadata(text: str) -> dict[str, str | None]:
    normalized_text = _normalize_metadata_text(text)
    return {
        "show_name": _first_candidate(_extract_show_name_candidates(normalized_text)),
        "performance_date": _select_single_candidate(_extract_performance_date_candidates(normalized_text)),
        "performance_time": _select_single_candidate(_extract_performance_time_candidates(normalized_text)),
        "venue_name": _first_candidate(_extract_venue_candidates(normalized_text)),
    }


def _select_single_candidate(values: list[str]) -> str | None:
    if not values:
        return None
    if len(values) == 1:
        return values[0]
    return None


def _first_candidate(values: list[str]) -> str | None:
    return values[0] if values else None


def _extract_show_name_candidates(text: str) -> list[str]:
    normalized_text = _normalize_metadata_text(text)
    candidates: list[str] = list(_extract_known_show_candidates(normalized_text))
    candidates.extend(_extract_show_name_candidates_near_dates(normalized_text))
    candidates.extend(_extract_show_name_candidates_after_dates(normalized_text))
    for line in _clean_ticket_lines(normalized_text):
        if not _looks_like_show_line(line):
            continue
        candidates.append(line)
    return _dedupe_strings(candidates)


def _extract_venue_candidates(text: str) -> list[str]:
    normalized_text = _normalize_metadata_text(text)
    known_candidates = _extract_known_venue_candidates(normalized_text)
    if known_candidates:
        return known_candidates

    candidates: list[str] = []
    for line in _clean_ticket_lines(normalized_text):
        if not VENUE_KEYWORD_RE.search(line):
            continue
        if _looks_like_disclaimer_line(line):
            continue
        if SHOW_NOISE_RE.search(line):
            continue
        if not re.search(r"[A-Za-z]{3,}", line):
            continue
        candidates.append(line)
    return _dedupe_strings(candidates)


def _clean_ticket_lines(text: str) -> list[str]:
    raw_lines = re.split(r"[\r\n]+", text or "")
    cleaned: list[str] = []
    for raw_line in raw_lines:
        line = re.sub(r"\s+", " ", raw_line).strip(" |")
        if not line:
            continue
        cleaned.append(line)
    return cleaned


def _normalize_metadata_text(text: str) -> str:
    normalized_lines: list[str] = []
    for raw_line in re.split(r"[\r\n]+", text or ""):
        line = raw_line
        line = re.sub(r"(?<=[a-z])(?=[A-Z])", " ", line)
        line = re.sub(r"(?<=\d{4})(?=\d{1,2}(?::|\.)\d{2}\s*[APap]?)", " ", line)
        line = re.sub(r"(?<=[AP]M)(?=[A-Z])", " ", line)
        line = re.sub(r"(\d{1,2}:\d{2})(?=[A-Z])", r"\1 ", line)
        line = re.sub(r"\s+", " ", line).strip()
        if line:
            normalized_lines.append(line)
    return "\n".join(normalized_lines)


def _extract_show_name_candidates_near_dates(text: str) -> list[str]:
    candidates: list[str] = []
    for match in FULL_DATE_WITH_YEAR_RE.finditer(text):
        window = text[max(0, match.start() - 140) : match.start()]
        candidate = _extract_trailing_title_from_context(window)
        if candidate:
            candidates.append(candidate)
    return _dedupe_strings(candidates)


def _extract_show_name_candidates_after_dates(text: str) -> list[str]:
    candidates: list[str] = []
    for match in FULL_DATE_WITH_YEAR_RE.finditer(text):
        window = text[match.end() : min(len(text), match.end() + 120)]
        window = re.sub(r"^[^A-Za-z]+", "", window)
        candidate = _extract_leading_title_from_context(window)
        if candidate:
            candidates.append(candidate)
    return _dedupe_strings(candidates)


def _extract_trailing_title_from_context(context: str) -> str | None:
    tokens = re.findall(r"[A-Za-z0-9'’&.-]+", context)
    if not tokens:
        return None

    collected: list[str] = []
    seen_title_word = False
    for token in reversed(tokens):
        clean = token.strip()
        if not clean:
            continue
        lowered = clean.lower()
        if lowered in SHOW_LABEL_BLACKLIST:
            break
        if clean[0].isupper():
            collected.append(clean)
            seen_title_word = True
            continue
        if lowered in LOWERCASE_TITLE_CONNECTORS and seen_title_word:
            collected.append(lowered)
            continue
        break

    if not collected:
        return None
    candidate = " ".join(reversed(collected)).strip()
    if not _looks_like_show_line(candidate):
        return None
    return candidate


def _extract_leading_title_from_context(context: str) -> str | None:
    tokens = re.findall(r"[A-Za-z0-9'’&:-]+", context)
    if not tokens:
        return None

    collected: list[str] = []
    for token in tokens:
        clean = token.strip()
        if not clean:
            continue
        lowered = clean.lower()
        if lowered in SHOW_LABEL_BLACKLIST:
            continue
        if clean[0].isupper():
            collected.append(clean)
            continue
        if lowered in LOWERCASE_TITLE_CONNECTORS and collected:
            collected.append(lowered)
            continue
        break

    if not collected:
        return None
    candidate = " ".join(collected).strip()
    if not _looks_like_show_line(candidate):
        return None
    return candidate


def _extract_known_show_candidates(text: str) -> list[str]:
    candidates: list[tuple[int, str]] = []
    for pattern, canonical in KNOWN_SHOW_PATTERNS:
        match = pattern.search(text)
        if match:
            candidates.append((match.start(), canonical))
    candidates.sort(key=lambda item: item[0])
    return [value for _index, value in candidates]


def _extract_known_venue_candidates(text: str) -> list[str]:
    candidates: list[tuple[int, str]] = []
    for pattern, canonical in KNOWN_VENUE_HINTS:
        match = pattern.search(text)
        if match:
            candidates.append((match.start(), canonical))
    candidates.sort(key=lambda item: item[0])
    return [value for _index, value in candidates]


def _looks_like_disclaimer_line(line: str) -> bool:
    return any(phrase.lower() in line.lower() for phrase in DISCLAIMER_PHRASES)


def _looks_like_show_line(line: str) -> bool:
    if not line or len(line) < 3:
        return False
    lowered = line.strip().lower()
    if lowered in SHOW_LABEL_BLACKLIST:
        return False
    if lowered in {"level", "date", "time", "lead booker:", "return code: £"}:
        return False
    if "customer name" in lowered or "order reference" in lowered or "order id" in lowered:
        return False
    if lowered == "united kingdom":
        return False
    if MONTH_DAY_RE.search(line) or DAY_MONTH_RE.search(line):
        return False
    if TIME_12H_RE.search(line) or TIME_12H_COMPACT_RE.search(line) or TIME_24H_RE.search(line):
        return False
    if VENUE_KEYWORD_RE.search(line) and not any(pattern.search(line) for pattern, _canonical in KNOWN_SHOW_PATTERNS):
        return False
    if _looks_like_disclaimer_line(line):
        return False
    if SHOW_NOISE_RE.search(line):
        return False
    if re.fullmatch(r"[\d\s.]+", line):
        return False
    if re.search(r"\b[A-Z]{1,3}\d{1,3}\b", line):
        return False
    words = re.findall(r"[A-Za-z][A-Za-z'&.-]*", line)
    if not words:
        return False
    title_like_words = [word for word in words if word[0].isupper()]
    return len(title_like_words) >= max(1, len(words) // 2)


def _split_seat_label(seat_label: str) -> tuple[str, str]:
    generic_match = CANONICAL_GENERIC_ADMISSION_RE.fullmatch(seat_label.strip().upper())
    if generic_match:
        return "STANDING", generic_match.group(1)
    match = re.fullmatch(r"([A-Z]{1,3})(\d{1,3})", seat_label.strip().upper())
    if not match:
        raise TicketBundleError(f"Could not split seat label: {seat_label}")
    return match.group(1), match.group(2)


def _decode_qr_payload_for_page(page) -> str:
    fitz = _load_fitz_backend()
    cv2, np = _load_qr_decoder_backend()
    detector = cv2.QRCodeDetector()
    for scale in (2, 3, 4):
        pixmap = page.get_pixmap(matrix=fitz.Matrix(scale, scale), alpha=False)
        channels = max(1, pixmap.n)
        image = np.frombuffer(pixmap.samples, dtype=np.uint8).reshape(pixmap.height, pixmap.width, channels)
        if channels == 4:
            image = cv2.cvtColor(image, cv2.COLOR_BGRA2BGR)
        elif channels == 1:
            image = cv2.cvtColor(image, cv2.COLOR_GRAY2BGR)

        for candidate in _qr_decode_candidates(cv2, image):
            try:
                data, _points, _straight = detector.detectAndDecode(candidate)
            except Exception:
                data = ""
            if data and data.strip():
                return data.strip()

    raise TicketBundleError("Could not decode a QR code from one or more ticket pages.")


def _qr_decode_candidates(cv2, image):
    height, width = image.shape[:2]
    grayscale = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    candidates = [image, grayscale]
    if width > 10 and height > 10:
        candidates.extend(
            [
                image[height // 2 :, :],
                image[:, width // 2 :],
                image[height // 3 :, width // 3 :],
                grayscale[height // 2 :, :],
                grayscale[:, width // 2 :],
                grayscale[height // 3 :, width // 3 :],
            ]
        )
    return candidates


def _extract_printed_barcode_value_from_page(page) -> str | None:
    candidates: list[tuple[float, int, str]] = []
    try:
        words = page.get_text("words") or []
    except Exception:
        words = []

    for word in words:
        value = str(word[4]).strip()
        if not re.fullmatch(r"\d{10,24}", value):
            continue
        top = float(word[1])
        bottom = float(word[3])
        candidates.append((bottom, len(value), value))

    if not candidates:
        try:
            text = page.get_text("text") or ""
        except Exception:
            text = ""
        standalone = _dedupe_strings(re.findall(r"\b\d{10,24}\b", text))
        if len(standalone) == 1:
            return standalone[0]
        return None

    candidates.sort(key=lambda item: (item[0], item[1]), reverse=True)
    values = [value for _bottom, _length, value in candidates]
    unique_values = _dedupe_strings(values)
    if len(unique_values) == 1:
        return unique_values[0]

    bottom_most = [item for item in candidates if abs(item[0] - candidates[0][0]) < 5]
    bottom_unique = _dedupe_strings([value for _bottom, _length, value in bottom_most])
    if len(bottom_unique) == 1:
        return bottom_unique[0]

    return None


def build_pkpass_for_ticket(ticket: ParsedTicketPage) -> bytes:
    pass_payload = _build_pass_payload(ticket)
    serialized_pass = json.dumps(pass_payload, sort_keys=True, separators=(",", ":")).encode("utf-8")
    files = {
        "pass.json": serialized_pass,
        "icon.png": PASS_ICON_PNG,
        "icon@2x.png": PASS_ICON_PNG,
    }
    manifest = {name: hashlib.sha1(blob).hexdigest() for name, blob in files.items()}
    files["manifest.json"] = json.dumps(manifest, sort_keys=True, separators=(",", ":")).encode("utf-8")

    output = BytesIO()
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        for name, blob in files.items():
            archive.writestr(name, blob)
    output.seek(0)
    return output.getvalue()


def _build_pass_payload(ticket: ParsedTicketPage) -> dict[str, object]:
    serial_source = f"{ticket.show_name}|{ticket.performance_date}|{ticket.performance_time}|{ticket.seat_label}|{ticket.qr_payload}"
    serial_number = hashlib.sha1(serial_source.encode("utf-8")).hexdigest()
    barcode = {
        "format": "PKBarcodeFormatQR",
        "message": ticket.qr_payload,
        "messageEncoding": "iso-8859-1",
        "altText": ticket.seat_label,
    }
    return {
        "formatVersion": 1,
        "passTypeIdentifier": WALLET_PASS_TYPE_IDENTIFIER,
        "serialNumber": serial_number,
        "teamIdentifier": WALLET_TEAM_IDENTIFIER,
        "organizationName": WALLET_ORGANIZATION_NAME,
        "description": WALLET_PASS_DESCRIPTION,
        "logoText": ticket.show_name,
        "eventTicket": {
            "headerFields": [
                {"key": "venue", "label": "Venue", "value": ticket.venue_name},
            ],
            "primaryFields": [
                {"key": "show", "label": "Show", "value": ticket.show_name},
            ],
            "secondaryFields": [
                {"key": "date", "label": "Date", "value": ticket.performance_date},
                {"key": "time", "label": "Time", "value": ticket.performance_time},
            ],
            "auxiliaryFields": [
                {"key": "row", "label": "Row", "value": ticket.row},
                {"key": "seat", "label": "Seat", "value": ticket.seat},
            ],
        },
        "barcode": barcode,
        "barcodes": [barcode],
    }


def build_booking_groups(
    allocation_rows: list[dict[str, str]],
    seat_to_page: dict[str, int],
) -> list[BookingTicketGroup]:
    grouped_rows: dict[str, list[dict[str, str]]] = {}
    for idx, row in enumerate(allocation_rows):
        booking_ref = row["booking_reference"].strip()
        email_key = row["email"].strip().lower()
        customer_name_key = _normalize_customer_name_for_grouping(row["customer_name"])
        normalized_ref = _normalize_booking_reference_for_grouping(booking_ref)
        if normalized_ref:
            key = f"{normalized_ref}|{email_key}"
        elif email_key and customer_name_key:
            key = f"{email_key}|{customer_name_key}"
        else:
            key = email_key
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
                if page not in pages:
                    pages.append(page)

        groups.append(
            BookingTicketGroup(
                booking_reference=booking_ref,
                customer_name=customer_name,
                email=email,
                seat_labels=seat_labels,
                page_indexes=pages,
                missing_seats=missing,
            )
        )

    return groups


def build_bundle_zip(
    pdf_bytes: bytes,
    groups: list[BookingTicketGroup],
    parsed_pages: list[ParsedTicketPage] | None = None,
) -> bytes:
    complete_groups, _ = split_groups_for_output(groups)
    output = BytesIO()
    base_modified_at = _build_bundle_zip_base_modified_at()
    pass_artifacts = build_pkpass_artifacts(parsed_pages or [])
    written_pass_pages: set[int] = set()

    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as archive:
        manifest_csv = _build_manifest_csv(complete_groups)
        manifest_info = zipfile.ZipInfo("manifest.csv", date_time=_zip_info_datetime(base_modified_at, len(complete_groups)))
        manifest_info.compress_type = zipfile.ZIP_DEFLATED
        archive.writestr(manifest_info, manifest_csv)

        for index, group in enumerate(complete_groups):
            pdf_blob = build_group_pdf(pdf_bytes, group)
            filename = output_pdf_filename(group)
            file_info = zipfile.ZipInfo(filename, date_time=_zip_info_datetime(base_modified_at, index))
            file_info.compress_type = zipfile.ZIP_DEFLATED
            archive.writestr(file_info, pdf_blob)

            for pass_page_index in group.page_indexes:
                if pass_page_index in written_pass_pages:
                    continue
                artifact = pass_artifacts.get(pass_page_index)
                if artifact is None:
                    continue
                pass_name, pass_blob = artifact
                pass_info = zipfile.ZipInfo(
                    f"wallet/{pass_name}",
                    date_time=_zip_info_datetime(base_modified_at, len(complete_groups) + len(written_pass_pages) + 1),
                )
                pass_info.compress_type = zipfile.ZIP_DEFLATED
                archive.writestr(pass_info, pass_blob)
                written_pass_pages.add(pass_page_index)

    output.seek(0)
    return output.getvalue()


def _build_bundle_zip_base_modified_at() -> datetime:
    now = datetime.now().replace(microsecond=0)
    even_second = now.second - (now.second % 2)
    return now.replace(second=even_second)


def _zip_info_datetime(base_modified_at: datetime, index: int) -> tuple[int, int, int, int, int, int]:
    modified_at = base_modified_at - timedelta(seconds=index * 4)
    if modified_at.year < 1980:
        modified_at = modified_at.replace(year=1980, month=1, day=1, hour=0, minute=0, second=0)
    even_second = modified_at.second - (modified_at.second % 2)
    modified_at = modified_at.replace(second=even_second)
    return (
        modified_at.year,
        modified_at.month,
        modified_at.day,
        modified_at.hour,
        modified_at.minute,
        modified_at.second,
    )


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

        generic_labels = _extract_generic_admission_tokens(clean)
        if generic_labels:
            out.extend(generic_labels)
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
    generic_tokens = _extract_generic_admission_tokens(token)
    if generic_tokens:
        return generic_tokens
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
    for m in PACKED_SEAT_AFTER_SECTION_RE.finditer(text):
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


def _extract_generic_admission_tokens(text: str) -> list[str]:
    clean = text.strip().upper()
    if not clean:
        return []

    range_match = GENERIC_ADMISSION_RANGE_RE.fullmatch(clean)
    if range_match:
        start = int(range_match.group(1))
        end = int(range_match.group(2))
        lo, hi = sorted((start, end))
        return [_generic_admission_label(number) for number in range(lo, hi + 1)]

    single_match = GENERIC_ADMISSION_TOKEN_RE.fullmatch(clean)
    if single_match:
        return [_generic_admission_label(single_match.group(1))]

    canonical_match = CANONICAL_GENERIC_ADMISSION_RE.fullmatch(clean)
    if canonical_match:
        return [_generic_admission_label(canonical_match.group(1))]

    return []


def _generic_admission_label(number: int | str) -> str:
    return f"STANDING{int(number)}"


def _sorted_generic_admission_labels(seat_labels: set[str]) -> list[str]:
    generic_labels = [seat_label for seat_label in seat_labels if CANONICAL_GENERIC_ADMISSION_RE.fullmatch(seat_label)]
    return sorted(generic_labels, key=_generic_admission_sort_key)


def _generic_admission_sort_key(seat_label: str) -> int:
    match = CANONICAL_GENERIC_ADMISSION_RE.fullmatch(seat_label.strip().upper())
    if not match:
        return 0
    return int(match.group(1))


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


def _normalize_customer_name_for_grouping(value: str) -> str:
    clean = " ".join(value.strip().lower().split())
    if not clean:
        return ""
    if _looks_like_unknown_name(clean):
        return ""
    return clean


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
    if GENERIC_ADMISSION_TOKEN_RE.search(text) or CANONICAL_GENERIC_ADMISSION_RE.search(text):
        return True
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


def output_pkpass_filename(ticket: ParsedTicketPage) -> str:
    return f"{_safe_filename(ticket.row)}-{_safe_filename(ticket.seat)}.pkpass"


def build_output_pass_filenames(tickets: list[ParsedTicketPage]) -> list[str]:
    used: set[str] = set()
    output: list[str] = []

    for ticket in tickets:
        base = output_pkpass_filename(ticket)
        candidate = base
        duplicate_index = 2
        while candidate in used:
            stem = base[:-7] if base.lower().endswith(".pkpass") else base
            candidate = f"{stem}_{duplicate_index}.pkpass"
            duplicate_index += 1
        used.add(candidate)
        output.append(candidate)

    return output


def build_pkpass_artifacts(tickets: list[ParsedTicketPage]) -> dict[int, tuple[str, bytes]]:
    filenames = build_output_pass_filenames(tickets)
    return {
        ticket.page_index: (filename, build_pkpass_for_ticket(ticket))
        for ticket, filename in zip(tickets, filenames)
    }


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


def _load_fitz_backend():
    try:
        import fitz

        return fitz
    except Exception as exc:
        raise TicketBundleError("PyMuPDF is required for Wallet pass generation. Install with: pip install PyMuPDF") from exc


def _load_qr_decoder_backend():
    try:
        import cv2
        import numpy as np

        return cv2, np
    except Exception as exc:
        raise TicketBundleError(
            "OpenCV is required for QR decoding. Install with: pip install opencv-python-headless"
        ) from exc
