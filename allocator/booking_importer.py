from __future__ import annotations

import csv
import re
from dataclasses import dataclass
from io import StringIO

from allocator.models import Booking, BookingPreference


class BookingImportError(ValueError):
    pass


@dataclass(slots=True)
class ParsedBookingImport:
    bookings: list[Booking]
    preferences: list[BookingPreference]


def parse_bookings_csv(
    content: str,
    performance_id: int,
    starting_booking_id: int = 1,
    known_sections: list[str] | None = None,
) -> ParsedBookingImport:
    reader = csv.DictReader(StringIO(content))
    if not reader.fieldnames:
        raise BookingImportError("Bookings CSV has no header row")

    required = {"booking_reference", "customer_name", "quantity"}
    headers = {h.strip() for h in reader.fieldnames if h}
    missing = required - headers
    if missing:
        raise BookingImportError(f"Bookings CSV missing required columns: {sorted(missing)}")

    bookings: list[Booking] = []
    preferences: list[BookingPreference] = []

    for index, raw in enumerate(reader):
        row = {k.strip(): (v or "").strip() for k, v in raw.items() if k}

        booking_id = starting_booking_id + index
        preference_text = _merge_preference_text(row.get("preferences", ""), row.get("notes", ""))

        booking = Booking(
            id=booking_id,
            performance_id=performance_id,
            booking_reference=row["booking_reference"],
            customer_name=row["customer_name"],
            quantity=int(row["quantity"]),
            notes=row.get("notes", ""),
        )

        preference = normalize_preference_text(
            booking_id=booking_id,
            raw_text=preference_text,
            known_sections=known_sections or [],
        )

        bookings.append(booking)
        preferences.append(preference)

    return ParsedBookingImport(bookings=bookings, preferences=preferences)


def normalize_preference_text(
    booking_id: int,
    raw_text: str,
    known_sections: list[str],
) -> BookingPreference:
    text = raw_text.lower()

    wants_aisle = "aisle" in text
    wants_central = any(token in text for token in ["central", "center", "middle"])

    avoid_front_patterns = [
        "avoid front",
        "not too near the front",
        "not near the front",
        "away from front",
        "back half",
    ]
    avoid_front = any(p in text for p in avoid_front_patterns)
    wants_front = ("front" in text) and not avoid_front

    accessible_required = any(
        token in text
        for token in [
            "accessible",
            "wheelchair",
            "mobility",
            "step-free",
            "disabled access",
        ]
    )

    must_sit_together = True
    if "split" in text and "okay" in text:
        must_sit_together = False
    if "must sit together" in text or "together" in text:
        must_sit_together = True

    section_preference, mandatory = _extract_section_preference(text, known_sections)
    near_booking_reference = _extract_near_booking_reference(text)

    return BookingPreference(
        booking_id=booking_id,
        wants_aisle=wants_aisle,
        wants_central=wants_central,
        wants_front=wants_front,
        avoid_front=avoid_front,
        section_preference=section_preference,
        section_preference_mandatory=mandatory,
        must_sit_together=must_sit_together,
        accessible_required=accessible_required,
        near_booking_reference=near_booking_reference,
        raw_request_text=raw_text.strip(),
    )


def _merge_preference_text(preferences: str, notes: str) -> str:
    if preferences and notes:
        return f"{preferences}; {notes}"
    return preferences or notes or ""


def _extract_near_booking_reference(text: str) -> str | None:
    match = re.search(r"(?:near|with)\s+([a-z]{0,3}\d{2,})", text)
    if not match:
        return None
    return match.group(1).upper()


def _extract_section_preference(text: str, known_sections: list[str]) -> tuple[str | None, bool]:
    if not known_sections:
        return None, False

    normalized = [(section, section.lower()) for section in known_sections]
    mandatory_phrases = ["only", "must", "mandatory", "strictly", "require"]

    for original, lowered in normalized:
        if lowered in text:
            mandatory = any(phrase in text for phrase in mandatory_phrases)
            return original, mandatory

    return None, False
