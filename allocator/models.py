from __future__ import annotations

from dataclasses import dataclass, field
from datetime import UTC, date, datetime, time
from enum import Enum
from typing import Optional


class SeatStatus(str, Enum):
    AVAILABLE = "available"
    HELD = "held"
    SOLD = "sold"


class MatchStatus(str, Enum):
    PERFECT_MATCH = "Perfect match"
    GOOD_MATCH = "Good match"
    PARTIAL_MATCH = "Partial match"
    NEEDS_MANUAL_REVIEW = "Needs manual review"
    UNALLOCATED = "Unallocated"


@dataclass(slots=True)
class Theatre:
    id: int
    name: str
    city: str


@dataclass(slots=True)
class TheatreSeat:
    theatre_id: int
    section: str
    row: str
    seat_number: int
    seat_label: str
    is_aisle: bool = False
    is_accessible: bool = False
    x_position: Optional[float] = None
    y_position: Optional[float] = None
    adjacent_group_key: Optional[str] = None


@dataclass(slots=True)
class Performance:
    id: int
    theatre_id: int
    show_name: str
    performance_date: date
    performance_time: time
    supplier_reference: Optional[str] = None


@dataclass(slots=True)
class AvailableSeat:
    performance_id: int
    theatre_seat_id: str
    section: str
    row: str
    seat_number: int
    seat_label: str
    status: SeatStatus = SeatStatus.AVAILABLE
    source_import_id: Optional[str] = None


@dataclass(slots=True)
class Booking:
    id: int
    performance_id: int
    booking_reference: str
    customer_name: str
    quantity: int
    notes: str = ""
    created_at: datetime = field(default_factory=lambda: datetime.now(UTC))
    priority_override: Optional[int] = None


@dataclass(slots=True)
class BookingPreference:
    booking_id: int
    wants_aisle: bool = False
    wants_central: bool = False
    wants_front: bool = False
    avoid_front: bool = False
    section_preference: Optional[str] = None
    section_preference_mandatory: bool = False
    must_sit_together: bool = True
    accessible_required: bool = False
    near_booking_reference: Optional[str] = None
    raw_request_text: str = ""


@dataclass(slots=True)
class Allocation:
    booking_id: int
    assigned_seats: list[str] = field(default_factory=list)
    match_status: MatchStatus = MatchStatus.UNALLOCATED
    match_notes: str = ""
    manually_edited: bool = False


@dataclass(slots=True)
class SeatPlanInferenceWarning:
    code: str
    message: str
    row: Optional[str] = None
    section: Optional[str] = None


@dataclass(slots=True)
class SeatPlanIngestionResult:
    theatre: Theatre
    theatre_seats: list[TheatreSeat]
    warnings: list[SeatPlanInferenceWarning] = field(default_factory=list)
    requires_manual_review: bool = False
