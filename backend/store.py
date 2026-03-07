from __future__ import annotations

from dataclasses import dataclass, field

from allocator.models import Allocation, AvailableSeat, Booking, BookingPreference, Performance, Theatre, TheatreSeat


@dataclass(slots=True)
class InMemoryStore:
    theatre_counter: int = 1
    performance_counter: int = 1
    booking_counter: int = 1

    theatres: dict[int, Theatre] = field(default_factory=dict)
    theatre_seats: dict[int, list[TheatreSeat]] = field(default_factory=dict)
    performances: dict[int, Performance] = field(default_factory=dict)

    available_seats_by_performance: dict[int, list[AvailableSeat]] = field(default_factory=dict)
    bookings_by_performance: dict[int, list[Booking]] = field(default_factory=dict)
    preferences_by_booking_id: dict[int, BookingPreference] = field(default_factory=dict)
    allocations_by_performance: dict[int, dict[int, Allocation]] = field(default_factory=dict)

    def next_theatre_id(self) -> int:
        value = self.theatre_counter
        self.theatre_counter += 1
        return value

    def next_performance_id(self) -> int:
        value = self.performance_counter
        self.performance_counter += 1
        return value

    def next_booking_id_start(self, count: int) -> int:
        value = self.booking_counter
        self.booking_counter += count
        return value
