from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass

from allocator.models import Allocation, AvailableSeat, Booking, BookingPreference, MatchStatus, TheatreSeat


@dataclass(slots=True)
class AllocationRunResult:
    allocations: list[Allocation]
    unallocated_booking_ids: list[int]


def run_allocation(
    bookings: list[Booking],
    preferences: dict[int, BookingPreference],
    available_seats: list[AvailableSeat],
    theatre_seats: list[TheatreSeat],
) -> AllocationRunResult:
    seat_meta = {
        (seat.section, seat.row, seat.seat_number): seat
        for seat in theatre_seats
    }

    grouped_available: dict[tuple[str, str], list[AvailableSeat]] = defaultdict(list)
    for seat in available_seats:
        grouped_available[(seat.section, seat.row)].append(seat)

    for seats in grouped_available.values():
        seats.sort(key=lambda s: s.seat_number)

    consumed: set[str] = set()
    allocations: list[Allocation] = []
    unallocated: list[int] = []

    ordered_bookings = sorted(bookings, key=lambda b: _booking_sort_key(b, preferences.get(b.id)))

    for booking in ordered_bookings:
        pref = preferences.get(booking.id, BookingPreference(booking_id=booking.id))
        block, score = _find_best_block(booking, pref, grouped_available, consumed, seat_meta)

        if not block:
            allocations.append(
                Allocation(
                    booking_id=booking.id,
                    assigned_seats=[],
                    match_status=MatchStatus.UNALLOCATED,
                    match_notes="No valid contiguous block found.",
                )
            )
            unallocated.append(booking.id)
            continue

        for seat in block:
            consumed.add(_seat_key(seat))

        allocations.append(
            Allocation(
                booking_id=booking.id,
                assigned_seats=[seat.seat_label for seat in block],
                match_status=_status_from_score(score),
                match_notes=_match_notes(pref, score),
            )
        )

    return AllocationRunResult(allocations=allocations, unallocated_booking_ids=unallocated)


def _find_best_block(
    booking: Booking,
    pref: BookingPreference,
    grouped_available: dict[tuple[str, str], list[AvailableSeat]],
    consumed: set[str],
    seat_meta: dict[tuple[str, str, int], TheatreSeat],
) -> tuple[list[AvailableSeat], float]:
    best_block: list[AvailableSeat] = []
    best_score = float("-inf")

    for (section, _row), seats in grouped_available.items():
        if pref.section_preference_mandatory and pref.section_preference:
            if section.lower() != pref.section_preference.lower():
                continue

        for block in _iter_contiguous_blocks(seats, booking.quantity, consumed):
            if pref.accessible_required and not _block_has_accessible(block, seat_meta):
                continue

            score = _score_block(block, pref, seat_meta)
            if score > best_score:
                best_score = score
                best_block = block

    return best_block, best_score


def _iter_contiguous_blocks(
    sorted_row_seats: list[AvailableSeat],
    size: int,
    consumed: set[str],
) -> list[list[AvailableSeat]]:
    if size <= 0:
        return []

    candidates: list[list[AvailableSeat]] = []
    for i in range(0, len(sorted_row_seats) - size + 1):
        block = sorted_row_seats[i : i + size]
        if any(_seat_key(seat) in consumed for seat in block):
            continue

        contiguous = True
        for j in range(1, len(block)):
            if block[j].seat_number != block[j - 1].seat_number + 1:
                contiguous = False
                break
        if contiguous:
            candidates.append(block)

    return candidates


def _score_block(
    block: list[AvailableSeat],
    pref: BookingPreference,
    seat_meta: dict[tuple[str, str, int], TheatreSeat],
) -> float:
    score = 100.0

    if pref.section_preference:
        if block[0].section.lower() == pref.section_preference.lower():
            score += 18.0
        elif pref.section_preference_mandatory:
            return float("-inf")
        else:
            score -= 12.0

    if pref.wants_aisle:
        if _block_has_aisle(block, seat_meta):
            score += 12.0
        else:
            score -= 6.0

    if pref.wants_central:
        score += _centrality_score(block, seat_meta)

    if pref.wants_front:
        score += _front_score(block, seat_meta)

    if pref.avoid_front:
        score += _avoid_front_score(block, seat_meta)

    score += _fragmentation_penalty(block)
    return score


def _booking_sort_key(booking: Booking, pref: BookingPreference | None) -> tuple[int, int, int, str]:
    pref = pref or BookingPreference(booking_id=booking.id)

    strictness = 0
    if pref.accessible_required or pref.section_preference_mandatory:
        strictness = 1

    strength = 0
    if pref.wants_aisle or pref.wants_central or pref.wants_front or pref.avoid_front:
        strength = 1

    # Lower tuple sorts first. We want strictness first, then larger groups, then stronger preferences.
    return (-strictness, -booking.quantity, -strength, booking.booking_reference)


def _block_has_accessible(
    block: list[AvailableSeat],
    seat_meta: dict[tuple[str, str, int], TheatreSeat],
) -> bool:
    for seat in block:
        meta = seat_meta.get((seat.section, seat.row, seat.seat_number))
        if meta and meta.is_accessible:
            return True
    return False


def _block_has_aisle(
    block: list[AvailableSeat],
    seat_meta: dict[tuple[str, str, int], TheatreSeat],
) -> bool:
    for seat in block:
        meta = seat_meta.get((seat.section, seat.row, seat.seat_number))
        if meta and meta.is_aisle:
            return True
    return False


def _centrality_score(
    block: list[AvailableSeat],
    seat_meta: dict[tuple[str, str, int], TheatreSeat],
) -> float:
    distances: list[float] = []
    for seat in block:
        meta = seat_meta.get((seat.section, seat.row, seat.seat_number))
        if not meta or meta.x_position is None:
            continue
        distances.append(abs(meta.x_position))

    if not distances:
        return 0.0

    avg = sum(distances) / len(distances)
    return max(0.0, 14.0 - avg)


def _front_score(
    block: list[AvailableSeat],
    seat_meta: dict[tuple[str, str, int], TheatreSeat],
) -> float:
    depths: list[float] = []
    for seat in block:
        meta = seat_meta.get((seat.section, seat.row, seat.seat_number))
        if meta and meta.y_position is not None:
            depths.append(meta.y_position)

    if not depths:
        return 0.0

    avg_depth = sum(depths) / len(depths)
    return max(0.0, 10.0 - avg_depth)


def _avoid_front_score(
    block: list[AvailableSeat],
    seat_meta: dict[tuple[str, str, int], TheatreSeat],
) -> float:
    depths: list[float] = []
    for seat in block:
        meta = seat_meta.get((seat.section, seat.row, seat.seat_number))
        if meta and meta.y_position is not None:
            depths.append(meta.y_position)

    if not depths:
        return 0.0

    avg_depth = sum(depths) / len(depths)
    return min(10.0, avg_depth)


def _fragmentation_penalty(block: list[AvailableSeat]) -> float:
    # Mild preference for edge-aligned blocks to reduce tiny isolated leftovers.
    first = block[0].seat_number
    if first <= 2:
        return 3.0
    return 0.0


def _status_from_score(score: float) -> MatchStatus:
    if score >= 132:
        return MatchStatus.PERFECT_MATCH
    if score >= 118:
        return MatchStatus.GOOD_MATCH
    if score >= 100:
        return MatchStatus.PARTIAL_MATCH
    return MatchStatus.NEEDS_MANUAL_REVIEW


def _match_notes(pref: BookingPreference, score: float) -> str:
    notes: list[str] = []
    if pref.wants_aisle:
        notes.append("Aisle preference considered")
    if pref.wants_central:
        notes.append("Centrality preference considered")
    if pref.wants_front:
        notes.append("Front seating preference considered")
    if pref.avoid_front:
        notes.append("Avoid-front preference considered")
    if pref.section_preference:
        notes.append(f"Section preference: {pref.section_preference}")

    if not notes:
        notes.append("Allocated by availability and group continuity")

    notes.append(f"Score={score:.1f}")
    return "; ".join(notes)


def _seat_key(seat: AvailableSeat) -> str:
    return f"{seat.performance_id}:{seat.section}:{seat.row}:{seat.seat_number}"
