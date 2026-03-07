from __future__ import annotations

from dataclasses import dataclass
from typing import Iterable, Optional


@dataclass(slots=True)
class TicketStockRow:
    performance_id: int
    theatre: str
    show: str
    date: str
    time: str
    section: str
    row: Optional[str] = None
    seat_from: Optional[int] = None
    seat_to: Optional[int] = None
    seat_label: Optional[str] = None


@dataclass(slots=True)
class ExpandedTicketStockSeat:
    performance_id: int
    theatre: str
    show: str
    date: str
    time: str
    section: str
    row: str
    seat_number: int
    seat_label: str


def parse_seat_label(seat_label: str) -> tuple[str, int]:
    label = seat_label.strip()
    if not label:
        raise ValueError("seat_label cannot be empty")

    split_idx = 0
    for i, ch in enumerate(label):
        if ch.isdigit():
            split_idx = i
            break
    else:
        raise ValueError(f"seat_label '{seat_label}' has no numeric seat number")

    row = label[:split_idx].strip()
    seat_str = label[split_idx:].strip()
    if not row:
        raise ValueError(f"seat_label '{seat_label}' has no row prefix")
    if not seat_str.isdigit():
        raise ValueError(f"seat_label '{seat_label}' has non-numeric seat number")

    return row, int(seat_str)


def expand_ticket_stock_rows(rows: Iterable[TicketStockRow]) -> list[ExpandedTicketStockSeat]:
    expanded: list[ExpandedTicketStockSeat] = []

    for row in rows:
        if row.seat_label:
            parsed_row, seat_number = parse_seat_label(row.seat_label)
            expanded.append(
                ExpandedTicketStockSeat(
                    performance_id=row.performance_id,
                    theatre=row.theatre,
                    show=row.show,
                    date=row.date,
                    time=row.time,
                    section=row.section,
                    row=parsed_row,
                    seat_number=seat_number,
                    seat_label=f"{parsed_row}{seat_number}",
                )
            )
            continue

        if row.row is None or row.seat_from is None or row.seat_to is None:
            raise ValueError(
                "row + seat_from + seat_to are required when seat_label is not provided"
            )

        start = min(row.seat_from, row.seat_to)
        end = max(row.seat_from, row.seat_to)
        for seat_number in range(start, end + 1):
            expanded.append(
                ExpandedTicketStockSeat(
                    performance_id=row.performance_id,
                    theatre=row.theatre,
                    show=row.show,
                    date=row.date,
                    time=row.time,
                    section=row.section,
                    row=row.row,
                    seat_number=seat_number,
                    seat_label=f"{row.row}{seat_number}",
                )
            )

    expanded.sort(key=lambda s: (s.section, s.row, s.seat_number))
    return expanded
