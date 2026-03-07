from __future__ import annotations

import csv
from io import StringIO

from allocator.seat_plan_ingestion import SeatPlanRow
from allocator.seat_range_expander import TicketStockRow


REQUIRED_BASE_STOCK_COLUMNS = {
    "performance_id",
    "theatre",
    "show",
    "date",
    "time",
    "section",
}


class CsvImportError(ValueError):
    pass


def parse_ticket_stock_csv(content: str, delimiter: str = ",") -> list[TicketStockRow]:
    reader = csv.DictReader(StringIO(content), delimiter=delimiter)
    if not reader.fieldnames:
        raise CsvImportError("Ticket stock CSV has no header row")

    headers = {h.strip() for h in reader.fieldnames if h}
    missing = REQUIRED_BASE_STOCK_COLUMNS - headers
    if missing:
        raise CsvImportError(f"Ticket stock CSV missing required columns: {sorted(missing)}")

    rows: list[TicketStockRow] = []
    for raw in reader:
        row = {k.strip(): (v or "").strip() for k, v in raw.items() if k}

        seat_label = row.get("seat_label") or None
        seat_from = _to_int_or_none(row.get("seat_from"))
        seat_to = _to_int_or_none(row.get("seat_to"))

        rows.append(
            TicketStockRow(
                performance_id=int(row["performance_id"]),
                theatre=row["theatre"],
                show=row["show"],
                date=row["date"],
                time=row["time"],
                section=row["section"],
                row=row.get("row") or None,
                seat_from=seat_from,
                seat_to=seat_to,
                seat_label=seat_label,
            )
        )

    return rows


def parse_seat_plan_csv(content: str, delimiter: str = ",") -> list[SeatPlanRow]:
    reader = csv.DictReader(StringIO(content), delimiter=delimiter)
    if not reader.fieldnames:
        raise CsvImportError("Seat plan CSV has no header row")

    required = {"section", "row", "seat_number"}
    headers = {h.strip() for h in reader.fieldnames if h}
    missing = required - headers
    if missing:
        raise CsvImportError(f"Seat plan CSV missing required columns: {sorted(missing)}")

    rows: list[SeatPlanRow] = []
    for raw in reader:
        row = {k.strip(): (v or "").strip() for k, v in raw.items() if k}
        rows.append(
            SeatPlanRow(
                section=row["section"],
                row=row["row"],
                seat_number=int(row["seat_number"]),
                seat_label=row.get("seat_label") or None,
                is_accessible=(row.get("is_accessible", "").lower() in {"1", "true", "yes"}),
                x_position=_to_float_or_none(row.get("x_position")),
                y_position=_to_float_or_none(row.get("y_position")),
            )
        )

    return rows


def _to_int_or_none(value: str | None) -> int | None:
    if value is None:
        return None
    value = value.strip()
    return int(value) if value else None


def _to_float_or_none(value: str | None) -> float | None:
    if value is None:
        return None
    value = value.strip()
    return float(value) if value else None
