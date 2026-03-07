# Seat Data Foundations (MVP)

## Implemented now

- `allocator/models.py`: domain entities for theatre seats, performances, bookings, and allocation status.
- `allocator/seat_range_expander.py`: normalizes ticket stock into one seat per row.
- `allocator/seat_plan_ingestion.py`: structured seat-plan ingestion + metadata inference with manual-review warnings.

## Ticket stock accepted shapes

1. Range rows:
   - `performance_id, theatre, show, date, time, section, row, seat_from, seat_to`
2. Explicit seat labels:
   - `performance_id, theatre, show, date, time, section, seat_label`

Both normalize to individual seats with `row`, `seat_number`, and `seat_label`.

## Seat-plan ingestion strategy (v1)

1. Structured first:
   - CSV/TSV/JSON rows are ingested into `SeatPlanRow` records.
2. Auto inference:
   - Row-end seats => aisle.
   - Internal numbering gaps (e.g., 1,2,4,5) mark boundary seats near aisles.
   - Centrality proxy from seat number distance to row median.
   - Front/back proxy from row code (A closer than Z).
3. Confidence warnings:
   - Numbering gaps and single-seat rows are flagged for manual review.
4. Unstructured uploads:
   - OCR/PDF text is tokenized for probable seat labels.
   - Result is a draft that always requires manual confirmation.

## Why this order

Allocator quality depends on reliable per-seat metadata. These modules establish deterministic seat identity and inferred layout metadata before implementing booking import and assignment scoring.
