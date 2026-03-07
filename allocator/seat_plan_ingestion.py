from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass
from statistics import median
from typing import Iterable, Optional

from allocator.models import (
    SeatPlanInferenceWarning,
    SeatPlanIngestionResult,
    Theatre,
    TheatreSeat,
)
from allocator.seat_range_expander import parse_seat_label


@dataclass(slots=True)
class SeatPlanRow:
    section: str
    row: str
    seat_number: int
    seat_label: Optional[str] = None
    is_accessible: bool = False
    x_position: Optional[float] = None
    y_position: Optional[float] = None


@dataclass(slots=True)
class UnstructuredSeatPlanDraft:
    tokens: list[str]
    warnings: list[SeatPlanInferenceWarning]
    requires_manual_review: bool = True


class SeatPlanIngestor:
    """
    v1 seat-plan ingestion strategy:
    1) Prefer structured rows (CSV/TSV/JSON extracted from upload).
    2) Infer aisle/centrality/front-back/adjacency metadata deterministically.
    3) For unstructured text from OCR/PDF/image extraction, return a draft requiring review.
    """

    def ingest_structured_rows(
        self,
        theatre_id: int,
        theatre_name: str,
        city: str,
        rows: Iterable[SeatPlanRow],
    ) -> SeatPlanIngestionResult:
        theatre = Theatre(id=theatre_id, name=theatre_name, city=city)
        warnings: list[SeatPlanInferenceWarning] = []

        row_buckets: dict[tuple[str, str], list[SeatPlanRow]] = defaultdict(list)
        for entry in rows:
            if entry.seat_number <= 0:
                warnings.append(
                    SeatPlanInferenceWarning(
                        code="invalid-seat-number",
                        message=f"Ignored non-positive seat number: {entry.seat_number}",
                        row=entry.row,
                        section=entry.section,
                    )
                )
                continue
            row_buckets[(entry.section.strip(), entry.row.strip())].append(entry)

        inferred: list[TheatreSeat] = []
        for (section, row), bucket in row_buckets.items():
            bucket.sort(key=lambda x: x.seat_number)
            seat_numbers = [b.seat_number for b in bucket]
            seat_set = set(seat_numbers)
            gaps = self._find_gaps(seat_numbers)

            if gaps:
                warnings.append(
                    SeatPlanInferenceWarning(
                        code="row-number-gaps",
                        message=f"Detected numbering gaps in row {row}: {gaps}",
                        row=row,
                        section=section,
                    )
                )

            if len(bucket) == 1:
                warnings.append(
                    SeatPlanInferenceWarning(
                        code="single-seat-row",
                        message=(
                            f"Row {row} has one known seat. Aisle inference may be unreliable."
                        ),
                        row=row,
                        section=section,
                    )
                )

            median_seat = median(seat_numbers)
            for i, seat in enumerate(bucket):
                seat_label = seat.seat_label or f"{row}{seat.seat_number}"
                inferred.append(
                    TheatreSeat(
                        theatre_id=theatre_id,
                        section=section,
                        row=row,
                        seat_number=seat.seat_number,
                        seat_label=seat_label,
                        is_aisle=self._infer_aisle(i, bucket, seat_set),
                        is_accessible=seat.is_accessible,
                        x_position=seat.x_position,
                        y_position=seat.y_position,
                        adjacent_group_key=f"{section}:{row}:{seat.seat_number}",
                    )
                )

                # Store centrality/front-back proxy in positions when explicit values are missing.
                if seat.x_position is None:
                    inferred[-1].x_position = float(seat.seat_number - median_seat)
                if seat.y_position is None:
                    inferred[-1].y_position = self._row_depth_value(row)

        inferred.sort(key=lambda s: (s.section, s.row, s.seat_number))
        requires_review = any(w.code in {"single-seat-row", "row-number-gaps"} for w in warnings)
        return SeatPlanIngestionResult(
            theatre=theatre,
            theatre_seats=inferred,
            warnings=warnings,
            requires_manual_review=requires_review,
        )

    def ingest_unstructured_text(self, extracted_text: str) -> UnstructuredSeatPlanDraft:
        tokens = [t.strip(",.;:()[]{}") for t in extracted_text.split() if t.strip()]
        likely_seats: list[str] = []
        for token in tokens:
            try:
                parse_seat_label(token)
                likely_seats.append(token)
            except ValueError:
                continue

        warnings = [
            SeatPlanInferenceWarning(
                code="unstructured-plan",
                message=(
                    "Seat plan came from unstructured content (PDF/image OCR). "
                    "Auto-inference confidence is low; confirm aisle seats, row breaks, and inaccessible seats."
                ),
            )
        ]

        if not likely_seats:
            warnings.append(
                SeatPlanInferenceWarning(
                    code="no-seat-tokens-detected",
                    message="No seat-like tokens detected. Manual layout entry required.",
                )
            )

        return UnstructuredSeatPlanDraft(
            tokens=likely_seats,
            warnings=warnings,
            requires_manual_review=True,
        )

    @staticmethod
    def _find_gaps(seat_numbers: list[int]) -> list[int]:
        gaps: list[int] = []
        for i in range(1, len(seat_numbers)):
            prev = seat_numbers[i - 1]
            current = seat_numbers[i]
            if current - prev > 1:
                gaps.extend(range(prev + 1, current))
        return gaps

    @staticmethod
    def _infer_aisle(index: int, row_seats: list[SeatPlanRow], seat_set: set[int]) -> bool:
        current = row_seats[index].seat_number

        # Ends of row are aisle by default.
        if index == 0 or index == len(row_seats) - 1:
            return True

        # Internal numbering gaps often indicate a row split by an aisle.
        left_exists = (current - 1) in seat_set
        right_exists = (current + 1) in seat_set
        return not (left_exists and right_exists)

    @staticmethod
    def _row_depth_value(row: str) -> float:
        # Simple row-front proxy: earlier alphabet rows are closer to stage.
        value = 0.0
        for ch in row.upper():
            if "A" <= ch <= "Z":
                value += (ord(ch) - ord("A"))
            elif ch.isdigit():
                value += int(ch)
        return value
