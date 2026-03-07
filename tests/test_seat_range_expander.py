import unittest

from allocator.seat_range_expander import (
    TicketStockRow,
    expand_ticket_stock_rows,
    parse_seat_label,
)


class SeatRangeExpanderTests(unittest.TestCase):
    def test_expand_range_row(self) -> None:
        rows = [
            TicketStockRow(
                performance_id=1,
                theatre="London Palladium",
                show="Example Show",
                date="2026-03-20",
                time="19:30",
                section="Stalls",
                row="C",
                seat_from=1,
                seat_to=3,
            )
        ]

        expanded = expand_ticket_stock_rows(rows)
        self.assertEqual([s.seat_label for s in expanded], ["C1", "C2", "C3"])

    def test_expand_single_seat_label(self) -> None:
        rows = [
            TicketStockRow(
                performance_id=1,
                theatre="London Palladium",
                show="Example Show",
                date="2026-03-20",
                time="19:30",
                section="Stalls",
                seat_label="D12",
            )
        ]

        expanded = expand_ticket_stock_rows(rows)
        self.assertEqual(len(expanded), 1)
        self.assertEqual(expanded[0].row, "D")
        self.assertEqual(expanded[0].seat_number, 12)

    def test_parse_seat_label(self) -> None:
        row, number = parse_seat_label("AA14")
        self.assertEqual(row, "AA")
        self.assertEqual(number, 14)


if __name__ == "__main__":
    unittest.main()
