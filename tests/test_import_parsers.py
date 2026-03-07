import unittest

from allocator.import_parsers import parse_seat_plan_csv, parse_ticket_stock_csv


class ImportParsersTests(unittest.TestCase):
    def test_ticket_stock_parsing_range_and_label(self) -> None:
        csv_content = """performance_id,theatre,show,date,time,section,row,seat_from,seat_to,seat_label
1,London Palladium,Example Show,2026-03-20,19:30,Stalls,C,1,4,
1,London Palladium,Example Show,2026-03-20,19:30,Stalls,,,,D8
"""
        rows = parse_ticket_stock_csv(csv_content)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0].seat_from, 1)
        self.assertEqual(rows[1].seat_label, "D8")

    def test_seat_plan_parsing(self) -> None:
        csv_content = """section,row,seat_number,seat_label,is_accessible,x_position,y_position
Stalls,C,1,C1,true,10,20
"""
        rows = parse_seat_plan_csv(csv_content)
        self.assertEqual(len(rows), 1)
        self.assertTrue(rows[0].is_accessible)
        self.assertEqual(rows[0].x_position, 10.0)


if __name__ == "__main__":
    unittest.main()
