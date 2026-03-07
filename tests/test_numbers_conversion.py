import unittest

from backend.app import _numbers_table_score, _rows_to_csv


class NumbersConversionTests(unittest.TestCase):
    def test_numbers_table_score_prefers_email_rich_tables(self) -> None:
        email_table = [
            ["Name", "Email", "Seats"],
            ["Jane", "jane@example.com", "C1-C2"],
            ["Tom", "tom@example.com", "D4"],
        ]
        non_email_table = [
            ["A", "B", "C"],
            ["1", "2", "3"],
        ]

        self.assertGreater(_numbers_table_score(email_table), _numbers_table_score(non_email_table))

    def test_rows_to_csv_pads_to_rectangular_shape(self) -> None:
        rows = [["A", "B", "C"], ["1", "2"], ["x"]]
        csv_text = _rows_to_csv(rows)
        self.assertEqual(csv_text, "A,B,C\r\n1,2,\r\nx,,\r\n")


if __name__ == "__main__":
    unittest.main()
