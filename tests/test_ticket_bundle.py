import unittest

from allocator.ticket_bundle import (
    _extract_expected_seats_from_text,
    _extract_seat_tokens,
    _extract_seat_tokens_from_pdf_content_data,
    _normalize_seat_labels,
    build_booking_groups,
    output_pdf_filename,
    parse_allocation_csv,
    parse_seat_list,
)


class TicketBundleTests(unittest.TestCase):
    def test_parse_seat_list_with_ranges(self) -> None:
        parsed = parse_seat_list("C1-C3, C5; AA10")
        self.assertEqual(parsed, ["C1", "C2", "C3", "C5", "AA10"])

    def test_parse_seat_list_with_to_keyword(self) -> None:
        parsed = parse_seat_list("C1 to 5")
        self.assertEqual(parsed, ["C1", "C2", "C3", "C4", "C5"])

    def test_extract_seat_tokens_from_row_and_seat_labels(self) -> None:
        text = "LEVEL ROW SEAT\nROW J SEAT 11\n"
        parsed = _extract_seat_tokens(text)
        self.assertEqual(parsed, ["J11"])

    def test_extract_seat_tokens_from_reversed_split_boxes(self) -> None:
        text = "LEVEL ROW SEAT\n11J STALLS\n"
        parsed = _extract_seat_tokens(text)
        self.assertEqual(parsed, ["J11"])

    def test_extract_seat_tokens_from_split_lines(self) -> None:
        text = "LEVEL ROW SEAT\nJ\n11\nSTALLS\n"
        parsed = _extract_seat_tokens(text)
        self.assertEqual(parsed, ["J11"])

    def test_extract_seat_tokens_from_pdf_content_stream_literals(self) -> None:
        content = b"BT (J) Tj ET BT (11) Tj ET BT (STALLS) Tj ET"
        parsed = _extract_seat_tokens_from_pdf_content_data(content)
        self.assertEqual(parsed, ["J11"])

    def test_extract_expected_seats_from_text_handles_split_fields(self) -> None:
        text = "LEVEL ROW SEAT J 11 STALLS"
        expected = {"J11", "K9"}
        parsed = _extract_expected_seats_from_text(text, expected)
        self.assertEqual(parsed, ["J11"])

    def test_extract_expected_seats_from_packed_section_format(self) -> None:
        text = "Stalls-D-30SECTIONROWSEAT"
        expected = {"D30", "E31"}
        parsed = _extract_expected_seats_from_text(text, expected)
        self.assertEqual(parsed, ["D30"])

    def test_extract_seat_tokens_from_packed_section_format(self) -> None:
        text = "Order TGBGZP9D30Stalls Stalls-D-30SECTIONROWSEAT"
        parsed = _extract_seat_tokens(text)
        self.assertEqual(parsed, ["D30"])

    def test_extract_seat_tokens_from_compact_section_format(self) -> None:
        text = "Mon 16 Mar 2026 19:00 Dress Circle B1Paddington The Musical"
        parsed = _extract_seat_tokens(text)
        self.assertEqual(parsed, ["B1"])

    def test_extract_seat_tokens_ignores_postcode_tokens(self) -> None:
        text = "Stalls C4C4Stalls Criterion Theatre, London SW1V 9LB"
        parsed = _extract_seat_tokens(text)
        self.assertEqual(parsed, ["C4"])

    def test_normalize_seat_labels(self) -> None:
        parsed = _normalize_seat_labels({"j11", "12 K", "bad"})
        self.assertEqual(parsed, {"J11", "K12"})

    def test_parse_allocation_alias_columns(self) -> None:
        csv_content = """Order Number,Booker Name,E-mail,Allocated Seats\nB100,Jane,jane@example.com,C1 to C3\n"""
        rows = parse_allocation_csv(csv_content)
        self.assertEqual(rows[0]["booking_reference"], "B100")
        self.assertEqual(rows[0]["email"], "jane@example.com")
        self.assertEqual(rows[0]["seats_raw"], "C1 to C3")

    def test_parse_allocation_split_row_and_seat_columns(self) -> None:
        csv_content = (
            "Booking Ref,Customer Name,Email,Row,Seat Number\n"
            "B100,Jane,jane@example.com,C,12\n"
        )
        rows = parse_allocation_csv(csv_content)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["email"], "jane@example.com")
        self.assertEqual(rows[0]["seats_raw"], "C12")

    def test_parse_allocation_without_booking_ref(self) -> None:
        csv_content = """Customer Name,Email,Assigned Seats\nJane,jane@example.com,C1-C2\n"""
        rows = parse_allocation_csv(csv_content)
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["booking_reference"], "")
        self.assertEqual(rows[0]["email"], "jane@example.com")

    def test_parse_allocation_fills_down_email_and_booking_ref(self) -> None:
        csv_content = (
            "Booking Ref,Customer Name,Email,Assigned Seats\n"
            "B100,Rachel,rachel@example.com,C4\n"
            ",,,C5\n"
        )
        rows = parse_allocation_csv(csv_content)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["booking_reference"], "B100")
        self.assertEqual(rows[1]["booking_reference"], "B100")
        self.assertEqual(rows[1]["email"], "rachel@example.com")
        self.assertEqual(rows[1]["seats_raw"], "C5")

    def test_parse_allocation_keeps_email_rows_with_blank_seats(self) -> None:
        csv_content = (
            "Customer Name,Email,Assigned Seats\n"
            "Jan,jan@example.com,C1-C2\n"
            "Adam,adam@example.com,\n"
        )
        rows = parse_allocation_csv(csv_content)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[1]["email"], "adam@example.com")
        self.assertEqual(rows[1]["seats_raw"], "")

    def test_parse_allocation_without_headers_by_inference(self) -> None:
        csv_content = """intro,blah,blah\nJane,jane@example.com,C1 to C3\nTom,tom@example.com,D4-D5\n"""
        rows = parse_allocation_csv(csv_content)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["customer_name"], "Jane")
        self.assertEqual(rows[0]["email"], "jane@example.com")
        self.assertEqual(rows[0]["seats_raw"], "C1 to C3")

    def test_parse_allocation_without_headers_infers_split_row_and_seat_columns(self) -> None:
        csv_content = (
            "booking,email,row,seat\n"
            "B100,jane@example.com,C,12\n"
            "B101,tom@example.com,D,5\n"
        )
        rows = parse_allocation_csv(csv_content)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["seats_raw"], "C12")
        self.assertEqual(rows[1]["seats_raw"], "D5")

    def test_parse_allocation_without_headers_keeps_email_rows_with_blank_seats(self) -> None:
        csv_content = (
            "intro,blah,blah\n"
            "Jane,jane@example.com,C1 to C3\n"
            "Adam,adam@example.com,\n"
        )
        rows = parse_allocation_csv(csv_content)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[1]["email"], "adam@example.com")
        self.assertEqual(rows[1]["seats_raw"], "")

    def test_build_groups_maps_pages(self) -> None:
        rows = [
            {
                "booking_reference": "B001",
                "customer_name": "Jane",
                "email": "jane@example.com",
                "seats_raw": "C1-C2",
            }
        ]
        seat_to_page = {"C1": 0, "C2": 1}
        groups = build_booking_groups(rows, seat_to_page)

        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0].page_indexes, [0, 1])
        self.assertEqual(groups[0].missing_seats, [])
        self.assertEqual(output_pdf_filename(groups[0]), "Jane_tickets.pdf")

    def test_groups_by_email_when_booking_ref_missing(self) -> None:
        rows = [
            {
                "booking_reference": "",
                "customer_name": "Jane",
                "email": "jane@example.com",
                "seats_raw": "C1",
            },
            {
                "booking_reference": "",
                "customer_name": "Jane",
                "email": "jane@example.com",
                "seats_raw": "C2",
            },
        ]
        seat_to_page = {"C1": 0, "C2": 1}
        groups = build_booking_groups(rows, seat_to_page)
        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0].seat_labels, ["C1", "C2"])

    def test_groups_do_not_merge_different_emails_with_same_booking_ref(self) -> None:
        rows = [
            {
                "booking_reference": "AISLE",
                "customer_name": "Jan",
                "email": "jan@example.com",
                "seats_raw": "C1",
            },
            {
                "booking_reference": "AISLE",
                "customer_name": "Adam",
                "email": "adam@example.com",
                "seats_raw": "C2",
            },
        ]
        seat_to_page = {"C1": 0, "C2": 1}
        groups = build_booking_groups(rows, seat_to_page)
        self.assertEqual(len(groups), 2)
        self.assertEqual([g.email for g in groups], ["jan@example.com", "adam@example.com"])

    def test_groups_include_continuation_rows_with_same_email(self) -> None:
        rows = [
            {
                "booking_reference": "B100",
                "customer_name": "Rachel",
                "email": "rachel@example.com",
                "seats_raw": "C4",
            },
            {
                "booking_reference": "B100",
                "customer_name": "Rachel",
                "email": "rachel@example.com",
                "seats_raw": "C5",
            },
        ]
        seat_to_page = {"C4": 0, "C5": 1}
        groups = build_booking_groups(rows, seat_to_page)
        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0].seat_labels, ["C4", "C5"])
        self.assertEqual(groups[0].page_indexes, [0, 1])

    def test_groups_ignore_notes_like_booking_refs(self) -> None:
        rows = [
            {
                "booking_reference": "extra",
                "customer_name": "Sue Dillon",
                "email": "sue@example.com",
                "seats_raw": "C1-C2",
            },
            {
                "booking_reference": "central",
                "customer_name": "Sue Dillon",
                "email": "sue@example.com",
                "seats_raw": "C3-C4",
            },
        ]
        seat_to_page = {"C1": 0, "C2": 1, "C3": 2, "C4": 3}
        groups = build_booking_groups(rows, seat_to_page)
        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0].seat_labels, ["C1", "C2", "C3", "C4"])

    def test_output_filename_falls_back_to_email_when_name_unknown(self) -> None:
        rows = [
            {
                "booking_reference": "B100",
                "customer_name": "PS unknown",
                "email": "sue@example.com",
                "seats_raw": "C1",
            }
        ]
        seat_to_page = {"C1": 0}
        groups = build_booking_groups(rows, seat_to_page)
        self.assertEqual(output_pdf_filename(groups[0]), "sue_example.com.pdf")


if __name__ == "__main__":
    unittest.main()
