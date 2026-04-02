from io import BytesIO
import json
import zipfile
import unittest
from unittest.mock import patch

from allocator.ticket_bundle import (
    BookingTicketGroup,
    ParsedTicketPage,
    ParsedTicketPageResult,
    _extract_printed_barcode_value_from_page,
    _extract_show_name_candidates,
    _extract_venue_candidates,
    _extract_performance_date_candidates,
    _extract_performance_time_candidates,
    _extract_expected_seats_from_text,
    _extract_seat_tokens,
    _extract_seat_tokens_from_pdf_content_data,
    _normalize_seat_labels,
    TicketBundleError,
    build_bundle_zip,
    build_output_filenames,
    build_booking_groups,
    build_pkpass_for_ticket,
    extract_ticket_performance_metadata,
    parse_ticket_pdf_page_results,
    parse_ticket_pdf_pages,
    split_groups_for_output,
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

    def test_parse_seat_list_with_standing_labels(self) -> None:
        parsed = parse_seat_list("STANDING 1; Dance Floor GA 2; GA3")
        self.assertEqual(parsed, ["STANDING1", "STANDING2", "STANDING3"])

    def test_parse_seat_list_with_standing_range(self) -> None:
        parsed = parse_seat_list("Standing 1 to 3")
        self.assertEqual(parsed, ["STANDING1", "STANDING2", "STANDING3"])

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

    def test_extract_seat_tokens_from_reversed_split_lines_with_section_before(self) -> None:
        text = "ROYAL CIRCLE\n4\nD\n"
        parsed = _extract_seat_tokens(text)
        self.assertEqual(parsed, ["D4"])

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

    def test_extract_seat_tokens_from_old_vic_stage_stalls_format(self) -> None:
        text = (
            "One Flew Over the Cuckoo's Nest\n"
            "Thursday 09 April 2026 7:30 PM\n"
            "4446166\n"
            "6169165\n"
            "Stage Stalls\n"
            "K37\n"
            "55.00\n"
            "Group20+\n"
        )
        parsed = _extract_seat_tokens(text)
        self.assertEqual(parsed, ["K37"])

    def test_extract_seat_tokens_from_old_vic_stage_stalls_packed_format(self) -> None:
        text = (
            "One Flew Over the Cuckoo's Nest"
            "Thursday 09 April 20267:30 PM"
            "4446166 6169165"
            "Stage StallsK37 55.00"
            "Group20+"
        )
        parsed = _extract_seat_tokens(text)
        self.assertEqual(parsed, ["K37"])

    def test_extract_seat_tokens_from_old_vic_stage_stalls_packed_after_ticket_number(self) -> None:
        text = (
            "Bird & Bird is the proud Production Sponsor"
            "of One Flew Over the Cuckoo’s Nest4446166 6169165Stage StallsK37 55.00Group20+"
        )
        parsed = _extract_seat_tokens(text)
        self.assertEqual(parsed, ["K37"])

    def test_extract_seat_tokens_from_old_vic_stalls_format(self) -> None:
        text = (
            "One Flew Over the Cuckoo's Nest\n"
            "Thursday 09 April 2026 7:30 PM\n"
            "4446194\n"
            "6169165\n"
            "Stalls\n"
            "J12\n"
            "55.00\n"
            "Group20+\n"
        )
        parsed = _extract_seat_tokens(text)
        self.assertEqual(parsed, ["J12"])

    def test_extract_seat_tokens_from_old_vic_stalls_packed_format(self) -> None:
        text = (
            "One Flew Over the Cuckoo's Nest"
            "Thursday 09 April 20267:30 PM"
            "4446194 6169165"
            "StallsJ12 55.00"
            "Group20+"
        )
        parsed = _extract_seat_tokens(text)
        self.assertEqual(parsed, ["J12"])

    def test_extract_seat_tokens_from_old_vic_stalls_packed_after_ticket_number(self) -> None:
        text = (
            "Bird & Bird is the proud Production Sponsor"
            "of One Flew Over the Cuckoo’s Nest4446194 6169165StallsJ12 55.00Group20+"
        )
        parsed = _extract_seat_tokens(text)
        self.assertEqual(parsed, ["J12"])

    def test_extract_seat_tokens_from_kx_platform_format(self) -> None:
        text = (
            "Order TF40DJ6 Platform1 -F-8 Starlight Express Friday, 27 March, 2026 7:00 pm "
            "Troubadour Wembley Park Theatre SECTION Platform1 ROW F SEAT 8"
        )
        parsed = _extract_seat_tokens(text)
        self.assertEqual(parsed, ["F8"])

    def test_extract_performance_date_candidates_prefers_month_day(self) -> None:
        parsed = _extract_performance_date_candidates("Paddington Sat 15 March 2026 7:30pm")
        self.assertEqual(parsed, ["Mar 15"])

    def test_extract_performance_time_candidates_formats_24_hour_time(self) -> None:
        parsed = _extract_performance_time_candidates("Paddington Sat 15 March 2026 19:30")
        self.assertEqual(parsed, ["7.30"])

    def test_extract_performance_candidates_from_old_vic_format(self) -> None:
        text = "One Flew Over the Cuckoo's Nest Thursday 09 April 2026 7:30 PM Stage Stalls K37"
        self.assertEqual(_extract_performance_date_candidates(text), ["Apr 9"])
        self.assertEqual(_extract_performance_time_candidates(text), ["7.30"])

    def test_extract_show_name_candidates_from_old_vic_packed_text(self) -> None:
        text = (
            "Event Donate to us Order number Ticket number "
            "One Flew Over the Cuckoo's NestThursday 09 April 20267:30 PM "
            "Bird & Bird is the proud Production Sponsor"
        )
        self.assertEqual(_extract_show_name_candidates(text)[0], "One Flew Over the Cuckoo's Nest")

    def test_extract_show_name_candidates_known_ticket_formats(self) -> None:
        self.assertEqual(
            _extract_show_name_candidates("ABBA Voyage\nFRI 10 APR 2026\nStart Time 7:45 PM\nABBA Arena\nGroups Ticket")[0],
            "ABBA Voyage",
        )
        self.assertEqual(
            _extract_show_name_candidates("Mon 16 Mar 2026 19:00Savoy Theatre, LondonNabarroOrder 37638599\nDress Circle B1Paddington The Musical")[0],
            "Paddington The Musical",
        )
        self.assertEqual(
            _extract_show_name_candidates("Thursday, 19 March, 2026 2:30 pmHamiltonHamilton\nOrder ID: TFFLMR4")[0],
            "Hamilton",
        )

    def test_extract_venue_candidates_prefers_known_old_vic_hint_over_disclaimer(self) -> None:
        text = (
            "oldvictheatre.com Registered Charity No. 1072590 "
            "Tickets cannot be sold on for commercial gain by any outlet other than The Old Vic "
            "or one of its authorised agents."
        )
        self.assertEqual(_extract_venue_candidates(text), ["The Old Vic"])

    def test_extract_venue_candidates_known_theatre_hints(self) -> None:
        self.assertEqual(_extract_venue_candidates("ABBA Arena, Pudding Mill Lane"), ["ABBA Arena"])
        self.assertEqual(_extract_venue_candidates("Criterion Theatre, 218-223 Piccadilly"), ["Criterion Theatre"])
        self.assertEqual(_extract_venue_candidates("Rosebery Avenue, London EC1R 4TN"), ["Sadler's Wells Theatre"])
        self.assertEqual(_extract_venue_candidates("boxoffice@sohoplace.org"), ["@sohoplace"])

    def test_extract_performance_time_candidates_ignores_prices_and_contact_hours(self) -> None:
        titanique_text = "Thu 19 March 2026 @ 19:30 Ticket commission £1.25 restoration levy £1.50"
        self.assertEqual(_extract_performance_time_candidates(titanique_text), ["7.30"])

        sinatra_text = "Wednesday, 19 August, 2026 7:30 pm contact 020 7206 1174 Mon-Fri 9am -5.30pm"
        self.assertEqual(_extract_performance_time_candidates(sinatra_text), ["7.30"])

    def test_extract_ticket_performance_metadata_from_pdf_text(self) -> None:
        class FakePage:
            def extract_text(self) -> str:
                return "Paddington The Musical\nSat 15 March 2026\n7:30pm"

        class FakeReader:
            def __init__(self, _stream) -> None:
                self.pages = [FakePage()]

        with patch("allocator.ticket_bundle._load_pdf_backend", return_value=(FakeReader, object)):
            parsed = extract_ticket_performance_metadata(b"%PDF-pretend")

        self.assertEqual(
            parsed,
            {"performance_date": "Mar 15", "performance_time": "7.30", "confidence": True},
        )

    def test_extract_ticket_performance_metadata_marks_ambiguous_time_low_confidence(self) -> None:
        class FakePage:
            def extract_text(self) -> str:
                return "Sat 15 March 2026\nDoors 6:30pm\nPerformance 7:30pm"

        class FakeReader:
            def __init__(self, _stream) -> None:
                self.pages = [FakePage()]

        with patch("allocator.ticket_bundle._load_pdf_backend", return_value=(FakeReader, object)):
            parsed = extract_ticket_performance_metadata(b"%PDF-pretend")

        self.assertEqual(
            parsed,
            {"performance_date": "Mar 15", "performance_time": None, "confidence": False},
        )

    def test_parse_ticket_pdf_pages_extracts_structured_page_fields(self) -> None:
        page_text = "Paddington The Musical\nSat 15 March 2026\n7:30pm\nSavoy Theatre\nROW D\nSEAT 30\n"

        class FakePage:
            def get_text(self, mode: str):
                if mode == "text":
                    return page_text
                if mode == "words":
                    return []
                raise AssertionError(f"Unexpected mode: {mode}")

        class FakeDocument:
            page_count = 1

            def load_page(self, index: int):
                self.last_index = index
                return FakePage()

        class FakeFitz:
            @staticmethod
            def open(*_args, **_kwargs):
                return FakeDocument()

        with patch("allocator.ticket_bundle._load_fitz_backend", return_value=FakeFitz):
            with patch("allocator.ticket_bundle._decode_qr_payload_for_page", return_value="decoded-qr-payload"):
                parsed = parse_ticket_pdf_pages(b"%PDF-pretend", expected_seats={"D30"})

        self.assertEqual(len(parsed), 1)
        self.assertEqual(
            parsed[0],
            ParsedTicketPage(
                page_index=0,
                show_name="Paddington The Musical",
                performance_date="Mar 15",
                performance_time="7.30",
                venue_name="Savoy Theatre",
                row="D",
                seat="30",
                seat_label="D30",
                qr_payload="decoded-qr-payload",
            ),
        )

    def test_parse_ticket_pdf_page_results_keeps_pdf_matching_when_wallet_decode_fails(self) -> None:
        page_text = "Paddington The Musical\nSat 15 March 2026\n7:30pm\nSavoy Theatre\nROW D\nSEAT 30\n"

        class FakePage:
            def get_text(self, mode: str):
                if mode == "text":
                    return page_text
                if mode == "words":
                    return []
                raise AssertionError(f"Unexpected mode: {mode}")

        class FakeDocument:
            page_count = 1

            def load_page(self, index: int):
                self.last_index = index
                return FakePage()

        class FakeFitz:
            @staticmethod
            def open(*_args, **_kwargs):
                return FakeDocument()

        with patch("allocator.ticket_bundle._load_fitz_backend", return_value=FakeFitz):
            with patch(
                "allocator.ticket_bundle._decode_qr_payload_for_page",
                side_effect=TicketBundleError("Could not decode a QR code from one or more ticket pages."),
            ):
                parsed = parse_ticket_pdf_page_results(b"%PDF-pretend", expected_seats={"D30"})

        self.assertEqual(
            parsed,
            [
                ParsedTicketPageResult(
                    page_index=0,
                    show_name="Paddington The Musical",
                    performance_date="Mar 15",
                    performance_time="7.30",
                    venue_name="Savoy Theatre",
                    row="D",
                    seat="30",
                    seat_label="D30",
                    qr_payload=None,
                    wallet_error="Could not decode a QR code from one or more ticket pages.",
                )
            ],
        )

    def test_parse_ticket_pdf_page_results_extracts_abba_voyage_seat_from_positioned_words(self) -> None:
        page_text = (
            "ABBA Voyage\n"
            "FRI 10 APR 2026\n"
            "Start Time 7:45 PM\n"
            "ABBA Arena\n"
            "Groups Ticket\n"
            "via Gate A - Arena Right\n"
            "Block K\n"
            "SECTION\n"
            "U16s accompanied by an adult\n"
        )

        class FakePage:
            def get_text(self, mode: str):
                if mode == "text":
                    return page_text
                if mode == "words":
                    return [
                        (106.0, 50.0, 143.0, 67.0, "ABBA", 0, 0, 0),
                        (147.0, 50.0, 196.0, 67.0, "Voyage", 0, 0, 1),
                        (118.0, 191.0, 165.0, 212.0, "Block", 6, 0, 0),
                        (170.0, 191.0, 183.0, 212.0, "K", 6, 0, 1),
                        (137.0, 211.0, 164.0, 218.0, "SECTION", 7, 0, 0),
                        (96.0, 259.0, 205.0, 268.0, "U16s", 8, 0, 0),
                        (115.0, 223.0, 126.0, 244.0, "A", 13, 0, 0),
                        (172.0, 223.0, 182.0, 244.0, "1", 13, 0, 1),
                    ]
                raise AssertionError(f"Unexpected mode: {mode}")

        class FakeDocument:
            page_count = 1

            def load_page(self, _index: int):
                return FakePage()

        class FakeFitz:
            @staticmethod
            def open(*_args, **_kwargs):
                return FakeDocument()

        with patch("allocator.ticket_bundle._load_fitz_backend", return_value=FakeFitz):
            with patch(
                "allocator.ticket_bundle._decode_qr_payload_for_page",
                side_effect=TicketBundleError("Could not decode a QR code from one or more ticket pages."),
            ):
                parsed = parse_ticket_pdf_page_results(b"%PDF-pretend")

        self.assertEqual(len(parsed), 1)
        self.assertEqual(parsed[0].seat_label, "A1")
        self.assertEqual(parsed[0].show_name, "ABBA Voyage")
        self.assertEqual(parsed[0].venue_name, "ABBA Arena")
        self.assertEqual(parsed[0].performance_date, "Apr 10")
        self.assertEqual(parsed[0].performance_time, "7.45")

    def test_parse_ticket_pdf_page_results_assigns_abba_dance_floor_pages_without_expected_seats(self) -> None:
        seated_text = (
            "ABBA Voyage\nFRI 10 APR 2026\nStart Time 7:45 PM\nABBA Arena\nGroups Ticket\n"
            "via Gate A - Arena Right\nBlock K\nSECTION\nU16s accompanied by an adult\n"
        )
        dance_floor_text = (
            "ABBA Voyage\nFRI 10 APR 2026\nStart Time 7:45 PM\nABBA Arena\nGroups Ticket\n"
            "via Gate B\nDance Floor\nSECTION\nU16s accompanied by an adult\nGA\n"
        )

        class FakePage:
            def __init__(self, text: str, words: list[tuple]):
                self.text = text
                self.words = words

            def get_text(self, mode: str):
                if mode == "text":
                    return self.text
                if mode == "words":
                    return self.words
                raise AssertionError(f"Unexpected mode: {mode}")

        pages = [
            FakePage(
                seated_text,
                [
                    (106.0, 50.0, 143.0, 67.0, "ABBA", 0, 0, 0),
                    (147.0, 50.0, 196.0, 67.0, "Voyage", 0, 0, 1),
                    (118.0, 191.0, 165.0, 212.0, "Block", 6, 0, 0),
                    (170.0, 191.0, 183.0, 212.0, "K", 6, 0, 1),
                    (137.0, 211.0, 164.0, 218.0, "SECTION", 7, 0, 0),
                    (115.0, 223.0, 126.0, 244.0, "A", 13, 0, 0),
                    (172.0, 223.0, 182.0, 244.0, "1", 13, 0, 1),
                ],
            ),
            FakePage(
                dance_floor_text,
                [
                    (106.0, 50.0, 143.0, 67.0, "ABBA", 0, 0, 0),
                    (147.0, 50.0, 196.0, 67.0, "Voyage", 0, 0, 1),
                    (117.0, 153.0, 185.0, 169.0, "via", 5, 0, 0),
                    (152.0, 153.0, 185.0, 169.0, "Gate", 5, 0, 1),
                    (101.0, 191.0, 200.0, 212.0, "Dance", 6, 0, 0),
                    (152.0, 191.0, 200.0, 212.0, "Floor", 6, 0, 1),
                    (165.0, 223.0, 189.0, 244.0, "GA", 13, 0, 0),
                ],
            ),
        ]

        class FakeDocument:
            page_count = 2

            def load_page(self, index: int):
                return pages[index]

        class FakeFitz:
            @staticmethod
            def open(*_args, **_kwargs):
                return FakeDocument()

        with patch("allocator.ticket_bundle._load_fitz_backend", return_value=FakeFitz):
            with patch(
                "allocator.ticket_bundle._decode_qr_payload_for_page",
                side_effect=TicketBundleError("Could not decode a QR code from one or more ticket pages."),
            ):
                parsed = parse_ticket_pdf_page_results(b"%PDF-pretend")

        self.assertEqual(len(parsed), 2)
        self.assertEqual(parsed[0].seat_label, "A1")
        self.assertEqual(parsed[1].seat_label, "STANDING1")
        self.assertEqual(parsed[1].row, "STANDING")
        self.assertEqual(parsed[1].seat, "1")

    def test_parse_ticket_pdf_page_results_assigns_abba_dance_floor_pages_to_expected_standing_labels(self) -> None:
        dance_floor_text = (
            "ABBA Voyage\nFRI 10 APR 2026\nStart Time 7:45 PM\nABBA Arena\nGroups Ticket\n"
            "via Gate B\nDance Floor\nSECTION\nU16s accompanied by an adult\nGA\n"
        )

        class FakePage:
            def get_text(self, mode: str):
                if mode == "text":
                    return dance_floor_text
                if mode == "words":
                    return [
                        (106.0, 50.0, 143.0, 67.0, "ABBA", 0, 0, 0),
                        (147.0, 50.0, 196.0, 67.0, "Voyage", 0, 0, 1),
                        (101.0, 191.0, 200.0, 212.0, "Dance", 6, 0, 0),
                        (152.0, 191.0, 200.0, 212.0, "Floor", 6, 0, 1),
                        (165.0, 223.0, 189.0, 244.0, "GA", 13, 0, 0),
                    ]
                raise AssertionError(f"Unexpected mode: {mode}")

        class FakeDocument:
            page_count = 2

            def load_page(self, _index: int):
                return FakePage()

        class FakeFitz:
            @staticmethod
            def open(*_args, **_kwargs):
                return FakeDocument()

        with patch("allocator.ticket_bundle._load_fitz_backend", return_value=FakeFitz):
            with patch(
                "allocator.ticket_bundle._decode_qr_payload_for_page",
                side_effect=TicketBundleError("Could not decode a QR code from one or more ticket pages."),
            ):
                parsed = parse_ticket_pdf_page_results(b"%PDF-pretend", expected_seats={"standing 8", "standing 7"})

        self.assertEqual([page.seat_label for page in parsed], ["STANDING7", "STANDING8"])

    def test_parse_ticket_pdf_pages_extracts_old_vic_show_and_venue(self) -> None:
        page_text = (
            "Event\n"
            "Donate to usOrder numberTicket numberTerms and conditions Venue and Travel info Click hereAllocation\n"
            "£No need to print your ticket,  we will scan it from your phone\n"
            "Join as a MemberVisit our Education Hub\n"
            "oldvictheatre.comRegistered Charity No. 1072590Tickets cannot be sold on for commercial gain by any outlet other than The Old Vic or one of its authorised agents.\n"
            "WATERLOOSTATIONBACKSTAGEJoin us Backstage for a pre-theatre meal or post-show drink"
            "One Flew Over the Cuckoo's NestThursday 09 April 20267:30 PM"
            "Royal Bank of Canada Principal Partner:Bringing you more"
            "Bird & Bird is the proud Production Sponsorof One Flew Over the Cuckoo’s Nest"
            "4446191 6169165StallsJ7 55.00Group20+\n"
        )

        class FakePage:
            def get_text(self, mode: str):
                if mode == "text":
                    return page_text
                if mode == "words":
                    return []
                raise AssertionError(f"Unexpected mode: {mode}")

        class FakeDocument:
            page_count = 1

            def load_page(self, _index: int):
                return FakePage()

        class FakeFitz:
            @staticmethod
            def open(*_args, **_kwargs):
                return FakeDocument()

        with patch("allocator.ticket_bundle._load_fitz_backend", return_value=FakeFitz):
            with patch("allocator.ticket_bundle._decode_qr_payload_for_page", return_value="decoded-qr-payload"):
                parsed = parse_ticket_pdf_pages(b"%PDF-pretend", expected_seats={"J7"})

        self.assertEqual(parsed[0].show_name, "One Flew Over the Cuckoo's Nest")
        self.assertEqual(parsed[0].venue_name, "The Old Vic")

    def test_parse_ticket_pdf_pages_extracts_devil_wears_prada_metadata(self) -> None:
        page_text = (
            "Wednesday, 15 April, 2026 2:30 pmThe Devil Wears Prada\n"
            "Order TF51ZSTC\n"
            "21STALLS\n"
            "STALLS-C-21SECTION\n"
            "ROW\n"
            "SEAT529057474466180458\n"
            "Dominion Theatre\n"
            "Accessible by 9 shallow steps or a platform lift\n"
        )

        class FakePage:
            def get_text(self, mode: str):
                if mode == "text":
                    return page_text
                if mode == "words":
                    return []
                raise AssertionError(f"Unexpected mode: {mode}")

        class FakeDocument:
            page_count = 1

            def load_page(self, _index: int):
                return FakePage()

        class FakeFitz:
            @staticmethod
            def open(*_args, **_kwargs):
                return FakeDocument()

        with patch("allocator.ticket_bundle._load_fitz_backend", return_value=FakeFitz):
            with patch("allocator.ticket_bundle._decode_qr_payload_for_page", return_value="decoded-qr-payload"):
                parsed = parse_ticket_pdf_pages(b"%PDF-pretend")

        self.assertEqual(parsed[0].show_name, "The Devil Wears Prada")
        self.assertEqual(parsed[0].venue_name, "Dominion Theatre")
        self.assertEqual(parsed[0].performance_time, "2.30")
        self.assertEqual(parsed[0].seat_label, "C21")

    def test_parse_ticket_pdf_pages_requires_decoded_qr_payload(self) -> None:
        page_text = "Paddington The Musical\nSat 15 March 2026\n7:30pm\nSavoy Theatre\nROW D\nSEAT 30\n"

        class FakePage:
            def get_text(self, mode: str):
                if mode == "text":
                    return page_text
                if mode == "words":
                    return []
                raise AssertionError(f"Unexpected mode: {mode}")

        class FakeDocument:
            page_count = 1

            def load_page(self, _index: int):
                return FakePage()

        class FakeFitz:
            @staticmethod
            def open(*_args, **_kwargs):
                return FakeDocument()

        with patch("allocator.ticket_bundle._load_fitz_backend", return_value=FakeFitz):
            with patch(
                "allocator.ticket_bundle._decode_qr_payload_for_page",
                side_effect=TicketBundleError("Could not decode a QR code from one or more ticket pages."),
            ):
                with self.assertRaises(TicketBundleError):
                    parse_ticket_pdf_pages(b"%PDF-pretend", expected_seats={"D30"})

    def test_build_pkpass_for_ticket_creates_unsigned_wallet_bundle(self) -> None:
        ticket = ParsedTicketPage(
            page_index=0,
            show_name="Paddington The Musical",
            performance_date="Mar 15",
            performance_time="7.30",
            venue_name="Savoy Theatre",
            row="D",
            seat="30",
            seat_label="D30",
            qr_payload="decoded-qr-payload",
        )

        blob = build_pkpass_for_ticket(ticket)

        with zipfile.ZipFile(BytesIO(blob)) as archive:
            names = sorted(archive.namelist())
            self.assertEqual(names, ["icon.png", "icon@2x.png", "manifest.json", "pass.json"])
            pass_json = json.loads(archive.read("pass.json"))
            manifest_json = json.loads(archive.read("manifest.json"))

        self.assertEqual(pass_json["description"], "Theatre ticket")
        self.assertEqual(pass_json["eventTicket"]["headerFields"][0]["value"], "Savoy Theatre")
        self.assertEqual(pass_json["eventTicket"]["primaryFields"][0]["value"], "Paddington The Musical")
        self.assertEqual(pass_json["eventTicket"]["secondaryFields"][0]["value"], "Mar 15")
        self.assertEqual(pass_json["eventTicket"]["secondaryFields"][1]["value"], "7.30")
        self.assertEqual(pass_json["eventTicket"]["auxiliaryFields"][0]["value"], "D")
        self.assertEqual(pass_json["eventTicket"]["auxiliaryFields"][1]["value"], "30")
        self.assertEqual(pass_json["barcode"]["format"], "PKBarcodeFormatQR")
        self.assertEqual(pass_json["barcode"]["message"], "decoded-qr-payload")
        self.assertNotIn("signature", names)
        self.assertEqual(sorted(manifest_json.keys()), ["icon.png", "icon@2x.png", "pass.json"])

    def test_extract_printed_barcode_value_from_page_prefers_unique_long_numeric_word(self) -> None:
        class FakePage:
            def get_text(self, mode: str):
                if mode == "words":
                    return [
                        (0, 10, 10, 20, "37638599", 0, 0, 0),
                        (0, 300, 10, 310, "431033772571", 0, 0, 0),
                    ]
                if mode == "text":
                    return "Order 37638599\n431033772571"
                raise AssertionError(mode)

        self.assertEqual(_extract_printed_barcode_value_from_page(FakePage()), "431033772571")

    def test_extract_printed_barcode_value_from_page_returns_none_for_multiple_bottom_candidates(self) -> None:
        class FakePage:
            def get_text(self, mode: str):
                if mode == "words":
                    return [
                        (0, 300, 10, 310, "123456789012", 0, 0, 0),
                        (20, 301, 30, 311, "987654321098", 0, 0, 1),
                    ]
                if mode == "text":
                    return "123456789012 987654321098"
                raise AssertionError(mode)

        self.assertIsNone(_extract_printed_barcode_value_from_page(FakePage()))

    def test_extract_seat_tokens_from_compact_token_before_order(self) -> None:
        text = "RoStalls A18Order 38238468"
        parsed = _extract_seat_tokens(text)
        self.assertEqual(parsed, ["A18"])

    def test_extract_seat_tokens_ignores_postcode_tokens(self) -> None:
        text = "Stalls C4C4Stalls Criterion Theatre, London SW1V 9LB"
        parsed = _extract_seat_tokens(text)
        self.assertEqual(parsed, ["C4"])

    def test_normalize_seat_labels(self) -> None:
        parsed = _normalize_seat_labels({"j11", "12 K", "bad"})
        self.assertEqual(parsed, {"J11", "K12"})

    def test_normalize_seat_labels_handles_standing_labels(self) -> None:
        parsed = _normalize_seat_labels({"standing 8", "GA 7", "bad"})
        self.assertEqual(parsed, {"STANDING7", "STANDING8"})

    def test_normalize_seat_labels_handles_kx_platform_prefix(self) -> None:
        parsed = _normalize_seat_labels({"Platform1 -F-8", "Platform1 -F-17"})
        self.assertEqual(parsed, {"F8", "F17"})

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

    def test_parse_allocation_with_multiple_emails_in_single_cell(self) -> None:
        csv_content = (
            "Booking Ref,Customer Name,Email,Assigned Seats\n"
            "B200,Jo Porter,danporter.contact@gmail.com theporterfamily1@sky.com,D64\n"
        )
        rows = parse_allocation_csv(csv_content)
        self.assertEqual(len(rows), 2)
        self.assertEqual(rows[0]["email"], "danporter.contact@gmail.com")
        self.assertEqual(rows[1]["email"], "theporterfamily1@sky.com")
        self.assertEqual(rows[0]["seats_raw"], "D64")
        self.assertEqual(rows[1]["seats_raw"], "D64")

    def test_parse_allocation_fills_down_all_emails_from_multi_email_row(self) -> None:
        csv_content = (
            "Booking Ref,Customer Name,Email,Assigned Seats\n"
            "B200,Jo Porter,danporter.contact@gmail.com theporterfamily1@sky.com,D64\n"
            ",,,D65\n"
        )
        rows = parse_allocation_csv(csv_content)
        self.assertEqual(len(rows), 4)
        self.assertEqual([r["email"] for r in rows], [
            "danporter.contact@gmail.com",
            "theporterfamily1@sky.com",
            "danporter.contact@gmail.com",
            "theporterfamily1@sky.com",
        ])
        self.assertEqual([r["seats_raw"] for r in rows], ["D64", "D64", "D65", "D65"])

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

    def test_parse_allocation_without_headers_splits_multiple_emails(self) -> None:
        csv_content = (
            "BOOK2001,Jo Porter,danporter.contact@gmail.com theporterfamily1@sky.com,D64\n"
            "BOOK2002,Jane,jane@example.com,C1\n"
        )
        rows = parse_allocation_csv(csv_content)
        self.assertEqual(len(rows), 3)
        self.assertEqual(rows[0]["email"], "danporter.contact@gmail.com")
        self.assertEqual(rows[1]["email"], "theporterfamily1@sky.com")
        self.assertEqual(rows[0]["seats_raw"], "D64")

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

    def test_build_bundle_zip_sets_non_zero_descending_timestamps(self) -> None:
        groups = [
            BookingTicketGroup(
                booking_reference="B100",
                customer_name="Alice Example",
                email="alice@example.com",
                seat_labels=["A1"],
                page_indexes=[0],
                missing_seats=[],
            ),
            BookingTicketGroup(
                booking_reference="B101",
                customer_name="Bob Example",
                email="bob@example.com",
                seat_labels=["A2"],
                page_indexes=[1],
                missing_seats=[],
            ),
            BookingTicketGroup(
                booking_reference="B102",
                customer_name="Cara Example",
                email="cara@example.com",
                seat_labels=["A3"],
                page_indexes=[2],
                missing_seats=[],
            ),
        ]

        with patch("allocator.ticket_bundle.build_group_pdf", side_effect=[b"pdf-one", b"pdf-two", b"pdf-three"]):
            zip_blob = build_bundle_zip(b"%PDF-pretend", groups)

        with zipfile.ZipFile(BytesIO(zip_blob)) as archive:
            names = archive.namelist()
            self.assertEqual(names[0], "manifest.csv")
            pdf_infos = [archive.getinfo(output_pdf_filename(group)) for group in groups]

        for info in pdf_infos:
            self.assertGreaterEqual(info.date_time[0], 1980)
            self.assertNotEqual(info.date_time, (1980, 1, 1, 0, 0, 0))

        self.assertGreater(pdf_infos[0].date_time, pdf_infos[1].date_time)
        self.assertGreater(pdf_infos[1].date_time, pdf_infos[2].date_time)

        with zipfile.ZipFile(BytesIO(zip_blob)) as archive:
            manifest_info = archive.getinfo("manifest.csv")

        self.assertLess(manifest_info.date_time, pdf_infos[-1].date_time)

    def test_build_bundle_zip_includes_wallet_passes_when_parsed_pages_supplied(self) -> None:
        groups = [
            BookingTicketGroup(
                booking_reference="B100",
                customer_name="Alice Example",
                email="alice@example.com",
                seat_labels=["D30"],
                page_indexes=[0],
                missing_seats=[],
            )
        ]
        parsed_pages = [
            ParsedTicketPage(
                page_index=0,
                show_name="Paddington The Musical",
                performance_date="Mar 15",
                performance_time="7.30",
                venue_name="Savoy Theatre",
                row="D",
                seat="30",
                seat_label="D30",
                qr_payload="decoded-qr-payload",
            )
        ]

        with patch("allocator.ticket_bundle.build_group_pdf", return_value=b"pdf-one"):
            zip_blob = build_bundle_zip(b"%PDF-pretend", groups, parsed_pages=parsed_pages)

        with zipfile.ZipFile(BytesIO(zip_blob)) as archive:
            self.assertIn("Alice_Example_tickets.pdf", archive.namelist())
            self.assertIn("wallet/D-30.pkpass", archive.namelist())

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

    def test_build_groups_preserves_seat_order_when_source_pages_are_reversed(self) -> None:
        rows = [
            {
                "booking_reference": "B001",
                "customer_name": "Jane",
                "email": "jane@example.com",
                "seats_raw": "L1-L4",
            }
        ]
        seat_to_page = {"L1": 3, "L2": 2, "L3": 1, "L4": 0}
        groups = build_booking_groups(rows, seat_to_page)

        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0].seat_labels, ["L1", "L2", "L3", "L4"])
        self.assertEqual(groups[0].page_indexes, [3, 2, 1, 0])

    def test_build_groups_maps_kx_platform_seat_values(self) -> None:
        rows = [
            {
                "booking_reference": "B001",
                "customer_name": "Jane",
                "email": "jane@example.com",
                "seats_raw": "Platform1 -F-8",
            }
        ]
        seat_to_page = {"F8": 0}
        groups = build_booking_groups(rows, seat_to_page)

        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0].seat_labels, ["F8"])
        self.assertEqual(groups[0].page_indexes, [0])
        self.assertEqual(groups[0].missing_seats, [])

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

    def test_build_groups_maps_standing_labels(self) -> None:
        rows = [
            {
                "booking_reference": "B100",
                "customer_name": "Rachel",
                "email": "rachel@example.com",
                "seats_raw": "Standing 1; Standing 2",
            }
        ]
        seat_to_page = {"STANDING1": 4, "STANDING2": 5}
        groups = build_booking_groups(rows, seat_to_page)

        self.assertEqual(len(groups), 1)
        self.assertEqual(groups[0].seat_labels, ["STANDING1", "STANDING2"])
        self.assertEqual(groups[0].page_indexes, [4, 5])
        self.assertEqual(groups[0].missing_seats, [])

    def test_split_groups_for_output_excludes_incomplete_groups(self) -> None:
        rows = [
            {
                "booking_reference": "B100",
                "customer_name": "Rachel",
                "email": "rachel@example.com",
                "seats_raw": "C4",
            },
            {
                "booking_reference": "B101",
                "customer_name": "Adam",
                "email": "adam@example.com",
                "seats_raw": "C5",
            },
            {
                "booking_reference": "B102",
                "customer_name": "Sue",
                "email": "sue@example.com",
                "seats_raw": "",
            },
        ]
        seat_to_page = {"C4": 0}
        groups = build_booking_groups(rows, seat_to_page)
        complete, excluded = split_groups_for_output(groups)

        self.assertEqual([g.email for g in complete], ["rachel@example.com"])
        self.assertEqual([g.email for g in excluded], ["adam@example.com", "sue@example.com"])

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

    def test_build_output_filenames_disambiguates_duplicate_customer_names(self) -> None:
        rows = [
            {
                "booking_reference": "B200",
                "customer_name": "Jo Porter",
                "email": "danporter.contact@gmail.com",
                "seats_raw": "D64",
            },
            {
                "booking_reference": "B200",
                "customer_name": "Jo Porter",
                "email": "theporterfamily1@sky.com",
                "seats_raw": "D64",
            },
        ]
        seat_to_page = {"D64": 0}
        groups = build_booking_groups(rows, seat_to_page)
        names = build_output_filenames(groups)
        self.assertEqual(names[0], "Jo_Porter_tickets.pdf")
        self.assertEqual(names[1], "Jo_Porter_tickets_theporterfamily1.pdf")


if __name__ == "__main__":
    unittest.main()
