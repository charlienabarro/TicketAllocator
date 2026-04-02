import asyncio
from io import BytesIO
import unittest
import zipfile
from unittest.mock import patch

from fastapi import UploadFile

from allocator.ticket_bundle import ParsedTicketPage, ParsedTicketPageResult
from backend.app import download_preview_file, generate_ticket_bundle, preview_download_cache, preview_ticket_bundle


class TicketBundleWalletBackendTests(unittest.TestCase):
    def setUp(self) -> None:
        preview_download_cache.clear()
        self.parsed_pages = [
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
        self.parsed_page_results = [
            ParsedTicketPageResult(
                page_index=0,
                show_name="Paddington The Musical",
                performance_date="Mar 15",
                performance_time="7.30",
                venue_name="Savoy Theatre",
                row="D",
                seat="30",
                seat_label="D30",
                qr_payload="decoded-qr-payload",
                wallet_error=None,
            )
        ]

    def tearDown(self) -> None:
        preview_download_cache.clear()

    def _allocation_upload(self) -> UploadFile:
        return UploadFile(
            filename="allocation.csv",
            file=BytesIO(b"Customer Name,Email,Assigned Seats\nJane,jane@example.com,D30\n"),
        )

    def _ticket_upload(self) -> UploadFile:
        return UploadFile(
            filename="tickets.pdf",
            file=BytesIO(b"%PDF-pretend"),
        )

    async def _read_streaming_body(self, response) -> bytes:
        chunks: list[bytes] = []
        async for chunk in response.body_iterator:
            chunks.append(chunk)
        return b"".join(chunks)

    def test_preview_returns_wallet_pass_links(self) -> None:
        with patch("backend.app.WALLET_FEATURE_ENABLED", True):
            with patch("backend.app.parse_ticket_pdf_page_results", return_value=self.parsed_page_results):
                with patch("backend.app.build_group_pdf", return_value=b"pdf-one"):
                    payload = asyncio.run(
                        preview_ticket_bundle(
                            allocation_csv=self._allocation_upload(),
                            tickets_pdf=self._ticket_upload(),
                        )
                    )

        self.assertEqual(len(payload["rows"]), 1)
        wallet_passes = payload["rows"][0]["wallet_passes"]
        self.assertEqual(len(wallet_passes), 1)
        self.assertEqual(wallet_passes[0]["pass_file"], "D-30.pkpass")
        self.assertIn("data:application/vnd.apple.pkpass;base64,", wallet_passes[0]["pass_data_url"])

        pass_path = wallet_passes[0]["pass_url"]
        preview_id = pass_path.split("/")[3]
        file_name = pass_path.split("/")[-1]
        pass_response = download_preview_file(preview_id, file_name)
        self.assertEqual(pass_response.media_type, "application/vnd.apple.pkpass")

    def test_preview_reports_wallet_failures_without_blocking_pdf_preview(self) -> None:
        failed_result = ParsedTicketPageResult(
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

        with patch("backend.app.WALLET_FEATURE_ENABLED", True):
            with patch("backend.app.parse_ticket_pdf_page_results", return_value=[failed_result]):
                with patch("backend.app.build_group_pdf", return_value=b"pdf-one"):
                    payload = asyncio.run(
                        preview_ticket_bundle(
                            allocation_csv=self._allocation_upload(),
                            tickets_pdf=self._ticket_upload(),
                        )
                    )

        self.assertEqual(len(payload["rows"]), 1)
        self.assertEqual(payload["rows"][0]["wallet_passes"], [])
        self.assertEqual(
            payload["rows"][0]["wallet_failures"],
            [
                {
                    "seat_label": "D30",
                    "issue": "Could not decode a QR code from one or more ticket pages.",
                }
            ],
        )
        self.assertEqual(payload["stats"]["wallet_pass_count"], 0)
        self.assertEqual(payload["stats"]["wallet_failure_count"], 1)
        self.assertEqual(payload["wallet_failures"][0]["seat_label"], "D30")

    def test_preview_disables_wallet_work_when_feature_flag_is_off(self) -> None:
        with patch("backend.app.WALLET_FEATURE_ENABLED", False):
            with patch("backend.app.parse_ticket_pdf_page_results", return_value=self.parsed_page_results) as parse_results:
                with patch("backend.app.build_group_pdf", return_value=b"pdf-one"):
                    payload = asyncio.run(
                        preview_ticket_bundle(
                            allocation_csv=self._allocation_upload(),
                            tickets_pdf=self._ticket_upload(),
                        )
                    )

        parse_results.assert_called_once()
        self.assertEqual(payload["rows"][0]["wallet_passes"], [])
        self.assertEqual(payload["rows"][0]["wallet_failures"], [])
        self.assertEqual(payload["wallet_failures"], [])
        self.assertEqual(payload["stats"]["wallet_pass_count"], 0)
        self.assertEqual(payload["stats"]["wallet_failure_count"], 0)

    def test_generate_zip_includes_wallet_pass_bundle(self) -> None:
        with patch("backend.app.WALLET_FEATURE_ENABLED", True):
            with patch("backend.app.parse_ticket_pdf_pages", return_value=self.parsed_pages):
                with patch("allocator.ticket_bundle.build_group_pdf", return_value=b"pdf-one"):
                    response = asyncio.run(
                        generate_ticket_bundle(
                            allocation_csv=self._allocation_upload(),
                            tickets_pdf=self._ticket_upload(),
                        )
                    )

        zip_blob = asyncio.run(self._read_streaming_body(response))
        with zipfile.ZipFile(BytesIO(zip_blob)) as archive:
            names = archive.namelist()

        self.assertIn("Jane_tickets.pdf", names)
        self.assertIn("wallet/D-30.pkpass", names)

    def test_generate_zip_skips_wallet_when_feature_flag_is_off(self) -> None:
        with patch("backend.app.WALLET_FEATURE_ENABLED", False):
            with patch("backend.app.parse_ticket_pdf_page_results", return_value=self.parsed_page_results):
                with patch("allocator.ticket_bundle.build_group_pdf", return_value=b"pdf-one"):
                    response = asyncio.run(
                        generate_ticket_bundle(
                            allocation_csv=self._allocation_upload(),
                            tickets_pdf=self._ticket_upload(),
                        )
                    )

        zip_blob = asyncio.run(self._read_streaming_body(response))
        with zipfile.ZipFile(BytesIO(zip_blob)) as archive:
            names = archive.namelist()

        self.assertIn("Jane_tickets.pdf", names)
        self.assertNotIn("wallet/D-30.pkpass", names)


if __name__ == "__main__":
    unittest.main()
