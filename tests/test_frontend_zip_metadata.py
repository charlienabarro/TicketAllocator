from pathlib import Path
import re
import unittest


FRONTEND_APP_PATH = Path(__file__).resolve().parents[1] / "frontend" / "app.js"


class FrontendZipMetadataTests(unittest.TestCase):
    def test_direct_folder_save_writes_rows_in_reverse_so_first_numbers_row_finishes_last(self) -> None:
        source = FRONTEND_APP_PATH.read_text(encoding="utf-8")

        self.assertIn("for (const row of [...previewRows].reverse()) {", source)
        self.assertIn("// Write PDF last so Finder's newest-first sorting mirrors the Numbers-file order.", source)
        self.assertIn("previousPdfModifiedAt = await ensureFileHandleModifiedAfter(fileHandle, blob, previousPdfModifiedAt);", source)

    def test_direct_folder_save_rewrites_until_pdf_modified_time_is_newer(self) -> None:
        source = FRONTEND_APP_PATH.read_text(encoding="utf-8")

        self.assertIn("async function ensureFileHandleModifiedAfter(fileHandle, contents, previousModifiedAt, type = \"\") {", source)
        self.assertIn("currentModifiedAt = await getFileHandleModifiedAt(fileHandle);", source)
        self.assertRegex(source, r"if \(previousModifiedAt == null \|\| currentModifiedAt > previousModifiedAt\)")
        self.assertIn("await delay(25 * (attempt + 1));", source)
        self.assertIn("await writeFileHandleContents(fileHandle, contents, type);", source)

    def test_client_zip_builder_uses_encoded_modified_timestamps(self) -> None:
        source = FRONTEND_APP_PATH.read_text(encoding="utf-8")

        self.assertIn("function _encodeZipDosTimestamp(value)", source)
        self.assertIn("const modifiedAt = _encodeZipDosTimestamp(entry.modifiedAt || new Date());", source)

        self.assertRegex(source, r"lv\.setUint16\(10, modifiedAt\.time, true\)")
        self.assertRegex(source, r"lv\.setUint16\(12, modifiedAt\.date, true\)")
        self.assertRegex(source, r"cv\.setUint16\(12, modifiedAt\.time, true\)")
        self.assertRegex(source, r"cv\.setUint16\(14, modifiedAt\.date, true\)")

    def test_client_zip_entries_get_ordered_row_based_timestamps(self) -> None:
        source = FRONTEND_APP_PATH.read_text(encoding="utf-8")

        self.assertIn("const baseModifiedAt = _buildZipBaseModifiedAt();", source)
        self.assertIn("const pdfModifiedAt = _buildZipEntryModifiedAt(baseModifiedAt, rowIndex, 0);", source)
        self.assertIn("const emailModifiedAt = _buildZipEntryModifiedAt(baseModifiedAt, rowIndex, 1);", source)
        self.assertRegex(source, r"modifiedAt\.setSeconds\(modifiedAt\.getSeconds\(\) - \(rowIndex \* 4\) - \(entryOffset \* 2\)\);")

    def test_wallet_passes_are_written_to_selected_directory(self) -> None:
        source = FRONTEND_APP_PATH.read_text(encoding="utf-8")

        self.assertIn("const WALLET_FEATURE_ENABLED = false;", source)
        self.assertIn('const walletDirHandle = WALLET_FEATURE_ENABLED', source)
        self.assertIn("const walletPasses = WALLET_FEATURE_ENABLED && Array.isArray(row?.wallet_passes) ? row.wallet_passes : [];", source)
        self.assertIn("const passHandle = await walletDirHandle.getFileHandle(passFileName, { create: true });", source)

    def test_client_zip_builder_includes_wallet_pass_entries(self) -> None:
        source = FRONTEND_APP_PATH.read_text(encoding="utf-8")

        self.assertIn("function toDataUrlPkpassBlob(dataUrl)", source)
        self.assertIn("const passModifiedAt = _buildZipEntryModifiedAt(baseModifiedAt, rowIndex, 2 + passIndex);", source)
        self.assertIn("entries.push({ name: `${folderPrefix}/wallet/${passName}`, data: passBytes, modifiedAt: passModifiedAt });", source)

    def test_preview_renders_wallet_failures_clearly(self) -> None:
        source = FRONTEND_APP_PATH.read_text(encoding="utf-8")

        self.assertIn("const walletFailures = WALLET_FEATURE_ENABLED && Array.isArray(row.wallet_failures) ? row.wallet_failures : [];", source)
        self.assertIn("if (!WALLET_FEATURE_ENABLED) {", source)
        self.assertIn("function renderWalletFailures(failures)", source)
        self.assertIn("renderWalletFailures(data.wallet_failures || []);", source)


if __name__ == "__main__":
    unittest.main()
