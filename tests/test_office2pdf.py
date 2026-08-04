from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import office2pdf


def write_test_pdf(path: Path, marker: bytes = b"test") -> None:
    payload = b"%PDF-1.4\n" + marker + (b"x" * 180) + b"\n%%EOF\n"
    path.write_bytes(payload)


class ConversionTests(unittest.TestCase):
    def setUp(self) -> None:
        self.tempdir = tempfile.TemporaryDirectory()
        self.root = Path(self.tempdir.name)
        self.source = self.root / "newsletter.docx"
        self.source.write_bytes(b"placeholder source")
        self.output = self.root / "output"
        self.output.mkdir()

    def tearDown(self) -> None:
        self.tempdir.cleanup()

    @staticmethod
    def _successful_fake_soffice(args, env, timeout):
        out_dir = Path(args[args.index("--outdir") + 1])
        source = Path(args[-1])
        write_test_pdf(out_dir / f"{source.stem}.pdf", b"new-pdf")
        return 0, "converted", False

    @mock.patch("office2pdf._run_soffice")
    def test_overwrite_uses_new_staged_pdf(self, run_mock):
        run_mock.side_effect = self._successful_fake_soffice
        destination = self.output / "newsletter.pdf"
        write_test_pdf(destination, b"old-pdf")
        old_bytes = destination.read_bytes()

        result = office2pdf.convert_one(
            self.source, self.output, "soffice", 30, 0, True
        )

        self.assertEqual(result.status, office2pdf.ConversionStatus.SUCCESS)
        self.assertNotEqual(destination.read_bytes(), old_bytes)
        self.assertIn(b"new-pdf", destination.read_bytes())

    @mock.patch("office2pdf._run_soffice", return_value=(0, "claimed success", False))
    def test_stale_destination_cannot_create_false_success(self, _run_mock):
        destination = self.output / "newsletter.pdf"
        write_test_pdf(destination, b"old-pdf")
        old_bytes = destination.read_bytes()

        result = office2pdf.convert_one(
            self.source, self.output, "soffice", 30, 0, True
        )

        self.assertEqual(result.status, office2pdf.ConversionStatus.FAILED)
        self.assertEqual(destination.read_bytes(), old_bytes)

    @mock.patch("office2pdf._run_soffice", side_effect=OSError("launch blocked"))
    def test_launch_error_becomes_structured_failure(self, _run_mock):
        result = office2pdf.convert_one(
            self.source, self.output, "soffice", 30, 0, False
        )
        self.assertEqual(result.status, office2pdf.ConversionStatus.FAILED)
        self.assertIn("launch blocked", result.message)

    def test_valid_existing_pdf_is_skipped(self):
        destination = self.output / "newsletter.pdf"
        write_test_pdf(destination)

        result = office2pdf.convert_one(
            self.source, self.output, "soffice", 30, 0, False
        )

        self.assertEqual(result.status, office2pdf.ConversionStatus.SKIPPED)
        self.assertEqual(result.output, destination)

    def test_invalid_existing_pdf_is_not_silently_skipped(self):
        destination = self.output / "newsletter.pdf"
        destination.write_bytes(b"not a pdf")

        result = office2pdf.convert_one(
            self.source, self.output, "soffice", 30, 0, False
        )

        self.assertEqual(result.status, office2pdf.ConversionStatus.FAILED)
        self.assertIn("not a valid PDF", result.message)

    def test_invalid_runtime_values_are_contained(self):
        timeout_result = office2pdf.convert_one(
            self.source, self.output, "soffice", 0, 0, False
        )
        retry_result = office2pdf.convert_one(
            self.source, self.output, "soffice", 10, -1, False
        )
        self.assertEqual(timeout_result.status, office2pdf.ConversionStatus.FAILED)
        self.assertEqual(retry_result.status, office2pdf.ConversionStatus.FAILED)


class NativeWorkerSafetyTests(unittest.TestCase):
    def test_native_timeout_is_a_real_process_deadline(self) -> None:
        with tempfile.TemporaryDirectory() as tempdir:
            root = Path(tempdir)
            sleeper = root / "sleep_worker.py"
            sleeper.write_text(
                "import time\ntime.sleep(5)\n",
                encoding="utf-8",
            )
            source = root / "source.docx"
            source.write_bytes(b"placeholder")
            destination = root / "result.pdf"

            with mock.patch(
                "office2pdf._native_worker_command",
                return_value=[sys.executable, str(sleeper)],
            ):
                started = __import__("time").monotonic()
                with self.assertRaises(TimeoutError):
                    office2pdf._run_native_with_timeout(
                        office2pdf.Backend.MS_WORD,
                        source,
                        destination,
                        timeout=1,
                    )
                elapsed = __import__("time").monotonic() - started

            # The old ThreadPoolExecutor implementation returned only after
            # the five-second sleeper finished despite its one-second timeout.
            self.assertLess(elapsed, 2.5)

    @mock.patch("office2pdf.choose_backend", return_value=office2pdf.Backend.MS_WORD)
    @mock.patch("office2pdf._run_native_with_timeout", side_effect=RuntimeError("COM failed"))
    @mock.patch("office2pdf._run_soffice")
    def test_native_failure_falls_back_to_libreoffice(
        self,
        run_soffice,
        _run_native,
        _choose_backend,
    ) -> None:
        with tempfile.TemporaryDirectory() as tempdir:
            root = Path(tempdir)
            source = root / "newsletter.docx"
            source.write_bytes(b"placeholder")
            output = root / "output"
            output.mkdir()

            def successful_soffice(args, env, timeout):
                out_dir = Path(args[args.index("--outdir") + 1])
                write_test_pdf(out_dir / "newsletter.pdf", b"fallback")
                return 0, "converted", False

            run_soffice.side_effect = successful_soffice
            # The recursive fallback calls choose_backend(..., False). Make
            # that second call return LibreOffice rather than the patched word.
            _choose_backend.side_effect = [
                office2pdf.Backend.MS_WORD,
                office2pdf.Backend.LIBREOFFICE,
            ]

            result = office2pdf.convert_one(
                source,
                output,
                "soffice",
                timeout=10,
                retries=0,
                overwrite=False,
                prefer_native=True,
            )

            self.assertEqual(result.status, office2pdf.ConversionStatus.SUCCESS)
            self.assertEqual(result.backend, "libreoffice")
            self.assertIn("LibreOffice fallback", result.message)
            self.assertTrue((output / "newsletter.pdf").exists())

    def test_native_office_is_opt_in_by_default(self) -> None:
        parser = office2pdf.build_parser()
        args = parser.parse_args(["example.docx"])
        self.assertFalse(args.prefer_native)
        args = parser.parse_args(["example.docx", "--native-office"])
        self.assertTrue(args.prefer_native)


class DiscoveryTests(unittest.TestCase):
    def test_discovery_is_case_insensitive_sorted_and_excludes_lock_files(self):
        with tempfile.TemporaryDirectory() as tempdir:
            root = Path(tempdir)
            (root / "zeta.DOCX").write_bytes(b"x")
            (root / "Alpha.XLSX").write_bytes(b"x")
            (root / "~$Alpha.xlsx").write_bytes(b"x")
            (root / ".hidden.docx").write_bytes(b"x")
            (root / "notes.txt").write_bytes(b"x")

            found = office2pdf.discover_inputs([str(root)], recursive=False)

            self.assertEqual([path.name for path in found], ["Alpha.XLSX", "zeta.DOCX"])

    def test_collision_detection_refuses_same_stem(self):
        with tempfile.TemporaryDirectory() as tempdir:
            root = Path(tempdir)
            first = root / "report.docx"
            second = root / "report.xlsx"
            collisions = office2pdf.find_destination_collisions(
                [first, second], str(root / "out")
            )
            self.assertEqual(len(collisions), 1)


class ParserTests(unittest.TestCase):
    def test_parser_rejects_nonpositive_jobs_and_timeout(self):
        parser = office2pdf.build_parser()
        with self.assertRaises(SystemExit):
            parser.parse_args(["example.docx", "--jobs", "0"])
        with self.assertRaises(SystemExit):
            parser.parse_args(["example.docx", "--timeout", "0"])
        with self.assertRaises(SystemExit):
            parser.parse_args(["example.docx", "--retries", "-1"])


def _soffice_available() -> str | None:
    try:
        return office2pdf.find_soffice()
    except office2pdf.LibreOfficeNotFoundError:
        return None


SOFFICE = _soffice_available()


@unittest.skipUnless(SOFFICE, "LibreOffice not found on this system")
class RealLibreOfficeEndToEndTests(unittest.TestCase):
    """Runs the same guarantees as ConversionTests above, but against a real
    soffice binary instead of a mocked `_run_soffice`. The mocked tests above
    are the fast day-to-day suite; this class is the belt-and-suspenders
    check that nothing about the real subprocess/staging/validation pipeline
    broke, and that it's still true end to end, not just true of the mock."""

    def setUp(self) -> None:
        self.work = Path(tempfile.mkdtemp(prefix="o2p_e2e_"))
        src_txt = self.work / "doc.txt"
        src_txt.write_text("a real test document\n")
        subprocess.run(
            [SOFFICE, "--headless", "--convert-to", "docx", "--outdir", str(self.work), str(src_txt)],
            check=True, capture_output=True, timeout=60,
        )
        self.src_docx = self.work / "doc.docx"
        assert self.src_docx.exists()

    def tearDown(self) -> None:
        shutil.rmtree(self.work, ignore_errors=True)

    def test_real_conversion_produces_validated_pdf(self) -> None:
        result = office2pdf.convert_one(self.src_docx, self.work, SOFFICE, 60, 1, False)
        self.assertEqual(result.status, office2pdf.ConversionStatus.SUCCESS)
        self.assertEqual(result.backend, "libreoffice")
        ok, reason = office2pdf.validate_pdf(result.output)
        self.assertTrue(ok, reason)

    def test_real_skip_on_valid_existing_pdf(self) -> None:
        office2pdf.convert_one(self.src_docx, self.work, SOFFICE, 60, 1, False)
        result = office2pdf.convert_one(self.src_docx, self.work, SOFFICE, 60, 1, False)
        self.assertEqual(result.status, office2pdf.ConversionStatus.SKIPPED)

    def test_real_invalid_existing_pdf_is_not_silently_skipped(self) -> None:
        (self.work / "doc.pdf").write_bytes(b"not really a pdf")
        result = office2pdf.convert_one(self.src_docx, self.work, SOFFICE, 60, 1, False)
        self.assertEqual(result.status, office2pdf.ConversionStatus.FAILED)

    def test_real_failed_overwrite_does_not_corrupt_existing_good_pdf(self) -> None:
        """Regression test for the original P0 bug: a failed --overwrite
        attempt must never (a) report success, or (b) alter the existing
        good PDF -- exercised here against a real soffice binary that then
        gets swapped for a broken one, not a mock."""
        good = office2pdf.convert_one(self.src_docx, self.work, SOFFICE, 60, 1, False)
        self.assertEqual(good.status, office2pdf.ConversionStatus.SUCCESS)
        existing_bytes = good.output.read_bytes()

        bad = office2pdf.convert_one(self.src_docx, self.work, "/bin/false", 10, 0, True)
        self.assertEqual(bad.status, office2pdf.ConversionStatus.FAILED)
        self.assertEqual(good.output.read_bytes(), existing_bytes)

    def test_real_overwrite_replaces_with_new_valid_pdf(self) -> None:
        office2pdf.convert_one(self.src_docx, self.work, SOFFICE, 60, 1, False)
        result = office2pdf.convert_one(self.src_docx, self.work, SOFFICE, 60, 1, True)
        self.assertEqual(result.status, office2pdf.ConversionStatus.SUCCESS)

    def test_real_full_batch_via_cli_main(self) -> None:
        rc = office2pdf.main([str(self.src_docx)])
        self.assertEqual(rc, 0)
        self.assertTrue((self.work / "doc.pdf").exists())

    def test_native_backend_is_a_safe_noop_on_this_platform(self) -> None:
        """choose_backend() must only ever pick a native backend on Windows
        with the matching app installed; everywhere else (including here)
        it must transparently fall back to LibreOffice."""
        if sys.platform.startswith("win"):
            self.skipTest("this check is only meaningful off Windows")
        self.assertEqual(
            office2pdf.choose_backend(self.src_docx, prefer_native=True),
            office2pdf.Backend.LIBREOFFICE,
        )


def _native_word_available() -> bool:
    if not sys.platform.startswith("win"):
        return False
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return office2pdf._native_office_available().get("word", False)


NATIVE_WORD_AVAILABLE = _native_word_available()


@unittest.skipUnless(SOFFICE, "LibreOffice not found -- needed to author the .docx fixture for this test")
@unittest.skipUnless(
    NATIVE_WORD_AVAILABLE, "Windows + pywin32 + Microsoft Word not detected on this machine"
)
class NativeWordEndToEndTests(unittest.TestCase):
    """Exercises the real Word COM automation path (_convert_with_word), not
    a mock. This is the one backend that could not be verified during
    development -- it only proves anything the moment it actually runs on a
    Windows machine with Word installed. LibreOffice is still used here
    (via SOFFICE) purely to author the .docx fixture; the conversion under
    test never touches it -- soffice_bin is passed as None below, so a pass
    means Word did the export, not a silent LibreOffice fallback."""

    def setUp(self) -> None:
        self.work = Path(tempfile.mkdtemp(prefix="o2p_native_word_"))
        src_txt = self.work / "doc.txt"
        src_txt.write_text("a real test document for native Word export\n")
        subprocess.run(
            [SOFFICE, "--headless", "--convert-to", "docx", "--outdir", str(self.work), str(src_txt)],
            check=True, capture_output=True, timeout=60,
        )
        self.src_docx = self.work / "doc.docx"
        assert self.src_docx.exists()

    def tearDown(self) -> None:
        shutil.rmtree(self.work, ignore_errors=True)

    def test_choose_backend_selects_word_for_docx(self) -> None:
        self.assertEqual(
            office2pdf.choose_backend(self.src_docx, prefer_native=True),
            office2pdf.Backend.MS_WORD,
        )

    def test_real_word_export_produces_validated_pdf(self) -> None:
        # soffice_bin is deliberately None: a SUCCESS result with
        # backend == "word" proves Word did the export, not that it
        # silently fell back to LibreOffice.
        result = office2pdf.convert_one(
            self.src_docx, self.work, None, 60, 0, False, prefer_native=True
        )
        self.assertEqual(result.status, office2pdf.ConversionStatus.SUCCESS, result.message)
        self.assertEqual(result.backend, "word")
        ok, reason = office2pdf.validate_pdf(result.output)
        self.assertTrue(ok, reason)

    def test_real_word_skip_on_valid_existing_pdf(self) -> None:
        office2pdf.convert_one(self.src_docx, self.work, None, 60, 0, False, prefer_native=True)
        result = office2pdf.convert_one(self.src_docx, self.work, None, 60, 0, False, prefer_native=True)
        self.assertEqual(result.status, office2pdf.ConversionStatus.SKIPPED)

    def test_real_word_overwrite_replaces_with_new_valid_pdf(self) -> None:
        office2pdf.convert_one(self.src_docx, self.work, None, 60, 0, False, prefer_native=True)
        result = office2pdf.convert_one(self.src_docx, self.work, None, 60, 0, True, prefer_native=True)
        self.assertEqual(result.status, office2pdf.ConversionStatus.SUCCESS)
        self.assertEqual(result.backend, "word")

    def test_legacy_no_native_office_flag_still_forces_libreoffice(self) -> None:
        out_dir = self.work / "lo_out"
        rc = office2pdf.main([str(self.src_docx), "--no-native-office", "-o", str(out_dir)])
        self.assertEqual(rc, 0)
        self.assertTrue((out_dir / "doc.pdf").exists())


def _native_excel_available() -> bool:
    if not sys.platform.startswith("win"):
        return False
    try:
        import pythoncom  # noqa: F401
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return office2pdf._native_office_available().get("excel", False)


NATIVE_EXCEL_AVAILABLE = _native_excel_available()


@unittest.skipUnless(SOFFICE, "LibreOffice not found -- needed to author the .xlsx fixture for this test")
@unittest.skipUnless(
    NATIVE_EXCEL_AVAILABLE, "Windows + pywin32 + Microsoft Excel not detected on this machine"
)
class NativeExcelEndToEndTests(unittest.TestCase):
    """Exercises the real isolated Excel COM worker for the target XLSX path."""

    def setUp(self) -> None:
        self.work = Path(tempfile.mkdtemp(prefix="o2p_native_excel_"))
        source_csv = self.work / "accounts.csv"
        source_csv.write_text(
            "Category,Budget,Actual\nMaintenance,1250,1187.45\nUtilities,900,876.32\n",
            encoding="utf-8",
        )
        subprocess.run(
            [
                SOFFICE,
                "--headless",
                "--convert-to",
                "xlsx",
                "--outdir",
                str(self.work),
                str(source_csv),
            ],
            check=True,
            capture_output=True,
            timeout=60,
        )
        self.src_xlsx = self.work / "accounts.xlsx"
        assert self.src_xlsx.exists()

    def tearDown(self) -> None:
        shutil.rmtree(self.work, ignore_errors=True)

    def test_choose_backend_selects_excel_for_xlsx(self) -> None:
        self.assertEqual(
            office2pdf.choose_backend(self.src_xlsx, prefer_native=True),
            office2pdf.Backend.MS_EXCEL,
        )

    def test_real_excel_export_produces_validated_pdf(self) -> None:
        result = office2pdf.convert_one(
            self.src_xlsx,
            self.work,
            None,
            60,
            0,
            False,
            prefer_native=True,
        )
        self.assertEqual(result.status, office2pdf.ConversionStatus.SUCCESS, result.message)
        self.assertEqual(result.backend, "excel")
        ok, reason = office2pdf.validate_pdf(result.output)
        self.assertTrue(ok, reason)

    def test_real_excel_overwrite_replaces_with_valid_pdf(self) -> None:
        first = office2pdf.convert_one(
            self.src_xlsx, self.work, None, 60, 0, False, prefer_native=True
        )
        self.assertEqual(first.status, office2pdf.ConversionStatus.SUCCESS, first.message)
        result = office2pdf.convert_one(
            self.src_xlsx, self.work, None, 60, 0, True, prefer_native=True
        )
        self.assertEqual(result.status, office2pdf.ConversionStatus.SUCCESS, result.message)
        self.assertEqual(result.backend, "excel")


if __name__ == "__main__":
    unittest.main()
