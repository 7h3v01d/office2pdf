from __future__ import annotations

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


if __name__ == "__main__":
    unittest.main()
