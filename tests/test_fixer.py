"""Tests for the DOCX safe-fix engine. Uses round-trip: scan -> fix -> re-scan."""

from __future__ import annotations

import os
import struct
import zlib

from docx import Document
from lxml import etree

from ada_platform.models.issues import FixCategory, IssueType
from ada_platform.models.metadata import (
    DocumentMeta,
    ImageMeta,
    MetadataWorkbook,
)
from ada_platform.scanner.docx_scanner import DocxScanner
from ada_platform.fixer.docx_fixer import DocxFixer

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _write_tiny_png(path: str) -> None:
    def _chunk(ct: bytes, data: bytes) -> bytes:
        c = ct + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)
    sig = b"\x89PNG\r\n\x1a\n"
    ihdr = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    idat = zlib.compress(b"\x00\xff\xff\xff")
    with open(path, "wb") as f:
        f.write(sig + _chunk(b"IHDR", ihdr) + _chunk(b"IDAT", idat) + _chunk(b"IEND", b""))


def _save(doc: Document, path: str) -> str:
    doc.save(path)
    return path


class TestDocxFixerTitle:
    def test_fix_title_round_trip(self, tmp_path):
        # Create doc with no title
        doc = Document()
        doc.core_properties.title = ""
        doc.add_paragraph("Content.")
        in_path = _save(doc, str(tmp_path / "in.docx"))
        out_path = str(tmp_path / "out.docx")

        scanner = DocxScanner()
        fixer = DocxFixer()

        # Scan
        result = scanner.scan(in_path)
        title_issues = [i for i in result.issues if i.issue_type == IssueType.MISSING_TITLE]
        assert len(title_issues) == 1

        # Fix
        fix_result = fixer.fix(in_path, out_path, result)
        assert len(fix_result.auto_fixed) >= 1
        title_fixed = [i for i in fix_result.auto_fixed if i.issue_type == IssueType.MISSING_TITLE]
        assert len(title_fixed) == 1

        # Re-scan: title should be resolved
        result2 = scanner.scan(out_path)
        title_issues2 = [i for i in result2.issues if i.issue_type == IssueType.MISSING_TITLE]
        assert len(title_issues2) == 0


class TestDocxFixerLanguage:
    def test_fix_language_round_trip(self, tmp_path):
        doc = Document()
        doc.core_properties.title = "Test"
        # Strip all lang elements
        for lang in doc.element.findall(f".//{{{W_NS}}}lang"):
            lang.getparent().remove(lang)
        settings = doc.settings.element
        for lang in settings.findall(f".//{{{W_NS}}}lang"):
            lang.getparent().remove(lang)
        for tfl in settings.findall(f".//{{{W_NS}}}themeFontLang"):
            tfl.getparent().remove(tfl)

        in_path = _save(doc, str(tmp_path / "in.docx"))
        out_path = str(tmp_path / "out.docx")

        scanner = DocxScanner()
        result = scanner.scan(in_path)
        lang_before = [i for i in result.issues if i.issue_type == IssueType.MISSING_LANGUAGE]

        if lang_before:
            fix_result = DocxFixer().fix(in_path, out_path, result)
            lang_fixed = [i for i in fix_result.auto_fixed if i.issue_type == IssueType.MISSING_LANGUAGE]
            assert len(lang_fixed) == 1

            result2 = scanner.scan(out_path)
            lang_after = [i for i in result2.issues if i.issue_type == IssueType.MISSING_LANGUAGE]
            assert len(lang_after) == 0


class TestDocxFixerTableHeader:
    def test_fix_table_header_round_trip(self, tmp_path):
        doc = Document()
        doc.core_properties.title = "Test"
        table = doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "Header A"
        table.cell(0, 1).text = "Header B"
        table.cell(1, 0).text = "Data"
        table.cell(1, 1).text = "Data"
        in_path = _save(doc, str(tmp_path / "in.docx"))
        out_path = str(tmp_path / "out.docx")

        scanner = DocxScanner()
        result = scanner.scan(in_path)
        table_issues = [i for i in result.issues if i.issue_type == IssueType.TABLE_NO_HEADER]
        assert len(table_issues) == 1

        fix_result = DocxFixer().fix(in_path, out_path, result)
        table_fixed = [i for i in fix_result.auto_fixed if i.issue_type == IssueType.TABLE_NO_HEADER]
        assert len(table_fixed) == 1

        # Re-scan
        result2 = scanner.scan(out_path)
        table_issues2 = [i for i in result2.issues if i.issue_type == IssueType.TABLE_NO_HEADER]
        assert len(table_issues2) == 0


class TestDocxFixerAltText:
    def test_fix_alt_text_with_metadata(self, tmp_path):
        # Create doc with image, no alt text
        png_path = str(tmp_path / "test.png")
        _write_tiny_png(png_path)
        doc = Document()
        doc.core_properties.title = "Test"
        doc.add_picture(png_path)
        in_path = _save(doc, str(tmp_path / "report.docx"))
        out_path = str(tmp_path / "out.docx")

        metadata = MetadataWorkbook(
            images=[ImageMeta(
                image_id="Image 1",
                alt_text="Photo of community center",
                used_in="report.docx",
            )]
        )

        scanner = DocxScanner()
        result = scanner.scan(in_path, metadata=metadata)
        alt_issues = [i for i in result.issues if i.issue_type == IssueType.MISSING_ALT_TEXT]
        # Should be SAFE_AUTO because metadata provides alt text
        auto_alt = [i for i in alt_issues if i.fix_category == FixCategory.SAFE_AUTO]
        assert len(auto_alt) >= 1

        fix_result = DocxFixer().fix(in_path, out_path, result, metadata=metadata)
        alt_fixed = [i for i in fix_result.auto_fixed if i.issue_type == IssueType.MISSING_ALT_TEXT]
        assert len(alt_fixed) >= 1

        # Re-scan (without metadata this time, to verify alt text is actually in the file)
        result2 = scanner.scan(out_path)
        alt_issues2 = [i for i in result2.issues if i.issue_type == IssueType.MISSING_ALT_TEXT]
        assert len(alt_issues2) == 0


class TestDocxFixerHumanReview:
    def test_human_review_not_fixed(self, tmp_path):
        """Issues requiring human review should pass through unfixed."""
        png_path = str(tmp_path / "test.png")
        _write_tiny_png(png_path)
        doc = Document()
        doc.core_properties.title = "Test"
        doc.add_picture(png_path)
        in_path = _save(doc, str(tmp_path / "in.docx"))
        out_path = str(tmp_path / "out.docx")

        scanner = DocxScanner()
        result = scanner.scan(in_path)  # No metadata = HUMAN_REVIEW for alt text

        fix_result = DocxFixer().fix(in_path, out_path, result)
        review = [i for i in fix_result.needs_review if i.issue_type == IssueType.MISSING_ALT_TEXT]
        assert len(review) >= 1


class TestDocxFixerOutputFile:
    def test_output_file_created(self, tmp_path):
        doc = Document()
        doc.core_properties.title = ""
        doc.add_paragraph("Content.")
        in_path = _save(doc, str(tmp_path / "in.docx"))
        out_path = str(tmp_path / "out.docx")

        result = DocxScanner().scan(in_path)
        DocxFixer().fix(in_path, out_path, result)
        assert os.path.exists(out_path)

    def test_output_is_valid_docx(self, tmp_path):
        doc = Document()
        doc.core_properties.title = ""
        doc.add_paragraph("Content.")
        in_path = _save(doc, str(tmp_path / "in.docx"))
        out_path = str(tmp_path / "out.docx")

        result = DocxScanner().scan(in_path)
        DocxFixer().fix(in_path, out_path, result)

        # Should be openable as a valid DOCX
        doc2 = Document(out_path)
        assert doc2.core_properties.title  # Should have been fixed
