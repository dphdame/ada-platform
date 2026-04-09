"""Tests for the screen reader transcript simulator.

Tests both the simulator itself AND uses it as a QI gate
to verify that the fixer produces documents that read correctly.
"""

from __future__ import annotations

import os
import struct
import zlib

from docx import Document
from docx.shared import Pt
from lxml import etree

from ada_platform.qa.screen_reader_sim import (
    ReadingTranscript,
    TranscriptEntry,
    generate_transcript,
)
from ada_platform.scanner.docx_scanner import DocxScanner
from ada_platform.fixer.docx_fixer import DocxFixer
from ada_platform.metadata.excel_reader import ExcelMetadataReader
from ada_platform.models.metadata import (
    DocumentMeta,
    ImageMeta,
    LinkMeta,
    MetadataWorkbook,
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _write_tiny_png(path: str) -> None:
    def _chunk(ct, data):
        c = ct + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)
    with open(path, "wb") as f:
        f.write(b"\x89PNG\r\n\x1a\n")
        f.write(_chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)))
        f.write(_chunk(b"IDAT", zlib.compress(b"\x00\xff\xff\xff")))
        f.write(_chunk(b"IEND", b""))


def _make_defective_docx(tmp_path) -> str:
    """DOCX with multiple accessibility issues."""
    png_path = str(tmp_path / "img.png")
    _write_tiny_png(png_path)

    doc = Document()
    doc.core_properties.title = ""

    # Strip language
    for lang in doc.element.findall(f".//{{{W_NS}}}lang"):
        lang.getparent().remove(lang)
    settings = doc.settings.element
    for lang in settings.findall(f".//{{{W_NS}}}lang"):
        lang.getparent().remove(lang)
    for tfl in settings.findall(f".//{{{W_NS}}}themeFontLang"):
        tfl.getparent().remove(tfl)

    doc.add_heading("Introduction", level=1)
    doc.add_paragraph("Some text.")
    doc.add_heading("Details", level=3)  # Skip H2
    doc.add_picture(png_path, width=Pt(72))

    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Name"
    table.cell(0, 1).text = "Score"
    table.cell(1, 0).text = "Alice"
    table.cell(1, 1).text = "95"

    p = doc.add_paragraph()
    hl = etree.SubElement(p._element, f"{{{W_NS}}}hyperlink")
    r_id = doc.part.relate_to(
        "https://example.com",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hl.set(f"{{{R_NS}}}id", r_id)
    run = etree.SubElement(hl, f"{{{W_NS}}}r")
    t = etree.SubElement(run, f"{{{W_NS}}}t")
    t.text = "click here"

    path = str(tmp_path / "defective.docx")
    doc.save(path)
    return path


# ── Simulator unit tests ─────────────────────────────────────────


class TestTranscriptCleanDoc:
    def test_clean_doc_no_warnings(self, tmp_path):
        doc = Document()
        doc.core_properties.title = "Clean Report"
        doc.add_heading("Section One", level=1)
        doc.add_paragraph("Body text.")
        doc.add_heading("Section Two", level=2)
        doc.add_paragraph("More text.")
        path = str(tmp_path / "clean.docx")
        doc.save(path)

        transcript = generate_transcript(path)
        # May have language warning depending on locale, filter it
        non_lang_warnings = [w for w in transcript.warnings if w.issue != "missing_language"]
        assert len(non_lang_warnings) == 0, (
            f"Clean doc has warnings: {[w.text for w in non_lang_warnings]}"
        )

    def test_title_announced(self, tmp_path):
        doc = Document()
        doc.core_properties.title = "Budget Report 2026"
        doc.add_paragraph("Content.")
        path = str(tmp_path / "titled.docx")
        doc.save(path)

        transcript = generate_transcript(path)
        doc_entries = [e for e in transcript.entries if e.role == "document"]
        assert any("Budget Report 2026" in e.text for e in doc_entries)

    def test_headings_announced_with_level(self, tmp_path):
        doc = Document()
        doc.core_properties.title = "Test"
        doc.add_heading("Main", level=1)
        doc.add_heading("Sub", level=2)
        path = str(tmp_path / "headings.docx")
        doc.save(path)

        transcript = generate_transcript(path)
        headings = [e for e in transcript.entries if e.role == "heading"]
        assert len(headings) == 2
        assert headings[0].level == 1
        assert headings[0].text == "Main"
        assert headings[1].level == 2

    def test_table_headers_announced(self, tmp_path):
        doc = Document()
        doc.core_properties.title = "Test"
        table = doc.add_table(rows=2, cols=2)
        table.cell(0, 0).text = "Name"
        table.cell(0, 1).text = "Score"
        table.cell(1, 0).text = "Alice"
        table.cell(1, 1).text = "95"
        # Mark header
        first_row = table.rows[0]._tr
        tr_pr = first_row.find(f"{{{W_NS}}}trPr")
        if tr_pr is None:
            tr_pr = etree.SubElement(first_row, f"{{{W_NS}}}trPr")
            first_row.insert(0, tr_pr)
        etree.SubElement(tr_pr, f"{{{W_NS}}}tblHeader")

        path = str(tmp_path / "table.docx")
        doc.save(path)

        transcript = generate_transcript(path)
        headers = [e for e in transcript.entries if e.role == "table_header"]
        cells = [e for e in transcript.entries if e.role == "table_cell"]
        assert len(headers) == 2
        assert headers[0].text == "Name"
        # Data cells should include header prefix
        assert any("Name: Alice" in c.text for c in cells)
        assert any("Score: 95" in c.text for c in cells)


class TestTranscriptDefectiveDoc:
    def test_missing_title_warning(self, tmp_path):
        path = _make_defective_docx(tmp_path)
        transcript = generate_transcript(path)
        title_warnings = [w for w in transcript.warnings if w.issue == "missing_title"]
        assert len(title_warnings) == 1

    def test_missing_alt_text_warning(self, tmp_path):
        path = _make_defective_docx(tmp_path)
        transcript = generate_transcript(path)
        alt_warnings = [w for w in transcript.warnings if w.issue == "missing_alt_text"]
        assert len(alt_warnings) >= 1
        assert "no description" in alt_warnings[0].text.lower()

    def test_table_no_header_warning(self, tmp_path):
        path = _make_defective_docx(tmp_path)
        transcript = generate_transcript(path)
        table_warnings = [w for w in transcript.warnings if w.issue == "table_no_header"]
        assert len(table_warnings) == 1

    def test_weak_link_warning(self, tmp_path):
        path = _make_defective_docx(tmp_path)
        transcript = generate_transcript(path)
        link_warnings = [w for w in transcript.warnings if w.issue == "weak_link_text"]
        assert len(link_warnings) == 1
        assert "click here" in link_warnings[0].text.lower()


class TestTranscriptTextOutput:
    def test_text_only_readable(self, tmp_path):
        doc = Document()
        doc.core_properties.title = "My Report"
        doc.add_heading("Overview", level=1)
        doc.add_paragraph("This is the overview section.")
        path = str(tmp_path / "readable.docx")
        doc.save(path)

        transcript = generate_transcript(path)
        text = transcript.text_only
        assert "[Document: Title: My Report]" in text
        assert "[Heading level 1] Overview" in text
        assert "This is the overview section." in text

    def test_warnings_in_text(self, tmp_path):
        path = _make_defective_docx(tmp_path)
        transcript = generate_transcript(path)
        text = transcript.text_only
        assert "!!" in text  # Warning markers


# ── QI Gate: Before/After Comparison ─────────────────────────────


class TestQIGateBeforeAfter:
    """The critical test: does the fixer actually fix what the screen reader hears?"""

    def test_remediation_fixes_transcript_warnings(self, tmp_path):
        """
        Full QI gate:
        1. Generate transcript of defective DOCX -> expect warnings
        2. Run scanner + fixer with metadata
        3. Generate transcript of fixed DOCX -> warnings should be reduced
        4. Verify specific improvements
        """
        # Build defective doc
        defective_path = _make_defective_docx(tmp_path)
        out_path = str(tmp_path / "fixed.docx")

        # BEFORE transcript
        before = generate_transcript(defective_path)
        before_warnings = before.warnings
        assert len(before_warnings) >= 3, (
            f"Expected 3+ warnings before fix, got {len(before_warnings)}: "
            f"{[w.issue for w in before_warnings]}"
        )

        # Types we expect to find before fix
        before_issues = {w.issue for w in before_warnings}
        assert "missing_title" in before_issues
        assert "missing_alt_text" in before_issues
        assert "table_no_header" in before_issues

        # Fix with metadata
        metadata = MetadataWorkbook(
            documents=[DocumentMeta(filename="defective.docx", title="Accessible Report")],
            images=[ImageMeta(image_id="Image 1", alt_text="Photo of team meeting", used_in="defective.docx")],
        )
        scanner = DocxScanner()
        scan_result = scanner.scan(defective_path, metadata=metadata)
        fixer = DocxFixer()
        fixer.fix(defective_path, out_path, scan_result, metadata=metadata)

        # AFTER transcript
        after = generate_transcript(out_path)
        after_issues = {w.issue for w in after.warnings}

        # Title should be fixed
        assert "missing_title" not in after_issues, "Title still missing after fix"
        # Alt text should be fixed (metadata provided it)
        assert "missing_alt_text" not in after_issues, "Alt text still missing after fix"
        # Table header should be fixed
        assert "table_no_header" not in after_issues, "Table header still missing after fix"

        # Verify the transcript now reads correctly
        text = after.text_only
        assert "Accessible Report" in text  # Title announced
        assert "Photo of team meeting" in text  # Alt text announced
        assert "[Header] Name" in text  # Table header announced

        # Data cells should now include header context
        assert "Name: Alice" in text or "Name:Alice" in text

    def test_before_after_warning_count_decreases(self, tmp_path):
        """Quantitative check: fewer warnings after remediation."""
        defective_path = _make_defective_docx(tmp_path)
        out_path = str(tmp_path / "fixed.docx")

        before = generate_transcript(defective_path)

        metadata = MetadataWorkbook(
            documents=[DocumentMeta(filename="defective.docx", title="Fixed Report")],
            images=[ImageMeta(image_id="Image 1", alt_text="Chart showing results", used_in="defective.docx")],
        )
        scanner = DocxScanner()
        scan_result = scanner.scan(defective_path, metadata=metadata)
        DocxFixer().fix(defective_path, out_path, scan_result, metadata=metadata)

        after = generate_transcript(out_path)

        assert len(after.warnings) < len(before.warnings), (
            f"Warning count did not decrease: {len(before.warnings)} -> {len(after.warnings)}\n"
            f"Before: {[w.issue for w in before.warnings]}\n"
            f"After: {[w.issue for w in after.warnings]}"
        )

    def test_transcript_summary(self, tmp_path):
        doc = Document()
        doc.core_properties.title = "Clean"
        doc.add_paragraph("Text.")
        path = str(tmp_path / "c.docx")
        doc.save(path)

        transcript = generate_transcript(path)
        summary = transcript.summary()
        # May or may not have language warning
        assert "transcript" in summary.lower()
