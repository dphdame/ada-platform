"""End-to-end integration tests: full pipeline from defective DOCX to remediated output."""

from __future__ import annotations

import json
import os
import struct
import zlib

from docx import Document
from lxml import etree
from openpyxl import Workbook

from ada_platform.scanner.docx_scanner import DocxScanner
from ada_platform.fixer.docx_fixer import DocxFixer
from ada_platform.metadata.excel_reader import ExcelMetadataReader
from ada_platform.reporting.exception_report import (
    generate_json_report,
    generate_text_report,
)
from ada_platform.models.issues import IssueType, FixCategory

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
    """Create a DOCX with 5+ accessibility issues."""
    png_path = str(tmp_path / "test.png")
    _write_tiny_png(png_path)

    doc = Document()
    # Issue 1: No title
    doc.core_properties.title = ""

    # Issue 2: Strip language
    for lang in doc.element.findall(f".//{{{W_NS}}}lang"):
        lang.getparent().remove(lang)
    settings = doc.settings.element
    for lang in settings.findall(f".//{{{W_NS}}}lang"):
        lang.getparent().remove(lang)
    for tfl in settings.findall(f".//{{{W_NS}}}themeFontLang"):
        tfl.getparent().remove(tfl)

    # Issue 3: Heading skip (H1 -> H3)
    doc.add_heading("Introduction", level=1)
    doc.add_paragraph("Some text.")
    doc.add_heading("Subsection", level=3)

    # Issue 4: Image with no alt text
    from docx.shared import Pt
    doc.add_picture(png_path, width=Pt(100))

    # Issue 5: Table with no header row
    table = doc.add_table(rows=3, cols=2)
    table.cell(0, 0).text = "Name"
    table.cell(0, 1).text = "Value"
    table.cell(1, 0).text = "Alpha"
    table.cell(1, 1).text = "100"

    # Issue 6: Weak hyperlink
    p = doc.add_paragraph()
    hl = etree.SubElement(p._element, f"{{{W_NS}}}hyperlink")
    r_id = doc.part.relate_to(
        "https://example.com/report",
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


def _make_metadata(tmp_path) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "Documents"
    ws.append(["filename", "title", "language"])
    ws.append(["defective.docx", "Accessible Report 2026", "en-US"])

    ws2 = wb.create_sheet("Images")
    ws2.append(["image_id", "content_hash", "alt_text", "long_description", "decorative", "used_in"])
    ws2.append(["Image 1", "", "Photo of community meeting", "", False, "defective.docx"])

    ws3 = wb.create_sheet("Links")
    ws3.append(["url", "approved_text", "content_type"])
    ws3.append(["https://example.com/report", "2026 Annual Report (PDF)", "PDF"])

    ws4 = wb.create_sheet("Charts")
    ws4.append(["chart_id"])
    ws5 = wb.create_sheet("Tables")
    ws5.append(["table_id"])
    ws6 = wb.create_sheet("Slides")
    ws6.append(["slide_number"])

    path = str(tmp_path / "meta.xlsx")
    wb.save(path)
    return path


class TestFullPipeline:
    def test_end_to_end_with_metadata(self, tmp_path):
        """Full pipeline: defective DOCX + metadata -> remediated DOCX + report."""
        docx_path = _make_defective_docx(tmp_path)
        meta_path = _make_metadata(tmp_path)
        out_path = str(tmp_path / "fixed.docx")
        report_path = str(tmp_path / "report.json")

        # Read metadata
        meta_wb = ExcelMetadataReader().read(meta_path)

        # Scan
        scanner = DocxScanner()
        scan_result = scanner.scan(docx_path, metadata=meta_wb)
        assert scan_result.total_issues >= 5, (
            f"Expected 5+ issues, got {scan_result.total_issues}: "
            f"{[i.issue_type.value for i in scan_result.issues]}"
        )

        # Verify issue types present
        types_found = {i.issue_type for i in scan_result.issues}
        assert IssueType.MISSING_TITLE in types_found
        assert IssueType.TABLE_NO_HEADER in types_found
        assert IssueType.WEAK_LINK_TEXT in types_found

        # Fix
        fixer = DocxFixer()
        fix_result = fixer.fix(docx_path, out_path, scan_result, metadata=meta_wb)

        # Verify fixes applied
        assert len(fix_result.auto_fixed) >= 2  # title + table header at minimum
        assert os.path.exists(out_path)

        # Generate reports
        generate_json_report(fix_result, report_path)
        assert os.path.exists(report_path)

        with open(report_path) as f:
            report_data = json.load(f)
        assert "auto_fixed" in report_data
        assert "needs_review" in report_data

        # Text report
        text = generate_text_report(fix_result)
        assert "ACCESSIBILITY REPORT" in text
        # No XML/OOXML terms
        for term in ["xml", "ooxml", "w:lang", "tblHeader", "docPr"]:
            assert term.lower() not in text.lower()

        # Re-scan: auto-fixed issues should be resolved
        result2 = scanner.scan(out_path, metadata=meta_wb)
        fixed_types = {i.issue_type for i in fix_result.auto_fixed}
        remaining_types = {i.issue_type for i in result2.issues}
        for ft in fixed_types:
            if ft in (IssueType.MISSING_TITLE, IssueType.TABLE_NO_HEADER):
                assert ft not in remaining_types, (
                    f"{ft.value} was 'fixed' but still appears on re-scan"
                )

    def test_end_to_end_without_metadata(self, tmp_path):
        """Pipeline works without metadata; more issues go to human review."""
        docx_path = _make_defective_docx(tmp_path)
        out_path = str(tmp_path / "fixed.docx")

        scanner = DocxScanner()
        scan_result = scanner.scan(docx_path)

        fixer = DocxFixer()
        fix_result = fixer.fix(docx_path, out_path, scan_result)

        # Without metadata, alt text goes to human review
        review_types = {i.issue_type for i in fix_result.needs_review}
        assert IssueType.MISSING_ALT_TEXT in review_types

        # Title should still be auto-fixed (from filename)
        fixed_types = {i.issue_type for i in fix_result.auto_fixed}
        assert IssueType.MISSING_TITLE in fixed_types

    def test_clean_doc_passes_cleanly(self, tmp_path):
        """A clean document should have minimal or no issues."""
        doc = Document()
        doc.core_properties.title = "Clean Document"
        doc.add_heading("Section 1", level=1)
        doc.add_paragraph("Body text.")
        doc.add_heading("Section 2", level=2)
        doc.add_paragraph("More text.")
        path = str(tmp_path / "clean.docx")
        doc.save(path)

        scanner = DocxScanner()
        result = scanner.scan(path)

        # Should have no heading, table, link, or alt text issues
        structural = [
            i for i in result.issues
            if i.issue_type not in (IssueType.MISSING_LANGUAGE,)
        ]
        assert len(structural) == 0, (
            f"Clean doc has issues: {[i.issue_type.value for i in structural]}"
        )
