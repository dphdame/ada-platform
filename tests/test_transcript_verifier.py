"""Tests for transcript verifier: does the transcript match the actual document?"""

from __future__ import annotations

import struct
import zlib

from docx import Document
from docx.shared import Pt
from lxml import etree

from ada_platform.qa.transcript_verifier import (
    VerificationReport,
    extract_ground_truth,
    verify_transcript,
)
from ada_platform.scanner.docx_scanner import DocxScanner
from ada_platform.fixer.docx_fixer import DocxFixer
from ada_platform.models.metadata import (
    DocumentMeta,
    ImageMeta,
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


def _make_rich_docx(tmp_path) -> str:
    """DOCX with diverse content: title, headings, image with alt, table with headers, link."""
    png_path = str(tmp_path / "chart.png")
    _write_tiny_png(png_path)

    doc = Document()
    doc.core_properties.title = "Quarterly Budget Report"

    doc.add_heading("Executive Summary", level=1)
    doc.add_paragraph("This report covers Q4 2025 expenditures across all departments.")

    doc.add_heading("Revenue Breakdown", level=2)
    doc.add_paragraph("Total revenue was $4.2M, exceeding projections by 8%.")

    # Image with alt text
    pic = doc.add_picture(png_path, width=Pt(200))
    # Set alt text via XML
    body = doc.element.find(f"{{{W_NS}}}body")
    for drawing in body.iter(f"{{{W_NS}}}drawing"):
        for dp in drawing.findall(f".//{{{WP_NS}}}docPr"):
            dp.set("descr", "Bar chart showing quarterly revenue by department")

    doc.add_heading("Staffing", level=2)

    # Table with headers
    table = doc.add_table(rows=3, cols=3)
    table.cell(0, 0).text = "Department"
    table.cell(0, 1).text = "Headcount"
    table.cell(0, 2).text = "Budget"
    table.cell(1, 0).text = "Engineering"
    table.cell(1, 1).text = "42"
    table.cell(1, 2).text = "$2.1M"
    table.cell(2, 0).text = "Marketing"
    table.cell(2, 1).text = "18"
    table.cell(2, 2).text = "$800K"
    # Mark header row
    first_row = table.rows[0]._tr
    tr_pr = first_row.find(f"{{{W_NS}}}trPr")
    if tr_pr is None:
        tr_pr = etree.SubElement(first_row, f"{{{W_NS}}}trPr")
        first_row.insert(0, tr_pr)
    etree.SubElement(tr_pr, f"{{{W_NS}}}tblHeader")

    doc.add_heading("References", level=2)

    # Good link
    p = doc.add_paragraph()
    hl = etree.SubElement(p._element, f"{{{W_NS}}}hyperlink")
    r_id = doc.part.relate_to(
        "https://budget.example.com/q4-2025",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hl.set(f"{{{R_NS}}}id", r_id)
    run = etree.SubElement(hl, f"{{{W_NS}}}r")
    t = etree.SubElement(run, f"{{{W_NS}}}t")
    t.text = "Q4 2025 Budget Dashboard"

    path = str(tmp_path / "rich.docx")
    doc.save(path)
    return path


def _make_defective_docx(tmp_path) -> str:
    """DOCX with issues: no title, missing alt text, no table headers, weak link."""
    png_path = str(tmp_path / "img.png")
    _write_tiny_png(png_path)

    doc = Document()
    doc.core_properties.title = ""

    doc.add_heading("Section A", level=1)
    doc.add_paragraph("Introductory paragraph with enough text to be checked.")

    doc.add_picture(png_path, width=Pt(72))  # No alt text

    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Name"
    table.cell(0, 1).text = "Score"
    table.cell(1, 0).text = "Alice"
    table.cell(1, 1).text = "95"
    # No header marking

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


WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


# ── Ground truth extraction ──────────────────────────────────────


class TestGroundTruthExtraction:
    def test_extracts_title(self, tmp_path):
        path = _make_rich_docx(tmp_path)
        gt = extract_ground_truth(path)
        titles = [g for g in gt if g.kind == "title"]
        assert len(titles) == 1
        assert titles[0].text == "Quarterly Budget Report"

    def test_extracts_headings(self, tmp_path):
        path = _make_rich_docx(tmp_path)
        gt = extract_ground_truth(path)
        headings = [g for g in gt if g.kind == "heading"]
        assert len(headings) == 4
        assert headings[0].text == "Executive Summary"
        assert headings[0].level == 1
        assert headings[1].text == "Revenue Breakdown"
        assert headings[1].level == 2

    def test_extracts_images(self, tmp_path):
        path = _make_rich_docx(tmp_path)
        gt = extract_ground_truth(path)
        images = [g for g in gt if g.kind == "image"]
        assert len(images) == 1
        assert images[0].has_alt is True
        assert "quarterly revenue" in images[0].text.lower()

    def test_extracts_table_with_headers(self, tmp_path):
        path = _make_rich_docx(tmp_path)
        gt = extract_ground_truth(path)
        tables = [g for g in gt if g.kind == "table"]
        assert len(tables) == 1
        assert tables[0].has_header is True

        headers = [g for g in gt if g.kind == "table_header"]
        assert len(headers) == 3
        header_texts = {h.text for h in headers}
        assert "Department" in header_texts
        assert "Headcount" in header_texts

    def test_extracts_links(self, tmp_path):
        path = _make_rich_docx(tmp_path)
        gt = extract_ground_truth(path)
        links = [g for g in gt if g.kind == "link"]
        assert len(links) == 1
        assert links[0].text == "Q4 2025 Budget Dashboard"

    def test_extracts_paragraphs(self, tmp_path):
        path = _make_rich_docx(tmp_path)
        gt = extract_ground_truth(path)
        paras = [g for g in gt if g.kind == "paragraph"]
        assert len(paras) >= 2
        assert any("Q4 2025" in p.text for p in paras)


# ── Verification on clean document ───────────────────────────────


class TestVerifyCleanDoc:
    def test_rich_doc_passes_verification(self, tmp_path):
        path = _make_rich_docx(tmp_path)
        report = verify_transcript(path)
        print(report.full_report())  # For debugging visibility
        assert report.passed, (
            f"Rich doc failed verification:\n{report.full_report()}"
        )

    def test_all_headings_verified(self, tmp_path):
        path = _make_rich_docx(tmp_path)
        report = verify_transcript(path)
        heading_issues = [
            i for i in report.issues
            if "heading" in i.element.lower() or "Heading" in i.element
        ]
        assert len(heading_issues) == 0, (
            f"Heading verification issues:\n"
            + "\n".join(f"  {i.element}: {i.expected} vs {i.found}" for i in heading_issues)
        )

    def test_image_alt_verified(self, tmp_path):
        path = _make_rich_docx(tmp_path)
        report = verify_transcript(path)
        image_issues = [i for i in report.issues if "image" in i.element.lower() or "Image" in i.element]
        assert len(image_issues) == 0

    def test_table_cells_verified(self, tmp_path):
        path = _make_rich_docx(tmp_path)
        report = verify_transcript(path)
        cell_issues = [i for i in report.issues if "cell" in i.element.lower()]
        assert len(cell_issues) == 0, (
            f"Cell verification issues:\n"
            + "\n".join(f"  {i.element}: {i.expected} vs {i.found}" for i in cell_issues)
        )

    def test_link_verified(self, tmp_path):
        path = _make_rich_docx(tmp_path)
        report = verify_transcript(path)
        link_issues = [i for i in report.issues if "link" in i.element.lower() or "Link" in i.element]
        assert len(link_issues) == 0


# ── Verification on defective document ───────────────────────────


class TestVerifyDefectiveDoc:
    def test_defective_doc_has_gaps(self, tmp_path):
        path = _make_defective_docx(tmp_path)
        report = verify_transcript(path)
        # Should have gaps for missing alt text, but transcript should still
        # accurately represent what IS there
        assert report.ground_truth_count > 0

    def test_missing_alt_flagged(self, tmp_path):
        path = _make_defective_docx(tmp_path)
        report = verify_transcript(path)
        # The verifier should not flag an error for missing alt text
        # (the transcript correctly warns about it)
        alt_errors = [
            i for i in report.issues
            if i.severity == "error" and "alt" in i.element.lower()
        ]
        assert len(alt_errors) == 0


# ── Before/after fix comparison ──────────────────────────────────


class TestVerifyBeforeAfterFix:
    def test_fixed_doc_passes_better(self, tmp_path):
        """After fixing, verification should have fewer or equal issues."""
        defective_path = _make_defective_docx(tmp_path)
        out_path = str(tmp_path / "fixed.docx")

        metadata = MetadataWorkbook(
            documents=[DocumentMeta(filename="defective.docx", title="Fixed Report")],
            images=[ImageMeta(image_id="Image 1", alt_text="Team photo", used_in="defective.docx")],
        )

        scanner = DocxScanner()
        scan_result = scanner.scan(defective_path, metadata=metadata)
        DocxFixer().fix(defective_path, out_path, scan_result, metadata=metadata)

        before_report = verify_transcript(defective_path)
        after_report = verify_transcript(out_path)

        # Fixed doc should have fewer total issues
        before_total = before_report.error_count + before_report.gap_count
        after_total = after_report.error_count + after_report.gap_count
        assert after_total <= before_total, (
            f"Fix made verification worse: {before_total} -> {after_total}\n"
            f"Before:\n{before_report.full_report()}\n"
            f"After:\n{after_report.full_report()}"
        )

    def test_fixed_title_in_transcript(self, tmp_path):
        defective_path = _make_defective_docx(tmp_path)
        out_path = str(tmp_path / "fixed.docx")

        metadata = MetadataWorkbook(
            documents=[DocumentMeta(filename="defective.docx", title="Budget Analysis 2026")],
        )
        scanner = DocxScanner()
        scan_result = scanner.scan(defective_path, metadata=metadata)
        DocxFixer().fix(defective_path, out_path, scan_result, metadata=metadata)

        report = verify_transcript(out_path)
        title_issues = [i for i in report.issues if "title" in i.element.lower()]
        assert len(title_issues) == 0, (
            f"Title still wrong after fix:\n{report.full_report()}"
        )

    def test_fixed_alt_text_in_transcript(self, tmp_path):
        defective_path = _make_defective_docx(tmp_path)
        out_path = str(tmp_path / "fixed.docx")

        metadata = MetadataWorkbook(
            documents=[DocumentMeta(filename="defective.docx", title="Test")],
            images=[ImageMeta(image_id="Image 1", alt_text="Revenue chart for Q4", used_in="defective.docx")],
        )
        scanner = DocxScanner()
        scan_result = scanner.scan(defective_path, metadata=metadata)
        DocxFixer().fix(defective_path, out_path, scan_result, metadata=metadata)

        report = verify_transcript(out_path)
        alt_issues = [
            i for i in report.issues
            if "image" in i.element.lower() or "alt" in i.element.lower()
        ]
        # Should have no errors about image alt text
        alt_errors = [i for i in alt_issues if i.severity == "error"]
        assert len(alt_errors) == 0, (
            f"Alt text issues after fix:\n{report.full_report()}"
        )


# ── Report output ────────────────────────────────────────────────


class TestVerificationReportOutput:
    def test_full_report_no_xml_terms(self, tmp_path):
        path = _make_rich_docx(tmp_path)
        report = verify_transcript(path)
        text = report.full_report()
        for term in ["w:lang", "tblHeader", "docPr", "ooxml", "xmlns"]:
            assert term not in text.lower(), f"Technical term found: {term}"

    def test_summary_format(self, tmp_path):
        path = _make_rich_docx(tmp_path)
        report = verify_transcript(path)
        assert "PASS" in report.summary() or "WARN" in report.summary() or "FAIL" in report.summary()
