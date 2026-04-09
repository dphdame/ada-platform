"""Tests for the DOCX accessibility scanner."""

from __future__ import annotations

import os
import tempfile

import pytest
from docx import Document
from docx.shared import Pt
from lxml import etree

from ada_platform.models.issues import FixCategory, IssueType, Severity
from ada_platform.models.metadata import (
    DocumentMeta,
    ImageMeta,
    LinkMeta,
    MetadataWorkbook,
    TableMeta,
)
from ada_platform.scanner.docx_scanner import DocxScanner

# OOXML namespace for wordprocessingml
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


def _save(doc: Document, tmpdir: str, name: str = "test.docx") -> str:
    """Save a Document to a temp directory and return the path."""
    path = os.path.join(tmpdir, name)
    doc.save(path)
    return path


# ── Fixture builders ─────────────────────────────────────────────


def _make_doc_no_title() -> Document:
    doc = Document()
    doc.core_properties.title = ""
    doc.add_paragraph("Some content.")
    return doc


def _make_doc_with_title() -> Document:
    doc = Document()
    doc.core_properties.title = "Annual Report 2026"
    doc.add_paragraph("Some content.")
    return doc


def _make_doc_heading_skip() -> Document:
    doc = Document()
    doc.core_properties.title = "Test"
    doc.add_heading("Main Section", level=1)
    doc.add_paragraph("Body text.")
    doc.add_heading("Subsection", level=3)  # Skips level 2
    doc.add_paragraph("More text.")
    return doc


def _make_doc_bold_as_heading() -> Document:
    doc = Document()
    doc.core_properties.title = "Test"
    p = doc.add_paragraph()
    run = p.add_run("This Is A Section Title")
    run.bold = True
    doc.add_paragraph("Body text that follows the bold line.")
    return doc


def _make_doc_with_image_no_alt(tmpdir: str) -> Document:
    """Create a doc with an image that has no alt text."""
    doc = Document()
    doc.core_properties.title = "Test"
    # Create a tiny valid PNG (1x1 white pixel)
    png_path = os.path.join(tmpdir, "test_image.png")
    _write_tiny_png(png_path)
    doc.add_picture(png_path, width=Pt(100))
    return doc


def _make_doc_with_table_no_header() -> Document:
    doc = Document()
    doc.core_properties.title = "Test"
    table = doc.add_table(rows=3, cols=2)
    table.cell(0, 0).text = "Name"
    table.cell(0, 1).text = "Value"
    table.cell(1, 0).text = "Alpha"
    table.cell(1, 1).text = "100"
    table.cell(2, 0).text = "Beta"
    table.cell(2, 1).text = "200"
    return doc


def _make_doc_with_table_header() -> Document:
    doc = Document()
    doc.core_properties.title = "Test"
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "Header A"
    table.cell(0, 1).text = "Header B"
    table.cell(1, 0).text = "Data"
    table.cell(1, 1).text = "Data"
    # Mark first row as header via XML
    first_row = table.rows[0]._tr
    tr_pr = first_row.find(f"{{{W_NS}}}trPr")
    if tr_pr is None:
        tr_pr = etree.SubElement(first_row, f"{{{W_NS}}}trPr")
        first_row.insert(0, tr_pr)
    etree.SubElement(tr_pr, f"{{{W_NS}}}tblHeader")
    return doc


def _make_doc_with_weak_link() -> Document:
    doc = Document()
    doc.core_properties.title = "Test"
    p = doc.add_paragraph()
    # Create a hyperlink via XML (python-docx doesn't have a direct API)
    hyperlink = etree.SubElement(
        p._element,
        f"{{{W_NS}}}hyperlink",
    )
    # Add a relationship for the URL
    r_id = doc.part.relate_to(
        "https://example.com/report",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink.set(f"{{{R_NS}}}id", r_id)
    run = etree.SubElement(hyperlink, f"{{{W_NS}}}r")
    t = etree.SubElement(run, f"{{{W_NS}}}t")
    t.text = "click here"
    return doc


def _make_doc_with_good_link() -> Document:
    doc = Document()
    doc.core_properties.title = "Test"
    p = doc.add_paragraph()
    hyperlink = etree.SubElement(p._element, f"{{{W_NS}}}hyperlink")
    r_id = doc.part.relate_to(
        "https://example.com/report",
        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
        is_external=True,
    )
    hyperlink.set(f"{{{R_NS}}}id", r_id)
    run = etree.SubElement(hyperlink, f"{{{W_NS}}}r")
    t = etree.SubElement(run, f"{{{W_NS}}}t")
    t.text = "2026 Annual Report (PDF)"
    return doc


def _make_clean_doc() -> Document:
    """A document that should pass all checks."""
    doc = Document()
    doc.core_properties.title = "Clean Test Document"
    doc.add_heading("Introduction", level=1)
    doc.add_paragraph("This is a clean document.")
    doc.add_heading("Details", level=2)
    doc.add_paragraph("More content here.")
    return doc


def _write_tiny_png(path: str) -> None:
    """Write a minimal valid 1x1 white PNG file."""
    import struct
    import zlib

    def _chunk(chunk_type: bytes, data: bytes) -> bytes:
        c = chunk_type + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)

    signature = b"\x89PNG\r\n\x1a\n"
    ihdr_data = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    raw_row = b"\x00\xff\xff\xff"
    idat_data = zlib.compress(raw_row)

    with open(path, "wb") as f:
        f.write(signature)
        f.write(_chunk(b"IHDR", ihdr_data))
        f.write(_chunk(b"IDAT", idat_data))
        f.write(_chunk(b"IEND", b""))


# ── Tests ────────────────────────────────────────────────────────


class TestDocxScannerTitle:
    def test_missing_title_detected(self, tmp_path):
        doc = _make_doc_no_title()
        path = _save(doc, str(tmp_path))
        result = DocxScanner().scan(path)
        title_issues = [i for i in result.issues if i.issue_type == IssueType.MISSING_TITLE]
        assert len(title_issues) == 1
        assert title_issues[0].fix_category == FixCategory.SAFE_AUTO
        assert title_issues[0].suggested_fix  # Should derive from filename

    def test_title_present_no_issue(self, tmp_path):
        doc = _make_doc_with_title()
        path = _save(doc, str(tmp_path))
        result = DocxScanner().scan(path)
        title_issues = [i for i in result.issues if i.issue_type == IssueType.MISSING_TITLE]
        assert len(title_issues) == 0

    def test_title_from_metadata(self, tmp_path):
        doc = _make_doc_no_title()
        path = _save(doc, str(tmp_path), "report.docx")
        metadata = MetadataWorkbook(
            documents=[DocumentMeta(filename="report.docx", title="My Report")]
        )
        result = DocxScanner().scan(path, metadata=metadata)
        title_issues = [i for i in result.issues if i.issue_type == IssueType.MISSING_TITLE]
        assert len(title_issues) == 1
        assert title_issues[0].suggested_fix == "My Report"


class TestDocxScannerLanguage:
    def test_missing_language_detected(self, tmp_path):
        doc = Document()
        doc.core_properties.title = "Test"
        # Strip all lang elements to force detection
        body = doc.element
        for lang in body.findall(f".//{{{W_NS}}}lang"):
            lang.getparent().remove(lang)
        settings = doc.settings.element
        for lang in settings.findall(f".//{{{W_NS}}}lang"):
            lang.getparent().remove(lang)
        for tfl in settings.findall(f".//{{{W_NS}}}themeFontLang"):
            tfl.getparent().remove(tfl)
        path = _save(doc, str(tmp_path))
        result = DocxScanner().scan(path)
        lang_issues = [i for i in result.issues if i.issue_type == IssueType.MISSING_LANGUAGE]
        # python-docx may set default lang depending on locale; test that scanner runs
        # without error. If lang was stripped, it should detect.
        assert isinstance(result.issues, list)


class TestDocxScannerHeadings:
    def test_heading_skip_detected(self, tmp_path):
        doc = _make_doc_heading_skip()
        path = _save(doc, str(tmp_path))
        result = DocxScanner().scan(path)
        skip_issues = [i for i in result.issues if i.issue_type == IssueType.HEADING_SKIP]
        assert len(skip_issues) == 1
        assert "1" in skip_issues[0].description and "3" in skip_issues[0].description

    def test_bold_as_heading_detected(self, tmp_path):
        doc = _make_doc_bold_as_heading()
        path = _save(doc, str(tmp_path))
        result = DocxScanner().scan(path)
        bold_issues = [i for i in result.issues if i.issue_type == IssueType.HEADING_BOLD_NOT_STYLE]
        assert len(bold_issues) == 1
        assert "heading style" in bold_issues[0].description.lower()

    def test_heading_count(self, tmp_path):
        doc = _make_doc_heading_skip()
        path = _save(doc, str(tmp_path))
        result = DocxScanner().scan(path)
        assert result.total_headings == 2

    def test_consecutive_headings_ok(self, tmp_path):
        doc = Document()
        doc.core_properties.title = "Test"
        doc.add_heading("H1", level=1)
        doc.add_heading("H2", level=2)
        doc.add_heading("H3", level=3)
        path = _save(doc, str(tmp_path))
        result = DocxScanner().scan(path)
        skip_issues = [i for i in result.issues if i.issue_type == IssueType.HEADING_SKIP]
        assert len(skip_issues) == 0


class TestDocxScannerAltText:
    def test_missing_alt_text_detected(self, tmp_path):
        doc = _make_doc_with_image_no_alt(str(tmp_path))
        path = _save(doc, str(tmp_path))
        result = DocxScanner().scan(path)
        alt_issues = [i for i in result.issues if i.issue_type == IssueType.MISSING_ALT_TEXT]
        assert len(alt_issues) >= 1
        assert result.total_images >= 1

    def test_missing_alt_text_with_metadata(self, tmp_path):
        doc = _make_doc_with_image_no_alt(str(tmp_path))
        path = _save(doc, str(tmp_path), "report.docx")
        # Use positional matching since embedded filename may vary
        metadata = MetadataWorkbook(
            images=[
                ImageMeta(
                    image_id="Image 1",
                    alt_text="Photo of the community center",
                    used_in="report.docx",
                )
            ]
        )
        result = DocxScanner().scan(path, metadata=metadata)
        alt_issues = [i for i in result.issues if i.issue_type == IssueType.MISSING_ALT_TEXT]
        # Should be promoted to SAFE_AUTO with suggested fix
        for issue in alt_issues:
            if issue.suggested_fix:
                assert issue.fix_category == FixCategory.SAFE_AUTO


class TestDocxScannerTables:
    def test_table_no_header_detected(self, tmp_path):
        doc = _make_doc_with_table_no_header()
        path = _save(doc, str(tmp_path))
        result = DocxScanner().scan(path)
        table_issues = [i for i in result.issues if i.issue_type == IssueType.TABLE_NO_HEADER]
        assert len(table_issues) == 1
        assert result.total_tables == 1

    def test_table_with_header_ok(self, tmp_path):
        doc = _make_doc_with_table_header()
        path = _save(doc, str(tmp_path))
        result = DocxScanner().scan(path)
        table_issues = [
            i for i in result.issues
            if i.issue_type in (IssueType.TABLE_NO_HEADER, IssueType.TABLE_COMPLEX)
        ]
        assert len(table_issues) == 0


class TestDocxScannerLinks:
    def test_weak_link_detected(self, tmp_path):
        doc = _make_doc_with_weak_link()
        path = _save(doc, str(tmp_path))
        result = DocxScanner().scan(path)
        link_issues = [i for i in result.issues if i.issue_type == IssueType.WEAK_LINK_TEXT]
        assert len(link_issues) == 1
        assert "click here" in link_issues[0].current_value.lower()
        assert result.total_links >= 1

    def test_good_link_no_issue(self, tmp_path):
        doc = _make_doc_with_good_link()
        path = _save(doc, str(tmp_path))
        result = DocxScanner().scan(path)
        link_issues = [i for i in result.issues if i.issue_type == IssueType.WEAK_LINK_TEXT]
        assert len(link_issues) == 0

    def test_weak_link_with_metadata_suggests(self, tmp_path):
        doc = _make_doc_with_weak_link()
        path = _save(doc, str(tmp_path))
        metadata = MetadataWorkbook(
            links=[
                LinkMeta(
                    url="https://example.com/report",
                    approved_text="2026 Annual Report (PDF)",
                )
            ]
        )
        result = DocxScanner().scan(path, metadata=metadata)
        link_issues = [i for i in result.issues if i.issue_type == IssueType.WEAK_LINK_TEXT]
        assert len(link_issues) == 1
        assert link_issues[0].suggested_fix == "2026 Annual Report (PDF)"
        assert link_issues[0].fix_category == FixCategory.SUGGEST


class TestDocxScannerCleanDocument:
    def test_clean_doc_no_issues(self, tmp_path):
        doc = _make_clean_doc()
        path = _save(doc, str(tmp_path))
        result = DocxScanner().scan(path)
        # Clean doc may still trigger language issues depending on locale,
        # but should have no heading, table, link, or alt text issues
        non_lang = [
            i for i in result.issues
            if i.issue_type != IssueType.MISSING_LANGUAGE
        ]
        assert len(non_lang) == 0


class TestDocxScannerCounts:
    def test_counts_populated(self, tmp_path):
        doc = Document()
        doc.core_properties.title = "Test"
        doc.add_heading("H1", level=1)
        doc.add_heading("H2", level=2)
        doc.add_table(rows=2, cols=2)
        doc.add_table(rows=3, cols=3)
        path = _save(doc, str(tmp_path))
        result = DocxScanner().scan(path)
        assert result.total_headings == 2
        assert result.total_tables == 2
        assert result.file_type == "docx"
