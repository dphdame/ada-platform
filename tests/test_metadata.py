"""Tests for Excel metadata reader."""

from __future__ import annotations

import os

import pytest
from openpyxl import Workbook

from ada_platform.metadata.excel_reader import (
    ExcelMetadataReader,
    create_template_workbook,
)


def _create_test_workbook(path: str) -> str:
    """Create a fully populated test metadata workbook."""
    wb = Workbook()

    # Documents
    ws = wb.active
    ws.title = "Documents"
    ws.append(["filename", "title", "language", "business_unit", "template_type", "audience"])
    ws.append(["report.docx", "Annual Report 2026", "en-US", "PHD", "report", "public"])
    ws.append(["memo.docx", "Budget Memo", "en-US", "Finance", "memo", "internal"])

    # Images
    ws = wb.create_sheet("Images")
    ws.append(["image_id", "content_hash", "alt_text", "long_description", "decorative", "used_in"])
    ws.append(["logo.png", "abc123", "County seal", "Official seal of Santa Clara County", False, "report.docx"])
    ws.append(["divider.png", "def456", "", "", True, "report.docx"])
    ws.append(["Image 3", "", "Bar chart of enrollment", "", False, "report.docx"])

    # Charts
    ws = wb.create_sheet("Charts")
    ws.append(["chart_id", "chart_type", "title", "data_source", "x_axis", "y_axis", "key_takeaway", "short_alt", "long_description"])
    ws.append(["enroll_bar", "bar", "Enrollment by Year", "data/enrollment.csv", "Year", "Count", "Grew 15%", "Bar chart of enrollment growth", "Enrollment grew from 500 to 575 between 2020 and 2025."])

    # Links
    ws = wb.create_sheet("Links")
    ws.append(["url", "approved_text", "content_type"])
    ws.append(["https://example.com/report", "2026 Annual Report (PDF)", "PDF"])

    # Tables
    ws = wb.create_sheet("Tables")
    ws.append(["table_id", "has_header_row", "complex", "notes"])
    ws.append(["Table 1", True, False, ""])
    ws.append(["Table 2", True, True, "Multi-level headers"])

    # Slides
    ws = wb.create_sheet("Slides")
    ws.append(["slide_number", "title", "object_descriptions", "reading_order_notes", "document"])
    ws.append([1, "Welcome", "Logo and title", "", "deck.pptx"])
    ws.append([2, "Key Findings", "Chart and bullets", "Read chart before bullets", "deck.pptx"])

    wb.save(path)
    return path


class TestExcelMetadataReader:
    def test_reads_all_sheets(self, tmp_path):
        path = _create_test_workbook(str(tmp_path / "meta.xlsx"))
        reader = ExcelMetadataReader()
        wb = reader.read(path)

        assert len(wb.documents) == 2
        assert len(wb.images) == 3
        assert len(wb.charts) == 1
        assert len(wb.links) == 1
        assert len(wb.tables) == 2
        assert len(wb.slides) == 2

    def test_document_values(self, tmp_path):
        path = _create_test_workbook(str(tmp_path / "meta.xlsx"))
        wb = ExcelMetadataReader().read(path)
        doc = wb.get_document_meta("report.docx")
        assert doc is not None
        assert doc.title == "Annual Report 2026"
        assert doc.language == "en-US"
        assert doc.business_unit == "PHD"

    def test_image_values(self, tmp_path):
        path = _create_test_workbook(str(tmp_path / "meta.xlsx"))
        wb = ExcelMetadataReader().read(path)

        img = wb.get_image_alt("logo.png")
        assert img is not None
        assert img.alt_text == "County seal"
        assert img.content_hash == "abc123"

        divider = wb.get_image_alt("divider.png")
        assert divider is not None
        assert divider.decorative is True

    def test_image_hash_lookup(self, tmp_path):
        path = _create_test_workbook(str(tmp_path / "meta.xlsx"))
        wb = ExcelMetadataReader().read(path)
        img = wb.get_image_by_hash("abc123")
        assert img is not None
        assert img.image_id == "logo.png"

    def test_image_positional_lookup(self, tmp_path):
        path = _create_test_workbook(str(tmp_path / "meta.xlsx"))
        wb = ExcelMetadataReader().read(path)
        img = wb.get_image_by_position(3, document="report.docx")
        assert img is not None
        assert img.alt_text == "Bar chart of enrollment"

    def test_three_tier_resolve(self, tmp_path):
        path = _create_test_workbook(str(tmp_path / "meta.xlsx"))
        wb = ExcelMetadataReader().read(path)
        # Filename match
        img = wb.resolve_image(filename="logo.png")
        assert img.image_id == "logo.png"
        # Hash fallback
        img = wb.resolve_image(filename="unknown", content_hash="def456")
        assert img.image_id == "divider.png"
        # Position fallback
        img = wb.resolve_image(filename="x", content_hash="x", position=3, document="report.docx")
        assert img.image_id == "Image 3"

    def test_chart_values(self, tmp_path):
        path = _create_test_workbook(str(tmp_path / "meta.xlsx"))
        wb = ExcelMetadataReader().read(path)
        chart = wb.get_chart_info("enroll_bar")
        assert chart is not None
        assert chart.chart_type == "bar"
        assert chart.key_takeaway == "Grew 15%"

    def test_link_values(self, tmp_path):
        path = _create_test_workbook(str(tmp_path / "meta.xlsx"))
        wb = ExcelMetadataReader().read(path)
        link = wb.get_link_text("https://example.com/report")
        assert link is not None
        assert link.approved_text == "2026 Annual Report (PDF)"

    def test_table_values(self, tmp_path):
        path = _create_test_workbook(str(tmp_path / "meta.xlsx"))
        wb = ExcelMetadataReader().read(path)
        t1 = wb.get_table_info("Table 1")
        assert t1 is not None
        assert t1.has_header_row is True
        assert t1.complex is False
        t2 = wb.get_table_info("Table 2")
        assert t2.complex is True

    def test_slide_values(self, tmp_path):
        path = _create_test_workbook(str(tmp_path / "meta.xlsx"))
        wb = ExcelMetadataReader().read(path)
        s = wb.get_slide_info(1, document="deck.pptx")
        assert s is not None
        assert s.title == "Welcome"
        s2 = wb.get_slide_info(2, document="deck.pptx")
        assert s2.reading_order_notes == "Read chart before bullets"

    def test_missing_sheets_graceful(self, tmp_path):
        """A workbook with only Documents sheet should not error."""
        path = str(tmp_path / "partial.xlsx")
        wb_obj = Workbook()
        ws = wb_obj.active
        ws.title = "Documents"
        ws.append(["filename", "title"])
        ws.append(["test.docx", "Test"])
        wb_obj.save(path)

        wb = ExcelMetadataReader().read(path)
        assert len(wb.documents) == 1
        assert len(wb.images) == 0
        assert len(wb.charts) == 0

    def test_empty_rows_skipped(self, tmp_path):
        path = str(tmp_path / "empty_rows.xlsx")
        wb_obj = Workbook()
        ws = wb_obj.active
        ws.title = "Documents"
        ws.append(["filename", "title"])
        ws.append(["report.docx", "Report"])
        ws.append([None, None])  # Empty row
        ws.append(["", ""])  # Blank strings
        wb_obj.save(path)

        wb = ExcelMetadataReader().read(path)
        assert len(wb.documents) == 1


class TestCreateTemplateWorkbook:
    def test_creates_all_sheets(self, tmp_path):
        path = str(tmp_path / "template.xlsx")
        create_template_workbook(path)
        assert os.path.exists(path)

        wb = ExcelMetadataReader().read(path)
        # All sheets should exist but be empty (no data rows)
        assert len(wb.documents) == 0
        assert len(wb.images) == 0
        assert len(wb.charts) == 0
        assert len(wb.links) == 0
        assert len(wb.tables) == 0
        assert len(wb.slides) == 0

    def test_template_has_headers(self, tmp_path):
        from openpyxl import load_workbook as olw
        path = str(tmp_path / "template.xlsx")
        create_template_workbook(path)
        wb_check = olw(path)
        assert "Documents" in wb_check.sheetnames
        assert "Images" in wb_check.sheetnames
        assert "Charts" in wb_check.sheetnames
        assert "Links" in wb_check.sheetnames
        assert "Tables" in wb_check.sheetnames
        assert "Slides" in wb_check.sheetnames
        # Check Documents headers
        docs_ws = wb_check["Documents"]
        headers = [cell.value for cell in docs_ws[1]]
        assert "filename" in headers
        assert "title" in headers
        wb_check.close()
