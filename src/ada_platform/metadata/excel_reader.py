"""Read an Excel metadata workbook into MetadataWorkbook models."""

from __future__ import annotations

from typing import Optional

from openpyxl import Workbook, load_workbook

from ada_platform.models.metadata import (
    ChartMeta,
    DocumentMeta,
    ImageMeta,
    LinkMeta,
    MetadataWorkbook,
    SlideMeta,
    TableMeta,
)

# Expected sheet names and their required/optional columns
SHEET_SPECS: dict[str, list[str]] = {
    "Documents": ["filename", "title", "language", "business_unit", "template_type", "audience"],
    "Images": ["image_id", "content_hash", "alt_text", "long_description", "decorative", "used_in"],
    "Charts": [
        "chart_id", "chart_type", "title", "data_source",
        "x_axis", "y_axis", "key_takeaway", "short_alt", "long_description",
    ],
    "Links": ["url", "approved_text", "content_type"],
    "Tables": ["table_id", "has_header_row", "complex", "notes"],
    "Slides": ["slide_number", "title", "object_descriptions", "reading_order_notes", "document"],
}


def _str(val: object) -> str:
    """Convert a cell value to string, handling None."""
    if val is None:
        return ""
    return str(val).strip()


def _bool(val: object) -> bool:
    """Convert a cell value to bool."""
    if val is None:
        return False
    if isinstance(val, bool):
        return val
    s = str(val).strip().lower()
    return s in ("true", "yes", "1", "x")


def _int(val: object, default: int = 0) -> int:
    """Convert a cell value to int."""
    if val is None:
        return default
    try:
        return int(val)
    except (ValueError, TypeError):
        return default


def _read_header_map(ws) -> dict[str, int]:
    """Read the first row as column name -> column index mapping."""
    header_map: dict[str, int] = {}
    for col_idx, cell in enumerate(ws[1]):
        if cell.value:
            header_map[str(cell.value).strip().lower()] = col_idx
    return header_map


def _get(row, header_map: dict[str, int], col_name: str) -> object:
    """Get a cell value by column name from a row."""
    idx = header_map.get(col_name.lower())
    if idx is None or idx >= len(row):
        return None
    return row[idx].value


class ExcelMetadataReader:
    """Read a metadata workbook and return a MetadataWorkbook."""

    def read(self, file_path: str) -> MetadataWorkbook:
        """Read the Excel workbook at file_path."""
        wb = load_workbook(file_path, read_only=True, data_only=True)
        result = MetadataWorkbook()

        sheet_names = {s.lower(): s for s in wb.sheetnames}

        if "documents" in sheet_names:
            result.documents = self._read_documents(wb[sheet_names["documents"]])
        if "images" in sheet_names:
            result.images = self._read_images(wb[sheet_names["images"]])
        if "charts" in sheet_names:
            result.charts = self._read_charts(wb[sheet_names["charts"]])
        if "links" in sheet_names:
            result.links = self._read_links(wb[sheet_names["links"]])
        if "tables" in sheet_names:
            result.tables = self._read_tables(wb[sheet_names["tables"]])
        if "slides" in sheet_names:
            result.slides = self._read_slides(wb[sheet_names["slides"]])

        wb.close()
        return result

    def _read_documents(self, ws) -> list[DocumentMeta]:
        hm = _read_header_map(ws)
        items = []
        for row in ws.iter_rows(min_row=2):
            filename = _str(_get(row, hm, "filename"))
            if not filename:
                continue
            items.append(DocumentMeta(
                filename=filename,
                title=_str(_get(row, hm, "title")),
                language=_str(_get(row, hm, "language")) or "en-US",
                business_unit=_str(_get(row, hm, "business_unit")),
                template_type=_str(_get(row, hm, "template_type")),
                audience=_str(_get(row, hm, "audience")),
            ))
        return items

    def _read_images(self, ws) -> list[ImageMeta]:
        hm = _read_header_map(ws)
        items = []
        for row in ws.iter_rows(min_row=2):
            image_id = _str(_get(row, hm, "image_id"))
            if not image_id:
                continue
            items.append(ImageMeta(
                image_id=image_id,
                content_hash=_str(_get(row, hm, "content_hash")),
                alt_text=_str(_get(row, hm, "alt_text")),
                long_description=_str(_get(row, hm, "long_description")),
                decorative=_bool(_get(row, hm, "decorative")),
                used_in=_str(_get(row, hm, "used_in")),
            ))
        return items

    def _read_charts(self, ws) -> list[ChartMeta]:
        hm = _read_header_map(ws)
        items = []
        for row in ws.iter_rows(min_row=2):
            chart_id = _str(_get(row, hm, "chart_id"))
            if not chart_id:
                continue
            items.append(ChartMeta(
                chart_id=chart_id,
                chart_type=_str(_get(row, hm, "chart_type")),
                title=_str(_get(row, hm, "title")),
                data_source=_str(_get(row, hm, "data_source")),
                x_axis=_str(_get(row, hm, "x_axis")),
                y_axis=_str(_get(row, hm, "y_axis")),
                key_takeaway=_str(_get(row, hm, "key_takeaway")),
                short_alt=_str(_get(row, hm, "short_alt")),
                long_description=_str(_get(row, hm, "long_description")),
            ))
        return items

    def _read_links(self, ws) -> list[LinkMeta]:
        hm = _read_header_map(ws)
        items = []
        for row in ws.iter_rows(min_row=2):
            url = _str(_get(row, hm, "url"))
            if not url:
                continue
            items.append(LinkMeta(
                url=url,
                approved_text=_str(_get(row, hm, "approved_text")),
                content_type=_str(_get(row, hm, "content_type")),
            ))
        return items

    def _read_tables(self, ws) -> list[TableMeta]:
        hm = _read_header_map(ws)
        items = []
        for row in ws.iter_rows(min_row=2):
            table_id = _str(_get(row, hm, "table_id"))
            if not table_id:
                continue
            items.append(TableMeta(
                table_id=table_id,
                has_header_row=_bool(_get(row, hm, "has_header_row")),
                complex=_bool(_get(row, hm, "complex")),
                notes=_str(_get(row, hm, "notes")),
            ))
        return items

    def _read_slides(self, ws) -> list[SlideMeta]:
        hm = _read_header_map(ws)
        items = []
        for row in ws.iter_rows(min_row=2):
            slide_number = _int(_get(row, hm, "slide_number"))
            if not slide_number:
                continue
            items.append(SlideMeta(
                slide_number=slide_number,
                title=_str(_get(row, hm, "title")),
                object_descriptions=_str(_get(row, hm, "object_descriptions")),
                reading_order_notes=_str(_get(row, hm, "reading_order_notes")),
                document=_str(_get(row, hm, "document")),
            ))
        return items


def create_template_workbook(output_path: str) -> str:
    """Create an empty metadata workbook template with all 6 sheets."""
    wb = Workbook()

    # Documents sheet (default first sheet)
    ws = wb.active
    ws.title = "Documents"
    ws.append(SHEET_SPECS["Documents"])

    # Images sheet
    ws = wb.create_sheet("Images")
    ws.append(SHEET_SPECS["Images"])

    # Charts sheet
    ws = wb.create_sheet("Charts")
    ws.append(SHEET_SPECS["Charts"])

    # Links sheet
    ws = wb.create_sheet("Links")
    ws.append(SHEET_SPECS["Links"])

    # Tables sheet
    ws = wb.create_sheet("Tables")
    ws.append(SHEET_SPECS["Tables"])

    # Slides sheet
    ws = wb.create_sheet("Slides")
    ws.append(SHEET_SPECS["Slides"])

    wb.save(output_path)
    return output_path
