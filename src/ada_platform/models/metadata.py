"""Excel metadata workbook schema models."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class DocumentMeta:
    """Row from the Documents sheet."""

    filename: str
    title: str = ""
    language: str = "en-US"
    business_unit: str = ""
    template_type: str = ""
    audience: str = ""


@dataclass
class ImageMeta:
    """Row from the Images sheet."""

    image_id: str
    alt_text: str = ""
    long_description: str = ""
    decorative: bool = False
    used_in: str = ""
    content_hash: str = ""


@dataclass
class ChartMeta:
    """Row from the Charts sheet."""

    chart_id: str
    chart_type: str = ""
    title: str = ""
    data_source: str = ""
    x_axis: str = ""
    y_axis: str = ""
    key_takeaway: str = ""
    short_alt: str = ""
    long_description: str = ""


@dataclass
class LinkMeta:
    """Row from the Links sheet."""

    url: str
    approved_text: str = ""
    content_type: str = ""


@dataclass
class TableMeta:
    """Row from the Tables sheet."""

    table_id: str
    has_header_row: bool = True
    complex: bool = False
    notes: str = ""


@dataclass
class SlideMeta:
    """Row from the Slides sheet (Phase 2 readiness)."""

    slide_number: int
    title: str = ""
    object_descriptions: str = ""
    reading_order_notes: str = ""
    document: str = ""


@dataclass
class MetadataWorkbook:
    """Parsed contents of an Excel metadata workbook."""

    documents: list[DocumentMeta] = field(default_factory=list)
    images: list[ImageMeta] = field(default_factory=list)
    charts: list[ChartMeta] = field(default_factory=list)
    links: list[LinkMeta] = field(default_factory=list)
    tables: list[TableMeta] = field(default_factory=list)
    slides: list[SlideMeta] = field(default_factory=list)

    def get_document_meta(self, filename: str) -> Optional[DocumentMeta]:
        """Look up document metadata by filename."""
        for doc in self.documents:
            if doc.filename == filename:
                return doc
        return None

    def get_image_alt(self, image_id: str) -> Optional[ImageMeta]:
        """Look up image metadata by image_id (filename match)."""
        for img in self.images:
            if img.image_id == image_id:
                return img
        return None

    def get_image_by_hash(self, content_hash: str) -> Optional[ImageMeta]:
        """Look up image metadata by SHA-256 content hash."""
        if not content_hash:
            return None
        for img in self.images:
            if img.content_hash and img.content_hash == content_hash:
                return img
        return None

    def get_image_by_position(
        self, position: int, document: str = ""
    ) -> Optional[ImageMeta]:
        """Look up image metadata by position (1-based) within a document."""
        positional_id = f"Image {position}"
        for img in self.images:
            if img.image_id == positional_id:
                if not document or not img.used_in:
                    return img
                if document in img.used_in:
                    return img
        return None

    def resolve_image(
        self,
        *,
        filename: str = "",
        content_hash: str = "",
        position: int = 0,
        document: str = "",
    ) -> Optional[ImageMeta]:
        """Three-tier image lookup: filename, then content hash, then position."""
        if filename:
            result = self.get_image_alt(filename)
            if result:
                return result
        if content_hash:
            result = self.get_image_by_hash(content_hash)
            if result:
                return result
        if position > 0:
            result = self.get_image_by_position(position, document)
            if result:
                return result
        return None

    def get_chart_info(self, chart_id: str) -> Optional[ChartMeta]:
        """Look up chart metadata by chart_id."""
        for chart in self.charts:
            if chart.chart_id == chart_id:
                return chart
        return None

    def get_link_text(self, url: str) -> Optional[LinkMeta]:
        """Look up approved link text by URL."""
        for link in self.links:
            if link.url == url:
                return link
        return None

    def get_table_info(self, table_id: str) -> Optional[TableMeta]:
        """Look up table metadata by table_id."""
        for table in self.tables:
            if table.table_id == table_id:
                return table
        return None

    def get_slide_info(
        self, slide_number: int, document: str = ""
    ) -> Optional[SlideMeta]:
        """Look up slide metadata by slide number and optionally document."""
        for slide in self.slides:
            if slide.slide_number == slide_number:
                if not document or not slide.document or slide.document == document:
                    return slide
        return None
