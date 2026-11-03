"""Screen reader transcript simulator for DOCX files.

Walks the DOCX XML body in document order, producing a line-by-line
transcript approximating what NVDA/JAWS/VoiceOver would announce.

This is NOT a real screen reader. It simulates the structural
announcements to verify that:
  - Headings are announced with their level
  - Images announce their alt text (or flag missing alt)
  - Tables announce header cells before data cells
  - Hyperlinks announce their display text and destination
  - Document metadata (title, language) is present
  - Reading order follows the XML body sequence

Use this as a QI gate: if the transcript looks wrong, the document
will sound wrong to a real screen reader user.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Optional

from docx import Document
from lxml import etree

# OOXML namespaces
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A_NS = "http://schemas.openxmlformats.org/drawingml/2006/main"
MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006"
ADEC_NS = "http://schemas.microsoft.com/office/drawing/2017/decorative"


def _qn(ns: str, tag: str) -> str:
    return f"{{{ns}}}{tag}"


@dataclass
class TranscriptEntry:
    """A single announcement in the reading transcript."""

    role: str          # "document", "heading", "paragraph", "image", "table",
                       # "table_header", "table_cell", "link", "list_item", "warning"
    level: int = 0     # Heading level, table row/col, list nesting
    text: str = ""     # What would be spoken
    location: str = "" # e.g., "Paragraph 3", "Table 1, Row 1, Col 2"
    issue: str = ""    # Non-empty if something is wrong


@dataclass
class ReadingTranscript:
    """Complete reading transcript of a document."""

    file_path: str
    entries: list[TranscriptEntry] = field(default_factory=list)

    @property
    def warnings(self) -> list[TranscriptEntry]:
        """Entries that flag accessibility problems."""
        return [e for e in self.entries if e.issue]

    @property
    def text_only(self) -> str:
        """Plain text transcript as a screen reader user would hear it."""
        lines = []
        for e in self.entries:
            if e.role == "document":
                lines.append(f"[Document: {e.text}]")
            elif e.role == "heading":
                lines.append(f"[Heading level {e.level}] {e.text}")
            elif e.role == "paragraph":
                if e.text.strip():
                    lines.append(e.text)
            elif e.role == "image":
                lines.append(f"[Image] {e.text}")
            elif e.role == "table":
                lines.append(f"[Table: {e.text}]")
            elif e.role == "table_header":
                lines.append(f"  [Header] {e.text}")
            elif e.role == "table_cell":
                lines.append(f"  [Cell] {e.text}")
            elif e.role == "link":
                lines.append(f"[Link: {e.text}]")
            elif e.role == "list_item":
                lines.append(f"  {'  ' * e.level}* {e.text}")
            elif e.role == "warning":
                lines.append(f"  !! {e.text}")
        return "\n".join(lines)

    def summary(self) -> str:
        """One-line summary of transcript quality."""
        n_warn = len(self.warnings)
        if n_warn == 0:
            return "Transcript clean: no accessibility warnings."
        return f"Transcript has {n_warn} warning(s) that would affect screen reader users."


def generate_transcript(file_path: str) -> ReadingTranscript:
    """Generate a screen reader transcript for a DOCX file."""
    doc = Document(file_path)
    transcript = ReadingTranscript(file_path=file_path)

    # Document-level announcements
    _announce_document_metadata(doc, transcript)

    # Walk the body
    body = doc.element.find(_qn(W_NS, "body"))
    if body is None:
        return transcript

    para_index = 0
    table_index = 0
    image_index = 0

    for child in body:
        tag = etree.QName(child).localname

        if tag == "p":
            para_index += 1
            image_index = _process_paragraph(
                child, doc, transcript, para_index, image_index
            )

        elif tag == "tbl":
            table_index += 1
            _process_table(child, doc, transcript, table_index)

        elif tag == "sdt":
            # Structured document tag (e.g., TOC) — walk inner content
            sdt_content = child.find(_qn(W_NS, "sdtContent"))
            if sdt_content is not None:
                for sdt_child in sdt_content:
                    sdt_tag = etree.QName(sdt_child).localname
                    if sdt_tag == "p":
                        para_index += 1
                        image_index = _process_paragraph(
                            sdt_child, doc, transcript, para_index, image_index
                        )

    return transcript


def _announce_document_metadata(
    doc: Document, transcript: ReadingTranscript
) -> None:
    """Announce document title and language as a screen reader would on open."""
    title = (doc.core_properties.title or "").strip()
    if title:
        transcript.entries.append(TranscriptEntry(
            role="document",
            text=f"Title: {title}",
            location="Document properties",
        ))
    else:
        transcript.entries.append(TranscriptEntry(
            role="warning",
            text="No document title. Screen reader will announce the filename instead.",
            location="Document properties",
            issue="missing_title",
        ))

    # Check language
    has_lang = False
    settings = doc.settings.element if doc.settings else None
    if settings is not None:
        for el in settings.findall(f".//{_qn(W_NS, 'lang')}"):
            for attr_ns in [_qn(W_NS, "val"), "val"]:
                if el.get(attr_ns, ""):
                    has_lang = True
                    lang_val = el.get(attr_ns)
                    transcript.entries.append(TranscriptEntry(
                        role="document",
                        text=f"Language: {lang_val}",
                        location="Document settings",
                    ))
                    break
            if has_lang:
                break
        if not has_lang:
            for el in settings.findall(f".//{_qn(W_NS, 'themeFontLang')}"):
                val = el.get(_qn(W_NS, "val"), "")
                if val:
                    has_lang = True
                    transcript.entries.append(TranscriptEntry(
                        role="document",
                        text=f"Language: {val}",
                        location="Document settings",
                    ))
                    break

    if not has_lang:
        # Check body-level lang
        for el in doc.element.findall(f".//{_qn(W_NS, 'lang')}"):
            for attr_ns in [_qn(W_NS, "val"), "val", _qn(W_NS, "bidi"), _qn(W_NS, "eastAsia")]:
                if el.get(attr_ns, ""):
                    has_lang = True
                    break
            if has_lang:
                break

    if not has_lang:
        transcript.entries.append(TranscriptEntry(
            role="warning",
            text="No document language set. Screen reader may mispronounce text.",
            location="Document settings",
            issue="missing_language",
        ))


def _process_paragraph(
    para_el: etree._Element,
    doc: Document,
    transcript: ReadingTranscript,
    para_index: int,
    image_index: int,
) -> int:
    """Process a paragraph element. Returns updated image_index."""
    location = f"Paragraph {para_index}"

    # Check if it's a heading
    ppr = para_el.find(_qn(W_NS, "pPr"))
    style_val = ""
    if ppr is not None:
        p_style = ppr.find(_qn(W_NS, "pStyle"))
        if p_style is not None:
            style_val = p_style.get(_qn(W_NS, "val"), "")

    # Check for numbered/bulleted list
    num_pr = None
    if ppr is not None:
        num_pr = ppr.find(_qn(W_NS, "numPr"))

    # Extract full text
    text_parts: list[str] = []
    for t_el in para_el.iter(_qn(W_NS, "t")):
        if t_el.text:
            text_parts.append(t_el.text)
    full_text = "".join(text_parts).strip()

    # Process any images in this paragraph
    for drawing in para_el.findall(f".//{_qn(W_NS, 'drawing')}"):
        image_index += 1
        _process_image(drawing, doc, transcript, image_index)

    # Process any hyperlinks in this paragraph
    for hl in para_el.findall(_qn(W_NS, "hyperlink")):
        _process_hyperlink(hl, doc, transcript, para_index)

    # Announce the paragraph
    if style_val.startswith("Heading"):
        try:
            level = int(style_val.replace("Heading", "").strip())
        except ValueError:
            level = 0
        transcript.entries.append(TranscriptEntry(
            role="heading",
            level=level,
            text=full_text,
            location=location,
        ))
    elif num_pr is not None:
        ilvl = num_pr.find(_qn(W_NS, "ilvl"))
        level = int(ilvl.get(_qn(W_NS, "val"), "0")) if ilvl is not None else 0
        transcript.entries.append(TranscriptEntry(
            role="list_item",
            level=level,
            text=full_text,
            location=location,
        ))
    elif full_text:
        transcript.entries.append(TranscriptEntry(
            role="paragraph",
            text=full_text,
            location=location,
        ))

    return image_index


def _process_image(
    drawing_el: etree._Element,
    doc: Document,
    transcript: ReadingTranscript,
    image_index: int,
) -> None:
    """Process a drawing/image element."""
    location = f"Image {image_index}"

    # Get alt text and decorative flag from docPr
    alt_text = ""
    is_decorative = False
    for dp in drawing_el.findall(f".//{{{WP_NS}}}docPr"):
        alt_text = dp.get("descr", "").strip()
        for dec_el in dp.iter(f"{{{ADEC_NS}}}decorative"):
            if dec_el.get("val") in ("1", "true"):
                is_decorative = True
        break

    if is_decorative:
        return  # Screen readers skip images marked decorative in OOXML

    if alt_text:
        # Check if alt text is a file path (bad)
        if "/" in alt_text or "\\" in alt_text or alt_text.endswith(
            (".png", ".jpg", ".jpeg", ".gif", ".bmp", ".svg", ".webp", ".emf")
        ):
            transcript.entries.append(TranscriptEntry(
                role="image",
                text=f'"{alt_text}"',
                location=location,
            ))
            transcript.entries.append(TranscriptEntry(
                role="warning",
                text=f"Alt text appears to be a file path, not a description: \"{alt_text}\"",
                location=location,
                issue="filepath_alt_text",
            ))
        else:
            transcript.entries.append(TranscriptEntry(
                role="image",
                text=alt_text,
                location=location,
            ))
    else:
        transcript.entries.append(TranscriptEntry(
            role="warning",
            text="Image has no description. Screen reader will say 'image' with no context.",
            location=location,
            issue="missing_alt_text",
        ))


def _process_table(
    tbl_el: etree._Element,
    doc: Document,
    transcript: ReadingTranscript,
    table_index: int,
) -> None:
    """Process a table element, announcing headers and cells."""
    rows = tbl_el.findall(_qn(W_NS, "tr"))
    n_rows = len(rows)
    n_cols = 0
    if rows:
        n_cols = len(rows[0].findall(_qn(W_NS, "tc")))

    transcript.entries.append(TranscriptEntry(
        role="table",
        text=f"{n_rows} rows, {n_cols} columns",
        location=f"Table {table_index}",
    ))

    # Check if first row is marked as header
    first_row_is_header = False
    if rows:
        tr_pr = rows[0].find(_qn(W_NS, "trPr"))
        if tr_pr is not None:
            if tr_pr.find(_qn(W_NS, "tblHeader")) is not None:
                first_row_is_header = True

    if not first_row_is_header and n_rows > 1:
        transcript.entries.append(TranscriptEntry(
            role="warning",
            text="Table has no header row marked. Screen reader cannot associate data with column names.",
            location=f"Table {table_index}",
            issue="table_no_header",
        ))

    # Walk rows and cells
    header_texts: list[str] = []
    for row_idx, row in enumerate(rows):
        cells = row.findall(_qn(W_NS, "tc"))
        is_header_row = (row_idx == 0 and first_row_is_header)

        for col_idx, cell in enumerate(cells):
            # Extract cell text
            cell_texts = []
            for t_el in cell.iter(_qn(W_NS, "t")):
                if t_el.text:
                    cell_texts.append(t_el.text)
            cell_text = "".join(cell_texts).strip()

            cell_location = f"Table {table_index}, Row {row_idx + 1}, Col {col_idx + 1}"

            if is_header_row:
                header_texts.append(cell_text)
                transcript.entries.append(TranscriptEntry(
                    role="table_header",
                    text=cell_text,
                    location=cell_location,
                ))
            else:
                # A screen reader announces: "[header]: [value]"
                prefix = ""
                if header_texts and col_idx < len(header_texts):
                    prefix = f"{header_texts[col_idx]}: "
                transcript.entries.append(TranscriptEntry(
                    role="table_cell",
                    text=f"{prefix}{cell_text}",
                    location=cell_location,
                ))


def _process_hyperlink(
    hl_el: etree._Element,
    doc: Document,
    transcript: ReadingTranscript,
    para_index: int,
) -> None:
    """Process a hyperlink element."""
    # Get display text
    text_parts = []
    for t_el in hl_el.iter(_qn(W_NS, "t")):
        if t_el.text:
            text_parts.append(t_el.text)
    display_text = "".join(text_parts).strip()

    if not display_text:
        return

    # Get URL
    r_id = hl_el.get(_qn(R_NS, "id"), "")
    url = ""
    if r_id:
        try:
            rel = doc.part.rels[r_id]
            url = rel.target_ref
        except (KeyError, AttributeError):
            pass

    location = f"Paragraph {para_index}"

    if url:
        transcript.entries.append(TranscriptEntry(
            role="link",
            text=f"{display_text} (goes to: {url})",
            location=location,
        ))
    else:
        transcript.entries.append(TranscriptEntry(
            role="link",
            text=display_text,
            location=location,
        ))

    # Flag weak link text
    weak_patterns = {"click here", "here", "link", "read more", "more", "learn more"}
    if display_text.lower().strip() in weak_patterns:
        transcript.entries.append(TranscriptEntry(
            role="warning",
            text=f'Link text "{display_text}" is not descriptive. '
                 f"Screen reader users browsing by links will only hear this word.",
            location=location,
            issue="weak_link_text",
        ))
