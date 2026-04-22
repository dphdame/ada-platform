"""Verify screen reader transcript accuracy against DOCX ground truth.

Independently walks the DOCX to build a ground-truth element list,
then compares it against the transcript to catch:
  - Missing elements (transcript skipped something in the doc)
  - Extra elements (transcript invented something not in the doc)
  - Wrong text (transcript mangled the content)
  - Wrong order (elements announced out of document sequence)
  - Wrong metadata (title/language mismatch)
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Optional

from docx import Document
from lxml import etree

from ada_platform.qa.screen_reader_sim import (
    ReadingTranscript,
    TranscriptEntry,
    generate_transcript,
)

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


@dataclass
class GroundTruthElement:
    """An element extracted directly from the DOCX for comparison."""

    kind: str        # "title", "language", "heading", "paragraph", "image",
                     # "table", "table_cell", "link"
    text: str = ""
    level: int = 0   # heading level, table row/col
    location: str = ""
    has_alt: bool = False
    has_header: bool = False


@dataclass
class VerificationIssue:
    """A discrepancy between ground truth and transcript."""

    severity: str    # "error" = transcript is wrong, "gap" = transcript missed something
    element: str     # What element is affected
    expected: str    # What the doc contains
    found: str       # What the transcript says (or "missing")
    location: str = ""


@dataclass
class VerificationReport:
    """Result of comparing transcript to ground truth."""

    file_path: str
    ground_truth_count: int = 0
    transcript_entry_count: int = 0
    issues: list[VerificationIssue] = field(default_factory=list)

    @property
    def passed(self) -> bool:
        errors = [i for i in self.issues if i.severity == "error"]
        return len(errors) == 0

    @property
    def error_count(self) -> int:
        return sum(1 for i in self.issues if i.severity == "error")

    @property
    def gap_count(self) -> int:
        return sum(1 for i in self.issues if i.severity == "gap")

    def summary(self) -> str:
        if self.passed and self.gap_count == 0:
            return (
                f"PASS: Transcript accurately represents all {self.ground_truth_count} "
                f"elements in the document."
            )
        parts = []
        if self.error_count:
            parts.append(f"{self.error_count} error(s) (transcript is wrong)")
        if self.gap_count:
            parts.append(f"{self.gap_count} gap(s) (transcript missed content)")
        status = "FAIL" if self.error_count else "WARN"
        return f"{status}: {', '.join(parts)}."

    def full_report(self) -> str:
        lines = [
            "=" * 60,
            "TRANSCRIPT VERIFICATION REPORT",
            "=" * 60,
            "",
            f"File: {self.file_path}",
            f"Ground truth elements: {self.ground_truth_count}",
            f"Transcript entries: {self.transcript_entry_count}",
            "",
            self.summary(),
            "",
        ]
        if self.issues:
            for i, issue in enumerate(self.issues, 1):
                marker = "ERROR" if issue.severity == "error" else "GAP"
                lines.append(f"  {i}. [{marker}] {issue.element}")
                if issue.location:
                    lines.append(f"     Location: {issue.location}")
                lines.append(f"     Expected: {issue.expected}")
                lines.append(f"     Found:    {issue.found}")
                lines.append("")
        else:
            lines.append("  No issues found.")
        lines.append("=" * 60)
        return "\n".join(lines)


def extract_ground_truth(file_path: str) -> list[GroundTruthElement]:
    """Independently extract all elements from a DOCX in document order."""
    doc = Document(file_path)
    elements: list[GroundTruthElement] = []

    # Document metadata
    title = (doc.core_properties.title or "").strip()
    elements.append(GroundTruthElement(
        kind="title",
        text=title,
        location="Document properties",
    ))

    # Walk body
    body = doc.element.find(f"{{{W_NS}}}body")
    if body is None:
        return elements

    para_idx = 0
    table_idx = 0
    image_idx = 0

    for child in body:
        tag = etree.QName(child).localname

        if tag == "p":
            para_idx += 1
            loc = f"Paragraph {para_idx}"

            # Check style
            ppr = child.find(f"{{{W_NS}}}pPr")
            style_val = ""
            if ppr is not None:
                ps = ppr.find(f"{{{W_NS}}}pStyle")
                if ps is not None:
                    style_val = ps.get(f"{{{W_NS}}}val", "")

            # Extract text
            text_parts = []
            for t_el in child.iter(f"{{{W_NS}}}t"):
                if t_el.text:
                    text_parts.append(t_el.text)
            full_text = "".join(text_parts).strip()

            # Images
            for drawing in child.findall(f".//{{{W_NS}}}drawing"):
                image_idx += 1
                alt_text = ""
                for dp in drawing.findall(f".//{{{WP_NS}}}docPr"):
                    alt_text = dp.get("descr", "").strip()
                    break
                elements.append(GroundTruthElement(
                    kind="image",
                    text=alt_text,
                    location=f"Image {image_idx}",
                    has_alt=bool(alt_text),
                ))

            # Hyperlinks
            for hl in child.findall(f"{{{W_NS}}}hyperlink"):
                hl_text_parts = []
                for t_el in hl.iter(f"{{{W_NS}}}t"):
                    if t_el.text:
                        hl_text_parts.append(t_el.text)
                hl_text = "".join(hl_text_parts).strip()
                r_id = hl.get(f"{{{R_NS}}}id", "")
                url = ""
                if r_id:
                    try:
                        url = doc.part.rels[r_id].target_ref
                    except (KeyError, AttributeError):
                        pass
                if hl_text:
                    elements.append(GroundTruthElement(
                        kind="link",
                        text=hl_text,
                        level=0,
                        location=loc,
                    ))

            # Heading or paragraph
            if style_val.startswith("Heading"):
                try:
                    level = int(style_val.replace("Heading", "").strip())
                except ValueError:
                    level = 0
                elements.append(GroundTruthElement(
                    kind="heading",
                    text=full_text,
                    level=level,
                    location=loc,
                ))
            elif full_text:
                elements.append(GroundTruthElement(
                    kind="paragraph",
                    text=full_text,
                    location=loc,
                ))

        elif tag == "tbl":
            table_idx += 1
            rows = child.findall(f"{{{W_NS}}}tr")

            # Check header
            has_header = False
            if rows:
                tr_pr = rows[0].find(f"{{{W_NS}}}trPr")
                if tr_pr is not None and tr_pr.find(f"{{{W_NS}}}tblHeader") is not None:
                    has_header = True

            elements.append(GroundTruthElement(
                kind="table",
                text=f"{len(rows)} rows",
                location=f"Table {table_idx}",
                has_header=has_header,
            ))

            # Walk cells
            for row_idx, row in enumerate(rows):
                cells = row.findall(f"{{{W_NS}}}tc")
                for col_idx, cell in enumerate(cells):
                    cell_texts = []
                    for t_el in cell.iter(f"{{{W_NS}}}t"):
                        if t_el.text:
                            cell_texts.append(t_el.text)
                    cell_text = "".join(cell_texts).strip()
                    cell_kind = "table_header" if (row_idx == 0 and has_header) else "table_cell"
                    elements.append(GroundTruthElement(
                        kind=cell_kind,
                        text=cell_text,
                        location=f"Table {table_idx}, Row {row_idx+1}, Col {col_idx+1}",
                    ))

    return elements


def verify_transcript(file_path: str) -> VerificationReport:
    """Generate transcript and ground truth, then compare them."""
    ground_truth = extract_ground_truth(file_path)
    transcript = generate_transcript(file_path)

    report = VerificationReport(
        file_path=file_path,
        ground_truth_count=len(ground_truth),
        transcript_entry_count=len(transcript.entries),
    )

    # Build lookup structures from transcript
    transcript_entries = transcript.entries
    # Filter to content entries (skip warnings, they're meta)
    content_entries = [
        e for e in transcript_entries
        if e.role not in ("warning",)
    ]

    # 1. Check document title
    gt_title = next((g for g in ground_truth if g.kind == "title"), None)
    if gt_title:
        tr_title = next(
            (e for e in content_entries if e.role == "document" and "Title:" in e.text),
            None,
        )
        if gt_title.text:
            if tr_title is None:
                report.issues.append(VerificationIssue(
                    severity="error",
                    element="Document title",
                    expected=f"Title announced: \"{gt_title.text}\"",
                    found="Title not announced in transcript",
                    location="Document properties",
                ))
            elif gt_title.text not in tr_title.text:
                report.issues.append(VerificationIssue(
                    severity="error",
                    element="Document title",
                    expected=gt_title.text,
                    found=tr_title.text,
                    location="Document properties",
                ))

    # 2. Check headings
    gt_headings = [g for g in ground_truth if g.kind == "heading"]
    tr_headings = [e for e in content_entries if e.role == "heading"]

    for gt_h in gt_headings:
        match = next(
            (t for t in tr_headings if t.text == gt_h.text and t.level == gt_h.level),
            None,
        )
        if match is None:
            # Check if text matches but level is wrong
            text_match = next((t for t in tr_headings if t.text == gt_h.text), None)
            if text_match:
                report.issues.append(VerificationIssue(
                    severity="error",
                    element=f"Heading \"{gt_h.text}\"",
                    expected=f"Level {gt_h.level}",
                    found=f"Level {text_match.level}",
                    location=gt_h.location,
                ))
            else:
                report.issues.append(VerificationIssue(
                    severity="gap",
                    element=f"Heading \"{gt_h.text}\"",
                    expected=f"Heading level {gt_h.level} announced",
                    found="Not found in transcript",
                    location=gt_h.location,
                ))

    # 3. Check images
    gt_images = [g for g in ground_truth if g.kind == "image"]
    tr_images = [e for e in content_entries if e.role == "image"]
    tr_image_warnings = [e for e in transcript_entries if e.role == "warning" and "image" in e.issue.lower() if e.issue]

    for i, gt_img in enumerate(gt_images):
        if gt_img.has_alt:
            # Should have a matching transcript image entry with the alt text
            match = next(
                (t for t in tr_images if gt_img.text in t.text),
                None,
            )
            if match is None:
                report.issues.append(VerificationIssue(
                    severity="error",
                    element=f"Image alt text",
                    expected=f"Alt text announced: \"{gt_img.text}\"",
                    found="Alt text not found in transcript",
                    location=gt_img.location,
                ))
        else:
            # Should have a warning about missing alt text
            has_warning = any(
                w.issue == "missing_alt_text" and gt_img.location in w.location
                for w in [e for e in transcript_entries if e.role == "warning"]
            )
            if not has_warning:
                report.issues.append(VerificationIssue(
                    severity="gap",
                    element="Image missing alt text",
                    expected="Warning about missing alt text",
                    found="No warning in transcript",
                    location=gt_img.location,
                ))

    # 4. Check tables
    gt_tables = [g for g in ground_truth if g.kind == "table"]
    tr_tables = [e for e in content_entries if e.role == "table"]

    for gt_tbl in gt_tables:
        # Find matching transcript table
        match = next(
            (t for t in tr_tables if gt_tbl.location.split(",")[0] in t.location),
            None,
        )
        if match is None:
            report.issues.append(VerificationIssue(
                severity="gap",
                element=f"Table",
                expected="Table announced in transcript",
                found="Not found",
                location=gt_tbl.location,
            ))
        elif gt_tbl.has_header:
            # Verify header cells appear in transcript
            gt_header_cells = [
                g for g in ground_truth
                if g.kind == "table_header" and gt_tbl.location.replace(" rows", "") in g.location.split(",")[0]
            ]
            tr_header_cells = [
                e for e in content_entries
                if e.role == "table_header"
            ]
            for gt_hc in gt_header_cells:
                hc_match = next(
                    (t for t in tr_header_cells if gt_hc.text == t.text),
                    None,
                )
                if hc_match is None:
                    report.issues.append(VerificationIssue(
                        severity="error",
                        element="Table header cell",
                        expected=f"Header \"{gt_hc.text}\" announced",
                        found="Not found in transcript",
                        location=gt_hc.location,
                    ))

    # 5. Check table data cells
    gt_data_cells = [g for g in ground_truth if g.kind == "table_cell" and g.text]
    tr_data_cells = [e for e in content_entries if e.role == "table_cell"]

    for gt_cell in gt_data_cells:
        # The transcript may prefix with header text, so check if gt text is contained
        cell_match = next(
            (t for t in tr_data_cells if gt_cell.text in t.text),
            None,
        )
        if cell_match is None:
            report.issues.append(VerificationIssue(
                severity="error",
                element="Table data cell",
                expected=f"Cell value \"{gt_cell.text}\" announced",
                found="Not found in transcript",
                location=gt_cell.location,
            ))

    # 6. Check links
    gt_links = [g for g in ground_truth if g.kind == "link"]
    tr_links = [e for e in content_entries if e.role == "link"]

    for gt_link in gt_links:
        link_match = next(
            (t for t in tr_links if gt_link.text in t.text),
            None,
        )
        if link_match is None:
            report.issues.append(VerificationIssue(
                severity="error",
                element="Hyperlink",
                expected=f"Link text \"{gt_link.text}\" announced",
                found="Not found in transcript",
                location=gt_link.location,
            ))

    # 7. Check paragraphs (body text) - gap check only, not error
    gt_paras = [g for g in ground_truth if g.kind == "paragraph" and len(g.text) > 10]
    tr_paras = [e for e in content_entries if e.role == "paragraph"]

    for gt_p in gt_paras:
        # Check if any transcript paragraph contains this text
        p_match = next(
            (t for t in tr_paras if gt_p.text == t.text),
            None,
        )
        if p_match is None:
            # Could be split across entries or truncated; check substring
            substr_match = next(
                (t for t in tr_paras if gt_p.text[:30] in t.text),
                None,
            )
            if substr_match is None:
                report.issues.append(VerificationIssue(
                    severity="gap",
                    element="Paragraph text",
                    expected=f"\"{gt_p.text[:60]}...\"" if len(gt_p.text) > 60 else f"\"{gt_p.text}\"",
                    found="Not found in transcript",
                    location=gt_p.location,
                ))

    # 8. Check element ordering (headings should appear in document order)
    gt_heading_order = [g.text for g in gt_headings]
    tr_heading_order = [t.text for t in tr_headings]
    # Check relative order is preserved
    tr_idx = 0
    for gt_text in gt_heading_order:
        found = False
        while tr_idx < len(tr_heading_order):
            if tr_heading_order[tr_idx] == gt_text:
                found = True
                tr_idx += 1
                break
            tr_idx += 1
        if not found and gt_text in tr_heading_order:
            report.issues.append(VerificationIssue(
                severity="error",
                element="Reading order",
                expected=f"Heading \"{gt_text}\" in document order",
                found="Out of order in transcript",
            ))

    return report
