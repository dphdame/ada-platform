"""DOCX accessibility scanner."""

from __future__ import annotations

import hashlib
import os
import re
from typing import Optional

from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE as RT
from lxml import etree

from ada_platform.models.issues import (
    FixCategory,
    Issue,
    IssueType,
    WEAK_LINK_PATTERNS,
)
from ada_platform.models.metadata import MetadataWorkbook
from ada_platform.models.scan_result import ScanResult
from ada_platform.scanner.base import BaseScanner

# OOXML namespaces
_NSMAP = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "wp": "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "pic": "http://schemas.openxmlformats.org/drawingml/2006/picture",
    "mc": "http://schemas.openxmlformats.org/markup-compatibility/2006",
    "wps": "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
}

# URL-like pattern for detecting bare URLs used as link text
_URL_PATTERN = re.compile(r"^https?://", re.IGNORECASE)

# File path pattern for detecting file paths in alt text
_FILEPATH_PATTERN = re.compile(
    r"[\\/]|\.(?:png|jpg|jpeg|gif|bmp|tiff|svg|webp|emf|wmf)$",
    re.IGNORECASE,
)


def _qn(tag: str) -> str:
    """Expand a namespace-prefixed tag name to Clark notation."""
    prefix, local = tag.split(":")
    return f"{{{_NSMAP[prefix]}}}{local}"


class DocxScanner(BaseScanner):
    """Scan DOCX files for accessibility issues."""

    def scan(
        self,
        file_path: str,
        metadata: Optional[MetadataWorkbook] = None,
    ) -> ScanResult:
        doc = Document(file_path)
        filename = os.path.basename(file_path)

        issues: list[Issue] = []
        issues.extend(self._check_title(doc, filename, metadata))
        issues.extend(self._check_language(doc))

        heading_issues, heading_count = self._check_headings(doc)
        issues.extend(heading_issues)

        alt_issues, image_count = self._check_alt_text(doc, filename, metadata)
        issues.extend(alt_issues)

        table_issues, table_count = self._check_tables(doc, metadata)
        issues.extend(table_issues)

        link_issues, link_count = self._check_links(doc, metadata)
        issues.extend(link_issues)

        return ScanResult(
            file_path=file_path,
            file_type="docx",
            issues=issues,
            total_images=image_count,
            total_tables=table_count,
            total_links=link_count,
            total_headings=heading_count,
        )

    # ── Check 1: Document title ──────────────────────────────────

    def _check_title(
        self,
        doc: Document,
        filename: str,
        metadata: Optional[MetadataWorkbook],
    ) -> list[Issue]:
        title = (doc.core_properties.title or "").strip()
        if title:
            return []

        suggested = None
        fix_cat = FixCategory.SAFE_AUTO

        if metadata:
            doc_meta = metadata.get_document_meta(filename)
            if doc_meta and doc_meta.title:
                suggested = doc_meta.title

        if not suggested:
            # Derive from filename: remove extension, replace separators
            base = os.path.splitext(filename)[0]
            suggested = base.replace("_", " ").replace("-", " ").title()

        return [
            Issue.create(
                IssueType.MISSING_TITLE,
                location="Document properties",
                description="This document has no title set in its properties.",
                fix_category=fix_cat,
                suggested_fix=suggested,
            )
        ]

    # ── Check 2: Document language ───────────────────────────────

    def _check_language(self, doc: Document) -> list[Issue]:
        # Check w:settings for w:themeFontLang or document-level w:lang
        settings = doc.settings.element if doc.settings else None
        if settings is not None:
            # Look for any lang attribute in settings
            lang_els = settings.findall(f".//{_qn('w:lang')}")
            for el in lang_els:
                val = el.get(_qn("w:val")) or el.get("val", "")
                if val:
                    return []
            # Also check themeFontLang
            theme_els = settings.findall(f".//{_qn('w:themeFontLang')}")
            for el in theme_els:
                if el.get(_qn("w:val")) or el.get("val", ""):
                    return []

        # Check styles for default paragraph lang
        body = doc.element
        lang_els = body.findall(f".//{_qn('w:lang')}")
        for el in lang_els:
            for attr in [_qn("w:val"), "val", _qn("w:bidi"), _qn("w:eastAsia")]:
                if el.get(attr, ""):
                    return []

        return [
            Issue.create(
                IssueType.MISSING_LANGUAGE,
                location="Document settings",
                description="No language is set for this document.",
                suggested_fix="en-US",
            )
        ]

    # ── Check 3: Heading structure ───────────────────────────────

    def _check_headings(self, doc: Document) -> tuple[list[Issue], int]:
        issues: list[Issue] = []
        heading_count = 0
        last_level = 0

        for i, para in enumerate(doc.paragraphs, start=1):
            style_name = para.style.name if para.style else ""

            # Check for actual heading styles
            if style_name.startswith("Heading"):
                heading_count += 1
                try:
                    level = int(style_name.replace("Heading", "").strip())
                except ValueError:
                    continue

                if last_level > 0 and level > last_level + 1:
                    issues.append(
                        Issue.create(
                            IssueType.HEADING_SKIP,
                            location=f"Paragraph {i}",
                            description=(
                                f"Heading level jumps from {last_level} to "
                                f"{level}. Screen readers expect consecutive levels."
                            ),
                            current_value=f"Heading {level}",
                            suggested_fix=f"Heading {last_level + 1}",
                        )
                    )
                last_level = level
                continue

            # Check for bold-as-heading pattern:
            # entire paragraph is bold, short, Normal style, not empty
            if style_name in ("Normal", "Body Text", ""):
                text = para.text.strip()
                if not text or len(text) > 100:
                    continue
                if not para.runs:
                    continue
                all_bold = all(run.bold for run in para.runs if run.text.strip())
                if all_bold and para.runs:
                    issues.append(
                        Issue.create(
                            IssueType.HEADING_BOLD_NOT_STYLE,
                            location=f"Paragraph {i}",
                            description=(
                                f'"{text}" appears to be a heading but uses '
                                f"bold formatting instead of a heading style."
                            ),
                            current_value=text,
                        )
                    )

        return issues, heading_count

    # ── Check 4: Image alt text ──────────────────────────────────

    def _check_alt_text(
        self,
        doc: Document,
        filename: str,
        metadata: Optional[MetadataWorkbook],
    ) -> tuple[list[Issue], int]:
        issues: list[Issue] = []
        image_count = 0
        body = doc.element

        # Find all drawing and picture elements
        drawings = body.findall(f".//{_qn('w:drawing')}")
        for drawing in drawings:
            image_count += 1
            position = image_count

            # Get alt text from docPr element
            doc_pr_els = drawing.findall(
                f".//{{{_NSMAP['wp']}}}docPr"
            )
            alt_text = ""
            image_name = ""
            for dp in doc_pr_els:
                alt_text = dp.get("descr", "").strip()
                image_name = dp.get("name", "").strip()
                break

            # Try to get the embedded image filename from blipFill
            embedded_name = ""
            blips = drawing.findall(
                f".//{{{_NSMAP['a']}}}blip"
            )
            for blip in blips:
                embed_id = blip.get(f"{{{_NSMAP['r']}}}embed", "")
                if embed_id:
                    try:
                        rel = doc.part.rels[embed_id]
                        embedded_name = os.path.basename(rel.target_ref)
                    except (KeyError, AttributeError):
                        pass
                break

            # Compute content hash for three-tier matching
            content_hash = ""
            if embedded_name:
                try:
                    for rel in doc.part.rels.values():
                        if (
                            hasattr(rel, "target_ref")
                            and os.path.basename(rel.target_ref) == embedded_name
                        ):
                            blob = rel.target_part.blob
                            content_hash = hashlib.sha256(blob).hexdigest()
                            break
                except (AttributeError, KeyError):
                    pass

            # Check if alt text is missing or is a file path
            if not alt_text:
                # Try metadata three-tier lookup
                meta_img = None
                if metadata:
                    meta_img = metadata.resolve_image(
                        filename=embedded_name or image_name,
                        content_hash=content_hash,
                        position=position,
                        document=filename,
                    )

                if meta_img and meta_img.decorative:
                    continue  # Decorative images don't need alt text
                elif meta_img and meta_img.alt_text:
                    issues.append(
                        Issue.create(
                            IssueType.MISSING_ALT_TEXT,
                            location=f"Image {position}",
                            description=(
                                f"Image {position} needs a description."
                            ),
                            fix_category=FixCategory.SAFE_AUTO,
                            suggested_fix=meta_img.alt_text,
                            metadata_key=meta_img.image_id,
                        )
                    )
                else:
                    issues.append(
                        Issue.create(
                            IssueType.MISSING_ALT_TEXT,
                            location=f"Image {position}",
                            description=(
                                f"Image {position} needs a description."
                            ),
                        )
                    )

            elif _FILEPATH_PATTERN.search(alt_text):
                # Alt text looks like a file path
                meta_img = None
                if metadata:
                    meta_img = metadata.resolve_image(
                        filename=embedded_name or image_name,
                        content_hash=content_hash,
                        position=position,
                        document=filename,
                    )

                suggested = meta_img.alt_text if meta_img and meta_img.alt_text else None
                fix_cat = (
                    FixCategory.SAFE_AUTO
                    if suggested
                    else FixCategory.HUMAN_REVIEW
                )

                issues.append(
                    Issue.create(
                        IssueType.FILEPATH_ALT_TEXT,
                        location=f"Image {position}",
                        description=(
                            f"Image {position} has a file path as its "
                            f"description instead of meaningful text."
                        ),
                        fix_category=fix_cat,
                        current_value=alt_text,
                        suggested_fix=suggested,
                    )
                )

        return issues, image_count

    # ── Check 5: Table headers ───────────────────────────────────

    def _check_tables(
        self,
        doc: Document,
        metadata: Optional[MetadataWorkbook],
    ) -> tuple[list[Issue], int]:
        issues: list[Issue] = []
        table_count = 0

        for table in doc.tables:
            table_count += 1
            position = table_count

            # Check for tblHeader on first row
            tbl_xml = table._tbl
            rows = tbl_xml.findall(_qn("w:tr"))
            if not rows:
                continue

            first_row = rows[0]
            tr_pr = first_row.find(_qn("w:trPr"))
            has_header = False
            if tr_pr is not None:
                tbl_header = tr_pr.find(_qn("w:tblHeader"))
                if tbl_header is not None:
                    has_header = True

            # Check metadata for complex table flag
            table_id = f"Table {position}"
            meta_table = metadata.get_table_info(table_id) if metadata else None

            if meta_table and meta_table.complex:
                if not has_header:
                    issues.append(
                        Issue.create(
                            IssueType.TABLE_COMPLEX,
                            location=f"Table {position}",
                            description=(
                                f"Table {position} has a complex structure "
                                f"and needs manual review for accessibility."
                            ),
                        )
                    )
            elif not has_header:
                issues.append(
                    Issue.create(
                        IssueType.TABLE_NO_HEADER,
                        location=f"Table {position}",
                        description=(
                            f"Table {position} does not mark its first "
                            f"row as a header row."
                        ),
                    )
                )

        return issues, table_count

    # ── Check 6: Hyperlink text ──────────────────────────────────

    def _check_links(
        self,
        doc: Document,
        metadata: Optional[MetadataWorkbook],
    ) -> tuple[list[Issue], int]:
        issues: list[Issue] = []
        link_count = 0
        body = doc.element

        hyperlinks = body.findall(f".//{_qn('w:hyperlink')}")
        for hl in hyperlinks:
            link_count += 1
            position = link_count

            # Get the display text
            runs = hl.findall(f".//{_qn('w:t')}")
            display_text = "".join(
                r.text for r in runs if r.text
            ).strip()

            if not display_text:
                continue

            # Get the URL
            r_id = hl.get(_qn("r:id"), "")
            url = ""
            if r_id:
                try:
                    rel = doc.part.rels[r_id]
                    url = rel.target_ref
                except (KeyError, AttributeError):
                    pass

            # Check for weak link text
            is_weak = False
            lower_text = display_text.lower().strip()

            if lower_text in WEAK_LINK_PATTERNS:
                is_weak = True
            elif _URL_PATTERN.match(display_text):
                is_weak = True

            if not is_weak:
                continue

            # Check metadata for approved text
            suggested = None
            fix_cat = FixCategory.HUMAN_REVIEW
            if metadata and url:
                link_meta = metadata.get_link_text(url)
                if link_meta and link_meta.approved_text:
                    suggested = link_meta.approved_text
                    fix_cat = FixCategory.SUGGEST

            issues.append(
                Issue.create(
                    IssueType.WEAK_LINK_TEXT,
                    location=f"Link {position}",
                    description=(
                        f'Link {position} uses "{display_text}" as its text. '
                        f"Descriptive link text helps screen reader users "
                        f"understand where the link goes."
                    ),
                    fix_category=fix_cat,
                    current_value=display_text,
                    suggested_fix=suggested,
                )
            )

        return issues, link_count
