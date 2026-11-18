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
    "adec": "http://schemas.microsoft.com/office/drawing/2017/decorative",
}

# URL-like pattern for detecting bare URLs used as link text
_URL_PATTERN = re.compile(r"^https?://", re.IGNORECASE)

# File path pattern for detecting file paths in alt text
_FILEPATH_PATTERN = re.compile(
    r"(?:^|[\s\"'])(?:[A-Za-z]:\\|/(?:Users|tmp|var|home|sites|images|assets|figures)/|(?:figures|assets|images|media|uploads)/|\S+\.(?:png|jpg|jpeg|gif|bmp|tiff|svg|webp|emf|wmf)\b)",
    re.IGNORECASE,
)

# Redundant prefix pattern in alt text (Siteimprove & County Digital Accessibility standard)
_REDUNDANT_ALT_PREFIX_RE = re.compile(
    r"^\s*(?:an?\s+|the\s+)?(?:image|picture|photo|photograph|graphic|icon|screenshot)\s+of\s+",
    re.IGNORECASE,
)

# Maximum character count for regular alt text (Siteimprove / County standard: 125 target, 141 hard max)
_MAX_REGULAR_ALT_LENGTH = 141

# Literal tokens that authors mistakenly type into alt text instead of setting empty "" + decorative
_DECORATIVE_LITERAL_TOKENS = {"decorative", "[decorative]", "decoration", "null", '""', "''"}


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

        issues.extend(self._check_alignment(doc))
        issues.extend(self._check_colors(doc))

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
        h1_count = 0

        for i, para in enumerate(doc.paragraphs, start=1):
            style_name = para.style.name if para.style else ""

            # Check for actual heading styles
            if style_name.startswith("Heading"):
                heading_count += 1
                try:
                    level = int(style_name.replace("Heading", "").strip())
                except ValueError:
                    continue

                if heading_count == 1 and level > 1:
                    issues.append(
                        Issue.create(
                            IssueType.HEADING_FIRST_NOT_H1,
                            location=f"Paragraph {i}",
                            description=(
                                f"First heading is Heading {level}. Documents should "
                                f"begin with a single Heading 1 for the main title."
                            ),
                            current_value=f"Heading {level}",
                            suggested_fix="Heading 1",
                        )
                    )

                if level == 1:
                    h1_count += 1
                    if h1_count > 1:
                        issues.append(
                            Issue.create(
                                IssueType.HEADING_MULTIPLE_H1,
                                location=f"Paragraph {i}",
                                description=(
                                    "Document contains multiple Heading 1 paragraphs. "
                                    "Use a single Heading 1 for the document title and "
                                    "Heading 2 for major sections."
                                ),
                                current_value="Heading 1",
                                suggested_fix="Heading 2",
                            )
                        )

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

            # Get alt text and decorative flag from docPr element
            doc_pr_els = drawing.findall(
                f".//{{{_NSMAP['wp']}}}docPr"
            )
            alt_text = ""
            image_name = ""
            is_decorative = False
            for dp in doc_pr_els:
                alt_text = dp.get("descr", "").strip()
                image_name = dp.get("name", "").strip()
                for dec_el in dp.iter(f"{{{_NSMAP['adec']}}}decorative"):
                    if dec_el.get("val") in ("1", "true"):
                        is_decorative = True
                break

            if is_decorative:
                if alt_text:
                    issues.append(
                        Issue.create(
                            IssueType.DECORATIVE_IMAGE_HAS_ALT,
                            location=f"Image {position}",
                            description=(
                                f'Image {position} is marked decorative but still has '
                                f'non-empty alt text ("{alt_text}"). Decorative images '
                                f'must have an empty description ("") so screen readers skip them.'
                            ),
                            current_value=alt_text,
                            suggested_fix="",
                        )
                    )
                continue  # In-place decorative images with "" are skipped by screen readers

            if alt_text.lower() in _DECORATIVE_LITERAL_TOKENS:
                issues.append(
                    Issue.create(
                        IssueType.DECORATIVE_IMAGE_HAS_ALT,
                        location=f"Image {position}",
                        description=(
                            f'Image {position} has literal "{alt_text}" as its description. '
                            f'Decorative images must have an empty description ("") and be '
                            f"marked decorative so screen readers do not read the word aloud."
                        ),
                        current_value=alt_text,
                        suggested_fix="",
                    )
                )
                continue

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
            else:
                # Check for redundant "image of" / "photo of" / "picture of" prefix
                m_pref = _REDUNDANT_ALT_PREFIX_RE.match(alt_text)
                if m_pref:
                    remainder = alt_text[m_pref.end() :].strip()
                    cleaned = (
                        remainder[0].upper() + remainder[1:]
                        if remainder
                        else alt_text
                    )
                    issues.append(
                        Issue.create(
                            IssueType.ALT_TEXT_REDUNDANT_PREFIX,
                            location=f"Image {position}",
                            description=(
                                f'Image {position} description starts with a redundant '
                                f'phrase ("{m_pref.group(0).strip()}"). Screen readers '
                                f"already announce that an element is an image."
                            ),
                            current_value=alt_text,
                            suggested_fix=cleaned,
                        )
                    )

                if len(alt_text) > _MAX_REGULAR_ALT_LENGTH:
                    issues.append(
                        Issue.create(
                            IssueType.ALT_TEXT_TOO_LONG,
                            location=f"Image {position}",
                            description=(
                                f"Image {position} description is {len(alt_text)} characters "
                                f"(recommended concise limit is under {_MAX_REGULAR_ALT_LENGTH} "
                                f"characters). Keep alt text concise and place detailed chart "
                                f"or infographic data in the surrounding text or a table."
                            ),
                            current_value=alt_text,
                        )
                    )

        return issues, image_count

    # ── Check 5: Table headers, merged cells & empty cells ───────

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

            # Check for merged cells (horizontal gridSpan > 1 or vertical vMerge)
            has_merged = False
            for gs in tbl_xml.findall(f".//{_qn('w:gridSpan')}"):
                raw_span = gs.get(_qn("w:val")) or gs.get("val", "1")
                try:
                    if int(raw_span) > 1:
                        has_merged = True
                        break
                except ValueError:
                    pass
            if not has_merged and tbl_xml.findall(f".//{_qn('w:vMerge')}"):
                has_merged = True

            if has_merged:
                issues.append(
                    Issue.create(
                        IssueType.TABLE_MERGED_CELLS,
                        location=f"Table {position}",
                        description=(
                            f"Table {position} contains merged cells. Screen readers "
                            f"lose row and column context when cells are merged; "
                            f"split merged cells into a simple flat grid."
                        ),
                    )
                )

            # Check for empty cells in populated tables (County standard: use "NA" or "no data")
            non_empty_cells = 0
            empty_cells = 0
            for tr in rows:
                for tc in tr.findall(_qn("w:tc")):
                    cell_text = "".join(
                        t.text for t in tc.findall(f".//{_qn('w:t')}") if t.text
                    ).strip()
                    if cell_text:
                        non_empty_cells += 1
                    else:
                        empty_cells += 1

            if non_empty_cells > 0 and empty_cells > 0:
                issues.append(
                    Issue.create(
                        IssueType.TABLE_EMPTY_CELLS,
                        location=f"Table {position}",
                        description=(
                            f"Table {position} has {empty_cells} blank cell(s). "
                            f"Enter 'NA' or 'no data' in empty cells so screen reader "
                            f"users know no value was accidentally skipped."
                        ),
                        suggested_fix="NA",
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

    # ── Check 6b: Text alignment (no justified body text) ────────

    def _check_alignment(self, doc: Document) -> list[Issue]:
        issues: list[Issue] = []
        for p_idx, para in enumerate(doc.paragraphs, start=1):
            if not para.text.strip():
                continue
            p_pr = para._element.find(_qn("w:pPr"))
            if p_pr is None:
                continue
            jc = p_pr.find(_qn("w:jc"))
            if jc is None:
                continue
            val = (jc.get(_qn("w:val")) or jc.get("val", "")).lower()
            if val in ("both", "distribute", "justify"):
                issues.append(
                    Issue.create(
                        IssueType.TEXT_JUSTIFIED,
                        location=f"Paragraph {p_idx}",
                        description=(
                            f"Paragraph {p_idx} uses justified alignment, which "
                            f"creates uneven spacing between words for readers with "
                            f"dyslexia or low vision. Use left-aligned text instead."
                        ),
                        current_value="justified",
                        suggested_fix="left",
                    )
                )
        return issues

    # ── Check 7: Color contrast & color-only prose ───────────────

    def _check_colors(self, doc: Document) -> list[Issue]:
        from ada_platform.remediation_engine import (
            COLOR_WORDS_RE,
            audit_color_contrast,
        )

        issues: list[Issue] = []
        for item in audit_color_contrast(doc):
            if not item.passed:
                issues.append(
                    Issue.create(
                        IssueType.LOW_COLOR_CONTRAST,
                        location=item.location,
                        description=(
                            f"{item.role} text ({item.fg_hex} on {item.bg_hex}) has "
                            f"contrast ratio {item.ratio}:1, below the WCAG 2.1 AA "
                            f"{item.required_ratio}:1 minimum."
                        ),
                        fix_category=FixCategory.SUGGEST,
                        current_value=f"{item.fg_hex} on {item.bg_hex} ({item.ratio}:1)",
                        suggested_fix=item.recommended_fg,
                        metadata_key=item.pair_id,
                    )
                )

        for p_idx, p in enumerate(doc.paragraphs, start=1):
            txt = p.text
            for m in COLOR_WORDS_RE.finditer(txt):
                snippet = txt[max(0, m.start() - 20) : m.end() + 20].strip()
                issues.append(
                    Issue.create(
                        IssueType.COLOR_ONLY_INFO,
                        location=f"Paragraph {p_idx}",
                        description=(
                            f'Text refers to color alone ("...{snippet}..."). '
                            f"Ensure meaning is also conveyed via labels, patterns, or symbols."
                        ),
                        current_value=m.group(0),
                    )
                )
        return issues
