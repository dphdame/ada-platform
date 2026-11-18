"""HTML and Drupal WYSIWYG accessibility scanner."""

from __future__ import annotations

import os
import re
from pathlib import Path
from typing import Optional

from lxml import etree, html

from ada_platform.models.issues import (
    FixCategory,
    Issue,
    IssueType,
    WEAK_LINK_PATTERNS,
)
from ada_platform.models.metadata import MetadataWorkbook
from ada_platform.models.scan_result import ScanResult
from ada_platform.remediation_engine import (
    COLOR_WORDS_RE,
    contrast_ratio,
    normalize_hex,
    recommend_accessible_fg,
)
from ada_platform.scanner.base import BaseScanner

_URL_PATTERN = re.compile(r"^https?://", re.IGNORECASE)
_FILEPATH_PATTERN = re.compile(
    r"(?:^|[\s\"'])(?:[A-Za-z]:\\|/(?:Users|tmp|var|home|sites|images|assets|figures)/|(?:figures|assets|images|media|uploads)/|\S+\.(?:png|jpg|jpeg|gif|bmp|tiff|svg|webp|emf|wmf)\b)",
    re.IGNORECASE,
)
_REDUNDANT_ALT_PREFIX_RE = re.compile(
    r"^\s*(?:an?\s+|the\s+)?(?:image|picture|photo|photograph|graphic|icon|screenshot)\s+of\s+",
    re.IGNORECASE,
)
_MAX_REGULAR_ALT_LENGTH = 141
_DECORATIVE_LITERAL_TOKENS = {
    "decorative",
    "[decorative]",
    "decoration",
    "spacer",
    "null",
    '""',
    "''",
}

_FULL_DOC_RE = re.compile(r"<\s*(?:!doctype\s+html|html\b)", re.IGNORECASE)
_CSS_COLOR_PROP_RE = re.compile(
    r"(?:^|;)\s*(color|background-color|background)\s*:\s*([^;]+)",
    re.IGNORECASE,
)
_HEX_COLOR_RE = re.compile(r"#(?:[0-9a-fA-F]{6}|[0-9a-fA-F]{3})\b")
_RGB_COLOR_RE = re.compile(
    r"rgb\(\s*(\d{1,3})\s*,\s*(\d{1,3})\s*,\s*(\d{1,3})\s*\)",
    re.IGNORECASE,
)


def _parse_css_hex(raw_val: str) -> Optional[str]:
    """Extract a 6-digit #RRGGBB hex color from a CSS value if present."""
    m_hex = _HEX_COLOR_RE.search(raw_val)
    if m_hex:
        token = m_hex.group(0)
        if len(token) == 4:
            token = "#" + "".join(ch * 2 for ch in token[1:])
        return normalize_hex(token)
    m_rgb = _RGB_COLOR_RE.search(raw_val)
    if m_rgb:
        r, g, b = (max(0, min(255, int(x))) for x in m_rgb.groups())
        return f"#{r:02X}{g:02X}{g:02X}"
    return None


class HtmlScanner(BaseScanner):
    """Scan HTML files or Drupal WYSIWYG HTML fragments for accessibility issues."""

    def scan(
        self,
        file_path: str,
        metadata: Optional[MetadataWorkbook] = None,
    ) -> ScanResult:
        raw_html = Path(file_path).read_text(encoding="utf-8", errors="replace")
        return self.scan_html_string(
            raw_html,
            file_path=file_path,
            metadata=metadata,
        )

    def scan_html_string(
        self,
        raw_html: str,
        *,
        file_path: str = "fragment.html",
        metadata: Optional[MetadataWorkbook] = None,
        is_fragment: Optional[bool] = None,
    ) -> ScanResult:
        filename = os.path.basename(file_path)
        treat_as_fragment = (
            is_fragment
            if is_fragment is not None
            else (not bool(_FULL_DOC_RE.search(raw_html)))
        )

        doc = html.fromstring(raw_html) if raw_html.strip() else html.Element("div")

        issues: list[Issue] = []
        if not treat_as_fragment:
            issues.extend(self._check_title(doc, raw_html, filename, metadata))
            issues.extend(self._check_language(doc, raw_html))

        heading_issues, heading_count = self._check_headings(
            doc, treat_as_fragment=treat_as_fragment
        )
        issues.extend(heading_issues)

        alt_issues, image_count = self._check_images(doc, filename, metadata)
        issues.extend(alt_issues)

        table_issues, table_count = self._check_tables(doc)
        issues.extend(table_issues)

        link_issues, link_count = self._check_links(doc, metadata)
        issues.extend(link_issues)

        issues.extend(self._check_alignment_and_spacers(doc))
        issues.extend(self._check_colors(doc))

        return ScanResult(
            file_path=file_path,
            file_type="html",
            issues=issues,
            total_images=image_count,
            total_tables=table_count,
            total_links=link_count,
            total_headings=heading_count,
        )

    # ── 1. Page <title> (full HTML documents) ─────────────────────

    def _check_title(
        self,
        doc: html.HtmlElement,
        raw_html: str,
        filename: str,
        metadata: Optional[MetadataWorkbook],
    ) -> list[Issue]:
        titles = doc.xpath("//head/title | //title")
        if titles:
            title_text = (titles[0].text_content() or "").strip()
            if title_text:
                return []

        suggested = None
        if metadata:
            doc_meta = metadata.get_document_meta(filename)
            if doc_meta and doc_meta.title:
                suggested = doc_meta.title
        if not suggested:
            h1s = doc.xpath("//h1")
            if h1s and (h1s[0].text_content() or "").strip():
                suggested = h1s[0].text_content().strip()
        if not suggested:
            base = os.path.splitext(filename)[0]
            suggested = base.replace("_", " ").replace("-", " ").title()

        return [
            Issue.create(
                IssueType.MISSING_TITLE,
                location="Page header (<title>)",
                description="This web page has no <title> element set in its header.",
                fix_category=FixCategory.SAFE_AUTO,
                suggested_fix=suggested,
            )
        ]

    # ── 2. Document language (<html lang="...">) ──────────────────

    def _check_language(self, doc: html.HtmlElement, raw_html: str) -> list[Issue]:
        m_lang = re.search(
            r"<\s*html\b[^>]*\blang\s*=\s*['\"]([^'\"]+)['\"]",
            raw_html,
            re.IGNORECASE,
        )
        if m_lang and m_lang.group(1).strip():
            return []
        root_tree = doc.getroottree().getroot()
        if root_tree is not None and (root_tree.get("lang") or "").strip():
            return []

        return [
            Issue.create(
                IssueType.MISSING_LANGUAGE,
                location="Root <html> element",
                description="No page language attribute (lang) is set on the root <html> element.",
                fix_category=FixCategory.SAFE_AUTO,
                suggested_fix="en",
            )
        ]

    # ── 3. Heading hierarchy (<h1>..<h6> & bold-as-heading) ──────

    def _check_headings(
        self,
        doc: html.HtmlElement,
        *,
        treat_as_fragment: bool = False,
    ) -> tuple[list[Issue], int]:
        issues: list[Issue] = []
        heading_count = 0
        last_level = 0
        h1_count = 0

        for el in doc.iter():
            if not isinstance(el.tag, str):
                continue
            tag = el.tag.lower()
            if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
                heading_count += 1
                level = int(tag[1])
                txt = " ".join((el.text_content() or "").split())

                if heading_count == 1 and level > 1 and not treat_as_fragment:
                    issues.append(
                        Issue.create(
                            IssueType.HEADING_FIRST_NOT_H1,
                            location=f"Heading {heading_count} (<{tag}>)",
                            description=(
                                f'First heading is <{tag}> ("{txt[:60]}"). Web pages '
                                f"should start with a single <h1> for the main page title."
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
                                location=f"Heading {heading_count} (<h1>)",
                                description=(
                                    f'Page contains multiple <h1> headings ("{txt[:60]}"). '
                                    f"Use a single <h1> for the page title and <h2> for major sections."
                                ),
                                current_value="Heading 1",
                                suggested_fix="Heading 2",
                            )
                        )

                if last_level > 0 and level > last_level + 1:
                    issues.append(
                        Issue.create(
                            IssueType.HEADING_SKIP,
                            location=f"Heading {heading_count} (<{tag}>)",
                            description=(
                                f'Heading level jumps from H{last_level} to H{level} at "{txt[:60]}". '
                                f"Screen readers expect sequential heading levels (H1 -> H2 -> H3)."
                            ),
                            current_value=f"Heading {level}",
                            suggested_fix=f"Heading {last_level + 1}",
                        )
                    )
                last_level = level

        # Check for <p><strong>Short Title</strong></p> used as a fake heading
        for p_idx, p_el in enumerate(doc.xpath("//p"), start=1):
            p_txt = " ".join((p_el.text_content() or "").split())
            if not p_txt or len(p_txt) > 90:
                continue
            children = [c for c in p_el if isinstance(c.tag, str) and c.tag.lower() != "br"]
            if (
                len(children) == 1
                and children[0].tag.lower() in ("strong", "b")
                and not (p_el.text or "").strip()
                and not (children[0].tail or "").strip()
            ):
                issues.append(
                    Issue.create(
                        IssueType.HEADING_BOLD_NOT_STYLE,
                        location=f"Paragraph {p_idx}",
                        description=(
                            f'"{p_txt}" appears to be a section heading but uses bold text '
                            f"inside a paragraph instead of a semantic heading (<h2> or <h3>)."
                        ),
                        current_value=p_txt,
                        suggested_fix="Heading 2",
                    )
                )

        return issues, heading_count

    # ── 4. Images (<img> alt, decorative "", <=141 chars, prefix) ─

    def _check_images(
        self,
        doc: html.HtmlElement,
        filename: str,
        metadata: Optional[MetadataWorkbook],
    ) -> tuple[list[Issue], int]:
        issues: list[Issue] = []
        imgs = doc.xpath("//img")

        for idx, img in enumerate(imgs, start=1):
            src = (img.get("src") or "").strip()
            src_name = os.path.basename(src.split("?")[0]) if src else f"Image {idx}"
            has_alt_attr = "alt" in img.attrib
            raw_alt = img.get("alt") or ""
            alt_text = raw_alt.strip()
            role = (img.get("role") or "").strip().lower()
            aria_hidden = (img.get("aria-hidden") or "").strip().lower()
            is_role_decorative = role in ("presentation", "none") or aria_hidden == "true"

            # Case A: Marked decorative via role/aria-hidden OR has literal "decorative" in alt
            if is_role_decorative:
                if alt_text:
                    issues.append(
                        Issue.create(
                            IssueType.DECORATIVE_IMAGE_HAS_ALT,
                            location=f"Image {idx}",
                            description=(
                                f'Image {idx} ({src_name}) is marked decorative but has non-empty '
                                f'alt text ("{alt_text}"). Decorative images must have alt="" '
                                f"so screen readers skip them."
                            ),
                            current_value=alt_text,
                            suggested_fix="",
                        )
                    )
                continue

            if alt_text.lower() in _DECORATIVE_LITERAL_TOKENS:
                issues.append(
                    Issue.create(
                        IssueType.DECORATIVE_IMAGE_HAS_ALT,
                        location=f"Image {idx}",
                        description=(
                            f'Image {idx} ({src_name}) has literal "{alt_text}" as its alt text. '
                            f'Decorative images must use empty alt="" so screen readers do not '
                            f"announce the word aloud."
                        ),
                        current_value=alt_text,
                        suggested_fix="",
                    )
                )
                continue

            # Case B: Explicit empty alt="" (and not whitespace-only alt="   ")
            if has_alt_attr and raw_alt == "":
                # In HTML, alt="" is the standard way to mark a decorative image
                continue

            # Case C: Missing alt attribute altogether (<img src="...">) or whitespace-only alt="  "
            if not has_alt_attr or not alt_text:
                meta_img = None
                if metadata:
                    meta_img = metadata.resolve_image(
                        filename=src_name,
                        content_hash="",
                        position=idx,
                        document=filename,
                    )
                if meta_img and meta_img.decorative:
                    issues.append(
                        Issue.create(
                            IssueType.DECORATIVE_IMAGE_HAS_ALT,
                            location=f"Image {idx}",
                            description=(
                                f'Image {idx} ({src_name}) is decorative in metadata but is '
                                f'missing empty alt="".'
                            ),
                            suggested_fix="",
                        )
                    )
                elif meta_img and meta_img.alt_text:
                    issues.append(
                        Issue.create(
                            IssueType.MISSING_ALT_TEXT,
                            location=f"Image {idx}",
                            description=f"Image {idx} ({src_name}) is missing an alt text description.",
                            fix_category=FixCategory.SAFE_AUTO,
                            suggested_fix=meta_img.alt_text,
                            metadata_key=meta_img.image_id,
                        )
                    )
                else:
                    issues.append(
                        Issue.create(
                            IssueType.MISSING_ALT_TEXT,
                            location=f"Image {idx}",
                            description=(
                                f"Image {idx} ({src_name}) has no alt attribute. Add concise "
                                f'descriptive alt text (<= 125-141 chars) or alt="" if decorative.'
                            ),
                        )
                    )
                continue

            # Case D: File path as alt text
            if _FILEPATH_PATTERN.search(alt_text):
                issues.append(
                    Issue.create(
                        IssueType.FILEPATH_ALT_TEXT,
                        location=f"Image {idx}",
                        description=(
                            f"Image {idx} ({src_name}) uses a file name or path as its alt text "
                            f'("{alt_text}") instead of a meaningful description.'
                        ),
                        current_value=alt_text,
                    )
                )
                continue

            # Case E: Redundant "image of" / "photo of" / "picture of" prefix
            m_pref = _REDUNDANT_ALT_PREFIX_RE.match(alt_text)
            if m_pref:
                rem = alt_text[m_pref.end() :].strip()
                cleaned = rem[0].upper() + rem[1:] if rem else alt_text
                issues.append(
                    Issue.create(
                        IssueType.ALT_TEXT_REDUNDANT_PREFIX,
                        location=f"Image {idx}",
                        description=(
                            f'Image {idx} alt text starts with a redundant phrase '
                            f'("{m_pref.group(0).strip()}"). Screen readers already announce '
                            f"that an element is an image."
                        ),
                        current_value=alt_text,
                        suggested_fix=cleaned,
                    )
                )

            # Case F: Alt text exceeds 141-character Siteimprove ceiling
            if len(alt_text) > _MAX_REGULAR_ALT_LENGTH:
                issues.append(
                    Issue.create(
                        IssueType.ALT_TEXT_TOO_LONG,
                        location=f"Image {idx}",
                        description=(
                            f"Image {idx} alt text is {len(alt_text)} characters (exceeds the "
                            f"{_MAX_REGULAR_ALT_LENGTH}-character maximum; 125 characters is the "
                            f"recommended screen-reader target). Keep alt text concise and place "
                            f"full chart or infographic details in adjacent page text or a table."
                        ),
                        current_value=alt_text,
                    )
                )

        return issues, len(imgs)

    # ── 5. Tables (<th> headers, colspan/rowspan merges, empty cells)

    def _check_tables(self, doc: html.HtmlElement) -> tuple[list[Issue], int]:
        issues: list[Issue] = []
        tables = doc.xpath("//table")

        for idx, tbl in enumerate(tables, start=1):
            rows = tbl.xpath(".//tr")
            if not rows:
                continue

            ths = tbl.xpath(".//th")
            if not ths:
                issues.append(
                    Issue.create(
                        IssueType.TABLE_NO_HEADER,
                        location=f"Table {idx}",
                        description=(
                            f"Table {idx} has no header cells (<th>). In Drupal WYSIWYG, "
                            f"enable the Header Row toggle so column labels are tagged for screen readers."
                        ),
                    )
                )

            # Merged cells: colspan > 1 or rowspan > 1
            has_merged = False
            for cell in tbl.xpath(".//th | .//td"):
                for attr in ("colspan", "rowspan"):
                    val = (cell.get(attr) or "1").strip()
                    try:
                        if int(val) > 1:
                            has_merged = True
                            break
                    except ValueError:
                        pass
                if has_merged:
                    break

            if has_merged:
                issues.append(
                    Issue.create(
                        IssueType.TABLE_MERGED_CELLS,
                        location=f"Table {idx}",
                        description=(
                            f"Table {idx} uses merged cells (colspan or rowspan). "
                            f"Merged cells confuse screen reader table navigation; "
                            f"flatten the table structure instead."
                        ),
                    )
                )

            # Empty cells in populated tables
            non_empty = 0
            empty_cnt = 0
            for cell in tbl.xpath(".//th | .//td"):
                txt = " ".join((cell.text_content() or "").replace("\xa0", " ").split())
                if txt:
                    non_empty += 1
                else:
                    empty_cnt += 1

            if non_empty > 0 and empty_cnt > 0:
                issues.append(
                    Issue.create(
                        IssueType.TABLE_EMPTY_CELLS,
                        location=f"Table {idx}",
                        description=(
                            f"Table {idx} has {empty_cnt} blank cell(s). Enter 'NA' or "
                            f"'no data' in empty cells so screen reader users know no value was skipped."
                        ),
                        suggested_fix="NA",
                    )
                )

        return issues, len(tables)

    # ── 6. Links (<a href>) & Drupal Jump Links ──────────────────

    def _check_links(
        self,
        doc: html.HtmlElement,
        metadata: Optional[MetadataWorkbook],
    ) -> tuple[list[Issue], int]:
        issues: list[Issue] = []
        links = doc.xpath("//a[@href]")

        for idx, a_el in enumerate(links, start=1):
            href = (a_el.get("href") or "").strip()
            disp = " ".join((a_el.text_content() or "").split())
            aria_label = (a_el.get("aria-label") or "").strip()
            effective_text = disp or aria_label

            if not effective_text:
                # Check if link wraps an image with alt text
                img_alts = [
                    (im.get("alt") or "").strip()
                    for im in a_el.xpath(".//img")
                    if (im.get("alt") or "").strip()
                ]
                if img_alts:
                    effective_text = img_alts[0]

            if not effective_text:
                issues.append(
                    Issue.create(
                        IssueType.WEAK_LINK_TEXT,
                        location=f"Link {idx}",
                        description=(
                            f'Link {idx} (pointing to "{href}") has no visible text or aria-label.'
                        ),
                        current_value="",
                    )
                )
                continue

            lower_text = effective_text.lower().strip()
            is_weak = (
                lower_text in WEAK_LINK_PATTERNS
                or bool(_URL_PATTERN.match(effective_text))
            )
            if not is_weak:
                continue

            suggested = None
            fix_cat = FixCategory.HUMAN_REVIEW
            if metadata and href:
                link_meta = metadata.get_link_text(href)
                if link_meta and link_meta.approved_text:
                    suggested = link_meta.approved_text
                    fix_cat = FixCategory.SUGGEST

            issues.append(
                Issue.create(
                    IssueType.WEAK_LINK_TEXT,
                    location=f"Link {idx}",
                    description=(
                        f'Link {idx} uses "{effective_text}" as its link text. '
                        f"Use descriptive text that explains the destination out of context."
                    ),
                    fix_category=fix_cat,
                    current_value=effective_text,
                    suggested_fix=suggested,
                )
            )

        return issues, len(links)

    # ── 7. Alignment (no text-align: justify) & spacer <p> ───────

    def _check_alignment_and_spacers(self, doc: html.HtmlElement) -> list[Issue]:
        issues: list[Issue] = []

        for idx, el in enumerate(doc.xpath("//*[@style or @align]"), start=1):
            style = (el.get("style") or "").lower()
            align = (el.get("align") or "").lower().strip()
            if align == "justify" or re.search(r"text-align\s*:\s*justify", style):
                issues.append(
                    Issue.create(
                        IssueType.TEXT_JUSTIFIED,
                        location=f"Element {idx} (<{el.tag}>)",
                        description=(
                            f"Element <{el.tag}> uses justified text alignment, which creates "
                            f"uneven word spacing for readers with dyslexia or low vision. "
                            f"Use left-aligned text instead."
                        ),
                        current_value="justify",
                        suggested_fix="left",
                    )
                )

        for p_idx, p_el in enumerate(doc.xpath("//p"), start=1):
            txt = (p_el.text_content() or "").replace("\xa0", "").strip()
            has_child_media = len(p_el.xpath(".//img | .//iframe | .//svg | .//input | .//a")) > 0
            if not txt and not has_child_media:
                issues.append(
                    Issue.create(
                        IssueType.EMPTY_SPACER_PARAGRAPH,
                        location=f"Paragraph {p_idx}",
                        description=(
                            f"Paragraph {p_idx} is an empty spacer paragraph. Screen readers "
                            f"may announce 'blank' for empty paragraphs; use CSS margin/padding instead."
                        ),
                    )
                )

        return issues

    # ── 8. Inline Color Contrast & Color-Only Prose ──────────────

    def _check_colors(self, doc: html.HtmlElement) -> list[Issue]:
        issues: list[Issue] = []

        for idx, el in enumerate(doc.xpath("//*[@style]"), start=1):
            style = el.get("style") or ""
            fg_hex = None
            bg_hex = "#FFFFFF"
            for m in _CSS_COLOR_PROP_RE.finditer(style):
                prop = m.group(1).lower()
                val = m.group(2)
                parsed = _parse_css_hex(val)
                if not parsed:
                    continue
                if prop == "color":
                    fg_hex = parsed
                else:
                    bg_hex = parsed
            if fg_hex and (el.text_content() or "").strip():
                ratio = contrast_ratio(fg_hex, bg_hex)
                is_heading = isinstance(el.tag, str) and el.tag.lower() in ("h1", "h2")
                req = 3.0 if is_heading else 4.5
                if ratio < req:
                    issues.append(
                        Issue.create(
                            IssueType.LOW_COLOR_CONTRAST,
                            location=f"Element {idx} (<{el.tag}>)",
                            description=(
                                f"Inline style color ({fg_hex} on {bg_hex}) has contrast ratio "
                                f"{ratio}:1, below the WCAG 2.1 AA {req}:1 minimum."
                            ),
                            fix_category=FixCategory.SUGGEST,
                            current_value=f"{fg_hex} on {bg_hex} ({ratio}:1)",
                            suggested_fix=recommend_accessible_fg(bg_hex, req),
                        )
                    )

        for p_idx, p_el in enumerate(doc.xpath("//p | //li"), start=1):
            txt = " ".join((p_el.text_content() or "").split())
            for m in COLOR_WORDS_RE.finditer(txt):
                snippet = txt[max(0, m.start() - 20) : m.end() + 20].strip()
                issues.append(
                    Issue.create(
                        IssueType.COLOR_ONLY_INFO,
                        location=f"Text block {p_idx}",
                        description=(
                            f'Text refers to color or visual location alone ("...{snippet}..."). '
                            f"Ensure meaning is also conveyed via text labels or symbols."
                        ),
                        current_value=m.group(0),
                    )
                )

        return issues
