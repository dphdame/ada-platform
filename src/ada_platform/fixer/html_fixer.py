"""HTML and Drupal WYSIWYG safe-fix engine."""

from __future__ import annotations

import re
from pathlib import Path
from typing import Optional

from lxml import etree, html

from ada_platform.fixer.base import BaseFixer
from ada_platform.models.issues import FixCategory, Issue, IssueType
from ada_platform.models.metadata import MetadataWorkbook
from ada_platform.models.scan_result import FixResult, ScanResult

_FULL_DOC_RE = re.compile(r"<\s*(?:!doctype\s+html|html\b)", re.IGNORECASE)


class HtmlFixer(BaseFixer):
    """Apply SAFE_AUTO accessibility fixes to HTML pages and Drupal WYSIWYG fragments."""

    def fix(
        self,
        file_path: str,
        output_path: str,
        scan_result: ScanResult,
        metadata: Optional[MetadataWorkbook] = None,
    ) -> FixResult:
        raw_html = Path(file_path).read_text(encoding="utf-8", errors="replace")
        fixed_html, fix_result = self.fix_html_string(
            raw_html,
            scan_result=scan_result,
            file_path=file_path,
            output_path=output_path,
            metadata=metadata,
        )
        Path(output_path).write_text(fixed_html, encoding="utf-8")
        return fix_result

    def fix_html_string(
        self,
        raw_html: str,
        scan_result: ScanResult,
        *,
        file_path: str = "fragment.html",
        output_path: str = "fragment_fixed.html",
        metadata: Optional[MetadataWorkbook] = None,
    ) -> tuple[str, FixResult]:
        is_full_doc = bool(_FULL_DOC_RE.search(raw_html))
        if is_full_doc:
            doc = html.document_fromstring(raw_html)
        else:
            doc = html.fromstring(f"<div data-ada-wrapper='1'>{raw_html}</div>")

        auto_fixed: list[Issue] = []
        needs_review: list[Issue] = []
        skipped: list[Issue] = []

        for issue in scan_result.issues:
            if issue.fixed:
                continue

            if issue.fix_category == FixCategory.SAFE_AUTO and (
                issue.suggested_fix is not None
                or issue.issue_type
                in (
                    IssueType.TABLE_NO_HEADER,
                    IssueType.DECORATIVE_IMAGE_HAS_ALT,
                    IssueType.EMPTY_SPACER_PARAGRAPH,
                )
            ):
                if self._apply_fix(doc, issue):
                    issue.fixed = True
                    auto_fixed.append(issue)
                else:
                    skipped.append(issue)
            elif issue.fix_category == FixCategory.HUMAN_REVIEW:
                needs_review.append(issue)
            else:
                skipped.append(issue)

        if is_full_doc:
            out_str = html.tostring(
                doc,
                encoding="unicode",
                method="html",
                doctype="<!DOCTYPE html>",
            )
        else:
            inner_parts: list[str] = []
            if doc.text:
                inner_parts.append(doc.text)
            for child in doc:
                inner_parts.append(
                    html.tostring(child, encoding="unicode", method="html")
                )
            out_str = "".join(inner_parts)

        return (
            out_str,
            FixResult(
                file_path=file_path,
                output_path=output_path,
                auto_fixed=auto_fixed,
                needs_review=needs_review,
                skipped=skipped,
            ),
        )

    def _apply_fix(self, doc: html.HtmlElement, issue: Issue) -> bool:
        handlers = {
            IssueType.MISSING_TITLE: self._fix_title,
            IssueType.MISSING_LANGUAGE: self._fix_language,
            IssueType.MISSING_ALT_TEXT: self._fix_alt_text,
            IssueType.FILEPATH_ALT_TEXT: self._fix_alt_text,
            IssueType.ALT_TEXT_REDUNDANT_PREFIX: self._fix_alt_text,
            IssueType.DECORATIVE_IMAGE_HAS_ALT: self._fix_decorative_image,
            IssueType.TABLE_NO_HEADER: self._fix_table_header,
            IssueType.TEXT_JUSTIFIED: self._fix_text_justified,
            IssueType.EMPTY_SPACER_PARAGRAPH: self._fix_empty_spacer,
        }
        handler = handlers.get(issue.issue_type)
        if handler is None:
            return False
        try:
            return handler(doc, issue)
        except Exception:
            return False

    def _fix_title(self, doc: html.HtmlElement, issue: Issue) -> bool:
        root = doc.getroottree().getroot() if doc.getroottree() else doc
        heads = root.xpath("//head")
        if heads:
            head = heads[0]
        else:
            head = etree.Element("head")
            root.insert(0, head)
        titles = head.xpath("./title")
        if titles:
            titles[0].text = issue.suggested_fix or "Document"
        else:
            t_el = etree.SubElement(head, "title")
            t_el.text = issue.suggested_fix or "Document"
        return True

    def _fix_language(self, doc: html.HtmlElement, issue: Issue) -> bool:
        root = doc.getroottree().getroot() if doc.getroottree() else doc
        root.set("lang", issue.suggested_fix or "en")
        return True

    def _fix_alt_text(self, doc: html.HtmlElement, issue: Issue) -> bool:
        pos = self._parse_position(issue.location)
        imgs = doc.xpath("//img")
        if pos is None or pos > len(imgs):
            return False
        imgs[pos - 1].set("alt", issue.suggested_fix or "")
        return True

    def _fix_decorative_image(self, doc: html.HtmlElement, issue: Issue) -> bool:
        pos = self._parse_position(issue.location)
        imgs = doc.xpath("//img")
        if pos is None or pos > len(imgs):
            return False
        img = imgs[pos - 1]
        img.set("alt", "")
        img.set("role", "presentation")
        return True

    def _fix_table_header(self, doc: html.HtmlElement, issue: Issue) -> bool:
        pos = self._parse_position(issue.location)
        tables = doc.xpath("//table")
        if pos is None or pos > len(tables):
            return False
        tbl = tables[pos - 1]
        rows = tbl.xpath(".//tr")
        if not rows:
            return False
        first_row = rows[0]
        for cell in first_row.xpath("./td | ./th"):
            cell.tag = "th"
            if not cell.get("scope"):
                cell.set("scope", "col")
        return True

    def _fix_text_justified(self, doc: html.HtmlElement, issue: Issue) -> bool:
        pos = self._parse_position(issue.location)
        els = doc.xpath("//*[@style or @align]")
        if pos is None or pos > len(els):
            return False
        el = els[pos - 1]
        if (el.get("align") or "").lower() == "justify":
            el.set("align", "left")
        if el.get("style"):
            new_style = re.sub(
                r"text-align\s*:\s*justify",
                "text-align: left",
                el.get("style") or "",
                flags=re.IGNORECASE,
            )
            el.set("style", new_style)
        return True

    def _fix_empty_spacer(self, doc: html.HtmlElement, issue: Issue) -> bool:
        updated = False
        for p_el in list(doc.xpath("//p")):
            txt = (p_el.text_content() or "").replace("\xa0", "").strip()
            has_child = len(p_el.xpath(".//img | .//iframe | .//svg | .//input | .//a")) > 0
            if not txt and not has_child and p_el.getparent() is not None:
                p_el.getparent().remove(p_el)
                updated = True
        return updated

    def _parse_position(self, location: str) -> Optional[int]:
        m = re.search(r"(\d+)", location)
        return int(m.group(1)) if m else None
