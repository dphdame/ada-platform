"""DOCX safe-fix engine: applies high-confidence accessibility fixes."""

from __future__ import annotations

import os
import sys
from typing import Optional

from docx import Document
from lxml import etree

from ada_platform.models.issues import FixCategory, Issue, IssueType
from ada_platform.models.metadata import MetadataWorkbook
from ada_platform.models.scan_result import FixResult, ScanResult
from ada_platform.fixer.base import BaseFixer

# OOXML namespace
W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
WP_NS = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"


def _qn(tag: str) -> str:
    prefix, local = tag.split(":")
    ns = {"w": W_NS, "wp": WP_NS}
    return f"{{{ns[prefix]}}}{local}"


class DocxFixer(BaseFixer):
    """Apply SAFE_AUTO fixes to DOCX files."""

    def fix(
        self,
        file_path: str,
        output_path: str,
        scan_result: ScanResult,
        metadata: Optional[MetadataWorkbook] = None,
    ) -> FixResult:
        doc = Document(file_path)

        auto_fixed: list[Issue] = []
        needs_review: list[Issue] = []
        skipped: list[Issue] = []

        for issue in scan_result.issues:
            if issue.fixed:
                continue

            if issue.fix_category == FixCategory.SAFE_AUTO and (
                issue.suggested_fix
                or issue.issue_type
                in (IssueType.TABLE_NO_HEADER, IssueType.DECORATIVE_IMAGE_HAS_ALT)
            ):
                success = self._apply_fix(doc, issue)
                if success:
                    issue.fixed = True
                    auto_fixed.append(issue)
                else:
                    skipped.append(issue)
            elif issue.fix_category == FixCategory.HUMAN_REVIEW:
                needs_review.append(issue)
            else:
                skipped.append(issue)

        # Clean metadata (author, comments) using existing script if available
        self._clean_metadata(doc)

        doc.save(output_path)

        return FixResult(
            file_path=file_path,
            output_path=output_path,
            auto_fixed=auto_fixed,
            needs_review=needs_review,
            skipped=skipped,
        )

    def _apply_fix(self, doc: Document, issue: Issue) -> bool:
        """Dispatch to the appropriate fix method. Returns True on success."""
        handlers = {
            IssueType.MISSING_TITLE: self._fix_title,
            IssueType.MISSING_LANGUAGE: self._fix_language,
            IssueType.MISSING_ALT_TEXT: self._fix_alt_text,
            IssueType.FILEPATH_ALT_TEXT: self._fix_alt_text,
            IssueType.ALT_TEXT_REDUNDANT_PREFIX: self._fix_alt_text,
            IssueType.DECORATIVE_IMAGE_HAS_ALT: self._fix_decorative_image,
            IssueType.TABLE_NO_HEADER: self._fix_table_header,
            IssueType.TEXT_JUSTIFIED: self._fix_text_justified,
        }
        handler = handlers.get(issue.issue_type)
        if handler is None:
            return False
        try:
            handler(doc, issue)
            return True
        except Exception:
            return False

    def _fix_title(self, doc: Document, issue: Issue) -> None:
        """Set document title from suggested_fix."""
        doc.core_properties.title = issue.suggested_fix

    def _fix_language(self, doc: Document, issue: Issue) -> None:
        """Set document language in both styles and settings."""
        lang_val = issue.suggested_fix or "en-US"

        # 1. Add w:lang to document defaults in styles
        styles_el = doc.styles.element
        doc_defaults = styles_el.find(_qn("w:docDefaults"))
        if doc_defaults is None:
            doc_defaults = etree.SubElement(styles_el, _qn("w:docDefaults"))
            styles_el.insert(0, doc_defaults)

        rpr_default = doc_defaults.find(_qn("w:rPrDefault"))
        if rpr_default is None:
            rpr_default = etree.SubElement(doc_defaults, _qn("w:rPrDefault"))

        rpr = rpr_default.find(_qn("w:rPr"))
        if rpr is None:
            rpr = etree.SubElement(rpr_default, _qn("w:rPr"))

        lang = rpr.find(_qn("w:lang"))
        if lang is None:
            lang = etree.SubElement(rpr, _qn("w:lang"))

        lang.set(_qn("w:val"), lang_val)
        lang.set(_qn("w:eastAsia"), lang_val)
        lang.set(_qn("w:bidi"), lang_val)

        # 2. Also set themeFontLang in settings (scanner checks here too)
        settings = doc.settings.element
        theme_lang = settings.find(_qn("w:themeFontLang"))
        if theme_lang is None:
            theme_lang = etree.SubElement(settings, _qn("w:themeFontLang"))
        theme_lang.set(_qn("w:val"), lang_val)

    def _fix_alt_text(self, doc: Document, issue: Issue) -> None:
        """Set alt text on an image identified by position."""
        # Parse position from location (e.g., "Image 3")
        position = self._parse_position(issue.location)
        if position is None:
            return

        body = doc.element
        drawings = body.findall(f".//{_qn('w:drawing')}")

        if position > len(drawings):
            return

        drawing = drawings[position - 1]
        doc_pr_els = drawing.findall(f".//{{{WP_NS}}}docPr")
        for dp in doc_pr_els:
            dp.set("descr", issue.suggested_fix)
            break

    def _fix_decorative_image(self, doc: Document, issue: Issue) -> None:
        """Ensure decorative image has descr='' and <adec:decorative val='1'/>."""
        position = self._parse_position(issue.location)
        if position is None:
            return
        body = doc.element
        drawings = body.findall(f".//{_qn('w:drawing')}")
        if position > len(drawings):
            return
        drawing = drawings[position - 1]
        a_ns = "http://schemas.openxmlformats.org/drawingml/2006/main"
        adec_ns = "http://schemas.microsoft.com/office/drawing/2017/decorative"
        for dp in drawing.findall(f".//{{{WP_NS}}}docPr"):
            dp.set("descr", "")
            dp.set("title", "")
            has_dec = any(
                el.get("val") in ("1", "true")
                for el in dp.iter(f"{{{adec_ns}}}decorative")
            )
            if not has_dec:
                ext_lst = dp.find(f"{{{a_ns}}}extLst")
                if ext_lst is None:
                    ext_lst = etree.SubElement(dp, f"{{{a_ns}}}extLst")
                ext = etree.SubElement(
                    ext_lst,
                    f"{{{a_ns}}}ext",
                    uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}",
                )
                dec = etree.SubElement(ext, f"{{{adec_ns}}}decorative")
                dec.set("val", "1")
            break

    def _fix_table_header(self, doc: Document, issue: Issue) -> None:
        """Mark the first row of a table as a header row."""
        position = self._parse_position(issue.location)
        if position is None:
            return

        if position > len(doc.tables):
            return

        table = doc.tables[position - 1]
        tbl_xml = table._tbl
        rows = tbl_xml.findall(_qn("w:tr"))
        if not rows:
            return

        first_row = rows[0]
        tr_pr = first_row.find(_qn("w:trPr"))
        if tr_pr is None:
            tr_pr = etree.SubElement(first_row, _qn("w:trPr"))
            first_row.insert(0, tr_pr)

        existing = tr_pr.find(_qn("w:tblHeader"))
        if existing is None:
            etree.SubElement(tr_pr, _qn("w:tblHeader"))

    def _fix_text_justified(self, doc: Document, issue: Issue) -> None:
        """Convert justified paragraph alignment to left-aligned."""
        position = self._parse_position(issue.location)
        if position is None or position > len(doc.paragraphs):
            return
        para = doc.paragraphs[position - 1]
        p_pr = para._element.find(_qn("w:pPr"))
        if p_pr is not None:
            jc = p_pr.find(_qn("w:jc"))
            if jc is not None:
                jc.set(_qn("w:val"), "left")

    def _parse_position(self, location: str) -> Optional[int]:
        """Extract position number from location string like 'Image 3' or 'Table 1'."""
        parts = location.strip().split()
        if len(parts) >= 2:
            try:
                return int(parts[-1])
            except ValueError:
                pass
        return None

    def _clean_metadata(self, doc: Document) -> None:
        """Clean document metadata using existing script if importable."""
        try:
            # Try importing the existing clean_docx_metadata script
            scripts_dir = os.path.expanduser("~/.claude/scripts")
            if scripts_dir not in sys.path:
                sys.path.insert(0, scripts_dir)
            from clean_docx_metadata import clean_metadata
            clean_metadata(doc)
        except (ImportError, ModuleNotFoundError):
            # Fallback: basic metadata cleanup
            props = doc.core_properties
            if not props.author or props.author == "python-docx":
                props.author = ""
            props.comments = ""
            props.revision = 1
