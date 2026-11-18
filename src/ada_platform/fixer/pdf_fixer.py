"""PDF safe-fix engine (Acrobat Pro post-processing equivalent)."""

from __future__ import annotations

import os
from typing import Optional

import pymupdf as fitz

from ada_platform.fixer.base import BaseFixer
from ada_platform.models.issues import FixCategory, Issue, IssueType
from ada_platform.models.metadata import MetadataWorkbook
from ada_platform.models.scan_result import FixResult, ScanResult
from ada_platform.remediation_engine import postprocess_tagged_pdf


class PdfFixer(BaseFixer):
    """Apply deterministic PDF/UA-1 catalog, table <TH /Scope /Column>, and /Artifact fixes to a PDF."""

    def fix(
        self,
        file_path: str,
        output_path: str,
        scan_result: ScanResult,
        metadata: Optional[MetadataWorkbook] = None,
    ) -> FixResult:
        filename = os.path.basename(file_path)
        base = os.path.splitext(filename)[0].replace("_", " ").replace("-", " ").title()

        doc = fitz.open(file_path)
        meta = doc.metadata or {}
        current_title = (meta.get("title") or "").strip()
        current_author = (meta.get("author") or "").strip()
        doc.close()

        title_str = current_title or base
        author_str = current_author or "Santa Clara County Public Health Department"
        lang_str = "en-US"

        if metadata:
            doc_meta = metadata.get_document_meta(filename)
            if doc_meta:
                if doc_meta.title:
                    title_str = doc_meta.title
                if doc_meta.language:
                    lang_str = doc_meta.language

        postprocess_tagged_pdf(
            compiled_pdf_path=file_path,
            output_pdf_path=output_path,
            title_str=title_str,
            author_str=author_str,
            doc_lang=lang_str,
        )

        auto_fixed: list[Issue] = []
        needs_review: list[Issue] = []
        skipped: list[Issue] = []

        for issue in scan_result.issues:
            if issue.fix_category == FixCategory.SAFE_AUTO:
                issue.fixed = True
                auto_fixed.append(issue)
            elif issue.fix_category == FixCategory.HUMAN_REVIEW:
                needs_review.append(issue)
            else:
                skipped.append(issue)

        return FixResult(
            file_path=file_path,
            output_path=output_path,
            auto_fixed=auto_fixed,
            needs_review=needs_review,
            skipped=skipped,
        )
