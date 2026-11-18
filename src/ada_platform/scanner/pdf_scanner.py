"""PDF accessibility scanner (Acrobat Pro Full Check equivalent)."""

from __future__ import annotations

import os
import re
from typing import Optional

import pymupdf as fitz

from ada_platform.models.issues import FixCategory, Issue, IssueType
from ada_platform.models.metadata import MetadataWorkbook
from ada_platform.models.scan_result import ScanResult
from ada_platform.remediation_engine import _graphic_count, run_pdf_ua_audit
from ada_platform.scanner.base import BaseScanner


class PdfScanner(BaseScanner):
    """Scan PDF files for WCAG 2.1 AA / PDF/UA-1 & Adobe Acrobat Pro accessibility issues."""

    def scan(
        self,
        file_path: str,
        metadata: Optional[MetadataWorkbook] = None,
    ) -> ScanResult:
        qa = run_pdf_ua_audit(file_path)
        issues: list[Issue] = []

        # Map each failed check from the 16-check PDF/UA audit into a structured Issue
        for chk in qa.get("checks", []):
            if chk.get("passed"):
                continue
            name = chk.get("check", "")
            detail = chk.get("detail", "")

            if "Document title" in name or "DisplayDocTitle" in name:
                base = os.path.splitext(os.path.basename(file_path))[0]
                suggested = base.replace("_", " ").replace("-", " ").title()
                issues.append(
                    Issue.create(
                        IssueType.MISSING_TITLE,
                        location=f"PDF Catalog ({name})",
                        description=f"{name}: {detail}",
                        fix_category=FixCategory.SAFE_AUTO,
                        suggested_fix=suggested,
                    )
                )
            elif "Document language" in name:
                issues.append(
                    Issue.create(
                        IssueType.MISSING_LANGUAGE,
                        location="PDF Catalog (/Lang)",
                        description=f"{name}: {detail}",
                        fix_category=FixCategory.SAFE_AUTO,
                        suggested_fix="en-US",
                    )
                )
            elif "Table headers" in name:
                issues.append(
                    Issue.create(
                        IssueType.TABLE_NO_HEADER,
                        location="PDF Structure Tree (/Table)",
                        description=f"{name}: {detail}",
                        fix_category=FixCategory.SAFE_AUTO,
                        suggested_fix="Promote first-row /TD cells to /TH with /Scope /Column",
                    )
                )
            elif "Figure alt text" in name or "Alt text is substantive" in name:
                issues.append(
                    Issue.create(
                        IssueType.MISSING_ALT_TEXT,
                        location="PDF Structure Tree (/Figure)",
                        description=f"{name}: {detail}",
                        fix_category=FixCategory.HUMAN_REVIEW,
                    )
                )
            else:
                # Structural tag tree / MarkInfo / layout checks
                is_auto = "MarkInfo" in name or "Internal Tagged PDF" in name
                issues.append(
                    Issue.create(
                        IssueType.HEADING_SKIP,
                        location=f"PDF ({name})",
                        description=f"{name}: {detail}",
                        fix_category=(
                            FixCategory.SAFE_AUTO
                            if is_auto
                            else FixCategory.HUMAN_REVIEW
                        ),
                        suggested_fix="Apply PDF/UA catalog & tag post-processing" if is_auto else None,
                    )
                )

        # Count images, tables, links, headings in the PDF
        total_images = 0
        total_tables = 0
        total_links = 0
        total_headings = 0
        if os.path.exists(file_path):
            doc = fitz.open(file_path)
            try:
                total_images = sum(_graphic_count(p) for p in doc)
                total_links = sum(len(p.get_links()) for p in doc)
                total_headings = len(doc.get_toc())
                for xref in range(1, doc.xref_length()):
                    try:
                        obj = doc.xref_object(xref)
                    except Exception:
                        continue
                    if "/Type /StructElem" in obj and re.search(r"/S\s+/Table\b", obj):
                        total_tables += 1
            finally:
                doc.close()

        return ScanResult(
            file_path=file_path,
            file_type="pdf",
            issues=issues,
            total_images=total_images,
            total_tables=total_tables,
            total_links=total_links,
            total_headings=total_headings,
        )
