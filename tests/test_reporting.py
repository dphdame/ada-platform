"""Tests for exception report generator."""

from __future__ import annotations

import json
import os

from ada_platform.models.issues import FixCategory, Issue, IssueType
from ada_platform.models.scan_result import FixResult, ScanResult
from ada_platform.reporting.exception_report import (
    fix_result_to_dict,
    generate_json_report,
    generate_scan_text_report,
    generate_text_report,
    scan_result_to_dict,
)


def _make_fix_result() -> FixResult:
    return FixResult(
        file_path="input.docx",
        output_path="output.docx",
        auto_fixed=[
            Issue.create(
                IssueType.MISSING_TITLE,
                "Document properties",
                "This document has no title.",
                suggested_fix="Annual Report",
            ),
            Issue.create(
                IssueType.MISSING_LANGUAGE,
                "Document settings",
                "No language set.",
                suggested_fix="en-US",
            ),
        ],
        needs_review=[
            Issue.create(
                IssueType.MISSING_ALT_TEXT,
                "Image 1",
                "Image 1 needs a description.",
            ),
        ],
        skipped=[
            Issue.create(
                IssueType.HEADING_SKIP,
                "Paragraph 5",
                "Heading level jumps from 1 to 3.",
            ),
        ],
    )


def _make_scan_result() -> ScanResult:
    return ScanResult(
        file_path="test.docx",
        file_type="docx",
        issues=[
            Issue.create(
                IssueType.MISSING_TITLE,
                "Document properties",
                "No title.",
            ),
            Issue.create(
                IssueType.MISSING_ALT_TEXT,
                "Image 1",
                "Image needs description.",
            ),
        ],
        total_images=3,
        total_tables=1,
        total_links=2,
        total_headings=4,
    )


class TestFixResultToDict:
    def test_structure(self):
        d = fix_result_to_dict(_make_fix_result())
        assert d["file_path"] == "input.docx"
        assert d["output_path"] == "output.docx"
        assert len(d["auto_fixed"]) == 2
        assert len(d["needs_review"]) == 1
        assert len(d["skipped"]) == 1
        assert "summary" in d

    def test_issue_fields(self):
        d = fix_result_to_dict(_make_fix_result())
        issue = d["auto_fixed"][0]
        assert issue["type"] == "missing_title"
        assert issue["severity"] == "critical"
        assert issue["location"] == "Document properties"
        assert issue["suggested_fix"] == "Annual Report"


class TestScanResultToDict:
    def test_structure(self):
        d = scan_result_to_dict(_make_scan_result())
        assert d["file_type"] == "docx"
        assert d["total_issues"] == 2
        assert d["counts"]["images"] == 3
        assert len(d["issues"]) == 2


class TestJSONReport:
    def test_writes_valid_json(self, tmp_path):
        path = str(tmp_path / "report.json")
        generate_json_report(_make_fix_result(), path)
        assert os.path.exists(path)
        with open(path) as f:
            data = json.load(f)
        assert data["file_path"] == "input.docx"
        assert len(data["auto_fixed"]) == 2

    def test_scan_result_json(self, tmp_path):
        path = str(tmp_path / "scan.json")
        generate_json_report(_make_scan_result(), path)
        with open(path) as f:
            data = json.load(f)
        assert data["total_issues"] == 2


class TestTextReport:
    def test_contains_sections(self):
        text = generate_text_report(_make_fix_result())
        assert "FIXED AUTOMATICALLY (2 items)" in text
        assert "NEEDS YOUR REVIEW (1 items)" in text
        assert "INFORMATIONAL (1 items)" in text

    def test_contains_descriptions(self):
        text = generate_text_report(_make_fix_result())
        assert "no title" in text.lower()
        assert "Image 1" in text

    def test_contains_suggested_fix(self):
        text = generate_text_report(_make_fix_result())
        assert "Annual Report" in text

    def test_no_xml_terms(self):
        """Plain language: no OOXML or XML terminology."""
        text = generate_text_report(_make_fix_result())
        forbidden = ["xml", "ooxml", "w:lang", "tblHeader", "docPr", "StructTreeRoot"]
        lower = text.lower()
        for term in forbidden:
            assert term.lower() not in lower, f"Found technical term: {term}"

    def test_empty_result(self):
        fr = FixResult(file_path="clean.docx", output_path="clean_out.docx")
        text = generate_text_report(fr)
        assert "No accessibility issues found" in text


class TestScanTextReport:
    def test_contains_counts(self):
        text = generate_scan_text_report(_make_scan_result())
        assert "3 images" in text
        assert "1 tables" in text

    def test_contains_summary(self):
        text = generate_scan_text_report(_make_scan_result())
        assert "Found 2 issues" in text
