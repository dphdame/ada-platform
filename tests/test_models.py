"""Tests for data models: issues, scan_result, metadata."""

from ada_platform.models.issues import (
    FixCategory,
    Issue,
    IssueType,
    Severity,
    ISSUE_SEVERITY,
    ISSUE_DEFAULT_FIX,
    WEAK_LINK_PATTERNS,
)
from ada_platform.models.scan_result import (
    AuditLogEntry,
    FixResult,
    ScanResult,
)
from ada_platform.models.metadata import (
    ChartMeta,
    DocumentMeta,
    ImageMeta,
    LinkMeta,
    MetadataWorkbook,
    SlideMeta,
    TableMeta,
)


# ── IssueType enum ──────────────────────────────────────────────


class TestIssueType:
    def test_all_types_have_severity(self):
        for it in IssueType:
            assert it in ISSUE_SEVERITY, f"{it} missing from ISSUE_SEVERITY"

    def test_all_types_have_default_fix(self):
        for it in IssueType:
            assert it in ISSUE_DEFAULT_FIX, f"{it} missing from ISSUE_DEFAULT_FIX"

    def test_enum_values_are_strings(self):
        for it in IssueType:
            assert isinstance(it.value, str)


# ── Issue dataclass ──────────────────────────────────────────────


class TestIssue:
    def test_create_with_defaults(self):
        issue = Issue.create(
            IssueType.MISSING_TITLE,
            location="Document properties",
            description="This document has no title.",
        )
        assert issue.severity == Severity.CRITICAL
        assert issue.fix_category == FixCategory.SAFE_AUTO
        assert issue.fixed is False
        assert issue.current_value is None

    def test_create_with_override(self):
        issue = Issue.create(
            IssueType.MISSING_ALT_TEXT,
            location="Image 3",
            description="This image needs a description.",
            fix_category=FixCategory.SAFE_AUTO,
            suggested_fix="Bar chart showing revenue by quarter",
            metadata_key="chart_revenue_q4",
        )
        assert issue.fix_category == FixCategory.SAFE_AUTO
        assert issue.suggested_fix == "Bar chart showing revenue by quarter"
        assert issue.metadata_key == "chart_revenue_q4"

    def test_create_missing_language(self):
        issue = Issue.create(
            IssueType.MISSING_LANGUAGE,
            location="Document settings",
            description="No language set for this document.",
            suggested_fix="en-US",
        )
        assert issue.severity == Severity.CRITICAL
        assert issue.fix_category == FixCategory.SAFE_AUTO


# ── ScanResult ───────────────────────────────────────────────────


class TestScanResult:
    def _make_result(self) -> ScanResult:
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
                    IssueType.MISSING_LANGUAGE,
                    "Document settings",
                    "No language.",
                ),
                Issue.create(
                    IssueType.MISSING_ALT_TEXT,
                    "Image 1",
                    "Image needs description.",
                ),
                Issue.create(
                    IssueType.HEADING_SKIP,
                    "Paragraph 5",
                    "Heading level skipped.",
                ),
            ],
            total_images=3,
            total_tables=1,
            total_links=2,
            total_headings=4,
        )

    def test_safe_fixable(self):
        result = self._make_result()
        safe = result.safe_fixable
        assert len(safe) == 2
        types = {i.issue_type for i in safe}
        assert IssueType.MISSING_TITLE in types
        assert IssueType.MISSING_LANGUAGE in types

    def test_needs_review(self):
        result = self._make_result()
        review = result.needs_review
        assert len(review) == 1
        assert review[0].issue_type == IssueType.MISSING_ALT_TEXT

    def test_suggested(self):
        result = self._make_result()
        suggested = result.suggested
        assert len(suggested) == 1
        assert suggested[0].issue_type == IssueType.HEADING_SKIP

    def test_critical_count(self):
        result = self._make_result()
        assert result.critical_count == 3  # title, language, alt text

    def test_critical_count_after_fix(self):
        result = self._make_result()
        for issue in result.issues:
            if issue.issue_type == IssueType.MISSING_TITLE:
                issue.fixed = True
        assert result.critical_count == 2

    def test_total_issues(self):
        result = self._make_result()
        assert result.total_issues == 4

    def test_summary_with_issues(self):
        result = self._make_result()
        s = result.summary()
        assert "4 issues" in s
        assert "2 can be fixed automatically" in s
        assert "1 need your review" in s

    def test_summary_no_issues(self):
        result = ScanResult(file_path="clean.docx", file_type="docx")
        assert result.summary() == "No accessibility issues found."

    def test_timestamp_auto_set(self):
        result = ScanResult(file_path="t.docx", file_type="docx")
        assert result.timestamp  # non-empty ISO string


# ── FixResult ────────────────────────────────────────────────────


class TestFixResult:
    def test_summary_mixed(self):
        fr = FixResult(
            file_path="in.docx",
            output_path="out.docx",
            auto_fixed=[
                Issue.create(IssueType.MISSING_TITLE, "props", "No title.")
            ],
            needs_review=[
                Issue.create(IssueType.MISSING_ALT_TEXT, "Image 1", "Needs desc.")
            ],
        )
        s = fr.summary()
        assert "Fixed 1" in s
        assert "1 items need your review" in s

    def test_summary_no_changes(self):
        fr = FixResult(file_path="in.docx", output_path="out.docx")
        assert fr.summary() == "No changes needed."


# ── AuditLogEntry ────────────────────────────────────────────────


class TestAuditLogEntry:
    def test_creation(self):
        entry = AuditLogEntry(action="scan", detail="Scanned test.docx")
        assert entry.action == "scan"
        assert entry.timestamp


# ── MetadataWorkbook ─────────────────────────────────────────────


class TestMetadataWorkbook:
    def _make_workbook(self) -> MetadataWorkbook:
        return MetadataWorkbook(
            documents=[
                DocumentMeta(
                    filename="report.docx",
                    title="Annual Report 2026",
                    language="en-US",
                    business_unit="PHD",
                    template_type="report",
                    audience="public",
                ),
            ],
            images=[
                ImageMeta(
                    image_id="logo.png",
                    alt_text="County of Santa Clara seal",
                    content_hash="abc123",
                    used_in="report.docx",
                ),
                ImageMeta(
                    image_id="Image 2",
                    alt_text="Bar chart showing enrollment trends",
                    used_in="report.docx",
                ),
                ImageMeta(
                    image_id="divider.png",
                    decorative=True,
                    content_hash="def456",
                ),
            ],
            charts=[
                ChartMeta(
                    chart_id="enrollment_bar",
                    chart_type="bar",
                    title="Enrollment by Year",
                    x_axis="Year",
                    y_axis="Enrollment count",
                    key_takeaway="Enrollment grew 15% from 2020 to 2025.",
                ),
            ],
            links=[
                LinkMeta(
                    url="https://example.com/report",
                    approved_text="2026 Annual Report (PDF)",
                    content_type="PDF",
                ),
            ],
            tables=[
                TableMeta(table_id="budget_table", has_header_row=True),
                TableMeta(
                    table_id="complex_table",
                    has_header_row=True,
                    complex=True,
                    notes="Multi-level headers",
                ),
            ],
            slides=[
                SlideMeta(
                    slide_number=1,
                    title="Welcome",
                    document="deck.pptx",
                ),
            ],
        )

    def test_get_document_meta(self):
        wb = self._make_workbook()
        doc = wb.get_document_meta("report.docx")
        assert doc is not None
        assert doc.title == "Annual Report 2026"
        assert wb.get_document_meta("nonexistent.docx") is None

    def test_get_image_alt_by_id(self):
        wb = self._make_workbook()
        img = wb.get_image_alt("logo.png")
        assert img is not None
        assert img.alt_text == "County of Santa Clara seal"

    def test_get_image_by_hash(self):
        wb = self._make_workbook()
        img = wb.get_image_by_hash("abc123")
        assert img is not None
        assert img.image_id == "logo.png"
        assert wb.get_image_by_hash("nonexistent") is None
        assert wb.get_image_by_hash("") is None

    def test_get_image_by_position(self):
        wb = self._make_workbook()
        img = wb.get_image_by_position(2, document="report.docx")
        assert img is not None
        assert img.alt_text == "Bar chart showing enrollment trends"

    def test_resolve_image_filename_first(self):
        wb = self._make_workbook()
        img = wb.resolve_image(
            filename="logo.png", content_hash="abc123", position=1
        )
        assert img is not None
        assert img.image_id == "logo.png"

    def test_resolve_image_hash_fallback(self):
        wb = self._make_workbook()
        img = wb.resolve_image(
            filename="unknown.png", content_hash="abc123"
        )
        assert img is not None
        assert img.image_id == "logo.png"

    def test_resolve_image_position_fallback(self):
        wb = self._make_workbook()
        img = wb.resolve_image(
            filename="unknown.png",
            content_hash="nonexistent",
            position=2,
            document="report.docx",
        )
        assert img is not None
        assert img.image_id == "Image 2"

    def test_resolve_image_no_match(self):
        wb = self._make_workbook()
        assert wb.resolve_image(
            filename="nope", content_hash="nope", position=99
        ) is None

    def test_get_chart_info(self):
        wb = self._make_workbook()
        chart = wb.get_chart_info("enrollment_bar")
        assert chart is not None
        assert chart.chart_type == "bar"
        assert wb.get_chart_info("nonexistent") is None

    def test_get_link_text(self):
        wb = self._make_workbook()
        link = wb.get_link_text("https://example.com/report")
        assert link is not None
        assert link.approved_text == "2026 Annual Report (PDF)"
        assert wb.get_link_text("https://other.com") is None

    def test_get_table_info(self):
        wb = self._make_workbook()
        t = wb.get_table_info("budget_table")
        assert t is not None
        assert t.has_header_row is True
        assert t.complex is False

    def test_get_table_complex(self):
        wb = self._make_workbook()
        t = wb.get_table_info("complex_table")
        assert t is not None
        assert t.complex is True

    def test_get_slide_info(self):
        wb = self._make_workbook()
        s = wb.get_slide_info(1, document="deck.pptx")
        assert s is not None
        assert s.title == "Welcome"
        assert wb.get_slide_info(99) is None

    def test_empty_workbook(self):
        wb = MetadataWorkbook()
        assert wb.get_document_meta("any") is None
        assert wb.get_image_alt("any") is None
        assert wb.resolve_image(filename="x") is None


# ── Weak link patterns ───────────────────────────────────────────


class TestWeakLinkPatterns:
    def test_patterns_are_lowercase(self):
        for p in WEAK_LINK_PATTERNS:
            assert p == p.lower()

    def test_common_patterns_present(self):
        assert "click here" in WEAK_LINK_PATTERNS
        assert "here" in WEAK_LINK_PATTERNS
        assert "read more" in WEAK_LINK_PATTERNS
