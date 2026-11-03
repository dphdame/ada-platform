"""Issue types, fix categories, and the Issue dataclass."""

from __future__ import annotations

from dataclasses import dataclass
from enum import Enum
from typing import Optional


class IssueType(Enum):
    """Accessibility issue types detected by scanners."""

    MISSING_TITLE = "missing_title"
    MISSING_LANGUAGE = "missing_language"
    HEADING_SKIP = "heading_skip"
    HEADING_BOLD_NOT_STYLE = "heading_bold_not_style"
    MISSING_ALT_TEXT = "missing_alt_text"
    FILEPATH_ALT_TEXT = "filepath_alt_text"
    TABLE_NO_HEADER = "table_no_header"
    TABLE_COMPLEX = "table_complex"
    WEAK_LINK_TEXT = "weak_link_text"
    MISSING_FOOTNOTE_LINK = "missing_footnote_link"
    COLOR_ONLY_INFO = "color_only_info"
    LOW_COLOR_CONTRAST = "low_color_contrast"
    SMALL_FONT_SIZE = "small_font_size"
    EMPTY_SPACER_PARAGRAPH = "empty_spacer_paragraph"
    UNSPACED_SUPERSCRIPT = "unspaced_superscript"
    # PPTX-specific (reserved for Phase 2)
    SLIDE_MISSING_TITLE = "slide_missing_title"
    SLIDE_READING_ORDER = "slide_reading_order"


class FixCategory(Enum):
    """How an issue should be handled."""

    SAFE_AUTO = "safe_auto"
    SUGGEST = "suggest"
    HUMAN_REVIEW = "human_review"


class Severity(Enum):
    """Impact level of the issue."""

    CRITICAL = "critical"
    WARNING = "warning"
    INFO = "info"


# Default severity for each issue type
ISSUE_SEVERITY: dict[IssueType, Severity] = {
    IssueType.MISSING_TITLE: Severity.CRITICAL,
    IssueType.MISSING_LANGUAGE: Severity.CRITICAL,
    IssueType.HEADING_SKIP: Severity.WARNING,
    IssueType.HEADING_BOLD_NOT_STYLE: Severity.WARNING,
    IssueType.MISSING_ALT_TEXT: Severity.CRITICAL,
    IssueType.FILEPATH_ALT_TEXT: Severity.CRITICAL,
    IssueType.TABLE_NO_HEADER: Severity.CRITICAL,
    IssueType.TABLE_COMPLEX: Severity.WARNING,
    IssueType.WEAK_LINK_TEXT: Severity.WARNING,
    IssueType.MISSING_FOOTNOTE_LINK: Severity.INFO,
    IssueType.COLOR_ONLY_INFO: Severity.WARNING,
    IssueType.LOW_COLOR_CONTRAST: Severity.CRITICAL,
    IssueType.SMALL_FONT_SIZE: Severity.WARNING,
    IssueType.EMPTY_SPACER_PARAGRAPH: Severity.INFO,
    IssueType.UNSPACED_SUPERSCRIPT: Severity.WARNING,
    IssueType.SLIDE_MISSING_TITLE: Severity.CRITICAL,
    IssueType.SLIDE_READING_ORDER: Severity.WARNING,
}

# Default fix category for each issue type (without metadata)
ISSUE_DEFAULT_FIX: dict[IssueType, FixCategory] = {
    IssueType.MISSING_TITLE: FixCategory.SAFE_AUTO,
    IssueType.MISSING_LANGUAGE: FixCategory.SAFE_AUTO,
    IssueType.HEADING_SKIP: FixCategory.SUGGEST,
    IssueType.HEADING_BOLD_NOT_STYLE: FixCategory.SUGGEST,
    IssueType.MISSING_ALT_TEXT: FixCategory.HUMAN_REVIEW,
    IssueType.FILEPATH_ALT_TEXT: FixCategory.HUMAN_REVIEW,
    IssueType.TABLE_NO_HEADER: FixCategory.SAFE_AUTO,
    IssueType.TABLE_COMPLEX: FixCategory.HUMAN_REVIEW,
    IssueType.WEAK_LINK_TEXT: FixCategory.HUMAN_REVIEW,
    IssueType.MISSING_FOOTNOTE_LINK: FixCategory.SUGGEST,
    IssueType.COLOR_ONLY_INFO: FixCategory.HUMAN_REVIEW,
    IssueType.LOW_COLOR_CONTRAST: FixCategory.SUGGEST,
    IssueType.SMALL_FONT_SIZE: FixCategory.SAFE_AUTO,
    IssueType.EMPTY_SPACER_PARAGRAPH: FixCategory.SAFE_AUTO,
    IssueType.UNSPACED_SUPERSCRIPT: FixCategory.SAFE_AUTO,
    IssueType.SLIDE_MISSING_TITLE: FixCategory.SAFE_AUTO,
    IssueType.SLIDE_READING_ORDER: FixCategory.HUMAN_REVIEW,
}

# Weak link text patterns (case-insensitive match)
WEAK_LINK_PATTERNS: list[str] = [
    "click here",
    "here",
    "link",
    "read more",
    "more",
    "learn more",
    "this link",
    "this page",
]


@dataclass
class Issue:
    """A single accessibility issue found in a document."""

    issue_type: IssueType
    severity: Severity
    fix_category: FixCategory
    location: str
    description: str
    current_value: Optional[str] = None
    suggested_fix: Optional[str] = None
    metadata_key: Optional[str] = None
    fixed: bool = False

    @classmethod
    def create(
        cls,
        issue_type: IssueType,
        location: str,
        description: str,
        *,
        fix_category: Optional[FixCategory] = None,
        current_value: Optional[str] = None,
        suggested_fix: Optional[str] = None,
        metadata_key: Optional[str] = None,
    ) -> Issue:
        """Create an issue with default severity and fix category from lookups."""
        return cls(
            issue_type=issue_type,
            severity=ISSUE_SEVERITY.get(issue_type, Severity.WARNING),
            fix_category=fix_category or ISSUE_DEFAULT_FIX.get(
                issue_type, FixCategory.HUMAN_REVIEW
            ),
            location=location,
            description=description,
            current_value=current_value,
            suggested_fix=suggested_fix,
            metadata_key=metadata_key,
        )
