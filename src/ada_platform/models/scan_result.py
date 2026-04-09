"""Scan results, fix results, and audit log models."""

from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Optional

from ada_platform.models.issues import FixCategory, Issue, Severity


@dataclass
class ScanResult:
    """Result of scanning a document for accessibility issues."""

    file_path: str
    file_type: str
    issues: list[Issue] = field(default_factory=list)
    timestamp: str = field(
        default_factory=lambda: datetime.now(timezone.utc).isoformat()
    )
    total_images: int = 0
    total_tables: int = 0
    total_links: int = 0
    total_headings: int = 0

    @property
    def safe_fixable(self) -> list[Issue]:
        """Issues that can be auto-fixed with high confidence."""
        return [
            i for i in self.issues
            if i.fix_category == FixCategory.SAFE_AUTO and not i.fixed
        ]

    @property
    def suggested(self) -> list[Issue]:
        """Issues with a suggested fix that needs user confirmation."""
        return [
            i for i in self.issues
            if i.fix_category == FixCategory.SUGGEST and not i.fixed
        ]

    @property
    def needs_review(self) -> list[Issue]:
        """Issues that require human judgment."""
        return [
            i for i in self.issues
            if i.fix_category == FixCategory.HUMAN_REVIEW and not i.fixed
        ]

    @property
    def critical_count(self) -> int:
        """Number of unfixed critical issues."""
        return sum(
            1 for i in self.issues
            if i.severity == Severity.CRITICAL and not i.fixed
        )

    @property
    def total_issues(self) -> int:
        return len(self.issues)

    @property
    def unfixed_count(self) -> int:
        return sum(1 for i in self.issues if not i.fixed)

    def summary(self) -> str:
        """One-line plain-language summary."""
        parts = []
        n_auto = len(self.safe_fixable)
        n_review = len(self.needs_review)
        n_suggest = len(self.suggested)
        if n_auto:
            parts.append(f"{n_auto} can be fixed automatically")
        if n_suggest:
            parts.append(f"{n_suggest} have suggested fixes")
        if n_review:
            parts.append(f"{n_review} need your review")
        if not parts:
            return "No accessibility issues found."
        return f"Found {self.total_issues} issues: {', '.join(parts)}."


@dataclass
class FixResult:
    """Result of applying fixes to a document."""

    file_path: str
    output_path: str
    auto_fixed: list[Issue] = field(default_factory=list)
    needs_review: list[Issue] = field(default_factory=list)
    skipped: list[Issue] = field(default_factory=list)
    timestamp: str = field(
        default_factory=lambda: datetime.now(timezone.utc).isoformat()
    )

    def summary(self) -> str:
        """One-line plain-language summary."""
        parts = []
        if self.auto_fixed:
            parts.append(f"Fixed {len(self.auto_fixed)} issues automatically")
        if self.needs_review:
            parts.append(f"{len(self.needs_review)} items need your review")
        if self.skipped:
            parts.append(f"{len(self.skipped)} skipped")
        if not parts:
            return "No changes needed."
        return ". ".join(parts) + "."


@dataclass
class AuditLogEntry:
    """A single entry in the processing audit log."""

    action: str
    detail: str = ""
    issue_type: Optional[str] = None
    user_decision: Optional[str] = None
    timestamp: str = field(
        default_factory=lambda: datetime.now(timezone.utc).isoformat()
    )
