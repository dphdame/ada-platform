"""Generate JSON and plain-language accessibility reports."""

from __future__ import annotations

import json
from datetime import datetime, timezone
from typing import Any

from ada_platform.models.issues import Issue
from ada_platform.models.scan_result import FixResult, ScanResult


def scan_result_to_dict(result: ScanResult) -> dict[str, Any]:
    """Convert a ScanResult to a JSON-serializable dict."""
    return {
        "file_path": result.file_path,
        "file_type": result.file_type,
        "timestamp": result.timestamp,
        "summary": result.summary(),
        "total_issues": result.total_issues,
        "critical_count": result.critical_count,
        "counts": {
            "images": result.total_images,
            "tables": result.total_tables,
            "links": result.total_links,
            "headings": result.total_headings,
        },
        "issues": [_issue_to_dict(i) for i in result.issues],
    }


def fix_result_to_dict(result: FixResult) -> dict[str, Any]:
    """Convert a FixResult to a JSON-serializable dict."""
    return {
        "file_path": result.file_path,
        "output_path": result.output_path,
        "timestamp": result.timestamp,
        "summary": result.summary(),
        "auto_fixed": [_issue_to_dict(i) for i in result.auto_fixed],
        "needs_review": [_issue_to_dict(i) for i in result.needs_review],
        "skipped": [_issue_to_dict(i) for i in result.skipped],
    }


def _issue_to_dict(issue: Issue) -> dict[str, Any]:
    return {
        "type": issue.issue_type.value,
        "severity": issue.severity.value,
        "fix_category": issue.fix_category.value,
        "location": issue.location,
        "description": issue.description,
        "current_value": issue.current_value,
        "suggested_fix": issue.suggested_fix,
        "fixed": issue.fixed,
    }


def generate_json_report(result: FixResult | ScanResult, output_path: str) -> str:
    """Write a JSON report file and return the path."""
    if isinstance(result, FixResult):
        data = fix_result_to_dict(result)
    else:
        data = scan_result_to_dict(result)

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2, ensure_ascii=False)
    return output_path


def generate_text_report(result: FixResult) -> str:
    """Generate a plain-language text report from a FixResult."""
    lines: list[str] = []

    lines.append("=" * 60)
    lines.append("ACCESSIBILITY REPORT")
    lines.append("=" * 60)
    lines.append("")
    lines.append(f"Source: {result.file_path}")
    lines.append(f"Output: {result.output_path}")
    lines.append(f"Date:   {_format_timestamp(result.timestamp)}")
    lines.append("")
    lines.append(result.summary())
    lines.append("")

    if result.auto_fixed:
        lines.append("-" * 40)
        lines.append(f"FIXED AUTOMATICALLY ({len(result.auto_fixed)} items)")
        lines.append("-" * 40)
        for i, issue in enumerate(result.auto_fixed, 1):
            lines.append(f"  {i}. [{issue.location}] {issue.description}")
            if issue.suggested_fix:
                lines.append(f"     -> Set to: {issue.suggested_fix}")
        lines.append("")

    if result.needs_review:
        lines.append("-" * 40)
        lines.append(f"NEEDS YOUR REVIEW ({len(result.needs_review)} items)")
        lines.append("-" * 40)
        for i, issue in enumerate(result.needs_review, 1):
            lines.append(f"  {i}. [{issue.location}] {issue.description}")
            if issue.current_value:
                lines.append(f"     Current: {issue.current_value}")
            if issue.suggested_fix:
                lines.append(f"     Suggestion: {issue.suggested_fix}")
        lines.append("")

    if result.skipped:
        lines.append("-" * 40)
        lines.append(f"INFORMATIONAL ({len(result.skipped)} items)")
        lines.append("-" * 40)
        for i, issue in enumerate(result.skipped, 1):
            lines.append(f"  {i}. [{issue.location}] {issue.description}")
        lines.append("")

    if not result.auto_fixed and not result.needs_review and not result.skipped:
        lines.append("No accessibility issues found. Document looks good.")
        lines.append("")

    lines.append("=" * 60)
    return "\n".join(lines)


def generate_scan_text_report(result: ScanResult) -> str:
    """Generate a plain-language text report from a ScanResult (pre-fix)."""
    lines: list[str] = []

    lines.append("=" * 60)
    lines.append("ACCESSIBILITY SCAN RESULTS")
    lines.append("=" * 60)
    lines.append("")
    lines.append(f"File: {result.file_path}")
    lines.append(f"Type: {result.file_type.upper()}")
    lines.append(f"Date: {_format_timestamp(result.timestamp)}")
    lines.append("")
    lines.append(f"Found: {result.total_images} images, {result.total_tables} tables, "
                 f"{result.total_links} links, {result.total_headings} headings")
    lines.append("")
    lines.append(result.summary())
    lines.append("")

    if result.safe_fixable:
        lines.append(f"Can fix automatically: {len(result.safe_fixable)}")
        for issue in result.safe_fixable:
            lines.append(f"  - [{issue.location}] {issue.description}")

    if result.needs_review:
        lines.append(f"\nNeed your review: {len(result.needs_review)}")
        for issue in result.needs_review:
            lines.append(f"  - [{issue.location}] {issue.description}")

    if result.suggested:
        lines.append(f"\nSuggested fixes: {len(result.suggested)}")
        for issue in result.suggested:
            lines.append(f"  - [{issue.location}] {issue.description}")

    lines.append("")
    lines.append("=" * 60)
    return "\n".join(lines)


def _format_timestamp(ts: str) -> str:
    """Format an ISO timestamp to a readable date string."""
    try:
        dt = datetime.fromisoformat(ts)
        return dt.strftime("%Y-%m-%d %H:%M UTC")
    except (ValueError, TypeError):
        return ts
