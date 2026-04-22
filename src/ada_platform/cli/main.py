"""CLI for the ADA Document Accessibility Platform."""

from __future__ import annotations

import os
import sys

import click

from ada_platform.scanner.docx_scanner import DocxScanner
from ada_platform.fixer.docx_fixer import DocxFixer
from ada_platform.metadata.excel_reader import ExcelMetadataReader, create_template_workbook
from ada_platform.charts.chart_engine import ChartSpec, generate_chart_accessibility, read_csv_data
from ada_platform.reporting.exception_report import (
    generate_json_report,
    generate_text_report,
    generate_scan_text_report,
)
from ada_platform.qa.screen_reader_sim import generate_transcript
from ada_platform.qa.transcript_verifier import verify_transcript


@click.group()
@click.version_option(version="0.1.0", prog_name="ada-platform")
def cli():
    """ADA Document Accessibility Platform.

    Scan, fix, and report on document accessibility issues.
    """
    pass


@cli.command()
@click.argument("file_path", type=click.Path(exists=True))
@click.option("--metadata", "-m", type=click.Path(exists=True), help="Excel metadata workbook")
@click.option("--output", "-o", type=click.Path(), help="Output report path (JSON)")
def scan(file_path: str, metadata: str | None, output: str | None):
    """Scan a DOCX file for accessibility issues."""
    meta_wb = None
    if metadata:
        meta_wb = ExcelMetadataReader().read(metadata)
        click.echo(f"Loaded metadata from {metadata}")

    scanner = DocxScanner()
    result = scanner.scan(file_path, metadata=meta_wb)

    click.echo()
    click.echo(generate_scan_text_report(result))

    if output:
        generate_json_report(result, output)
        click.echo(f"\nJSON report saved to {output}")


@cli.command()
@click.argument("file_path", type=click.Path(exists=True))
@click.option("--metadata", "-m", type=click.Path(exists=True), help="Excel metadata workbook")
@click.option("--output", "-o", type=click.Path(), help="Output DOCX path", required=True)
@click.option("--report", "-r", type=click.Path(), help="Output report path (JSON)")
def fix(file_path: str, metadata: str | None, output: str, report: str | None):
    """Scan and fix a DOCX file, saving the remediated version."""
    meta_wb = None
    if metadata:
        meta_wb = ExcelMetadataReader().read(metadata)
        click.echo(f"Loaded metadata from {metadata}")

    scanner = DocxScanner()
    scan_result = scanner.scan(file_path, metadata=meta_wb)

    fixer = DocxFixer()
    fix_result = fixer.fix(file_path, output, scan_result, metadata=meta_wb)

    click.echo()
    click.echo(generate_text_report(fix_result))
    click.echo(f"\nRemediated file saved to {output}")

    if report:
        generate_json_report(fix_result, report)
        click.echo(f"JSON report saved to {report}")

    # Screen reader QI check
    t = generate_transcript(output)
    if t.warnings:
        click.echo(f"\nScreen reader QI: {t.summary()}")
        for w in t.warnings:
            click.echo(f"  [{w.location}] {w.text}")
    else:
        click.echo(f"\nScreen reader QI: {t.summary()}")


@cli.command()
@click.argument("file_path", type=click.Path(exists=True))
@click.option("--metadata", "-m", type=click.Path(exists=True), help="Excel metadata workbook")
def wizard(file_path: str, metadata: str | None):
    """Interactive accessibility wizard: scan, review, fix, export."""
    meta_wb = None
    if metadata:
        meta_wb = ExcelMetadataReader().read(metadata)
        click.echo(f"Loaded metadata from {metadata}")

    # Step 1: Scan
    click.echo("\nStep 1: Scanning document...")
    scanner = DocxScanner()
    scan_result = scanner.scan(file_path, metadata=meta_wb)
    click.echo(scan_result.summary())

    if scan_result.total_issues == 0:
        click.echo("\nNo issues found. Document looks good.")
        return

    # Step 2: Review
    click.echo(f"\n{len(scan_result.safe_fixable)} issues can be fixed automatically.")
    if scan_result.needs_review:
        click.echo(f"{len(scan_result.needs_review)} issues need your review:")
        for issue in scan_result.needs_review:
            click.echo(f"  - [{issue.location}] {issue.description}")

    # Step 3: Fix
    if scan_result.safe_fixable:
        proceed = click.confirm("\nApply automatic fixes?", default=True)
        if not proceed:
            click.echo("Skipping fixes.")
            return

    base, ext = os.path.splitext(file_path)
    output_path = f"{base}_accessible{ext}"
    report_path = f"{base}_report.json"

    fixer = DocxFixer()
    fix_result = fixer.fix(file_path, output_path, scan_result, metadata=meta_wb)

    # Step 4: Export
    click.echo(f"\n{fix_result.summary()}")
    click.echo(f"\nRemediated file: {output_path}")

    generate_json_report(fix_result, report_path)
    click.echo(f"Report: {report_path}")

    # Show text report
    click.echo()
    click.echo(generate_text_report(fix_result))

    # Step 5: Screen reader transcript QI
    click.echo("\nStep 5: Screen reader preview...")
    transcript = generate_transcript(output_path)
    if transcript.warnings:
        click.echo(f"\n{transcript.summary()}")
        click.echo("Remaining issues a screen reader user would encounter:")
        for w in transcript.warnings:
            click.echo(f"  [{w.location}] {w.text}")
    else:
        click.echo(transcript.summary())

    # Save transcript
    transcript_path = f"{base}_transcript.txt"
    with open(transcript_path, "w", encoding="utf-8") as f:
        f.write(transcript.text_only)
    click.echo(f"Transcript saved to {transcript_path}")


@cli.command()
@click.argument("file_path", type=click.Path(exists=True))
def transcript(file_path: str):
    """Generate a screen reader transcript for a DOCX file.

    Shows how the document would be read by assistive technology,
    including warnings for accessibility issues.
    """
    t = generate_transcript(file_path)
    click.echo(t.text_only)
    click.echo()
    click.echo(t.summary())
    if t.warnings:
        click.echo(f"\nWarnings ({len(t.warnings)}):")
        for w in t.warnings:
            click.echo(f"  [{w.location}] {w.text}")


@cli.command()
@click.argument("file_path", type=click.Path(exists=True))
def verify(file_path: str):
    """Verify that the screen reader transcript matches the document.

    Independently reads the DOCX and compares every element against
    the transcript to catch missing, wrong, or out-of-order content.
    """
    report = verify_transcript(file_path)
    click.echo(report.full_report())
    if report.passed:
        raise SystemExit(0)
    else:
        raise SystemExit(1)


@cli.command()
@click.argument("csv_path", type=click.Path(exists=True))
@click.option("--type", "chart_type", required=True, help="Chart type (bar, line, pie, etc.)")
@click.option("--title", required=True, help="Chart title")
@click.option("--x", "x_axis", required=True, help="X-axis label")
@click.option("--y", "y_axis", required=True, help="Y-axis label")
@click.option("--takeaway", default="", help="Key takeaway sentence")
def chart(csv_path: str, chart_type: str, title: str, x_axis: str, y_axis: str, takeaway: str):
    """Generate accessibility text for a chart from CSV data."""
    data = read_csv_data(csv_path)
    spec = ChartSpec(
        chart_type=chart_type,
        title=title,
        x_axis=x_axis,
        y_axis=y_axis,
        key_takeaway=takeaway,
        data=data,
    )
    result = generate_chart_accessibility(spec)

    click.echo("SHORT ALT TEXT:")
    click.echo(f"  {result.short_alt}")
    click.echo()
    click.echo("LONG DESCRIPTION:")
    click.echo(f"  {result.long_description}")
    click.echo()
    if result.data_table_headers:
        click.echo("DATA TABLE:")
        click.echo(f"  {' | '.join(result.data_table_headers)}")
        click.echo(f"  {'-+-'.join('-' * len(h) for h in result.data_table_headers)}")
        for row in result.data_table_rows:
            click.echo(f"  {' | '.join(row)}")


@cli.command("init-workbook")
@click.option("--output", "-o", type=click.Path(), default="metadata.xlsx", help="Output path")
def init_workbook(output: str):
    """Create an empty metadata workbook template."""
    create_template_workbook(output)
    click.echo(f"Created metadata workbook template: {output}")
    click.echo("Sheets: Documents, Images, Charts, Links, Tables, Slides")


if __name__ == "__main__":
    cli()
