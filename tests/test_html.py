"""Tests for the HTML and Drupal WYSIWYG accessibility scanner and fixer."""

from __future__ import annotations

from click.testing import CliRunner

from ada_platform.cli.main import cli
from ada_platform.fixer.html_fixer import HtmlFixer
from ada_platform.models.issues import IssueType
from ada_platform.scanner.html_scanner import HtmlScanner


class TestHtmlScannerAndFixer:
    def test_full_html_page_scan_and_auto_fix(self, tmp_path):
        raw_html = """<!DOCTYPE html>
<html>
<head></head>
<body>
  <h1>Public Health Immunization Clinic</h1>
  <h3>Hours and Walk-In Schedule</h3>
  <p style="text-align: justify;">Bring your immunization record to the clinic.</p>
  <p>&nbsp;</p>
  <img src="/sites/default/files/clinic.jpg">
  <img src="/sites/default/files/divider.png" alt="decorative">
  <img src="/sites/default/files/spacer.gif" alt="">
  <table>
    <tr><td>Day</td><td>Hours</td></tr>
    <tr><td>Monday</td><td>9:00 AM - 4:00 PM</td></tr>
    <tr><td>Tuesday</td><td></td></tr>
  </table>
  <p><a href="https://publichealth.sccgov.org/services">click here</a></p>
</body>
</html>"""
        in_file = tmp_path / "drupal_page.html"
        out_file = tmp_path / "drupal_page_fixed.html"
        in_file.write_text(raw_html, encoding="utf-8")

        scanner = HtmlScanner()
        res = scanner.scan(str(in_file))
        issue_types = {i.issue_type for i in res.issues}

        assert IssueType.MISSING_TITLE in issue_types
        assert IssueType.MISSING_LANGUAGE in issue_types
        assert IssueType.HEADING_SKIP in issue_types
        assert IssueType.TEXT_JUSTIFIED in issue_types
        assert IssueType.EMPTY_SPACER_PARAGRAPH in issue_types
        assert IssueType.MISSING_ALT_TEXT in issue_types
        assert IssueType.DECORATIVE_IMAGE_HAS_ALT in issue_types
        assert IssueType.TABLE_NO_HEADER in issue_types
        assert IssueType.TABLE_EMPTY_CELLS in issue_types
        assert IssueType.WEAK_LINK_TEXT in issue_types
        # Image 3 has explicit alt="" so only Image 1 is missing alt text
        missing_alts = [i for i in res.issues if i.issue_type == IssueType.MISSING_ALT_TEXT]
        assert len(missing_alts) == 1
        assert missing_alts[0].location == "Image 1"

        # Apply safe auto-fixes
        fix_res = HtmlFixer().fix(str(in_file), str(out_file), res)
        fixed_types = {i.issue_type for i in fix_res.auto_fixed}
        assert IssueType.MISSING_TITLE in fixed_types
        assert IssueType.MISSING_LANGUAGE in fixed_types
        assert IssueType.DECORATIVE_IMAGE_HAS_ALT in fixed_types
        assert IssueType.TABLE_NO_HEADER in fixed_types
        assert IssueType.TEXT_JUSTIFIED in fixed_types
        assert IssueType.EMPTY_SPACER_PARAGRAPH in fixed_types

        # Re-scan fixed file: all auto-fixed issue types must be gone
        res2 = scanner.scan(str(out_file))
        remaining = {i.issue_type for i in res2.issues}
        assert IssueType.MISSING_TITLE not in remaining
        assert IssueType.MISSING_LANGUAGE not in remaining
        assert IssueType.DECORATIVE_IMAGE_HAS_ALT not in remaining
        assert IssueType.TABLE_NO_HEADER not in remaining
        assert IssueType.TEXT_JUSTIFIED not in remaining
        assert IssueType.EMPTY_SPACER_PARAGRAPH not in remaining

    def test_drupal_wysiwyg_fragment_alt_limits_merged_cells_and_contrast(self):
        long_alt = "Image of " + ("a" * 135)  # 144 chars (> 141 limit)
        fragment = f"""
<h2>Service Overview</h2>
<p><strong>Eligibility Requirements</strong></p>
<p style="color: #BBBBBB; background-color: #FFFFFF;">Low contrast text in red.</p>
<img src="banner.png" alt="{long_alt}">
<table>
  <tr><th scope="col" colspan="2">Merged Header</th></tr>
  <tr><td>Group A</td><td>100</td></tr>
</table>
"""
        scanner = HtmlScanner()
        res = scanner.scan_html_string(fragment)
        types = {i.issue_type for i in res.issues}

        # Fragment should NOT complain about missing <title> or <html lang>
        assert IssueType.MISSING_TITLE not in types
        assert IssueType.MISSING_LANGUAGE not in types

        # Should catch bold-as-heading, redundant prefix, >141 char alt, merged cells, low contrast, color-only
        assert IssueType.HEADING_BOLD_NOT_STYLE in types
        assert IssueType.ALT_TEXT_REDUNDANT_PREFIX in types
        assert IssueType.ALT_TEXT_TOO_LONG in types
        assert IssueType.TABLE_MERGED_CELLS in types
        assert IssueType.LOW_COLOR_CONTRAST in types
        assert IssueType.COLOR_ONLY_INFO in types

    def test_cli_scan_and_fix_html(self, tmp_path):
        html_file = tmp_path / "page.html"
        fixed_file = tmp_path / "page_fixed.html"
        html_file.write_text(
            "<html><body><h1>Title</h1><table><tr><td>A</td></tr></table></body></html>",
            encoding="utf-8",
        )
        runner = CliRunner()
        r_scan = runner.invoke(cli, ["scan", str(html_file)])
        assert r_scan.exit_code == 0

        r_fix = runner.invoke(cli, ["fix", str(html_file), "-o", str(fixed_file)])
        assert r_fix.exit_code == 0
        assert fixed_file.is_file()
        assert "<th" in fixed_file.read_text(encoding="utf-8")

    def test_cli_scan_and_fix_pdf_acrobat_equivalent(self, tmp_path):
        from test_remediation_engine import _write_raw_word_tagged_pdf

        raw_pdf = tmp_path / "raw_word.pdf"
        fixed_pdf = tmp_path / "fixed_acrobat.pdf"
        _write_raw_word_tagged_pdf(str(raw_pdf), "Acrobat Pro Equivalent Test")

        runner = CliRunner()
        r_scan = runner.invoke(cli, ["scan", str(raw_pdf)])
        assert r_scan.exit_code == 0
        assert "Table headers (TH cells)" in r_scan.output

        r_fix = runner.invoke(cli, ["fix", str(raw_pdf), "-o", str(fixed_pdf)])
        assert r_fix.exit_code == 0
        assert fixed_pdf.is_file()

        # Re-scanning the fixed PDF should show 0 issues
        r_rescan = runner.invoke(cli, ["scan", str(fixed_pdf)])
        assert r_rescan.exit_code == 0
        assert "No accessibility issues found" in r_rescan.output

