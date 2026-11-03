"""Tests for the deterministic ADA remediation engine, color contrast checker, and interactive plan."""

from __future__ import annotations

from docx import Document
from docx.shared import Pt, RGBColor
from lxml import etree

from ada_platform.models.issues import IssueType
from ada_platform.qa.transcript_verifier import verify_transcript
from ada_platform.remediation_engine import (
    InteractiveRemediationPlan,
    apply_interactive_remediation,
    audit_color_contrast,
    contrast_ratio,
    inspect_document,
    recommend_accessible_fg,
    relative_luminance,
)
from ada_platform.scanner.docx_scanner import DocxScanner

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"


class TestColorContrastEngine:
    def test_wcag_luminance_and_ratio(self):
        assert relative_luminance("#000000") == 0.0
        assert relative_luminance("#FFFFFF") == 1.0
        assert contrast_ratio("#000000", "#FFFFFF") == 21.0
        # PHD brand colors against white/black
        assert contrast_ratio("#FFFFFF", "#3E5872") == 7.38
        assert contrast_ratio("#004F93", "#FFFFFF") == 8.27
        assert contrast_ratio("#000000", "#A3B8CD") == 10.3
        assert contrast_ratio("#000000", "#CFDAE5") == 14.81

    def test_low_contrast_detected_by_scanner(self, tmp_path):
        doc = Document()
        doc.core_properties.title = "Contrast Test"
        p = doc.add_paragraph()
        r = p.add_run("Low contrast light gray text on white")
        r.font.color.rgb = RGBColor(0xBB, 0xBB, 0xBB)  # ~1.92:1 on white
        r.font.size = Pt(11)
        path = str(tmp_path / "low_contrast.docx")
        doc.save(path)

        scan = DocxScanner().scan(path)
        contrast_issues = [
            i for i in scan.issues if i.issue_type == IssueType.LOW_COLOR_CONTRAST
        ]
        assert len(contrast_issues) == 1
        assert "1.92:1" in contrast_issues[0].description
        assert contrast_issues[0].suggested_fix == "#000000"

    def test_recommend_accessible_fg(self):
        assert recommend_accessible_fg("#FFFFFF") == "#000000"
        assert recommend_accessible_fg("#3E5872") == "#FFFFFF"


class TestInteractiveRemediationRoundTrip:
    def test_inspect_and_apply_remediation(self, tmp_path):
        doc = Document()
        doc.core_properties.title = "Stale Template Title"
        doc.add_heading("Seniors’ Health Brief", level=1)

        # Skipped heading H1 -> H3
        doc.add_heading("References", level=3)

        # Small font + unspaced superscript
        p = doc.add_paragraph()
        r1 = p.add_run("Stroke")
        r1.font.size = Pt(9)
        r2 = p.add_run("6")
        r2.font.superscript = True

        # Empty spacer paragraph
        doc.add_paragraph("")

        # Raw URL hyperlink
        p_link = doc.add_paragraph()
        hl = etree.SubElement(p_link._element, f"{{{W_NS}}}hyperlink")
        r_id = doc.part.relate_to(
            "https://www.who.int/news-room/fact-sheets/detail/ageing-and-health",
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink",
            is_external=True,
        )
        hl.set(f"{{{R_NS}}}id", r_id)
        r_el = etree.SubElement(hl, f"{{{W_NS}}}r")
        t_el = etree.SubElement(r_el, f"{{{W_NS}}}t")
        t_el.text = "https://www.who.int/news-room/fact-sheets/detail/ageing-and-health"

        # Table without header flag
        tbl = doc.add_table(rows=2, cols=2)
        tbl.cell(0, 0).text = "Indicator"
        tbl.cell(0, 1).text = "Percent"
        tbl.cell(1, 0).text = "Male"
        tbl.cell(1, 1).text = "45"

        in_path = str(tmp_path / "draft.docx")
        out_path = str(tmp_path / "draft_ADA.docx")
        doc.save(in_path)

        insp = inspect_document(in_path)
        assert insp.title_needs_attention is True
        assert insp.recommended_title == "Seniors’ Health Brief"
        assert len(insp.headings_and_styles) == 1
        assert insp.headings_and_styles[0].recommended_style == "Heading 2"
        assert len(insp.hyperlinks) == 1
        assert insp.hyperlinks[0].is_weak_or_url is True
        assert "World Health Organization" in insp.hyperlinks[0].suggested_text
        assert insp.hygiene.small_font_runs_count >= 1
        assert insp.hygiene.unspaced_superscripts_count == 1
        assert len(insp.hygiene.empty_spacer_paras) == 1

        plan = InteractiveRemediationPlan(
            document_title=insp.recommended_title,
            paragraph_style_overrides={
                item.para_index: item.recommended_style
                for item in insp.headings_and_styles
            },
            hyperlink_text_overrides={
                lk.link_index: lk.suggested_text
                for lk in insp.hyperlinks
                if lk.is_weak_or_url
            },
            table_header_flags={t.table_index: True for t in insp.tables},
        )

        stats = apply_interactive_remediation(in_path, out_path, plan)
        assert stats["headings_updated"] == 1
        assert stats["links_updated"] == 1
        assert stats["tables_updated"] == 1
        assert stats["spacers_removed"] == 1

        # Post-scan and transcript verification must pass with 0 issues
        post_scan = DocxScanner().scan(out_path)
        assert post_scan.total_issues == 0
        ver = verify_transcript(out_path)
        assert ver.passed is True
        assert ver.gap_count == 0
