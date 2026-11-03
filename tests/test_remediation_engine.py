"""Tests for the deterministic ADA remediation engine, color contrast checker, and interactive plan."""

from __future__ import annotations

from docx import Document
from docx.shared import Inches, Pt, RGBColor
from lxml import etree
import pymupdf as fitz

from ada_platform.models.issues import IssueType
from ada_platform.qa.screen_reader_sim import generate_transcript
from ada_platform.qa.transcript_verifier import verify_transcript
from ada_platform.remediation_engine import (
    ADEC_NS,
    REQUIRED_PDF_AUDIT_CHECKS,
    WP_NS,
    InteractiveRemediationPlan,
    apply_interactive_remediation,
    audit_color_contrast,
    build_tagged_pdf_from_docx,
    contrast_ratio,
    evaluate_deliverable_gate,
    inspect_document,
    recommend_accessible_fg,
    relative_luminance,
    run_pdf_ua_audit,
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

    def test_small_heading_requires_4_5_contrast_threshold(self, tmp_path):
        """A small heading (<14pt bold / <18pt regular) must require 4.5:1, not 3.0:1."""
        doc = Document()
        doc.core_properties.title = "Small Heading Contrast Test"
        h3 = doc.add_heading("", level=3)
        r = h3.add_run("Small 12pt bold heading with 3.54:1 gray")
        r.font.size = Pt(12)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0x88, 0x88, 0x88)  # 3.54:1 on white (< 4.5:1)

        items = audit_color_contrast(doc)
        h3_items = [i for i in items if i.role == "Heading 3"]
        assert len(h3_items) == 1
        assert h3_items[0].required_ratio == 4.5
        assert h3_items[0].passed is False

    def test_table_cell_second_run_low_contrast_is_caught(self, tmp_path):
        """A later low-contrast run inside a table cell must be evaluated and flagged."""
        doc = Document()
        doc.core_properties.title = "Table Multi-Run Contrast Test"
        tbl = doc.add_table(rows=2, cols=1)
        tbl.cell(0, 0).text = "Header"
        p_cell = tbl.cell(1, 0).paragraphs[0]
        r_good = p_cell.add_run("High contrast first run ")
        r_good.font.color.rgb = RGBColor(0x00, 0x00, 0x00)
        r_good.font.size = Pt(11)
        r_bad = p_cell.add_run("(low contrast second run)")
        r_bad.font.color.rgb = RGBColor(0xBB, 0xBB, 0xBB)  # 1.92:1 on white
        r_bad.font.size = Pt(11)

        items = audit_color_contrast(doc)
        failing = [i for i in items if not i.passed]
        assert len(failing) == 1
        assert failing[0].fg_hex == "#BBBBBB"
        assert failing[0].required_ratio == 4.5


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

    def test_multipage_decorative_images_stay_in_place_with_unique_ids(self, tmp_path):
        # 1x1 PNG bytes for test images
        png_bytes = (
            b"\x89PNG\r\n\x1a\n\x00\x00\x00\rIHDR\x00\x00\x00\x01\x00\x00\x00\x01"
            b"\x08\x02\x00\x00\x00\x90wS\xde\x00\x00\x00\x0cIDATx\x9cc\xf8\xcf\xc0"
            b"\x00\x00\x03\x01\x01\x00\xc9\xfe\x92\xef\x00\x00\x00\x00IEND\xaeB`\x82"
        )
        img_file = tmp_path / "dot.png"
        img_file.write_bytes(png_bytes)

        doc = Document()
        doc.core_properties.title = "Multi-Page Decorative Image Test"
        doc.add_heading("Page 1 Section", level=1)
        p1 = doc.add_paragraph("Paragraph with first decorative image on page 1.")
        p1.add_run().add_picture(str(img_file), width=Inches(0.5))

        doc.add_page_break()
        doc.add_heading("Page 2 Section", level=2)
        p2 = doc.add_paragraph("Paragraph with second decorative image on page 2.")
        p2.add_run().add_picture(str(img_file), width=Inches(0.5))

        # Force duplicate docPr IDs ("991") before remediation to test deduplication
        for dp in doc.element.iter(f"{{{WP_NS}}}docPr"):
            dp.set("id", "991")

        in_path = str(tmp_path / "multipage_dec.docx")
        out_path = str(tmp_path / "multipage_dec_ADA.docx")
        doc.save(in_path)

        plan = InteractiveRemediationPlan(
            document_title="Multi-Page Decorative Image Test",
            image_decisions={
                1: {"decorative": True, "alt_text": ""},
                2: {"decorative": True, "alt_text": ""},
            },
        )
        apply_interactive_remediation(in_path, out_path, plan)

        out_doc = Document(out_path)
        body_drawings = out_doc.element.body.findall(f".//{{{W_NS}}}drawing")
        assert len(body_drawings) == 2, "Both decorative images must remain in-place in <w:body>"

        ids = []
        for d in body_drawings:
            dp = d.find(f".//{{{WP_NS}}}docPr")
            assert dp is not None
            assert dp.get("descr") == ""
            dec_els = list(dp.iter(f"{{{ADEC_NS}}}decorative"))
            assert any(el.get("val") == "1" for el in dec_els)
            ids.append(int(dp.get("id")))

        assert len(ids) == len(set(ids)), f"wp:docPr IDs must be strictly unique, got {ids}"

        # Scanner, screen reader simulator, and verifier must all pass with 0 warnings/issues
        post_scan = DocxScanner().scan(out_path)
        assert post_scan.total_issues == 0
        sr = generate_transcript(out_path)
        assert len(sr.warnings) == 0
        ver = verify_transcript(out_path)
        assert ver.passed is True
        assert ver.gap_count == 0

    def test_non_english_language_propagates_to_docx_and_pdf(self, tmp_path):
        doc = Document()
        doc.add_heading("Informe de Salud de Adultos Mayores", level=1)
        doc.add_paragraph("Resumen de indicadores de salud pública del condado.")
        in_path = str(tmp_path / "es_doc.docx")
        out_docx = str(tmp_path / "es_doc_ADA.docx")
        out_pdf = str(tmp_path / "es_doc_ADA.pdf")
        doc.save(in_path)

        plan = InteractiveRemediationPlan(
            document_title="Informe de Salud de Adultos Mayores",
            document_language="es-US",
        )
        apply_interactive_remediation(in_path, out_docx, plan)

        qa_report = build_tagged_pdf_from_docx(
            out_docx,
            out_pdf,
            work_dir=str(tmp_path / "pdf_es_build"),
        )
        assert qa_report["all_passed"] is True

        pdf_doc = fitz.open(out_pdf)
        cat_lang = pdf_doc.xref_get_key(pdf_doc.pdf_catalog(), "Lang")
        pdf_doc.close()
        assert cat_lang[1].strip("()") == "es-US"

        # Verify language mismatch is caught by run_pdf_ua_audit
        mismatch_report = run_pdf_ua_audit(out_pdf, expected_lang="vi-VN")
        assert mismatch_report["all_passed"] is False

    def test_pdf_audit_fails_closed_and_deliverable_gate_blocks(self, tmp_path):
        # 1. Non-existent PDF fails closed
        missing_rep = run_pdf_ua_audit(str(tmp_path / "does_not_exist.pdf"))
        assert missing_rep["all_passed"] is False
        assert missing_rep["critical_failures"] >= 1

        # 2. Untagged blank PDF fails closed
        untagged_path = str(tmp_path / "untagged.pdf")
        raw_pdf = fitz.open()
        raw_pdf.new_page()
        raw_pdf.save(untagged_path)
        raw_pdf.close()

        untagged_rep = run_pdf_ua_audit(untagged_path)
        assert untagged_rep["all_passed"] is False
        assert untagged_rep["critical_failures"] >= 1

        # 3. Deliverable gate blocks certification whenever any check fails
        class DummyScan:
            total_issues = 1

        class DummySR:
            warnings = ["missing_alt_text"]

        class DummyVer:
            passed = False
            gap_count = 1

            def summary(self):
                return "FAIL: 1 error"

        gate = evaluate_deliverable_gate(DummyScan(), DummySR(), DummyVer(), untagged_rep)
        assert gate["passed"] is False
        assert len(gate["blocking_reasons"]) == 4
        assert "BLOCKED" in gate["status_label"]

        # 4. Deliverable gate rejects an incomplete 15-check set even if all_passed=True
        class CleanScan:
            total_issues = 0

        class CleanSR:
            warnings = []

        class CleanVer:
            passed = True
            gap_count = 0

            def summary(self):
                return "PASS"

        incomplete_15_report = {
            "all_passed": True,
            "critical_failures": 0,
            "warnings": 0,
            "checks": [
                {"check": name, "passed": True, "detail": "ok", "severity": "critical"}
                for name in REQUIRED_PDF_AUDIT_CHECKS[:-1]
            ],
        }
        gate_incomplete = evaluate_deliverable_gate(
            CleanScan(), CleanSR(), CleanVer(), incomplete_15_report
        )
        assert gate_incomplete["passed"] is False
        assert any("missing 1 required check" in r for r in gate_incomplete["blocking_reasons"])

        # 5. Deliverable gate rejects spoofed all_passed=True when any required check is False
        spoofed_report = {
            "all_passed": True,
            "critical_failures": 0,
            "warnings": 0,
            "checks": [
                {
                    "check": name,
                    "passed": (name != "Table headers (TH cells)"),
                    "detail": "test",
                    "severity": "critical",
                }
                for name in REQUIRED_PDF_AUDIT_CHECKS
            ],
        }
        gate_spoofed = evaluate_deliverable_gate(
            CleanScan(), CleanSR(), CleanVer(), spoofed_report
        )
        assert gate_spoofed["passed"] is False
        assert any("Table headers (TH cells)" in r for r in gate_spoofed["blocking_reasons"])
