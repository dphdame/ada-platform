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

    def test_non_english_language_propagates_to_docx_and_pdf(self, tmp_path, monkeypatch):
        import ada_platform.remediation_engine as rem_eng

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

        # On platforms without lualatex (e.g., Windows CI sandbox), use the Word backend with a simulated raw Word PDF
        try:
            rem_eng.resolve_lualatex_binary()
            has_lualatex = True
        except RuntimeError:
            has_lualatex = False

        if not has_lualatex:
            monkeypatch.setattr(
                rem_eng,
                "export_docx_to_tagged_pdf_via_word_windows",
                lambda _d, pdf_out: _write_raw_word_tagged_pdf(
                    pdf_out, "Informe de Salud de Adultos Mayores"
                ),
            )

        qa_report = build_tagged_pdf_from_docx(
            out_docx,
            out_pdf,
            work_dir=str(tmp_path / "pdf_es_build"),
            backend="latex" if has_lualatex else "word",
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

    def test_windows_word_backend_postprocesses_td_headers_and_empty_alt(self, tmp_path, monkeypatch):
        """Verify backend='word' converts DOCX via Word exporter and fixes raw Word PDF tags."""
        import ada_platform.remediation_engine as rem_eng

        doc = Document()
        doc.core_properties.title = "Windows Word Native Export Test"
        doc.add_heading("Windows Word Native Export Test", level=1)
        doc.add_paragraph("Body paragraph for Windows Word native PDF export.")
        tbl = doc.add_table(rows=2, cols=2)
        tbl.cell(0, 0).text = "Metric"
        tbl.cell(0, 1).text = "Value"
        tbl.cell(1, 0).text = "Seniors"
        tbl.cell(1, 1).text = "20%"

        in_docx = str(tmp_path / "win_in.docx")
        rem_docx = str(tmp_path / "win_ADA.docx")
        out_pdf = str(tmp_path / "win_ADA.pdf")
        doc.save(in_docx)

        plan = InteractiveRemediationPlan(
            document_title="Windows Word Native Export Test",
            document_language="en-US",
            table_header_flags={1: True},
        )
        apply_interactive_remediation(in_docx, rem_docx, plan)

        # Simulate raw Microsoft Word PDF export with the two classic raw Word PDF quirks:
        #   (a) Table header cells tagged as /TD without /Scope
        #   (b) A decorative /Shape StructElem with empty /Alt () ("Element 1" in Acrobat Pro)
        raw_unprocessed = str(tmp_path / "raw_word_unprocessed.pdf")
        _write_raw_word_tagged_pdf(raw_unprocessed, "Windows Word Native Export Test")
        raw_audit = run_pdf_ua_audit(raw_unprocessed, expected_lang="en-US")
        assert raw_audit["all_passed"] is False, "Raw Word PDF before post-processing must fail audit"

        monkeypatch.setattr(
            rem_eng,
            "export_docx_to_tagged_pdf_via_word_windows",
            lambda _docx_arg, pdf_arg: _write_raw_word_tagged_pdf(
                pdf_arg, "Windows Word Native Export Test"
            ),
        )

        report = build_tagged_pdf_from_docx(
            rem_docx,
            out_pdf,
            work_dir=str(tmp_path / "win_build"),
            backend="word",
        )
        assert report["all_passed"] is True
        assert [c["check"] for c in report["checks"]] == list(REQUIRED_PDF_AUDIT_CHECKS)

    def test_windows_powershell_com_script_invocation(self, tmp_path, monkeypatch):
        """Verify export_docx_to_tagged_pdf_via_word_windows invokes PowerShell COM ExportAsFixedFormat."""
        import subprocess
        import sys
        import ada_platform.remediation_engine as rem_eng

        # Ensure win32com import falls through to the built-in PowerShell COM automation path
        monkeypatch.setitem(sys.modules, "win32com", None)
        monkeypatch.setitem(sys.modules, "win32com.client", None)

        in_docx = tmp_path / "sample.docx"
        out_pdf = tmp_path / "sample_out.pdf"
        Document().save(str(in_docx))

        captured_cmds: list[list[str]] = []

        def _fake_powershell_run(cmd, **kwargs):
            captured_cmds.append(list(cmd))
            _write_raw_word_tagged_pdf(str(out_pdf), "PowerShell COM Export")
            return subprocess.CompletedProcess(cmd, 0, stdout="", stderr="")

        monkeypatch.setattr(rem_eng.subprocess, "run", _fake_powershell_run)

        rem_eng.export_docx_to_tagged_pdf_via_word_windows(str(in_docx), str(out_pdf))
        assert len(captured_cmds) == 1
        ps_cmd = captured_cmds[0]
        assert ps_cmd[0] == "powershell"
        assert "-NoProfile" in ps_cmd
        assert "-NonInteractive" in ps_cmd
        script_body = ps_cmd[-1]
        assert "New-Object -ComObject Word.Application" in script_body
        assert "ExportAsFixedFormat" in script_body
        assert out_pdf.is_file()

    def test_win32com_word_export_call_contract_and_auto_routing(self, tmp_path, monkeypatch):
        """Verify sys.platform == 'win32' auto-routes to Word COM via win32com.client.DispatchEx."""
        import types
        import sys
        import ada_platform.remediation_engine as rem_eng

        com_calls: list[tuple] = []

        class FakeWordDoc:
            def ExportAsFixedFormat(self, **kwargs):
                com_calls.append(("ExportAsFixedFormat", kwargs))
                _write_raw_word_tagged_pdf(kwargs["OutputFileName"], "Win32COM Auto Route Test")

            def Close(self, save_changes):
                com_calls.append(("Close", save_changes))

        class FakeDocuments:
            def Open(self, path, ReadOnly=True):
                com_calls.append(("Open", path, ReadOnly))
                return FakeWordDoc()

        class FakeWordApp:
            def __init__(self):
                self.Visible = True
                self.DisplayAlerts = 1
                self.Documents = FakeDocuments()

            def Quit(self):
                com_calls.append(("Quit", self.Visible, self.DisplayAlerts))

        fake_client = types.ModuleType("win32com.client")
        fake_client.DispatchEx = lambda prog_id: (
            com_calls.append(("DispatchEx", prog_id)) or FakeWordApp()
        )
        fake_win32com = types.ModuleType("win32com")
        fake_win32com.client = fake_client

        monkeypatch.setitem(sys.modules, "win32com", fake_win32com)
        monkeypatch.setitem(sys.modules, "win32com.client", fake_client)
        monkeypatch.setattr(rem_eng.sys, "platform", "win32")

        doc = Document()
        doc.core_properties.title = "Win32COM Auto Route Test"
        doc.add_heading("Win32COM Auto Route Test", level=1)
        doc.add_paragraph("Testing automatic Windows Word COM routing.")
        in_docx = str(tmp_path / "win32_auto.docx")
        out_pdf = str(tmp_path / "win32_auto.pdf")
        doc.save(in_docx)

        report = build_tagged_pdf_from_docx(
            in_docx,
            out_pdf,
            work_dir=str(tmp_path / "win32_work"),
            backend="auto",
        )
        assert report["all_passed"] is True
        assert com_calls[0] == ("DispatchEx", "Word.Application")
        assert com_calls[1][0] == "Open" and com_calls[1][2] is True
        export_kwargs = com_calls[2][1]
        assert export_kwargs["ExportFormat"] == 17
        assert export_kwargs["IncludeDocProps"] is True
        assert export_kwargs["CreateBookmarks"] == 1
        assert export_kwargs["DocStructureTags"] is True
        assert com_calls[3] == ("Close", 0)
        assert com_calls[4] == ("Quit", False, 0)


def _write_raw_word_tagged_pdf(pdf_path: str, title: str) -> None:
    """Create a realistic raw Word-exported tagged PDF with /TD header cells and empty /Alt () shape."""
    doc = fitz.open()
    page = doc.new_page(width=612, height=792)
    page.insert_text((72, 100), title, fontsize=16)
    page.insert_text((72, 130), "Body paragraph for Windows Word native PDF export.", fontsize=11)
    doc.set_toc([[1, title, 1]])

    stream_xref = page.get_contents()[0]
    orig_bytes = doc.xref_stream(stream_xref)
    doc.update_stream(stream_xref, b"/P <</MCID 0>> BDC\n" + orig_bytes + b"\nEMC\n")

    st_root = doc.get_new_xref()
    p_elem = doc.get_new_xref()
    shape_elem = doc.get_new_xref()
    tbl_elem = doc.get_new_xref()
    tr_elem = doc.get_new_xref()
    td1_elem = doc.get_new_xref()
    td2_elem = doc.get_new_xref()

    doc.update_object(td1_elem, f"<< /Type /StructElem /S /TD /P {tr_elem} 0 R >>")
    doc.update_object(td2_elem, f"<< /Type /StructElem /S /TD /P {tr_elem} 0 R >>")
    doc.update_object(
        tr_elem,
        f"<< /Type /StructElem /S /TR /P {tbl_elem} 0 R /K [{td1_elem} 0 R {td2_elem} 0 R] >>",
    )
    doc.update_object(
        tbl_elem,
        f"<< /Type /StructElem /S /Table /P {st_root} 0 R /K [{tr_elem} 0 R] >>",
    )
    doc.update_object(p_elem, f"<< /Type /StructElem /S /P /P {st_root} 0 R /K 0 >>")
    doc.update_object(
        shape_elem, f"<< /Type /StructElem /S /Shape /P {st_root} 0 R /Alt () >>"
    )
    doc.update_object(
        st_root,
        f"<< /Type /StructTreeRoot /K [{p_elem} 0 R {shape_elem} 0 R {tbl_elem} 0 R] >>",
    )
    doc.xref_set_key(doc.pdf_catalog(), "StructTreeRoot", f"{st_root} 0 R")
    doc.save(pdf_path)
    doc.close()
