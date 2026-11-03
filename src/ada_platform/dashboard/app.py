"""Streamlit dashboard for the ADA Document Accessibility Platform (Zero-AI Interactive Edition)."""

from __future__ import annotations

import json
import os
import tempfile
from pathlib import Path

import streamlit as st

from ada_platform.metadata.excel_reader import ExcelMetadataReader
from ada_platform.qa.screen_reader_sim import generate_transcript
from ada_platform.qa.transcript_verifier import verify_transcript
from ada_platform.remediation_engine import (
    PHD_BRAND_PALETTE,
    InteractiveRemediationPlan,
    apply_interactive_remediation,
    build_tagged_pdf_from_docx,
    contrast_ratio,
    evaluate_deliverable_gate,
    inspect_document,
)
from ada_platform.scanner.docx_scanner import DocxScanner


def _render_color_swatch_html(fg_hex: str, bg_hex: str, label: str) -> str:
    return (
        f'<span style="display:inline-block;padding:3px 10px;border-radius:4px;'
        f"background-color:{bg_hex};color:{fg_hex};border:1px solid #ccc;"
        f'font-weight:600;font-family:sans-serif;font-size:0.9rem;">'
        f"{label} ({fg_hex} on {bg_hex})</span>"
    )


def main():
    st.set_page_config(
        page_title="PHD ADA Document & PDF Compliance Tool",
        layout="wide",
    )
    st.title("ADA Document & Tagged PDF Compliance Tool")
    st.markdown(
        "Upload a Word document (`.docx`) to check accessibility, review and fix items "
        "directly on this page (including color contrast, decorative graphics, headings, "
        "and links), and download both an **ADA-compliant Word file (`.docx`)** and a "
        "**tagged PDF/UA-1 file (`.pdf`)**."
    )

    # ── Step 1: Upload ──────────────────────────────────────────────────────
    st.header("Step 1: Upload Word Document")
    col1, col2 = st.columns(2)
    with col1:
        docx_file = st.file_uploader(
            "Upload a Word (.docx) file",
            type=["docx"],
            help="Word document to inspect and convert into an accessible DOCX + tagged PDF",
        )
    with col2:
        metadata_file = st.file_uploader(
            "Upload metadata workbook (optional .xlsx)",
            type=["xlsx"],
            help="Optional Excel workbook with pre-approved alt text and link titles",
        )

    if not docx_file:
        st.info("Upload a `.docx` file above to begin.")
        return

    tmpdir = tempfile.mkdtemp(prefix="ada_interactive_")
    docx_path = os.path.join(tmpdir, docx_file.name)
    with open(docx_path, "wb") as f:
        f.write(docx_file.getvalue())

    meta_wb = None
    if metadata_file:
        meta_path = os.path.join(tmpdir, metadata_file.name)
        with open(meta_path, "wb") as f:
            f.write(metadata_file.getvalue())
        meta_wb = ExcelMetadataReader().read(meta_path)
        st.success(f"Loaded metadata workbook: {metadata_file.name}")

    # ── Step 2: Automated Scan & Color Contrast Audit ──────────────────────
    st.header("Step 2: Scan Summary & Color Contrast Check")
    scanner = DocxScanner()
    scan_result = scanner.scan(docx_path, metadata=meta_wb)
    inspection = inspect_document(docx_path)

    failing_colors = [c for c in inspection.contrast_items if not c.passed]
    weak_links = [lk for lk in inspection.hyperlinks if lk.is_weak_or_url]
    tables_missing_hdr = [t for t in inspection.tables if not t.has_header_flag]

    m1, m2, m3, m4, m5, m6 = st.columns(6)
    m1.metric("Images to Review", len(inspection.images))
    m2.metric("Heading / Style Issues", len(inspection.headings_and_styles))
    m3.metric("Links Needing Text", len(weak_links))
    m4.metric("Tables Missing Header", f"{len(tables_missing_hdr)}/{len(inspection.tables)}")
    m5.metric(
        "Color Contrast Pairs",
        f"{len(inspection.contrast_items) - len(failing_colors)}/{len(inspection.contrast_items)} PASS",
    )
    m6.metric(
        "Small Font / Footnotes",
        inspection.hygiene.small_font_runs_count + inspection.hygiene.unspaced_superscripts_count,
    )

    # Color Contrast Audit Section
    with st.expander(
        f"Color Contrast Audit ({len(inspection.contrast_items)} color pairs checked — "
        f"{len(failing_colors)} failing)",
        expanded=bool(failing_colors),
    ):
        st.markdown(
            "Every text and background color combination in your paragraphs, headings, "
            "links, and table cells is checked against the **WCAG 2.1 AA minimum** "
            "(`4.5:1` for regular text, `3.0:1` for large text $\\ge 18\\text{ pt}$ or $14\\text{ pt}$ bold)."
        )
        color_fg_overrides: dict[str, str] = {}
        for item in inspection.contrast_items:
            c_left, c_mid, c_right = st.columns([3, 2, 3])
            with c_left:
                st.markdown(
                    _render_color_swatch_html(item.fg_hex, item.bg_hex, item.role),
                    unsafe_allow_html=True,
                )
                st.caption(f"Sample ({item.location}): \"{item.sample_text}\"")
            with c_mid:
                badge = "PASS" if item.passed else "FAIL"
                st.markdown(
                    f"**Ratio:** `{item.ratio}:1` (min `{item.required_ratio}:1`) — **{badge}**"
                )
            with c_right:
                if not item.passed:
                    palette_options = {
                        f"{name} ({hex_val} -> {contrast_ratio(hex_val, item.bg_hex)}:1)": hex_val
                        for name, hex_val in PHD_BRAND_PALETTE.items()
                        if contrast_ratio(hex_val, item.bg_hex) >= item.required_ratio
                    }
                    chosen_label = st.selectbox(
                        f"Fix text color for {item.role}",
                        options=list(palette_options.keys()),
                        key=f"color_fix_{item.pair_id}",
                    )
                    color_fg_overrides[item.pair_id] = palette_options[chosen_label]
                else:
                    st.caption("Meets WCAG 2.1 AA contrast standard.")

        if inspection.color_only_warnings:
            st.warning(
                "Prose phrases that reference color alone (confirm labels/patterns are also present):\n"
                + "\n".join(f"- {w}" for w in inspection.color_only_warnings)
            )

    # ── Step 3: Interactive Review & Fixes (Zero AI) ────────────────────────
    st.header("Step 3: Review & Customize Fixes")

    # 3A. Document Metadata & Language
    st.subheader("3A. Document Title & Language")
    if inspection.title_needs_attention:
        st.warning(
            f"Current Word property title is **`{inspection.core_title or '(blank)'}`**, "
            f"which does not match the document's main heading (**`{inspection.first_h1_text}`**). "
            f"We pre-filled the main heading below."
        )
    col_t1, col_t2, col_t3 = st.columns([3, 2, 1])
    with col_t1:
        doc_title_input = st.text_input(
            "Document Title (announced first by screen readers and shown in the PDF title bar)",
            value=inspection.recommended_title,
        )
    with col_t2:
        doc_author_input = st.text_input(
            "Document Author / Organization",
            value=inspection.author,
        )
    with col_t3:
        doc_lang_input = st.selectbox(
            "Language",
            options=["en-US", "es-US", "vi-VN", "zh-CN", "tl-PH"],
            index=0,
        )

    # 3B. Images & Decorative Motifs
    st.subheader("3B. Images & Decorative Graphics")
    image_decisions: dict[int, dict[str, object]] = {}
    if not inspection.images:
        st.success("No body images requiring review.")
    else:
        for img in inspection.images:
            with st.container(border=True):
                ic1, ic2 = st.columns([1, 4])
                with ic1:
                    if img.image_bytes:
                        st.image(img.image_bytes, caption=f"{img.name} ({img.location})", width=120)
                    else:
                        st.markdown(f"**{img.name}** ({img.location})")
                with ic2:
                    default_dec = img.is_decorative or (img.is_floating and not img.current_alt)
                    is_dec = st.checkbox(
                        "Purely decorative (skip in screen readers in both Word and PDF)",
                        value=default_dec,
                        key=f"img_dec_{img.index}",
                    )
                    if is_dec:
                        st.caption(
                            "This graphic stays in its exact document position and is marked "
                            "decorative (`adec:decorative` in Word and `/Artifact` in PDF) so screen readers skip it."
                        )
                        image_decisions[img.index] = {"decorative": True, "alt_text": ""}
                    else:
                        alt_val = st.text_input(
                            f"Description (alt text) for {img.name}",
                            value=img.current_alt,
                            placeholder="Describe what this chart or image shows...",
                            key=f"img_alt_{img.index}",
                        )
                        image_decisions[img.index] = {"decorative": False, "alt_text": alt_val}

    # 3C. Headings & Paragraph Styles
    st.subheader("3C. Headings & Paragraph Styles")
    paragraph_style_overrides: dict[int, str] = {}
    style_choices = ["Heading 1", "Heading 2", "Heading 3", "Normal", "Note/Source", "Keep current"]
    if not inspection.headings_and_styles:
        st.success("Heading hierarchy and paragraph styles are clean.")
    else:
        for item in inspection.headings_and_styles:
            with st.container(border=True):
                hc1, hc2 = st.columns([3, 2])
                with hc1:
                    st.markdown(
                        f"**Paragraph {item.para_index}**: `\"{item.text_preview}\"`"
                    )
                    st.caption(f"Current style: `{item.current_style}` — {item.reason}")
                with hc2:
                    default_idx = (
                        style_choices.index(item.recommended_style)
                        if item.recommended_style in style_choices
                        else 0
                    )
                    chosen_style = st.selectbox(
                        f"Assign style for Paragraph {item.para_index}",
                        options=style_choices,
                        index=default_idx,
                        key=f"para_style_{item.para_index}",
                    )
                    if chosen_style != "Keep current":
                        paragraph_style_overrides[item.para_index] = chosen_style

    # 3D. Hyperlinks
    st.subheader("3D. Descriptive Hyperlink Text")
    hyperlink_text_overrides: dict[int, str] = {}
    if not weak_links:
        st.success("All hyperlinks already have descriptive display text.")
    else:
        st.markdown(
            "Screen readers read link text out loud. Replace raw URLs (`https://...`) "
            "or vague phrases (`click here`) with the destination title:"
        )
        for lk in weak_links:
            with st.container(border=True):
                lc1, lc2 = st.columns([2, 3])
                with lc1:
                    st.markdown(f"**Link {lk.link_index}** (Paragraph {lk.para_index})")
                    st.caption(f"Context: \"{lk.context_snippet}\"")
                    st.caption(f"Destination URL: `{lk.url}`")
                with lc2:
                    new_text = st.text_input(
                        f"Readable link title for Link {lk.link_index}",
                        value=lk.suggested_text,
                        key=f"link_txt_{lk.link_index}",
                    )
                    hyperlink_text_overrides[lk.link_index] = new_text

    # 3E. Tables & Formatting Hygiene
    st.subheader("3E. Tables, Font Sizes & Footnote Formatting")
    tc1, tc2 = st.columns(2)
    table_header_flags: dict[int, bool] = {}
    with tc1:
        mark_all_tables = st.checkbox(
            f"Mark Row 1 of all {len(inspection.tables)} table(s) as repeating header rows (<TH>) "
            "(keeps tables and shaded category rows as-is)",
            value=True,
        )
        for tbl in inspection.tables:
            table_header_flags[tbl.table_index] = mark_all_tables
        enforce_font_floor = st.checkbox(
            f"Raise small text ({inspection.hygiene.small_font_runs_count} runs under 11pt) to 11pt minimum",
            value=True,
        )
    with tc2:
        bracket_sups = st.checkbox(
            f"Add space and brackets to superscript footnotes ({inspection.hygiene.unspaced_superscripts_count} found, "
            "e.g., 'Stroke6' -> 'Stroke [6]')",
            value=True,
        )
        remove_spacers = st.checkbox(
            f"Remove empty spacer paragraphs ({len(inspection.hygiene.empty_spacer_paras)} found) "
            f"and hidden zero-width characters ({len(inspection.hygiene.zero_width_char_paras)} found)",
            value=True,
        )

    # ── Step 4: Apply Fixes & Build Both Word (.docx) and PDF (.pdf) ────────
    st.header("Step 4: Apply Fixes & Generate Accessible Word + PDF")
    if st.button("Apply Fixes & Build Accessible Word + Tagged PDF", type="primary"):
        base_stem = Path(docx_file.name).stem.replace(" ", "_")
        out_docx_name = f"{base_stem}_ADA.docx"
        out_pdf_name = f"{base_stem}_ADA.pdf"
        out_docx_path = os.path.join(tmpdir, out_docx_name)
        out_pdf_path = os.path.join(tmpdir, out_pdf_name)

        plan = InteractiveRemediationPlan(
            document_title=doc_title_input,
            document_author=doc_author_input,
            document_language=doc_lang_input,
            image_decisions=image_decisions,
            paragraph_style_overrides=paragraph_style_overrides,
            hyperlink_text_overrides=hyperlink_text_overrides,
            table_header_flags=table_header_flags,
            color_fg_overrides=color_fg_overrides,
            enforce_min_font_size=enforce_font_floor,
            min_font_size_pt=11.0,
            remove_empty_spacer_paragraphs=remove_spacers,
            strip_zero_width_chars=remove_spacers,
            bracket_and_space_superscripts=bracket_sups,
        )

        with st.spinner("Applying Word accessibility fixes and verifying screen-reader transcript..."):
            stats = apply_interactive_remediation(docx_path, out_docx_path, plan)
            post_scan = scanner.scan(out_docx_path, metadata=meta_wb)
            sr_transcript = generate_transcript(out_docx_path)
            verification = verify_transcript(out_docx_path)

        with st.spinner("Compiling tagged PDF/UA-1 and running 16-check PDF accessibility audit..."):
            pdf_qa_report = build_tagged_pdf_from_docx(
                out_docx_path,
                out_pdf_path,
                work_dir=os.path.join(tmpdir, "pdf_build"),
                language=plan.document_language,
            )

        gate = evaluate_deliverable_gate(
            post_scan=post_scan,
            sr_transcript=sr_transcript,
            verification=verification,
            pdf_qa_report=pdf_qa_report,
        )

        if gate["passed"]:
            st.success(
                f"{gate['status_label']} — Word issues: {post_scan.total_issues} | "
                f"Screen-reader verifier: {verification.summary()} | "
                f"PDF 16-check audit: ALL PASSED"
            )
        else:
            st.error(
                f"{gate['status_label']}\n\n"
                + "\n".join(f"- {reason}" for reason in gate["blocking_reasons"])
                + "\n\nResolve the remaining items above in Step 2 / Step 3 and re-run before publishing."
            )

        # Verification details
        v_col1, v_col2 = st.columns(2)
        with v_col1:
            with st.expander("Word Screen-Reader Transcript & Verification", expanded=not gate["passed"]):
                st.markdown(f"**{sr_transcript.summary()}**")
                st.markdown(f"**{verification.summary()}**")
                st.text(sr_transcript.text_only[:4000])
        with v_col2:
            with st.expander("Tagged PDF/UA-1 16-Check QA Report", expanded=not gate["passed"]):
                for chk in pdf_qa_report.get("checks", []):
                    icon = "PASS" if chk.get("passed") else "FAIL"
                    chk_name = chk.get("check") or chk.get("name") or "Check"
                    st.markdown(f"- **[{icon}] {chk_name}**: `{chk.get('detail')}`")

        # Download buttons (fail-closed: Accessible downloads require gate["passed"])
        st.subheader("Download Accessible Deliverables")
        d1, d2, d3 = st.columns(3)
        with d1:
            with open(out_docx_path, "rb") as f:
                st.download_button(
                    label=(
                        f"Download Accessible Word ({out_docx_name})"
                        if gate["passed"]
                        else "Blocked: Unresolved Accessibility Findings (.docx)"
                    ),
                    data=f.read(),
                    file_name=out_docx_name if gate["passed"] else f"{base_stem}_DRAFT_UNRESOLVED.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    disabled=not gate["passed"],
                )
        with d2:
            with open(out_pdf_path, "rb") as f:
                st.download_button(
                    label=(
                        f"Download Tagged PDF/UA-1 ({out_pdf_name})"
                        if gate["passed"]
                        else "Blocked: Unresolved Accessibility Findings (.pdf)"
                    ),
                    data=f.read(),
                    file_name=out_pdf_name if gate["passed"] else f"{base_stem}_DRAFT_UNRESOLVED.pdf",
                    mime="application/pdf",
                    type="primary",
                    disabled=not gate["passed"],
                )
        with d3:
            combined_audit = {
                "deliverable_gate": gate,
                "word_remediation_stats": stats,
                "word_remaining_issues": post_scan.total_issues,
                "screen_reader_verification": verification.summary(),
                "pdf_qa_audit": pdf_qa_report,
            }
            st.download_button(
                label="Download QA Audit Receipt (.json)",
                data=json.dumps(combined_audit, indent=2),
                file_name=f"{base_stem}_ADA_Audit.json",
                mime="application/json",
            )


if __name__ == "__main__":
    main()
