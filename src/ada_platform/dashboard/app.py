"""Streamlit dashboard for the ADA Document Accessibility Platform."""

from __future__ import annotations

import os
import tempfile

import streamlit as st

from ada_platform.scanner.docx_scanner import DocxScanner
from ada_platform.fixer.docx_fixer import DocxFixer
from ada_platform.metadata.excel_reader import ExcelMetadataReader
from ada_platform.reporting.exception_report import (
    generate_json_report,
    generate_text_report,
    generate_scan_text_report,
)
from ada_platform.qa.screen_reader_sim import generate_transcript
from ada_platform.qa.transcript_verifier import verify_transcript


def main():
    st.set_page_config(page_title="ADA Accessibility Platform", layout="wide")
    st.title("ADA Document Accessibility Platform")
    st.markdown("Upload a document to scan for accessibility issues and apply fixes.")

    # Step 1: Upload
    st.header("Step 1: Upload")
    col1, col2 = st.columns(2)
    with col1:
        docx_file = st.file_uploader(
            "Upload a DOCX file",
            type=["docx"],
            help="Word document to check for accessibility issues",
        )
    with col2:
        metadata_file = st.file_uploader(
            "Upload metadata workbook (optional)",
            type=["xlsx"],
            help="Excel workbook with reusable alt text, link text, and document metadata",
        )

    if not docx_file:
        st.info("Upload a DOCX file to get started.")
        return

    # Save uploaded files to temp directory
    tmpdir = tempfile.mkdtemp()
    docx_path = os.path.join(tmpdir, docx_file.name)
    with open(docx_path, "wb") as f:
        f.write(docx_file.getvalue())

    meta_wb = None
    if metadata_file:
        meta_path = os.path.join(tmpdir, metadata_file.name)
        with open(meta_path, "wb") as f:
            f.write(metadata_file.getvalue())
        meta_wb = ExcelMetadataReader().read(meta_path)
        st.success(f"Loaded metadata from {metadata_file.name}")

    # Step 2: Scan
    st.header("Step 2: Scan")
    scanner = DocxScanner()
    scan_result = scanner.scan(docx_path, metadata=meta_wb)

    st.metric("Total Issues", scan_result.total_issues)
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("Auto-fixable", len(scan_result.safe_fixable))
    with col2:
        st.metric("Suggested", len(scan_result.suggested))
    with col3:
        st.metric("Need Review", len(scan_result.needs_review))

    st.markdown(f"**{scan_result.summary()}**")

    # Show scan details
    with st.expander("Scan Details"):
        st.text(generate_scan_text_report(scan_result))

    if scan_result.total_issues == 0:
        st.success("No accessibility issues found. Document looks good.")
        return

    # Step 3: Review issues requiring human judgment
    if scan_result.needs_review:
        st.header("Step 3: Review")
        st.markdown("These issues need your input:")
        for issue in scan_result.needs_review:
            with st.container():
                st.markdown(f"**{issue.location}**: {issue.description}")
                if issue.current_value:
                    st.caption(f"Current: {issue.current_value}")

    # Step 4: Fix
    st.header("Step 4: Apply Fixes")
    if st.button("Apply automatic fixes", type="primary"):
        base, ext = os.path.splitext(docx_file.name)
        output_name = f"{base}_accessible{ext}"
        output_path = os.path.join(tmpdir, output_name)
        report_path = os.path.join(tmpdir, f"{base}_report.json")

        fixer = DocxFixer()
        fix_result = fixer.fix(docx_path, output_path, scan_result, metadata=meta_wb)

        st.success(fix_result.summary())

        # Show report
        with st.expander("Full Report"):
            st.text(generate_text_report(fix_result))

        # Generate JSON report
        generate_json_report(fix_result, report_path)

        # Step 5: Screen Reader Preview (QI)
        st.header("Step 5: Screen Reader Preview")
        sr_transcript = generate_transcript(output_path)

        if sr_transcript.warnings:
            st.warning(sr_transcript.summary())
            with st.expander("Screen reader warnings"):
                for w in sr_transcript.warnings:
                    st.markdown(f"- **{w.location}**: {w.text}")
        else:
            st.success(sr_transcript.summary())

        with st.expander("Full screen reader transcript"):
            st.text(sr_transcript.text_only)

        # Step 6: Transcript Verification (QI)
        st.header("Step 6: Transcript Verification")
        verification = verify_transcript(output_path)
        if verification.passed and verification.gap_count == 0:
            st.success(verification.summary())
        elif verification.passed:
            st.warning(verification.summary())
        else:
            st.error(verification.summary())

        with st.expander("Verification details"):
            if verification.issues:
                for issue in verification.issues:
                    icon = "X" if issue.severity == "error" else "?"
                    st.markdown(
                        f"- **[{icon}] {issue.element}** ({issue.location})\n"
                        f"  - Expected: {issue.expected}\n"
                        f"  - Found: {issue.found}"
                    )
            else:
                st.markdown("Every element in the document matches the transcript.")

        # Step 7: Download
        st.header("Step 7: Download")
        col1, col2, col3 = st.columns(3)
        with col1:
            with open(output_path, "rb") as f:
                st.download_button(
                    label="Download remediated DOCX",
                    data=f.read(),
                    file_name=output_name,
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
        with col2:
            with open(report_path, "rb") as f:
                st.download_button(
                    label="Download issue report (JSON)",
                    data=f.read(),
                    file_name=f"{base}_report.json",
                    mime="application/json",
                )
        with col3:
            st.download_button(
                label="Download screen reader transcript",
                data=sr_transcript.text_only,
                file_name=f"{base}_transcript.txt",
                mime="text/plain",
            )


if __name__ == "__main__":
    main()
