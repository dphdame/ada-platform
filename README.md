# ADA Document & Tagged PDF Compliance Platform (`ada-platform`)

Deterministic, zero-AI WCAG 2.1 AA and Tagged PDF remediation platform for Word (`.docx`), PDF (`.pdf`), and Drupal / CMS HTML (`.html`) deliverables.

- **Case Study Article:** [What Should Be AI, What Should Be Code, and What Needs a Human?](https://tooearlytosay.com/research/methodology/what-should-be-ai-code-or-human-ada-case-study/)
- **Optional AI-Augmented Skill (`SKILL.md`):** [Download `ada-compliance-skill.md`](https://tooearlytosay.com/tools/templates/raw/ada-compliance-skill.md) (for drafting `<= 141`-char alt text and rewriting vague links in Claude Code, Cursor, or Antigravity before running this deterministic fixer).

## Quick Start (macOS, Linux, or Windows CLI + Web UI)

```bash
git clone https://github.com/dphdame/ada-platform.git
cd ada-platform
python3 -m venv .venv && .venv/bin/pip install -e .

# 1. Scan & auto-fix a Word document (.docx)
ada-platform scan draft_brief.docx
ada-platform fix draft_brief.docx -o brief_accessible.docx

# 2. Scan & auto-fix a PDF (.pdf) — 16-check Adobe Acrobat Pro Full Check equivalent
ada-platform scan report.pdf
ada-platform fix report.pdf -o report_accessible.pdf

# 3. Scan & auto-fix a Drupal web page or HTML snippet (.html)
ada-platform scan service_page.html
ada-platform fix service_page.html -o service_page_accessible.html

# 4. Launch the interactive local Streamlit web application
ada-platform ui
```

## Key Capabilities (100% Deterministic Python — Zero Generative AI)

1. **Interactive Web Application (`ada-platform ui`)**:
   - **Document Metadata & Language**: Inspects and sets document title, author, and BCP-47 language tag (`en-US`, `es-US`, `vi-VN`, `zh-TW`, `tl-PH`).
   - **Heading Hierarchy**: Detects skipped heading levels and visually styled fake headings, with interactive dropdown overrides.
   - **WCAG 2.1 AA Color Contrast & Color-Only Prose**: Evaluates every text run across body paragraphs, headings, hyperlinks, and table cells against WCAG 2.1 AA contrast ratios (`4.5:1` for normal text; `3.0:1` for large text >= 18pt or >= 14pt bold) with color pickers for remediation.
   - **Hyperlink Descriptive Text**: Flags raw URLs and vague link text (`click here`, `read more`) for inline replacement.
   - **Table Accessibility**: Sets repeating header rows (`<w:tblHeader/>`) on multi-column tables.
   - **Images & Decorative Graphics**: Marks decorative graphics in-place with `<adec:decorative val="1"/>` and deduplicated `wp:docPr` IDs, or applies descriptive alt text.
2. **Native Windows Word Tagged PDF Export + Surgical PDF Post-Processing**:
   - On Windows (`sys.platform == "win32"`), exports the remediated `.docx` directly via Microsoft Word's native `ExportAsFixedFormat` (`DocStructureTags=True`, `CreateBookmarks=1`, `IncludeDocProps=True`) using `win32com` or built-in Windows PowerShell COM automation—preserving exact Word layout without requiring LaTeX.
   - Automatically post-processes the exported PDF (`postprocess_tagged_pdf`) to promote first-row table cells from `<TD>` to `<TH /Scope /Column>` and convert empty-alt decorative shapes (`/Alt ()`) to `/Artifact`.
   - Runs a fail-closed 16-check internal structural audit (`run_pdf_ua_audit`) and deliverable gate (`evaluate_deliverable_gate`) before certifying downloads.

---

## Quick Start on Windows

### Prerequisites
- **Windows 10 or Windows 11**
- **Python 3.11 or newer** (installed with "Add Python to PATH" checked)
- **Microsoft Word for Windows** (installed and activated, used automatically for native `.docx -> tagged .pdf` export)

### Option A: One-Click Windows Install & Launch
Double-click **`install_and_run_windows.bat`** in the root folder (or run it from Command Prompt / PowerShell). It will automatically:
1. Create an isolated `.venv` virtual environment,
2. Install `ada-platform` and all required dependencies,
3. Run `ada-platform doctor` and `ada-platform ui --check`, and
4. Open the interactive Streamlit web application in your browser at `http://localhost:8501`.

### Option B: Manual PowerShell Commands
Open PowerShell in the unzipped `ADA_Compliance_Platform` folder and run:

```powershell
# 1. Create and activate a virtual environment
python -m venv .venv
.\.venv\Scripts\Activate.ps1

# 2. Install the package (from the pre-built wheel or editable source)
python -m pip install --upgrade pip
python -m pip install -e .

# 3. Verify installation and environment readiness
ada-platform doctor
ada-platform ui --check

# 4. Launch the interactive web app
ada-platform ui
```

### Running the Full Windows Verification & Live Word COM Test
To run the automated test suite plus a live Microsoft Word COM `.docx -> tagged .pdf` test on Windows:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\windows_sandbox_test.ps1
```
