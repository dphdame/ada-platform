# Phase 1 Implementation Plan: ADA Document Accessibility Platform

**Scope:** DOCX workflow, Excel metadata layer, chart-description generation
**Created:** 2026-04-08

---

## 1. Directory Structure

```
/Users/victoriaperez/Projects/ADA_compliance/
├── SCOPE.md
├── PHASE1_PLAN.md
├── pyproject.toml
├── src/
│   └── ada_platform/
│       ├── __init__.py
│       ├── models/
│       │   ├── __init__.py
│       │   ├── issues.py             # IssueType enum, Issue dataclass, FixCategory enum
│       │   ├── scan_result.py        # ScanResult, FixResult, AuditLog models
│       │   └── metadata.py           # Excel workbook schema models
│       ├── scanner/
│       │   ├── __init__.py
│       │   ├── base.py               # Abstract BaseScanner (format-agnostic)
│       │   └── docx_scanner.py       # DOCX-specific scanner
│       ├── fixer/
│       │   ├── __init__.py
│       │   ├── base.py               # Abstract BaseFixer (format-agnostic)
│       │   └── docx_fixer.py         # DOCX safe-fix engine
│       ├── metadata/
│       │   ├── __init__.py
│       │   └── excel_reader.py       # openpyxl-based metadata workbook reader
│       ├── charts/
│       │   ├── __init__.py
│       │   └── chart_engine.py       # Chart description generator
│       ├── reporting/
│       │   ├── __init__.py
│       │   └── exception_report.py   # JSON + human-readable issue log
│       ├── cli/
│       │   ├── __init__.py
│       │   └── main.py               # Click CLI wizard
│       └── dashboard/
│           ├── __init__.py
│           └── app.py                # Streamlit dashboard
├── tests/
│   ├── conftest.py                   # Shared fixtures, programmatic DOCX builders
│   ├── test_scanner.py
│   ├── test_fixer.py
│   ├── test_metadata.py
│   ├── test_chart_engine.py
│   ├── test_reporting.py
│   └── test_cli.py
└── templates/
    └── metadata_workbook_template.xlsx
```

**Why this layout:** The `src/ada_platform/` namespace prevents import collisions. Each subdirectory (`scanner/`, `fixer/`, `metadata/`, `charts/`) is a self-contained module with an abstract base class in `base.py`. Phase 2 adds `scanner/pptx_scanner.py` and `fixer/pptx_fixer.py` without touching existing files. Phase 3 adds `scanner/pdf_scanner.py`. Models are shared across all formats.

---

## 2. Data Models

### 2a. Issue Types and Fix Categories (`models/issues.py`)

```python
from enum import Enum
from dataclasses import dataclass, field
from typing import Optional

class IssueType(Enum):
    MISSING_TITLE = "missing_title"
    MISSING_LANGUAGE = "missing_language"
    HEADING_SKIP = "heading_skip"                    # e.g., H1 -> H3
    HEADING_BOLD_NOT_STYLE = "heading_bold_not_style" # Bold text instead of heading style
    MISSING_ALT_TEXT = "missing_alt_text"
    FILEPATH_ALT_TEXT = "filepath_alt_text"           # Alt text is a file path
    TABLE_NO_HEADER = "table_no_header"
    WEAK_LINK_TEXT = "weak_link_text"                 # "click here", "link", bare URL
    MISSING_FOOTNOTE_LINK = "missing_footnote_link"
    COLOR_ONLY_INFO = "color_only_info"              # Flagged, not auto-fixable

class FixCategory(Enum):
    SAFE_AUTO = "safe_auto"        # High confidence, apply automatically
    SUGGEST = "suggest"            # Offer a suggested fix, user confirms
    HUMAN_REVIEW = "human_review"  # Cannot automate, needs human judgment

class Severity(Enum):
    CRITICAL = "critical"
    WARNING = "warning"
    INFO = "info"

@dataclass
class Issue:
    issue_type: IssueType
    severity: Severity
    fix_category: FixCategory
    location: str                        # "Paragraph 12", "Image 3", "Table 2"
    description: str                     # Plain language for non-technical users
    current_value: Optional[str] = None
    suggested_fix: Optional[str] = None
    metadata_key: Optional[str] = None   # Key into Excel workbook if applicable
    fixed: bool = False
```

### Fix Category Mapping

| IssueType | Default FixCategory | Condition for SAFE_AUTO |
|---|---|---|
| MISSING_TITLE | SAFE_AUTO | Title available from Excel or filename |
| MISSING_LANGUAGE | SAFE_AUTO | Always (set to "en-US") |
| HEADING_SKIP | SUGGEST | Suggest restructuring, user confirms |
| HEADING_BOLD_NOT_STYLE | SUGGEST | Suggest converting to heading style |
| MISSING_ALT_TEXT | SAFE_AUTO if Excel has it, else HUMAN_REVIEW | Excel Images sheet has matching entry |
| TABLE_NO_HEADER | SAFE_AUTO | Simple table (first row is clearly a header) |
| WEAK_LINK_TEXT | SUGGEST if Excel has it, else HUMAN_REVIEW | Excel Links sheet has approved text |

### 2b. Scan and Fix Results (`models/scan_result.py`)

```python
@dataclass
class ScanResult:
    file_path: str
    file_type: str                     # "docx", "pptx", "pdf"
    timestamp: str
    issues: list[Issue] = field(default_factory=list)
    total_images: int = 0
    total_tables: int = 0
    total_links: int = 0
    total_headings: int = 0

    @property
    def safe_fixable(self) -> list[Issue]:
        return [i for i in self.issues if i.fix_category == FixCategory.SAFE_AUTO]

    @property
    def needs_review(self) -> list[Issue]:
        return [i for i in self.issues if i.fix_category == FixCategory.HUMAN_REVIEW]

    @property
    def critical_count(self) -> int:
        return sum(1 for i in self.issues if i.severity == Severity.CRITICAL and not i.fixed)

@dataclass
class FixResult:
    file_path: str
    timestamp: str
    auto_fixed: list[Issue]
    needs_review: list[Issue]
    skipped: list[Issue]

@dataclass
class AuditLogEntry:
    timestamp: str
    action: str                        # "scan", "auto_fix", "user_override", "export"
    issue_type: Optional[str] = None
    detail: str = ""
    user_decision: Optional[str] = None
```

### 2c. Excel Metadata Schema (`models/metadata.py`)

**Documents sheet:**

| Column | Type | Purpose |
|---|---|---|
| `filename` | str | Document filename (match key) |
| `title` | str | Accessible document title |
| `language` | str | BCP 47 language tag (default: "en-US") |
| `business_unit` | str | Department/unit |
| `template_type` | str | "memo", "report", "brief", "letter" |
| `audience` | str | "internal", "public", "board" |

**Images sheet:**

| Column | Type | Purpose |
|---|---|---|
| `image_id` | str | Filename or hash-based identifier |
| `content_hash` | str | SHA-256 hash of image bytes (auto-computed, for matching embedded images that lose filenames) |
| `alt_text` | str | Reusable alt text |
| `long_description` | str | Extended description (optional) |
| `decorative` | bool | True = mark as decorative artifact |
| `used_in` | str | Comma-separated document filenames |

**Charts sheet:**

| Column | Type | Purpose |
|---|---|---|
| `chart_id` | str | Identifier |
| `chart_type` | str | "bar", "line", "pie", "scatter", etc. |
| `title` | str | Chart title |
| `data_source` | str | Path to CSV/Excel range |
| `x_axis` | str | What x-axis represents |
| `y_axis` | str | What y-axis represents |
| `key_takeaway` | str | One-sentence finding |
| `short_alt` | str | Generated or manually entered alt text |
| `long_description` | str | Narrative description |

**Links sheet:**

| Column | Type | Purpose |
|---|---|---|
| `url` | str | Canonical URL |
| `approved_text` | str | Compliant link text |
| `content_type` | str | "PDF", "video", "external site", etc. |

**Tables sheet:**

| Column | Type | Purpose |
|---|---|---|
| `table_id` | str | Identifier or caption text |
| `has_header_row` | bool | First row is headers |
| `complex` | bool | Needs manual review (merged cells, multi-level) |
| `notes` | str | Notes for human reviewer |

**Slides sheet** (included now for Phase 2 readiness):

| Column | Type | Purpose |
|---|---|---|
| `slide_number` | int | Slide position |
| `title` | str | Slide title (for missing-title fixes) |
| `object_descriptions` | str | Descriptions for non-text objects |
| `reading_order_notes` | str | Notes on intended reading order |
| `document` | str | Which PPTX file this applies to |

---

## 3. Module Architecture

### 3a. Scanner (`scanner/base.py` + `scanner/docx_scanner.py`)

**BaseScanner** interface: `scan(file_path, metadata=None) -> ScanResult`

**DocxScanner** implements six check methods, each returning `list[Issue]`:

1. `_check_title(doc, metadata)` -- checks `doc.core_properties.title`
2. `_check_language(doc)` -- inspects `doc.element` for `w:lang` via lxml
3. `_check_headings(doc)` -- walks paragraphs, checks heading hierarchy and bold-as-heading
4. `_check_alt_text(doc, metadata)` -- iterates inline shapes/drawings via lxml; cross-references Excel Images sheet
5. `_check_tables(doc, metadata)` -- inspects `tbl` elements for `<w:tblHeader/>`; cross-references Excel Tables sheet
6. `_check_links(doc, metadata)` -- extracts hyperlinks, checks for weak text; cross-references Excel Links sheet

Uses lxml for low-level XML inspection when python-docx API doesn't expose the attribute (same approach as existing `fix_docx_refs.py`).

### 3b. Fixer (`fixer/base.py` + `fixer/docx_fixer.py`)

**BaseFixer** interface: `fix(file_path, scan_result, metadata=None) -> FixResult`

**DocxFixer** applies SAFE_AUTO fixes only:

1. `_fix_title(doc, issue)` -- sets `doc.core_properties.title`
2. `_fix_language(doc)` -- inserts `w:lang val="en-US"` via lxml
3. `_fix_alt_text(doc, issue)` -- sets alt text on image via lxml
4. `_fix_table_headers(doc, issue)` -- sets `<w:tblHeader/>` on first row

Post-fix: calls `clean_metadata()` from existing `~/.claude/scripts/clean_docx_metadata.py`.

### 3c. Metadata Engine (`metadata/excel_reader.py`)

- Reads `.xlsx` via openpyxl, validates sheet names, skips missing sheets gracefully
- Returns `MetadataWorkbook` with `.documents`, `.images`, `.charts`, `.links`, `.tables`
- Lookup methods: `get_document_meta(filename)`, `get_image_alt(image_id)`, `get_chart_info(chart_id)`, `get_link_text(url)`, `get_table_info(table_id)`
- Image matching: three-tier (filename, SHA-256 content hash, positional fallback)

### 3d. Chart Engine (`charts/chart_engine.py`)

Format-agnostic. Takes chart metadata + source data (CSV/Excel), returns:

1. **Short alt text** (<125 chars): `"{chart_type} chart showing {y_axis} by {x_axis}. {key_takeaway}"`
2. **Long narrative description**: Structured paragraph with min, max, mean, trend direction
3. **Accessible data table**: List of row dicts insertable as DOCX/PPTX table

Returns `ChartAccessibility` dataclass. Template-based (no LLM); structured for future LLM upgrade.

### 3e. Exception Report (`reporting/exception_report.py`)

Two outputs from a `FixResult`:

1. **JSON report**: Full issue list with types, locations, severity, fix status
2. **Human-readable report**: Grouped by status ("Auto-Fixed (N items)", "Needs Your Review (N items)")

Summary line: `"Fixed 4 issues automatically. 2 items need your review."`

### 3f. CLI + Dashboard

**CLI (Click):**
```
ada-platform scan <file.docx> [--metadata workbook.xlsx] [--output report.json]
ada-platform fix <file.docx> [--metadata workbook.xlsx] [--output fixed.docx]
ada-platform wizard <file.docx>   # Interactive: scan -> review -> fix -> export
ada-platform chart <data.csv> --type bar --title "..." --x "..." --y "..."
ada-platform init-workbook         # Creates empty metadata template
```

**Dashboard (Streamlit):** Zero frontend code, file upload widget, step-by-step wizard, same modules as CLI.

---

## 4. Pipeline Flow

```
User uploads DOCX + optional Excel workbook
          |
          v
    ExcelMetadataReader.read(workbook) -> MetadataWorkbook
          |
          v
    DocxScanner.scan(docx, metadata) -> ScanResult
          |
          v
    DocxFixer.fix(docx, scan_result, metadata) -> FixResult
          |                                          |
          v                                          v
    ChartEngine.generate(chart_meta, data)    ExceptionReport.generate(fix_result)
          |                                          |
          v                                          v
    Insert chart descriptions                 JSON + plain-text report
    into fixed DOCX
          |
          v
    clean_metadata(doc)  # from existing script
          |
          v
    doc.save(output_path)
```

---

## 5. Build Order

Each step produces a testable artifact. No step proceeds until the previous passes tests.

| Step | Module | Depends On | Testable Artifact |
|---|---|---|---|
| 1 | Data models | None | Unit tests: instantiate all models, verify enums, verify ScanResult properties |
| 2 | DOCX Scanner | Step 1 | pytest: create defective DOCXs, verify correct IssueTypes detected |
| 3 | Excel Metadata Reader | Step 1 | pytest: create workbook, read back, verify all lookups |
| 4 | Scanner + Metadata integration | Steps 2, 3 | pytest: scanner with metadata promotes fix categories |
| 5 | DOCX Safe-Fix Engine | Steps 2, 3 | pytest: scan -> fix -> re-scan shows zero SAFE_AUTO issues |
| 6 | Chart Engine | Step 1 | pytest: CSV input produces alt text <125 chars, long description with stats |
| 7 | Exception Report | Steps 2, 5 | pytest: FixResult produces valid JSON + plain-language text |
| 8 | CLI | Steps 2-7 | pytest: Click CliRunner, verify exit codes and output |
| 9 | Streamlit Dashboard | Steps 2-7 | Manual: upload -> scan -> fix -> download walkthrough |
| 10 | Integration with existing scripts | Steps 5, 8 | End-to-end: DOCX through full pipeline |

---

## 6. Test Strategy

### Fixture Generation

Build all test DOCX files programmatically in `conftest.py` using python-docx. No committed binary files; tests are self-documenting.

| Fixture | How It's Built | What It Tests |
|---|---|---|
| Missing title | `Document()` with no `core_properties.title` | Scanner: MISSING_TITLE; Fixer: sets from metadata |
| Missing language | `Document()` with no `w:lang` element | Scanner: MISSING_LANGUAGE; Fixer: injects `w:lang` |
| Heading skip | Paragraphs with H1 then H3 (no H2) | Scanner: HEADING_SKIP at correct location |
| Bold-as-heading | Bold paragraph with Normal style | Scanner: HEADING_BOLD_NOT_STYLE |
| Missing alt text | `doc.add_picture()` with no alt text | Scanner: MISSING_ALT_TEXT; Fixer: applies from metadata |
| File path alt text | Image with alt = "figures/chart1.png" | Scanner: FILEPATH_ALT_TEXT |
| Table no headers | `doc.add_table()` with no `tblHeader` XML | Scanner: TABLE_NO_HEADER; Fixer: marks first row |
| Weak link text | Hyperlink with text "click here" | Scanner: WEAK_LINK_TEXT |
| Clean document | All issues resolved | Scanner: empty issues list |
| Metadata workbook | 6-sheet workbook via openpyxl | Reader: parses all sheets, lookups return correct values |
| Chart CSV | 2-column CSV (Year, Value) | Chart engine: alt text <125 chars, description has trend |

### Round-Trip Tests

Scan -> fix -> re-scan to confirm issues resolved. These verify scanner-fixer agreement.

### Regression Tests

Import `clean_metadata` from existing `clean_docx_metadata.py` and verify it still sets author to "Victoria Cholette" and clears comments.

---

## 7. Tech Stack

| Component | Library | Version |
|---|---|---|
| DOCX manipulation | python-docx | 1.1+ |
| DOCX XML access | lxml | 5.0+ |
| Excel reading | openpyxl | 3.1+ |
| CLI | Click | 8.1+ |
| Dashboard | Streamlit | 1.30+ |
| Chart data processing | Standard library (csv, statistics) | Python 3.11 stdlib |
| Testing | pytest | 8.0+ |

---

## 8. Risks and Mitigations

| Risk | Mitigation |
|---|---|
| python-docx doesn't expose language metadata or table header XML | Use lxml on `doc.element`, same approach as existing `fix_docx_refs.py` |
| Image matching between DOCX embedded images and Excel registry | Three-tier: filename first, SHA-256 content hash, positional fallback |
| Bold-as-heading detection false positives | Only flag: bold is entire paragraph, <100 chars, followed by body text |
| Chart descriptions too generic from templates | Structure engine for future LLM-based generator upgrade |
| Streamlit single-user limitation | Acceptable for Phase 1 internal tool; note multi-user would need Flask |

---

## 9. QI Gates

| Gate | When | Tool/Script | Pass Criteria |
|---|---|---|---|
| Model completeness | After Step 1 | pytest | All enums, dataclasses instantiate; properties compute correctly |
| Scanner accuracy | After Step 2 | pytest round-trip | Each fixture triggers exactly the expected IssueType |
| Metadata fidelity | After Step 3 | pytest | All 6 sheets parse; lookups match input data |
| Fix verification | After Step 5 | pytest scan-fix-rescan | Zero SAFE_AUTO issues after fix pass |
| Chart output | After Step 6 | pytest | Alt text <125 chars; long description includes min/max/trend |
| Plain language | After Step 7 | Manual review | No XML/OOXML terms in human-readable report |
| CLI exit codes | After Step 8 | pytest CliRunner | Correct exit codes for clean file, defective file, missing file |
| End-to-end | After Step 10 | Full pipeline test | DOCX with 5+ issues -> remediated DOCX + report |

---

## 10. Existing Scripts Integration

| Existing Script | How Integrated | When Called |
|---|---|---|
| `clean_docx_metadata.py` | Import `clean_metadata()` directly | Post-fix, before save |
| `fix_docx_refs.py` | Optional step when `.aux` file detected | Post-fix, LaTeX-origin DOCXs only |
| `ada_postprocess.py` | Not used in Phase 1 | Phase 3 (PDF pipeline) |
| `ada_qa_audit.py` | Not used in Phase 1 | Phase 3 (PDF pipeline) |
| `check_pptx_overflow.py` | Reference for Phase 2 scanner interface | Phase 2 (PPTX) |
| `SKILL.md` (ada-compliance) | Reference for WCAG checks and PHD standards | Scanner rule definitions |

---

## Design Decisions (Resolved 2026-04-08)

1. **Dashboard: Streamlit.** Single-user local install per machine. No shared server.
2. **Chart engine: Template-only.** No LLM in Phase 1. Structure for future upgrade.
3. **Image matching: Filename + content-hash + positional.** Three-tier: filename match first, then SHA-256 content hash of image bytes, then positional fallback.
4. **Excel workbook includes Slides sheet now.** Pre-built for Phase 2 PPTX readiness even though DOCX workflow doesn't use it.
5. **Deployment: Local only.** `pip install` + `ada-platform wizard` CLI or `streamlit run` locally.
