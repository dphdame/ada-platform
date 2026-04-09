# Document Accessibility Platform Scope

## Project Definition

Multi-format document accessibility platform for Word, PowerPoint, and PDF that helps non-technical staff produce more accessible deliverables through a guided dashboard, shared metadata from Excel, and format-specific remediation workflows.

Automates predictable accessibility tasks, reduces repetitive manual entry, and surfaces only issues requiring human judgment (chart interpretation, complex table structure, reading-order exceptions).

## Primary Goal

Let a staff user upload a draft document, optionally connect an Excel workbook with reusable metadata, review only unresolved accessibility issues in plain language, and download a remediated output plus an issue log.

Target environment: public-sector or administrative settings where authors are not accessibility specialists, documents are often template-based, and teams need a practical path to better compliance.

## Document Strategy

| Document Type | Primary Authoring Path | Final Output |
|---|---|---|
| Structured reports | LaTeX where feasible | Tagged PDF (PDF/UA-oriented) |
| Board communications, memos, letters, short briefs | DOCX | Accessible DOCX + tagged PDF export |
| Slide decks and presentations | PPTX | Accessible PPTX, optional PDF export |
| Existing PDFs needing review | PDF intake | Audit, guided remediation, structured replacement |

## Format Scope

### Word (DOCX)
- Document title, language metadata
- Heading styles
- Image alt text
- Simple table headers
- Weak hyperlink text
- Chart accessibility (descriptions from structured data, short alt text, longer description, accessible data table)

### PowerPoint (PPTX)
- Slide title checks
- Object reading order review
- Image and chart alt text
- Layout-based slide structure
- Common accessibility issues (PowerPoint checker parity)

### PDF
- Tag analysis, reading order
- Document title, language
- Figure alt text, heading structure
- LaTeX-generated tagged PDFs (preferred for reports)
- Legacy PDF: audit, issue surfacing, guided exception handling, replacement workflows

## Dashboard Design

Steps:
1. Upload a file
2. Optionally upload/connect Excel metadata workbook
3. Run automated scan
4. Auto-apply safe fixes
5. Review unresolved issues in plain language
6. Generate remediated outputs and issue log

Plain language throughout: "3 images need descriptions" not "3 Figure tags missing /Alt entry."

## Excel Metadata Layer

Reusable metadata registry to avoid re-entering the same accessibility content.

| Sheet | Purpose |
|---|---|
| Documents | Title, language, business unit, template type, audience |
| Images | Reusable alt text, decorative flags, image identifiers |
| Charts | Chart titles, types, source data ranges, summary fields, description rules |
| Links | Approved link text and canonical URLs |
| Tables | Header-row assumptions, simple/complex flags |
| Slides | PPTX-specific defaults for slide titles and object descriptions |

## Chart Accessibility Engine

Shared module across Word, PowerPoint, and PDF:
- Chart image generation
- Short alt text (type, subject, key takeaway)
- Longer narrative description when needed
- Optional accessible data table
- User can keep/edit generated text or replace chart image while retaining accessibility text

## Safe Automation (High Confidence)

- Document title and language metadata
- Known alt text from Excel mappings
- Heading styles / template mappings
- Simple table headers in Word
- Known slide titles from templates
- Chart text from structured data
- Flagging weak hyperlink text

## Human Review Required

- Decorative vs. informative image classification
- Chart text reflecting intended message
- PDF reading order in complex layouts
- PowerPoint object order matching narrative
- Complex table simplification decisions

## Key Product Modules

| Module | Purpose |
|---|---|
| Intake | Upload DOCX, PPTX, PDF, and optional Excel metadata workbook |
| File classifier | Determine document type, template type, likely workflow |
| Metadata engine | Read reusable values from Excel and map to document elements |
| Rule engine | Run accessibility checks by format |
| Safe-fix engine | Apply automatable fixes |
| Exception review UI | Present unresolved issues in plain language |
| Chart engine | Generate chart visuals, alt text, long descriptions, data tables |
| Export engine | Save remediated files and issue logs |
| Audit log | Record auto-fixes, unresolved issues, user overrides |

## Format-Specific Workflows

### Report Workflow
LaTeX templates for PDF generation; companion DOCX for review/commenting if needed.

### Memo/Board Communication Workflow
DOCX-first; output accessible DOCX + tagged PDF for distribution.

### Presentation Workflow
PPTX as editable source; PDF export after accessibility work.

### Legacy PDF Workflow
Analysis, issue reporting, guided repair, reconstruction recommendations.

## Phased Implementation

| Phase | Scope |
|---|---|
| Phase 1 | DOCX workflow, Excel metadata layer, chart-description generation |
| Phase 2 | PPTX workflow with reading-order review and alt-text support |
| Phase 3 | Report-focused LaTeX PDF pipeline and PDF audit workflow |
| Phase 4 | Legacy PDF support, analytics, template governance, admin reporting |

## Success Criteria

- Non-technical staff can use with minimal training
- Repetitive accessibility inputs reduced through Excel-driven metadata
- Generated files pass large share of automated accessibility checks before final human review
- Manual remediation narrowed to genuinely judgment-based issues

## Scope Boundaries

Does NOT promise one-click ADA/Section 508 compliance for every file and layout. Promises a practical workflow that automates predictable parts, standardizes metadata/chart descriptions, and routes ambiguous issues to human review.
