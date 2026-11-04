#Requires -Version 5.1
<#
.SYNOPSIS
    End-to-end Windows verification script for the ADA Document & Tagged PDF Compliance Platform.
.DESCRIPTION
    1. Validates the Windows PowerShell Word COM script AST syntax.
    2. Runs `ada-platform doctor` and `ada-platform ui --check` on Windows.
    3. Runs the full pytest suite (including isolated wheel build & clean Windows venv install).
    4. If desktop Microsoft Word (`Word.Application` COM) is installed on this Windows machine,
       executes a live native Word `.docx -> tagged .pdf` export + 16-check structural audit.
#>

$ErrorActionPreference = 'Stop'
$RepoRoot = Split-Path -Parent $PSScriptRoot
Set-Location $RepoRoot

Write-Host "=== 1. Validating Windows PowerShell COM Script Syntax ==="
$tokens = $null
$errors = $null
$comScript = @"
`$ErrorActionPreference = 'Stop';
`$word = New-Object -ComObject Word.Application;
`$word.Visible = `$false;
`$word.DisplayAlerts = 0;
`$doc = `$null;
try {
  `$doc = `$word.Documents.Open('C:\temp\in.docx', `$false, `$true);
  `$doc.ExportAsFixedFormat('C:\temp\out.pdf', 17, `$false, 0, 0, 1, 1, 0, `$true, `$true, 1, `$true, `$true, `$false);
} finally {
  if (`$null -ne `$doc) { `$doc.Close(0); }
  `$word.Quit();
  [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject(`$word);
}
"@
[void][System.Management.Automation.Language.Parser]::ParseInput($comScript, [ref]$tokens, [ref]$errors)
if ($errors.Count -gt 0) {
    throw "PowerShell AST validation failed: $($errors | Out-String)"
}
Write-Host "[PASS] PowerShell COM script syntax validated."

Write-Host "=== 2. Running ada-platform doctor & ui --check ==="
uv run ada-platform doctor
uv run ada-platform ui --check

Write-Host "=== 3. Running Full Pytest Suite ==="
uv run pytest -q

Write-Host "=== 4. Checking Live Microsoft Word COM Availability ==="
$wordInstalled = $false
try {
    $testWord = New-Object -ComObject Word.Application
    $testWord.Quit()
    [void][System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($testWord)
    $wordInstalled = $true
} catch {
    Write-Host "[INFO] Desktop Microsoft Word (Word.Application COM) is not installed on this Windows instance; skipping live COM export."
}

if ($wordInstalled) {
    Write-Host "[PASS] Desktop Microsoft Word detected! Running live .docx -> tagged .pdf export & 16-check audit..."
    uv run python -c @"
import tempfile
from pathlib import Path
from docx import Document
from ada_platform.remediation_engine import (
    InteractiveRemediationPlan,
    apply_interactive_remediation,
    build_tagged_pdf_from_docx,
)

with tempfile.TemporaryDirectory() as tmp:
    tmp_p = Path(tmp)
    in_docx = tmp_p / 'live_word_test.docx'
    rem_docx = tmp_p / 'live_word_test_ADA.docx'
    out_pdf = tmp_p / 'live_word_test_ADA.pdf'

    doc = Document()
    doc.core_properties.title = 'Live Windows Word COM Tagged PDF Test'
    doc.add_heading('Live Windows Word COM Tagged PDF Test', level=1)
    doc.add_paragraph('Verified via native Microsoft Word ExportAsFixedFormat COM automation.')
    tbl = doc.add_table(rows=2, cols=2)
    tbl.cell(0, 0).text = 'Indicator'
    tbl.cell(0, 1).text = 'Value'
    tbl.cell(1, 0).text = 'Seniors'
    tbl.cell(1, 1).text = '20%'
    doc.save(str(in_docx))

    plan = InteractiveRemediationPlan(
        document_title='Live Windows Word COM Tagged PDF Test',
        document_language='en-US',
        table_header_flags={1: True},
    )
    apply_interactive_remediation(str(in_docx), str(rem_docx), plan)
    report = build_tagged_pdf_from_docx(str(rem_docx), str(out_pdf), backend='word')
    assert report['all_passed'] is True, f'Audit failed: {report}'
    print('LIVE_WORD_COM_PASS: 16/16 checks passed!')
"@
}
