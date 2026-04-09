"""Tests for the CLI using Click's CliRunner."""

from __future__ import annotations

import json
import os
import struct
import zlib

from click.testing import CliRunner
from docx import Document
from openpyxl import Workbook

from ada_platform.cli.main import cli


def _write_tiny_png(path: str) -> None:
    def _chunk(ct, data):
        c = ct + data
        return struct.pack(">I", len(data)) + c + struct.pack(">I", zlib.crc32(c) & 0xFFFFFFFF)
    with open(path, "wb") as f:
        f.write(b"\x89PNG\r\n\x1a\n")
        f.write(_chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)))
        f.write(_chunk(b"IDAT", zlib.compress(b"\x00\xff\xff\xff")))
        f.write(_chunk(b"IEND", b""))


def _make_test_docx(tmp_path, name="test.docx") -> str:
    doc = Document()
    doc.core_properties.title = ""
    doc.add_paragraph("Content.")
    table = doc.add_table(rows=2, cols=2)
    table.cell(0, 0).text = "A"
    table.cell(0, 1).text = "B"
    path = str(tmp_path / name)
    doc.save(path)
    return path


def _make_test_metadata(tmp_path) -> str:
    wb = Workbook()
    ws = wb.active
    ws.title = "Documents"
    ws.append(["filename", "title"])
    ws.append(["test.docx", "Test Document Title"])
    ws2 = wb.create_sheet("Images")
    ws2.append(["image_id", "alt_text"])
    path = str(tmp_path / "meta.xlsx")
    wb.save(path)
    return path


def _make_test_csv(tmp_path) -> str:
    path = str(tmp_path / "data.csv")
    with open(path, "w") as f:
        f.write("Year,Count\n2020,100\n2021,120\n2022,140\n")
    return path


class TestScanCommand:
    def test_scan_basic(self, tmp_path):
        docx = _make_test_docx(tmp_path)
        runner = CliRunner()
        result = runner.invoke(cli, ["scan", docx])
        assert result.exit_code == 0
        assert "ACCESSIBILITY SCAN RESULTS" in result.output

    def test_scan_with_json_output(self, tmp_path):
        docx = _make_test_docx(tmp_path)
        report = str(tmp_path / "report.json")
        runner = CliRunner()
        result = runner.invoke(cli, ["scan", docx, "-o", report])
        assert result.exit_code == 0
        assert os.path.exists(report)
        with open(report) as f:
            data = json.load(f)
        assert data["file_type"] == "docx"

    def test_scan_with_metadata(self, tmp_path):
        docx = _make_test_docx(tmp_path)
        meta = _make_test_metadata(tmp_path)
        runner = CliRunner()
        result = runner.invoke(cli, ["scan", docx, "-m", meta])
        assert result.exit_code == 0
        assert "metadata" in result.output.lower()

    def test_scan_nonexistent_file(self):
        runner = CliRunner()
        result = runner.invoke(cli, ["scan", "/nonexistent.docx"])
        assert result.exit_code != 0


class TestFixCommand:
    def test_fix_creates_output(self, tmp_path):
        docx = _make_test_docx(tmp_path)
        out = str(tmp_path / "fixed.docx")
        runner = CliRunner()
        result = runner.invoke(cli, ["fix", docx, "-o", out])
        assert result.exit_code == 0
        assert os.path.exists(out)
        assert "ACCESSIBILITY REPORT" in result.output

    def test_fix_with_json_report(self, tmp_path):
        docx = _make_test_docx(tmp_path)
        out = str(tmp_path / "fixed.docx")
        report = str(tmp_path / "report.json")
        runner = CliRunner()
        result = runner.invoke(cli, ["fix", docx, "-o", out, "-r", report])
        assert result.exit_code == 0
        assert os.path.exists(report)


class TestWizardCommand:
    def test_wizard_auto_fix(self, tmp_path):
        docx = _make_test_docx(tmp_path)
        runner = CliRunner()
        result = runner.invoke(cli, ["wizard", docx], input="y\n")
        assert result.exit_code == 0
        assert "Scanning" in result.output
        # Should create *_accessible.docx
        accessible = str(tmp_path / "test_accessible.docx")
        assert os.path.exists(accessible)


class TestChartCommand:
    def test_chart_output(self, tmp_path):
        csv_path = _make_test_csv(tmp_path)
        runner = CliRunner()
        result = runner.invoke(cli, [
            "chart", csv_path,
            "--type", "bar",
            "--title", "Growth",
            "--x", "Year",
            "--y", "Count",
            "--takeaway", "Up 40%.",
        ])
        assert result.exit_code == 0
        assert "SHORT ALT TEXT" in result.output
        assert "LONG DESCRIPTION" in result.output
        assert "DATA TABLE" in result.output


class TestInitWorkbook:
    def test_creates_template(self, tmp_path):
        out = str(tmp_path / "template.xlsx")
        runner = CliRunner()
        result = runner.invoke(cli, ["init-workbook", "-o", out])
        assert result.exit_code == 0
        assert os.path.exists(out)
        assert "Documents" in result.output


class TestVersion:
    def test_version(self):
        runner = CliRunner()
        result = runner.invoke(cli, ["--version"])
        assert result.exit_code == 0
        assert "0.1.0" in result.output
