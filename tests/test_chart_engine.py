"""Tests for the chart accessibility engine."""

from __future__ import annotations

import os

from ada_platform.charts.chart_engine import (
    ChartAccessibility,
    ChartSpec,
    generate_chart_accessibility,
    read_csv_data,
)


def _make_enrollment_spec() -> ChartSpec:
    return ChartSpec(
        chart_type="bar",
        title="Enrollment by Year",
        x_axis="Year",
        y_axis="Count",
        key_takeaway="Enrollment grew 15% from 2020 to 2025.",
        data=[
            {"Year": "2020", "Count": "500"},
            {"Year": "2021", "Count": "520"},
            {"Year": "2022", "Count": "540"},
            {"Year": "2023", "Count": "555"},
            {"Year": "2024", "Count": "565"},
            {"Year": "2025", "Count": "575"},
        ],
    )


class TestShortAlt:
    def test_under_125_chars(self):
        spec = _make_enrollment_spec()
        result = generate_chart_accessibility(spec)
        assert len(result.short_alt) <= 125

    def test_includes_chart_type(self):
        spec = _make_enrollment_spec()
        result = generate_chart_accessibility(spec)
        assert "Bar" in result.short_alt or "bar" in result.short_alt

    def test_includes_axes(self):
        spec = _make_enrollment_spec()
        result = generate_chart_accessibility(spec)
        assert "Count" in result.short_alt
        assert "Year" in result.short_alt

    def test_includes_takeaway_if_fits(self):
        spec = ChartSpec(
            chart_type="line",
            title="Test",
            x_axis="X",
            y_axis="Y",
            key_takeaway="Up 10%.",
        )
        result = generate_chart_accessibility(spec)
        assert "Up 10%" in result.short_alt

    def test_no_data_still_works(self):
        spec = ChartSpec(
            chart_type="pie",
            title="Revenue Share",
            x_axis="",
            y_axis="",
        )
        result = generate_chart_accessibility(spec)
        assert "Pie" in result.short_alt
        assert "Revenue Share" in result.short_alt


class TestLongDescription:
    def test_mentions_min_max(self):
        spec = _make_enrollment_spec()
        result = generate_chart_accessibility(spec)
        assert "500" in result.long_description
        assert "575" in result.long_description

    def test_mentions_trend(self):
        spec = _make_enrollment_spec()
        result = generate_chart_accessibility(spec)
        assert "increasing" in result.long_description.lower()

    def test_mentions_average(self):
        spec = _make_enrollment_spec()
        result = generate_chart_accessibility(spec)
        assert "average" in result.long_description.lower()

    def test_mentions_title(self):
        spec = _make_enrollment_spec()
        result = generate_chart_accessibility(spec)
        assert "Enrollment by Year" in result.long_description

    def test_decreasing_trend(self):
        spec = ChartSpec(
            chart_type="line",
            title="Decline",
            x_axis="Year",
            y_axis="Value",
            data=[
                {"Year": "2020", "Value": "100"},
                {"Year": "2021", "Value": "80"},
                {"Year": "2022", "Value": "60"},
            ],
        )
        result = generate_chart_accessibility(spec)
        assert "decreasing" in result.long_description.lower()

    def test_includes_key_takeaway(self):
        spec = _make_enrollment_spec()
        result = generate_chart_accessibility(spec)
        assert "Enrollment grew 15%" in result.long_description

    def test_no_data_minimal_description(self):
        spec = ChartSpec(chart_type="bar", title="Empty", x_axis="X", y_axis="Y")
        result = generate_chart_accessibility(spec)
        assert "bar chart" in result.long_description.lower()


class TestDataTable:
    def test_headers_match_data_keys(self):
        spec = _make_enrollment_spec()
        result = generate_chart_accessibility(spec)
        assert result.data_table_headers == ["Year", "Count"]

    def test_row_count_matches_data(self):
        spec = _make_enrollment_spec()
        result = generate_chart_accessibility(spec)
        assert len(result.data_table_rows) == 6

    def test_row_values(self):
        spec = _make_enrollment_spec()
        result = generate_chart_accessibility(spec)
        assert result.data_table_rows[0] == ["2020", "500"]
        assert result.data_table_rows[-1] == ["2025", "575"]

    def test_no_data_empty_table(self):
        spec = ChartSpec(chart_type="bar", title="Empty", x_axis="X", y_axis="Y")
        result = generate_chart_accessibility(spec)
        assert result.data_table_headers == []
        assert result.data_table_rows == []


class TestReadCSV:
    def test_reads_csv_file(self, tmp_path):
        csv_path = str(tmp_path / "data.csv")
        with open(csv_path, "w") as f:
            f.write("Year,Count\n2020,500\n2021,520\n")
        data = read_csv_data(csv_path)
        assert len(data) == 2
        assert data[0]["Year"] == "2020"
        assert data[0]["Count"] == "500"


class TestEdgeCases:
    def test_non_numeric_y_values(self):
        spec = ChartSpec(
            chart_type="bar",
            title="Categories",
            x_axis="Name",
            y_axis="Category",
            data=[
                {"Name": "Alice", "Category": "A"},
                {"Name": "Bob", "Category": "B"},
            ],
        )
        result = generate_chart_accessibility(spec)
        # Should not crash; long description won't have stats
        assert result.short_alt
        assert result.long_description

    def test_dollar_values_parsed(self):
        spec = ChartSpec(
            chart_type="bar",
            title="Revenue",
            x_axis="Quarter",
            y_axis="Revenue",
            data=[
                {"Quarter": "Q1", "Revenue": "$1,000"},
                {"Quarter": "Q2", "Revenue": "$2,000"},
            ],
        )
        result = generate_chart_accessibility(spec)
        assert "1,000" in result.long_description
        assert "2,000" in result.long_description
