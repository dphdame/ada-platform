"""Template-based chart accessibility engine.

Generates short alt text, long narrative descriptions, and accessible
data tables from structured chart data. No LLM required.
"""

from __future__ import annotations

import csv
import statistics
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class ChartAccessibility:
    """Generated accessibility content for a chart."""

    short_alt: str
    long_description: str
    data_table_headers: list[str] = field(default_factory=list)
    data_table_rows: list[list[str]] = field(default_factory=list)


@dataclass
class ChartSpec:
    """Specification for generating chart accessibility content."""

    chart_type: str
    title: str
    x_axis: str
    y_axis: str
    key_takeaway: str = ""
    data: list[dict[str, str]] = field(default_factory=list)


def read_csv_data(csv_path: str) -> list[dict[str, str]]:
    """Read CSV file into list of row dicts."""
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        return list(reader)


def generate_chart_accessibility(spec: ChartSpec) -> ChartAccessibility:
    """Generate accessibility content from a chart specification."""
    short_alt = _generate_short_alt(spec)
    long_desc = _generate_long_description(spec)
    headers, rows = _generate_data_table(spec)

    return ChartAccessibility(
        short_alt=short_alt,
        long_description=long_desc,
        data_table_headers=headers,
        data_table_rows=rows,
    )


def _generate_short_alt(spec: ChartSpec) -> str:
    """Generate short alt text (<125 chars).

    Template: "{chart_type} chart showing {y_axis} by {x_axis}. {key_takeaway}"
    """
    chart_label = spec.chart_type.capitalize() if spec.chart_type else "Chart"
    parts = [f"{chart_label} chart"]

    if spec.y_axis and spec.x_axis:
        parts.append(f"showing {spec.y_axis} by {spec.x_axis}")
    elif spec.title:
        parts.append(f"titled {spec.title}")

    base = " ".join(parts)

    if spec.key_takeaway:
        candidate = f"{base}. {spec.key_takeaway}"
        if len(candidate) <= 125:
            return candidate

    # Truncate if needed
    if len(base) > 125:
        return base[:122] + "..."
    return base


def _generate_long_description(spec: ChartSpec) -> str:
    """Generate a narrative description with statistics from the data."""
    parts: list[str] = []

    # Opening sentence
    chart_label = spec.chart_type.capitalize() if spec.chart_type else "Chart"
    if spec.title:
        parts.append(f"This {chart_label.lower()} chart is titled \"{spec.title}.\"")
    else:
        parts.append(f"This is a {chart_label.lower()} chart.")

    if spec.x_axis and spec.y_axis:
        parts.append(
            f"The horizontal axis shows {spec.x_axis} and "
            f"the vertical axis shows {spec.y_axis}."
        )

    # Data-driven statistics
    if spec.data:
        numeric_values = _extract_numeric_values(spec)
        if numeric_values:
            vals = numeric_values
            n = len(vals)

            # Basic stats
            min_val = min(vals)
            max_val = max(vals)
            mean_val = statistics.mean(vals)

            # Find labels for min/max
            min_label = _find_label_for_value(spec, min_val)
            max_label = _find_label_for_value(spec, max_val)

            parts.append(
                f"The data includes {n} values "
                f"ranging from {_fmt(min_val)} to {_fmt(max_val)}."
            )

            if min_label and max_label:
                parts.append(
                    f"The lowest value is {_fmt(min_val)} ({min_label}) "
                    f"and the highest is {_fmt(max_val)} ({max_label})."
                )

            parts.append(f"The average is {_fmt(mean_val)}.")

            # Trend (first vs last)
            if n >= 2:
                first = vals[0]
                last = vals[-1]
                if last > first:
                    pct = ((last - first) / first * 100) if first != 0 else 0
                    parts.append(
                        f"The values show an increasing trend, "
                        f"rising {_fmt(pct)}% from start to end."
                    )
                elif last < first:
                    pct = ((first - last) / first * 100) if first != 0 else 0
                    parts.append(
                        f"The values show a decreasing trend, "
                        f"falling {_fmt(pct)}% from start to end."
                    )
                else:
                    parts.append("The values remain stable over the period.")

    if spec.key_takeaway:
        parts.append(f"Key takeaway: {spec.key_takeaway}")

    return " ".join(parts)


def _generate_data_table(
    spec: ChartSpec,
) -> tuple[list[str], list[list[str]]]:
    """Generate an accessible data table from the chart data."""
    if not spec.data:
        return [], []

    # Use dict keys as headers
    headers = list(spec.data[0].keys())
    rows = [[row.get(h, "") for h in headers] for row in spec.data]
    return headers, rows


def _extract_numeric_values(spec: ChartSpec) -> list[float]:
    """Extract numeric values from the y-axis column of the data."""
    if not spec.data:
        return []

    # Try y_axis column name first, then second column
    y_col = spec.y_axis
    first_row = spec.data[0]

    if y_col not in first_row:
        # Try to find a numeric column
        keys = list(first_row.keys())
        if len(keys) >= 2:
            y_col = keys[1]
        else:
            return []

    values: list[float] = []
    for row in spec.data:
        raw = row.get(y_col, "")
        try:
            values.append(float(str(raw).replace(",", "").replace("$", "")))
        except (ValueError, TypeError):
            continue
    return values


def _find_label_for_value(
    spec: ChartSpec, target: float
) -> Optional[str]:
    """Find the x-axis label corresponding to a numeric value."""
    if not spec.data:
        return None

    x_col = spec.x_axis
    first_row = spec.data[0]
    if x_col not in first_row:
        keys = list(first_row.keys())
        x_col = keys[0] if keys else ""

    y_col = spec.y_axis
    if y_col not in first_row:
        keys = list(first_row.keys())
        y_col = keys[1] if len(keys) >= 2 else ""

    for row in spec.data:
        raw = row.get(y_col, "")
        try:
            val = float(str(raw).replace(",", "").replace("$", ""))
            if abs(val - target) < 0.001:
                return row.get(x_col, "")
        except (ValueError, TypeError):
            continue
    return None


def _fmt(val: float) -> str:
    """Format a number: integers without decimals, floats to 1 decimal."""
    if val == int(val):
        return f"{int(val):,}"
    return f"{val:,.1f}"
