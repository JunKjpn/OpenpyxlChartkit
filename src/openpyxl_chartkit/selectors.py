from __future__ import annotations

from dataclasses import dataclass
from itertools import groupby
from typing import Any

from .specs import ChartSpec, SeriesSpec
from .table import DataRow, Table


@dataclass(frozen=True)
class SeriesRange:
    """A contiguous Excel row range rendered as one chart series."""

    title: str
    start_row: int
    end_row: int


@dataclass(frozen=True)
class ChartRange:
    title: str
    table: Table


def resolve_chart_ranges(table: Table, spec: ChartSpec) -> tuple[ChartRange, ...]:
    if spec.horizontal_grouping is None and spec.vertical_grouping is None:
        return (
            ChartRange(
                title=spec.title_template or "",
                table=table,
            ),
        )

    grouping = spec.horizontal_grouping
    if grouping is None:
        raise NotImplementedError("Vertical grouping is not supported yet.")

    table.require_columns(*grouping.columns)
    ranges: list[ChartRange] = []
    seen_keys: set[tuple[Any, ...]] = set()
    for key, rows in groupby(table.rows, key=lambda row: _group_key(row, grouping.columns)):
        grouped_rows = tuple(rows)
        if key in seen_keys:
            raise ValueError(f"Group {key!r} in sheet '{table.sheet_name}' is not contiguous.")
        seen_keys.add(key)
        _require_contiguous_rows(grouped_rows, table.sheet_name, key)
        title = f"{spec.title_template or "Chart"} ({_group_title(grouping.columns, key)})"
        ranges.append(
            ChartRange(
                title=title,
                table=Table(sheet_name=table.sheet_name, columns=table.columns, rows=grouped_rows),
            )
        )
    return tuple(ranges)

def resolve_series_ranges(table: Table, spec: SeriesSpec) -> tuple[SeriesRange, ...]:
    """Resolve one SeriesSpec into one or more contiguous worksheet ranges."""
    if not table.rows:
        raise ValueError(f"Sheet '{table.sheet_name}' contains no data rows.")

    grouping = spec.grouping
    if grouping is None:
        return (
            SeriesRange(
                title=spec.title or spec.y.title or spec.y.column,
                start_row=table.rows[0].excel_row,
                end_row=table.rows[-1].excel_row,
            ),
        )

    if grouping.source != "row":
        raise NotImplementedError("Only row-based series grouping is supported.")

    table.require_columns(*grouping.columns)
    ranges: list[SeriesRange] = []
    seen_keys: set[tuple[Any, ...]] = set()

    for key, rows in groupby(table.rows, key=lambda row: _group_key(row, grouping.columns)):
        grouped_rows = tuple(rows)
        if key in seen_keys:
            raise ValueError(
                f"Group {key!r} in sheet '{table.sheet_name}' is not contiguous. "
                "Use a helper sheet to normalize interleaved data."
            )
        seen_keys.add(key)
        _require_contiguous_rows(grouped_rows, table.sheet_name, key)

        ranges.append(
            SeriesRange(
                title=_group_title(columns=grouping.columns, key=key),
                start_row=grouped_rows[0].excel_row,
                end_row=grouped_rows[-1].excel_row,
            )
        )

    return tuple(ranges)


def _group_key(row: DataRow, columns: tuple[str, ...]) -> tuple[Any, ...]:
    return tuple(row.values[column] for column in columns)


def _require_contiguous_rows(rows: tuple[DataRow, ...], sheet_name: str, key: tuple[Any, ...]) -> None:
    row_numbers = [row.excel_row for row in rows]
    if row_numbers != list(range(row_numbers[0], row_numbers[-1] + 1)):
        raise ValueError(f"Group {key!r} in sheet '{sheet_name}' contains blank or non-contiguous rows.")


def _group_title(columns: tuple[str, ...], key: tuple[Any, ...]) -> str:
    return ", ".join(f"{column}={value}" for column, value in zip(columns, key))
