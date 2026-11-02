from __future__ import annotations

from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from ._utils import _read_headers


def _validate_arguments(
    source: Worksheet,
    value_columns: tuple[str, ...],
    variable_name: str,
    value_name: str,
    output_sheet_name: str,
) -> None:
    """Validate basic wide_to_long arguments."""
    if not output_sheet_name:
        raise ValueError("output_sheet_name must not be empty.")

    if output_sheet_name == source.title:
        raise ValueError("Output worksheet cannot be the source worksheet.")

    if not value_columns:
        raise ValueError("value_columns must contain at least one column.")

    if not variable_name or not value_name:
        raise ValueError("variable_name and value_name must not be empty.")

    if variable_name == value_name:
        raise ValueError("variable_name and value_name must be different.")


def _validate_columns(
    headers: tuple[str, ...],
    id_columns: tuple[str, ...],
    value_columns: tuple[str, ...],
    variable_name: str,
    value_name: str,
    sort_by: tuple[str, ...],
) -> tuple[dict[str, int], tuple[str, ...]]:
    """Validate column specifications and build column indexes."""
    column_index = {
        name: index
        for index, name in enumerate(headers)
    }

    if len(id_columns) != len(set(id_columns)):
        raise ValueError("id_columns contains duplicate column names.")

    if len(value_columns) != len(set(value_columns)):
        raise ValueError("value_columns contains duplicate column names.")

    overlap = set(id_columns) & set(value_columns)
    if overlap:
        raise ValueError(
            f"Columns cannot be both identifier and value columns: {sorted(overlap)}"
        )

    required_columns = (*id_columns, *value_columns)
    missing_columns = [
        name for name in required_columns
        if name not in column_index
    ]

    if missing_columns:
        raise ValueError(f"Missing columns: {missing_columns}")

    output_headers = (*id_columns, variable_name, value_name)

    if len(output_headers) != len(set(output_headers)):
        raise ValueError(
            "Output column names must be unique and must not conflict "
            "with identifier columns."
        )

    if len(sort_by) != len(set(sort_by)):
        raise ValueError("sort_by contains duplicate column names.")

    missing_sort_columns = [
        name for name in sort_by
        if name not in output_headers
    ]

    if missing_sort_columns:
        raise ValueError(
            f"Sort columns not found in output: {missing_sort_columns}"
        )

    return column_index, output_headers


def _build_long_rows(
    source: Worksheet,
    column_index: dict[str, int],
    id_columns: tuple[str, ...],
    value_columns: tuple[str, ...],
) -> list[list[Any]]:
    """Convert worksheet rows from wide format to long format."""
    output_rows = []

    for row in source.iter_rows(
        min_row=2,
        max_col=len(column_index),
        values_only=True,
    ):
        identifiers = tuple(
            row[column_index[name]]
            for name in id_columns
        )

        for column in value_columns:
            value = row[column_index[column]]

            if value is None:
                continue

            output_rows.append([
                *identifiers,
                column,
                value,
            ])

    return output_rows


def _build_sort_key(
    row: list[Any],
    sort_by: tuple[str, ...],
    output_index: dict[str, int],
    value_order: dict[str, int],
    variable_name: str,
) -> tuple:
    """Build a sort key for a long-format row."""
    keys = []

    for name in sort_by:
        index = output_index[name]

        if name == variable_name:
            keys.append(value_order[row[index]])
        else:
            keys.append((row[index] is None, row[index]))

    return tuple(keys)


def _sort_rows(
    rows: list[list[Any]],
    sort_by: tuple[str, ...],
    output_headers: tuple[str, ...],
    value_columns: tuple[str, ...],
    variable_name: str,
) -> None:
    """Sort long-format rows by the specified output columns."""
    output_index = {
        name: index
        for index, name in enumerate(output_headers)
    }

    value_order = {
        name: index
        for index, name in enumerate(value_columns)
    }

    rows.sort(
        key=lambda row: _build_sort_key(
            row=row,
            sort_by=sort_by,
            output_index=output_index,
            value_order=value_order,
            variable_name=variable_name,
        )
    )


def _create_output_worksheet(
    source: Worksheet,
    output_sheet_name: str,
    output_headers: tuple[str, ...],
    rows: list[list[Any]],
) -> Worksheet:
    """Create an output worksheet and write the converted rows."""
    workbook = source.parent

    if output_sheet_name in workbook.sheetnames:
        del workbook[output_sheet_name]

    output = workbook.create_sheet(output_sheet_name)
    output.append(output_headers)

    for row in rows:
        output.append(row)

    return output


def wide_to_long(
    source: Worksheet,
    id_columns: tuple[str, ...],
    value_columns: tuple[str, ...],
    variable_name: str = "Metric",
    value_name: str = "Value",
    output_sheet_name: str = "chart_data",
    sort_by: tuple[str, ...] = (),
) -> Worksheet:
    """Convert wide-format worksheet data into long-format data.

    Args:
        source: Source worksheet.
        id_columns: Columns whose values are preserved.
        value_columns: Columns to unpivot into variable/value pairs.
        variable_name: Header name for the original column names.
        value_name: Header name for the corresponding values.
        output_sheet_name: Name of the output worksheet.
        sort_by: Output columns used to sort the converted rows.

    Returns:
        The newly created long-format worksheet.

    Raises:
        ValueError: If headers or column specifications are invalid.
    """
    _validate_arguments(
        source,
        value_columns,
        variable_name,
        value_name,
        output_sheet_name,
    )

    headers = _read_headers(source)

    column_index, output_headers = _validate_columns(
        headers,
        id_columns,
        value_columns,
        variable_name,
        value_name,
        sort_by,
    )

    output_rows = _build_long_rows(
        source,
        column_index,
        id_columns,
        value_columns,
    )

    if sort_by:
        _sort_rows(
            output_rows,
            sort_by,
            output_headers,
            value_columns,
            variable_name,
        )

    return _create_output_worksheet(
        source,
        output_sheet_name,
        output_headers,
        output_rows,
    )
