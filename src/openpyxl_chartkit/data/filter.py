from __future__ import annotations

from typing import Any

from openpyxl.worksheet.worksheet import Worksheet

from ._utils import _read_headers


def _validate_arguments(
    source: Worksheet,
    output_sheet_name: str,
) -> None:
    """Validate filter_rows arguments."""
    if not output_sheet_name:
        raise ValueError("output_sheet_name must not be empty.")

    if output_sheet_name == source.title:
        raise ValueError("Output worksheet cannot be the source worksheet.")


def _create_output_worksheet(
    source: Worksheet,
    output_sheet_name: str,
    headers: tuple[str, ...],
) -> Worksheet:
    """Create an output worksheet with headers."""
    workbook = source.parent

    if output_sheet_name in workbook.sheetnames:
        del workbook[output_sheet_name]

    output = workbook.create_sheet(output_sheet_name)
    output.append(headers)

    return output


def _filter_rows(
    source: Worksheet,
    output: Worksheet,
    column_index: int,
    value: Any,
    max_column: int,
) -> None:
    """Copy rows matching the specified column value."""
    for row in source.iter_rows(
        min_row=2,
        max_col=max_column,
        values_only=True,
    ):
        if row[column_index] == value:
            output.append(row)


def filter_rows(
    source: Worksheet,
    column: str,
    value: Any,
    output_sheet_name: str = "filtered_data",
) -> Worksheet:
    """Filter worksheet rows by an exact column value.

    Args:
        source: Source worksheet.
        column: Header name of the column to filter.
        value: Value to match.
        output_sheet_name: Name of the output worksheet.

    Returns:
        A new worksheet containing the header and matching rows.

    Raises:
        ValueError: If the worksheet headers or arguments are invalid.
    """
    _validate_arguments(source, output_sheet_name)

    headers = _read_headers(source)

    if column not in headers:
        raise ValueError(
            f"Column '{column}' not found in worksheet '{source.title}'."
        )

    column_index = headers.index(column)

    output = _create_output_worksheet(
        source,
        output_sheet_name,
        headers,
    )

    _filter_rows(
        source,
        output,
        column_index,
        value,
        len(headers),
    )

    return output
