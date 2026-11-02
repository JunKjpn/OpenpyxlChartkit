from __future__ import annotations

from typing import Any

from openpyxl.worksheet.worksheet import Worksheet


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
    if not output_sheet_name:
        raise ValueError("output_sheet_name must not be empty.")

    if output_sheet_name == source.title:
        raise ValueError("Output worksheet cannot be the source worksheet.")

    # Read and validate headers.
    header_values = [cell.value for cell in source[1]]

    # Ignore trailing empty cells.
    while header_values and header_values[-1] is None:
        header_values.pop()

    if not header_values:
        raise ValueError(f"Worksheet '{source.title}' has no header row.")

    if any(
        not isinstance(header, str) or not header.strip()
        for header in header_values
    ):
        raise ValueError(
            f"Worksheet '{source.title}' has an empty or non-text header."
        )

    if len(header_values) != len(set(header_values)):
        raise ValueError(
            f"Worksheet '{source.title}' has duplicate headers."
        )

    headers = tuple(header_values)

    if column not in headers:
        raise ValueError(
            f"Column '{column}' not found in worksheet '{source.title}'."
        )

    column_index = headers.index(column)

    # Create output worksheet.
    workbook = source.parent

    if output_sheet_name in workbook.sheetnames:
        del workbook[output_sheet_name]

    output = workbook.create_sheet(output_sheet_name)
    output.append(headers)

    # Copy rows matching the specified value.
    for row in source.iter_rows(
        min_row=2,
        max_col=len(headers),
        values_only=True,
    ):
        if row[column_index] == value:
            output.append(row)

    return output