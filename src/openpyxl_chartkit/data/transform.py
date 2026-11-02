from __future__ import annotations

from openpyxl.worksheet.worksheet import Worksheet


def _build_sort_key(
    row: list,
    sort_by: tuple[str, ...],
    output_headers: tuple[str, ...],
    value_order: dict[str, int],
    variable_name: str,
) -> tuple:
    keys = []

    for name in sort_by:
        index = output_headers.index(name)

        if name == variable_name:
            keys.append(value_order[row[index]])
        else:
            keys.append((row[index] is None, row[index]))

    return tuple(keys)


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

    Returns:
        The newly created long-format worksheet.

    Raises:
        ValueError: If headers or column specifications are invalid.
    """
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
    column_index = {
        name: index
        for index, name in enumerate(headers)
    }

    # Validate column specifications.
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

    # Create output worksheet.
    workbook = source.parent

    if output_sheet_name in workbook.sheetnames:
        del workbook[output_sheet_name]

    output = workbook.create_sheet(output_sheet_name)
    output.append(output_headers)

    # Convert wide-format rows into long-format rows.
    output_rows = []

    for row in source.iter_rows(
            min_row=2,
            max_col=len(headers),
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

    if sort_by:
        value_order = {
            name: index
            for index, name in enumerate(value_columns)
        }

        output_rows.sort(
            key=lambda row: _build_sort_key(
                row=row,
                sort_by=sort_by,
                output_headers=output_headers,
                value_order=value_order,
                variable_name=variable_name,
            )
        )

    for row in output_rows:
        output.append(row)

    return output