from openpyxl.worksheet.worksheet import Worksheet


def _read_headers(source: Worksheet) -> tuple[str, ...]:
    """Read and validate worksheet headers."""
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

    return tuple(header_values)