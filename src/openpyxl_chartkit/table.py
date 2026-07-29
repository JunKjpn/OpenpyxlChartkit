from __future__ import annotations

from dataclasses import dataclass
from typing import Any

from openpyxl.worksheet.worksheet import Worksheet


@dataclass(frozen=True)
class DataRow:
    """A worksheet row together with its original Excel row number."""

    excel_row: int
    values: dict[str, Any]


@dataclass(frozen=True)
class Table:
    """Worksheet data that preserves the row numbers used by Excel charts."""

    sheet_name: str
    columns: tuple[str, ...]
    rows: tuple[DataRow, ...]

    @classmethod
    def from_worksheet(
        cls,
        worksheet: Worksheet,
        *,
        header_row: int = 1,
        skip_blank_rows: bool = True,
    ) -> Table:
        headers = tuple(cell.value for cell in worksheet[header_row])
        cls._validate_headers(headers, worksheet.title, header_row)

        rows: list[DataRow] = []
        for excel_row, values in enumerate(
            worksheet.iter_rows(
                min_row=header_row + 1,
                max_col=len(headers),
                values_only=True,
            ),
            start=header_row + 1,
        ):
            if skip_blank_rows and all(value is None for value in values):
                continue
            rows.append(DataRow(excel_row=excel_row, values=dict(zip(headers, values))))

        return cls(sheet_name=worksheet.title, columns=headers, rows=tuple(rows))

    def require_columns(self, *names: str) -> None:
        missing = [name for name in names if name not in self.columns]
        if missing:
            raise ValueError(
                f"Sheet '{self.sheet_name}' is missing columns: {missing}"
            )

    @staticmethod
    def _validate_headers(
        headers: tuple[Any, ...],
        sheet_name: str,
        header_row: int,
    ) -> None:
        if not headers:
            raise ValueError(f"Sheet '{sheet_name}' has no header row at {header_row}.")
        if any(not isinstance(header, str) or not header.strip() for header in headers):
            raise ValueError(
                f"Sheet '{sheet_name}' has an empty or non-text header."
            )

        duplicates = {header for header in headers if headers.count(header) > 1}
        if duplicates:
            raise ValueError(
                f"Sheet '{sheet_name}' has duplicate headers: {sorted(duplicates)}"
            )
