from __future__ import annotations

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference, ScatterChart, Series
from .specs import AxisSpec, ChartSpec, SeriesSpec
from .table import Table


class ChartRenderer:
    """Render the minimal ChartSpec subset into an existing workbook."""

    def render(self, workbook: Workbook, spec: ChartSpec) -> None:
        source = self._get_sheet(workbook, spec.source_sheet)
        output = self._get_or_create_sheet(workbook, spec.output_sheet)
        table = Table.from_worksheet(source)

        required_columns = [spec.x.column, *(series.y.column for series in spec.series)]
        table.require_columns(*required_columns)
        self._require_contiguous_rows(table)

        chart = self._make_chart(spec)
        x_values = self._reference(source, table, spec.x)

        for series_spec in spec.series:
            y_values = self._reference(source, table, series_spec.y)
            self._add_series(chart, spec.chart_type, x_values, y_values, series_spec)

        chart.title = spec.title_template or None
        first_y_axis = spec.series[0].y
        # openpyxl serializes scatter-chart axes in the opposite order from the
        # public attribute names.  Assign titles/settings by their visible axes.
        chart.y_axis.title = spec.x.title or spec.x.column
        chart.x_axis.title = first_y_axis.title or first_y_axis.column
        chart.width = spec.layout.width
        chart.height = spec.layout.height
        self._apply_axis_settings(chart.y_axis, spec.x)
        self._apply_axis_settings(chart.x_axis, first_y_axis)

        output.add_chart(chart, self._anchor(output, spec))

    @staticmethod
    def _get_sheet(workbook: Workbook, name: str):
        try:
            return workbook[name]
        except KeyError as error:
            raise ValueError(f"Source sheet '{name}' does not exist.") from error

    @staticmethod
    def _get_or_create_sheet(workbook: Workbook, name: str):
        return workbook[name] if name in workbook.sheetnames else workbook.create_sheet(name)

    @staticmethod
    def _require_contiguous_rows(table: Table) -> None:
        if not table.rows:
            raise ValueError(f"Sheet '{table.sheet_name}' contains no data rows.")
        row_numbers = [row.excel_row for row in table.rows]
        expected = list(range(row_numbers[0], row_numbers[-1] + 1))
        if row_numbers != expected:
            raise ValueError(
                "The minimal renderer requires contiguous data rows. "
                "Filtering and non-contiguous data will be supported by selectors."
            )

    @staticmethod
    def _make_chart(spec: ChartSpec):
        if spec.chart_type == "scatter":
            return ScatterChart()
        if spec.chart_type == "bar":
            return BarChart()
        raise ValueError(f"Unsupported chart type: {spec.chart_type}")

    @staticmethod
    def _reference(worksheet, table: Table, axis: AxisSpec) -> Reference:
        column_number = table.columns.index(axis.column) + 1
        return Reference(
            worksheet,
            min_col=column_number,
            min_row=table.rows[0].excel_row,
            max_row=table.rows[-1].excel_row,
        )

    @staticmethod
    def _add_series(chart, chart_type: str, x_values: Reference, y_values: Reference, spec: SeriesSpec) -> None:
        if chart_type == "bar":
            chart.add_data(y_values, titles_from_data=False)
            chart.set_categories(x_values)
            return

        series = Series(y_values, x_values, title=spec.title or spec.y.title or spec.y.column)
        if spec.style == "line":
            series.graphicalProperties.line.width = 1.5 * 12700
        elif spec.style == "circle":
            series.marker.symbol = "circle"
            series.marker.size = 4
            series.graphicalProperties.line.noFill = True
        chart.series.append(series)

    @staticmethod
    def _apply_axis_settings(axis, spec: AxisSpec) -> None:
        if spec.log_scale:
            axis.scaling.logBase = 10
        if spec.minimum is not None:
            axis.scaling.min = spec.minimum
        if spec.maximum is not None:
            axis.scaling.max = spec.maximum

    @staticmethod
    def _anchor(output, spec: ChartSpec) -> str:
        if spec.layout.placement == "right_up":
            return f"{output.cell(row=2, column=output.max_column + 2).coordinate}"
        return f"{output.cell(row=output.max_row + 2, column=1).coordinate}"
