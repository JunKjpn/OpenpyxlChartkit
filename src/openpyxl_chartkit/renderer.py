from __future__ import annotations

from openpyxl import Workbook
from openpyxl.chart import BarChart, Reference, ScatterChart, Series

from .specs import AxisSpec, ChartSpec, SeriesSpec
from .selectors import ChartRange, SeriesRange, resolve_chart_ranges, resolve_series_ranges
from .table import Table
from .layout import ChartLayoutEngine
from .session import ChartLayoutSession


class ChartRenderer:
    """Render the minimal ChartSpec subset into an existing workbook."""

    def render(self, workbook: Workbook, spec: ChartSpec, session: ChartLayoutSession | None = None) -> None:
        source = self._get_sheet(workbook=workbook, name=spec.source_sheet)
        output = self._get_or_create_sheet(workbook=workbook, name=spec.output_sheet)
        table = Table.from_worksheet(worksheet=source)

        required_columns = [spec.x.column, *(series.y.column for series in spec.series)]
        table.require_columns(*required_columns)

        chart_ranges = resolve_chart_ranges(table, spec)
        layout_engine = ChartLayoutEngine(spec.layout)

        if session is not None:
            grid_rows = max(chart_range.row for chart_range in chart_ranges) + 1
            origin_row, origin_column = session.allocate(
                rows=grid_rows,
                row_step=layout_engine.row_step,
            )
        else:
            origin_row, origin_column = 1, 1

        for chart_range in chart_ranges:
            chart = self._build_chart(source, chart_range, spec)

            output.add_chart(
                chart,
                layout_engine.anchor(
                    grid_row=chart_range.row,
                    grid_column=chart_range.column,
                    origin_row=origin_row,
                    origin_column=origin_column,
                ),
            )

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
    def _make_chart(spec: ChartSpec) -> ScatterChart | BarChart:
        if spec.chart_type == "scatter":
            return ScatterChart()
        if spec.chart_type == "bar":
            return BarChart()
        raise ValueError(f"Unsupported chart type: {spec.chart_type}")

    @staticmethod
    def _reference(worksheet, table: Table, axis: AxisSpec, series_range: SeriesRange) -> Reference:
        column_number = table.columns.index(axis.column) + 1
        return Reference(
            worksheet,
            min_col=column_number,
            min_row=series_range.start_row,
            max_row=series_range.end_row,
        )

    @staticmethod
    def _add_series(
            chart, chart_type: str, x_values: Reference, y_values: Reference, spec: SeriesSpec, title: str,
    ) -> None:
        if chart_type == "bar":
            chart.add_data(y_values, titles_from_data=False)
            chart.set_categories(x_values)
            return

        series = Series(y_values, x_values, title=title)
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

    def _build_chart(self, source, chart_range: ChartRange, spec: ChartSpec) -> ScatterChart | BarChart:
        chart = self._make_chart(spec)
        table = chart_range.table

        for series_spec in spec.series:
            for series_range in resolve_series_ranges(table, series_spec):
                x_values = self._reference(source, table, spec.x, series_range)
                y_values = self._reference(source, table, series_spec.y, series_range)
                self._add_series(chart, spec.chart_type, x_values, y_values, series_spec, series_range.title)

        if any(series.grouping is not None for series in spec.series):
            chart.legend.position = "b"

        chart.title = chart_range.title or None
        first_y_axis = spec.series[0].y
        chart.width = spec.layout.width
        chart.height = spec.layout.height
        chart.x_axis.title = spec.x.title or spec.x.column
        chart.y_axis.title = first_y_axis.title or first_y_axis.column

        self._apply_axis_settings(chart.x_axis, spec.x)
        self._apply_axis_settings(chart.y_axis, first_y_axis)
        return chart