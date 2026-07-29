import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from openpyxl_chartkit import (
    AxisSpec,
    ChartRenderer,
    ChartSpec,
    GroupingSpec,
    SeriesSpec,
)


class ChartRendererTests(unittest.TestCase):
    def test_render_adds_scatter_chart(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "data"
        sheet.append(["x", "y"])
        sheet.append([1, 2])
        sheet.append([2, 4])

        spec = ChartSpec(
            source_sheet="data",
            output_sheet="chart",
            chart_type="scatter",
            x=AxisSpec("x", title="X"),
            series=(SeriesSpec(AxisSpec("y", title="Y"), title="Y values"),),
            title_template="Sample chart",
        )

        ChartRenderer().render(workbook, spec)
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory) / "chart.xlsx"
            workbook.save(output)

            reopened = load_workbook(output)
            chart_sheet = reopened["chart"]
            self.assertEqual(len(chart_sheet._charts), 1)
            self.assertEqual(len(chart_sheet._charts[0].series), 1)
            self.assertEqual(
                chart_sheet._charts[0].x_axis.title.tx.rich.p[0].r[0].t,
                "Y",
            )

    def test_render_splits_series_by_multiple_columns(self):
        workbook = Workbook()
        sheet = workbook.active
        sheet.title = "data"
        sheet.append(["sample_num", "data_num", "x", "y"])
        sheet.append([1, 1, 1, 2])
        sheet.append([1, 1, 2, 4])
        sheet.append([1, 2, 1, 3])
        sheet.append([1, 2, 2, 6])

        spec = ChartSpec(
            source_sheet="data",
            output_sheet="chart",
            chart_type="scatter",
            x=AxisSpec("x"),
            series=(
                SeriesSpec(
                    AxisSpec("y"),
                    title="Stress",
                    grouping=GroupingSpec(("sample_num", "data_num")),
                ),
            ),
        )

        ChartRenderer().render(workbook, spec)
        chart = workbook["chart"]._charts[0]
        self.assertEqual(len(chart.series), 2)
        self.assertEqual(chart.series[0].tx.v, "sample_num=1, data_num=1")
        self.assertEqual(chart.series[1].tx.v, "sample_num=1, data_num=2")
