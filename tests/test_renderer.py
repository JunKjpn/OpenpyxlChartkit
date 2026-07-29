import tempfile
import unittest
from pathlib import Path

from openpyxl import Workbook, load_workbook

from openpyxl_chartkit import AxisSpec, ChartRenderer, ChartSpec, SeriesSpec


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
