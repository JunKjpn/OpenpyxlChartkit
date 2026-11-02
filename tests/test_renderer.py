from openpyxl import Workbook, load_workbook

from openpyxl_chartkit import AxisSpec, ChartRenderer, ChartSpec, GroupingSpec, SeriesSpec, ChartLayoutSession


def create_workbook_with_data(headers, rows):
    """Create a workbook with a data worksheet."""
    workbook = Workbook()
    sheet = workbook.active
    sheet.title = "data"
    sheet.append(headers)

    for row in rows:
        sheet.append(row)

    return workbook


def test_render_adds_scatter_chart(tmp_path):
    workbook = create_workbook_with_data(
        headers=["x", "y"],
        rows=[
            [1, 2],
            [2, 4],
        ],
    )

    spec = ChartSpec(
        source_sheet="data",
        output_sheet="chart",
        chart_type="scatter",
        x=AxisSpec("x", title="X"),
        series=(SeriesSpec(AxisSpec("y", title="Y"), title="Y values"),),
        title_template="Sample chart",
    )

    ChartRenderer().render(workbook, spec)

    output = tmp_path / "chart.xlsx"
    workbook.save(output)

    reopened = load_workbook(output)
    chart_sheet = reopened["chart"]

    assert len(chart_sheet._charts) == 1
    assert len(chart_sheet._charts[0].series) == 1

    # The Y-axis title is configured from AxisSpec("y", title="Y").
    assert chart_sheet._charts[0].y_axis.title.tx.rich.p[0].r[0].t == "Y"


def test_render_splits_series_by_multiple_columns():
    workbook = create_workbook_with_data(
        headers=["sample_num", "data_num", "x", "y"],
        rows=[
            [1, 1, 1, 2],
            [1, 1, 2, 4],
            [1, 2, 1, 3],
            [1, 2, 2, 6],
        ],
    )

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

    assert len(chart.series) == 2
    assert chart.series[0].tx.v == "sample_num=1, data_num=1"
    assert chart.series[1].tx.v == "sample_num=1, data_num=2"


def test_render_multiple_specs_without_overlap():
    workbook = create_workbook_with_data(
        headers=["x", "y1", "y2"],
        rows=[
            [1, 2, 3],
            [2, 4, 6],
            [3, 6, 9],
        ],
    )

    spec1 = ChartSpec(
        source_sheet="data",
        output_sheet="chart",
        chart_type="scatter",
        x=AxisSpec("x"),
        series=(SeriesSpec(AxisSpec("y1")),),
    )

    spec2 = ChartSpec(
        source_sheet="data",
        output_sheet="chart",
        chart_type="scatter",
        x=AxisSpec("x"),
        series=(SeriesSpec(AxisSpec("y2")),),
    )

    renderer = ChartRenderer()
    session = ChartLayoutSession()

    renderer.render(workbook, spec1, session=session)
    renderer.render(workbook, spec2, session=session)

    charts = workbook["chart"]._charts

    assert len(charts) == 2
    assert charts[0].anchor == "A1"
    assert charts[1].anchor == "A20"