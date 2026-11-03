import pytest
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


def test_render_horizontal_grouping_with_multiple_specs():
    workbook = create_workbook_with_data(
        headers=["year", "sales", "profit", "region"],
        rows=[
            [2023, 100, 20, "East"],
            [2024, 120, 25, "East"],
            [2023, 80, 15, "West"],
            [2024, 90, 18, "West"],
        ],
    )

    grouped_spec = ChartSpec(
        source_sheet="data",
        output_sheet="chart",
        chart_type="scatter",
        x=AxisSpec("year"),
        series=(SeriesSpec(AxisSpec("sales")),),
        horizontal_grouping=GroupingSpec(("region",)),
        title_template="Sales by region",
    )

    second_spec = ChartSpec(
        source_sheet="data",
        output_sheet="chart",
        chart_type="scatter",
        x=AxisSpec("year"),
        series=(SeriesSpec(AxisSpec("profit")),),
        title_template="Profit",
    )

    renderer = ChartRenderer()
    session = ChartLayoutSession(vertical_gap=2)

    renderer.render(workbook, grouped_spec, session=session)
    renderer.render(workbook, second_spec, session=session)

    charts = workbook["chart"]._charts

    assert len(charts) == 3

    # Horizontal grouping
    assert charts[0].anchor == "A1"
    assert charts[1].anchor == "G1"

    # The next ChartSpec is placed below the previous group.
    assert charts[2].anchor == "A20"


def test_render_vertical_grouping_with_multiple_specs():
    workbook = create_workbook_with_data(
        headers=["year", "sales", "profit", "region"],
        rows=[
            [2023, 100, 20, "East"],
            [2024, 120, 25, "East"],
            [2023, 80, 15, "West"],
            [2024, 90, 18, "West"],
        ],
    )

    grouped_spec = ChartSpec(
        source_sheet="data",
        output_sheet="chart",
        chart_type="scatter",
        x=AxisSpec("year"),
        series=(SeriesSpec(AxisSpec("sales")),),
        vertical_grouping=GroupingSpec(("region",)),
        title_template="Sales by region",
    )

    second_spec = ChartSpec(
        source_sheet="data",
        output_sheet="chart",
        chart_type="scatter",
        x=AxisSpec("year"),
        series=(SeriesSpec(AxisSpec("profit")),),
        title_template="Profit",
    )

    renderer = ChartRenderer()
    session = ChartLayoutSession(vertical_gap=2)

    renderer.render(workbook, grouped_spec, session=session)
    renderer.render(workbook, second_spec, session=session)

    charts = workbook["chart"]._charts

    assert len(charts) == 3

    # Vertical grouping
    assert charts[0].anchor == "A1"
    assert charts[1].anchor == "A18"

    # The next ChartSpec is placed below the previous group.
    assert charts[2].anchor == "A37"


def test_render_two_dimensional_grouping_with_multiple_specs():
    workbook = create_workbook_with_data(
        headers=["year", "sales", "region", "product"],
        rows=[
            [2023, 100, "East", "A"],
            [2024, 120, "East", "A"],
            [2023, 80, "East", "B"],
            [2024, 90, "East", "B"],
            [2023, 110, "West", "A"],
            [2024, 130, "West", "A"],
            [2023, 70, "West", "B"],
            [2024, 85, "West", "B"],
        ],
    )

    grouped_spec = ChartSpec(
        source_sheet="data",
        output_sheet="chart",
        chart_type="scatter",
        x=AxisSpec("year"),
        series=(SeriesSpec(AxisSpec("sales")),),
        horizontal_grouping=GroupingSpec(("product",)),
        vertical_grouping=GroupingSpec(("region",)),
        title_template="Sales",
    )

    second_spec = ChartSpec(
        source_sheet="data",
        output_sheet="chart",
        chart_type="scatter",
        x=AxisSpec("year"),
        series=(SeriesSpec(AxisSpec("sales")),),
        title_template="Total sales",
    )

    renderer = ChartRenderer()
    session = ChartLayoutSession(vertical_gap=2)

    renderer.render(workbook, grouped_spec, session=session)
    renderer.render(workbook, second_spec, session=session)

    charts = workbook["chart"]._charts

    assert len(charts) == 5

    # Two-dimensional grouping
    assert charts[0].anchor == "A1"
    assert charts[1].anchor == "G1"
    assert charts[2].anchor == "A18"
    assert charts[3].anchor == "G18"

    # The next ChartSpec is below the entire 2D group.
    assert charts[4].anchor == "A37"


def test_render_two_dimensional_grouping_with_uneven_groups():
    workbook = create_workbook_with_data(
        headers=["year", "sales", "region", "product"],
        rows=[
            [2023, 100, "East", "A"],
            [2024, 120, "East", "A"],
            [2023, 80, "East", "B"],
            [2024, 90, "East", "B"],
            [2023, 110, "West", "A"],
            [2024, 130, "West", "A"],
            [2023, 70, "West", "B"],
            [2024, 85, "West", "B"],
            [2023, 60, "West", "C"],
            [2024, 75, "West", "C"],
        ],
    )

    grouped_spec = ChartSpec(
        source_sheet="data",
        output_sheet="chart",
        chart_type="scatter",
        x=AxisSpec("year"),
        series=(SeriesSpec(AxisSpec("sales")),),
        horizontal_grouping=GroupingSpec(("product",)),
        vertical_grouping=GroupingSpec(("region",)),
    )

    second_spec = ChartSpec(
        source_sheet="data",
        output_sheet="chart",
        chart_type="scatter",
        x=AxisSpec("year"),
        series=(SeriesSpec(AxisSpec("sales")),),
    )

    renderer = ChartRenderer()
    session = ChartLayoutSession(vertical_gap=2)

    renderer.render(workbook, grouped_spec, session=session)
    renderer.render(workbook, second_spec, session=session)

    charts = workbook["chart"]._charts

    assert len(charts) == 6

    # East: 2 charts
    assert charts[0].anchor == "A1"
    assert charts[1].anchor == "G1"

    # West: 3 charts
    assert charts[2].anchor == "A18"
    assert charts[3].anchor == "G18"
    assert charts[4].anchor == "M18"

    # The next ChartSpec starts below the entire grid.
    assert charts[5].anchor == "A37"


def test_render_non_contiguous_groups_raise_error():
    workbook = create_workbook_with_data(
        headers=["year", "sales", "region"],
        rows=[
            [2023, 100, "East"],
            [2024, 120, "East"],
            [2023, 80, "West"],
            [2024, 90, "West"],
            [2023, 110, "East"],
            [2024, 130, "East"],
        ],
    )

    spec = ChartSpec(
        source_sheet="data",
        output_sheet="chart",
        chart_type="scatter",
        x=AxisSpec("year"),
        series=(SeriesSpec(AxisSpec("sales")),),
        horizontal_grouping=GroupingSpec(("region",)),
    )

    with pytest.raises(ValueError, match="not contiguous"):
        ChartRenderer().render(workbook, spec)

def test_render_uses_grouping_value_as_y_axis_title():
    workbook = create_workbook_with_data(
        headers=["metric", "x", "y"],
        rows=[
            ["G' MPa", 1, 10],
            ["G' MPa", 2, 20],
            ['G" MPa', 1, 30],
            ['G" MPa', 2, 40],
            ["tan-delta", 1, 50],
            ["tan-delta", 2, 60],
        ],
    )

    spec = ChartSpec(
        source_sheet="data",
        output_sheet="chart",
        chart_type="scatter",
        x=AxisSpec("x"),
        series=(SeriesSpec(AxisSpec("y")),),
        horizontal_grouping=GroupingSpec(("metric",)),
        y_axis_title_from="metric",
    )

    ChartRenderer().render(workbook, spec)

    charts = workbook["chart"]._charts

    assert len(charts) == 3

    titles = [
        chart.y_axis.title.tx.rich.p[0].r[0].t
        for chart in charts
    ]

    assert titles == ["G' MPa", 'G" MPa', "tan-delta"]

def test_render_raises_error_when_y_axis_title_value_is_not_unique():
    workbook = create_workbook_with_data(
        headers=["metric", "x", "y"],
        rows=[
            ["G' MPa", 1, 10],
            ['G" MPa', 2, 20],
        ],
    )

    spec = ChartSpec(
        source_sheet="data",
        output_sheet="chart",
        chart_type="scatter",
        x=AxisSpec("x"),
        series=(SeriesSpec(AxisSpec("y")),),
        y_axis_title_from="metric",
    )

    with pytest.raises(ValueError, match="exactly one distinct value"):
        ChartRenderer().render(workbook, spec)

def test_render_scatter_chart_with_column_grouping():
    workbook = create_workbook_with_data(
        headers=["Strain", "Sample A", "Sample B", "Sample C"],
        rows=[
            [1, 10, 12, 15],
            [2, 20, 22, 25],
            [3, 30, 32, 35],
        ],
    )

    spec = ChartSpec(
        source_sheet="data",
        output_sheet="chart",
        chart_type="scatter",
        x=AxisSpec("Strain"),
        series=(
            SeriesSpec(
                AxisSpec("Sample A"),
                grouping=GroupingSpec(
                    columns=("Sample A", "Sample B", "Sample C"),
                    source="column",
                ),
            ),
        ),
    )

    ChartRenderer().render(workbook, spec)

    chart = workbook["chart"]._charts[0]

    # Three series are generated.
    assert len(chart.series) == 3

    # Series titles correspond to column names.
    assert [series.tx.v for series in chart.series] == [
        "Sample A",
        "Sample B",
        "Sample C",
    ]

    # All series use the same X-axis range.
    for series in chart.series:
        assert series.xVal.numRef.f.endswith("$A$2:$A$4")

    # Each series uses its corresponding Y-axis column.
    assert chart.series[0].yVal.numRef.f.endswith("$B$2:$B$4")
    assert chart.series[1].yVal.numRef.f.endswith("$C$2:$C$4")
    assert chart.series[2].yVal.numRef.f.endswith("$D$2:$D$4")


def test_render_column_grouping_raises_error_for_missing_column():
    workbook = create_workbook_with_data(
        headers=["Strain", "Sample A", "Sample B"],
        rows=[
            [1, 10, 12],
            [2, 20, 22],
        ],
    )

    spec = ChartSpec(
        source_sheet="data",
        output_sheet="chart",
        chart_type="scatter",
        x=AxisSpec("Strain"),
        series=(
            SeriesSpec(
                AxisSpec("Sample A"),
                grouping=GroupingSpec(
                    columns=("Sample A", "Sample C"),
                    source="column",
                ),
            ),
        ),
    )

    with pytest.raises(ValueError, match="missing columns"):
        ChartRenderer().render(workbook, spec)


def test_render_column_grouping_does_not_require_y_axis_column():
    workbook = create_workbook_with_data(
        headers=["Strain", "Sample A", "Sample B"],
        rows=[
            [1, 10, 12],
            [2, 20, 22],
            [3, 30, 32],
        ],
    )

    spec = ChartSpec(
        source_sheet="data",
        output_sheet="chart",
        chart_type="scatter",
        x=AxisSpec("Strain"),
        series=(
            SeriesSpec(
                AxisSpec("UnusedColumn"),
                grouping=GroupingSpec(
                    columns=("Sample A", "Sample B"),
                    source="column",
                ),
            ),
        ),
    )

    ChartRenderer().render(workbook, spec)

    chart = workbook["chart"]._charts[0]

    assert len(chart.series) == 2
    assert chart.series[0].yVal.numRef.f.endswith("$B$2:$B$4")
    assert chart.series[1].yVal.numRef.f.endswith("$C$2:$C$4")


def test_render_column_grouping_raises_error_for_empty_data():
    workbook = create_workbook_with_data(
        headers=["Strain", "Sample A", "Sample B"],
        rows=[],
    )

    spec = ChartSpec(
        source_sheet="data",
        output_sheet="chart",
        chart_type="scatter",
        x=AxisSpec("Strain"),
        series=(
            SeriesSpec(
                AxisSpec("Sample A"),
                grouping=GroupingSpec(
                    columns=("Sample A", "Sample B"),
                    source="column",
                ),
            ),
        ),
    )

    with pytest.raises(ValueError, match="no data rows"):
        ChartRenderer().render(workbook, spec)
