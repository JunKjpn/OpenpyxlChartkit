from pathlib import Path

from openpyxl import load_workbook

from openpyxl_chartkit import AxisSpec, ChartRenderer, ChartSpec, GroupingSpec, LayoutSpec, SeriesSpec, filter_rows, \
    wide_to_long


def build_vis_sample(input_path: Path, output_path: Path) -> None:
    workbook = load_workbook(input_path)
    source = workbook["data"]

    # Filter DataNum = 2.
    filtered = filter_rows(
        source=source,
        column="DataNum",
        value=2,
        output_sheet_name="filtered_data",
    )

    # Convert wide-format data into long-format data.
    chart_data = wide_to_long(
        source=filtered,
        id_columns=("SampleNum", "Strain %"),
        value_columns=("G' MPa", 'G" MPa', "tan-delta"),
        variable_name="Metric",
        value_name="Value",
        output_sheet_name="chart_data",
        sort_by=("Metric", "SampleNum", "Strain %"),
    )

    # Define chart specification.
    spec = ChartSpec(
        source_sheet=chart_data.title,
        output_sheet="charts",
        chart_type="scatter",
        x=AxisSpec(
            column="Strain %",
            title="Strain (%)",
            log_scale=True,
        ),
        series=(
            SeriesSpec(
                y=AxisSpec(column="Value"),
                title="Value",
                style="line",
                grouping=GroupingSpec(
                    columns=("SampleNum",),
                ),
            ),
        ),
        horizontal_grouping=GroupingSpec(
            columns=("Metric",),
        ),
        layout=LayoutSpec(
            width=20,
            height=12,
        ),
        title_template="VIS Chart",
    )

    ChartRenderer().render(workbook, spec)

    del workbook["filtered_data"]

    output_path.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(output_path)


if __name__ == "__main__":
    project_root = Path(__file__).resolve().parents[1]

    input_path = project_root / "samples" / "vis_sample_data.xlsx"
    output_path = project_root / "outputs" / "vis_sample_chart.xlsx"

    build_vis_sample(
        input_path=input_path,
        output_path=output_path,
    )