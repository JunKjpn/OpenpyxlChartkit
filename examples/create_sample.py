from pathlib import Path

from openpyxl import load_workbook

from openpyxl_chartkit import (
    AxisSpec,
    ChartRenderer,
    ChartSpec,
    GroupingSpec,
    LayoutSpec,
    SeriesSpec,
)


def build_ss_sample(input_path: Path, output_path: Path) -> None:
    workbook = load_workbook(input_path)

    spec = ChartSpec(
        source_sheet="mean-data",
        output_sheet="mean-data",
        chart_type="scatter",
        x=AxisSpec(column="Strain %", title="Strain (%)"),
        series=(
            SeriesSpec(
                y=AxisSpec(column="Stress MPa", title="Stress (MPa)"),
                title="Stress MPa",
                style="line",
                grouping=GroupingSpec(columns=("SampleNum", "DataNum"))
            ),
        ),
        layout=LayoutSpec(width=20, height=15, placement="right_up"),
        title_template="Stress-Strain Chart",
    )

    ChartRenderer().render(workbook, spec)
    workbook.save(output_path)


if __name__ == "__main__":
    project_root = Path(__file__).resolve().parents[1]
    input_path = project_root / "samples" / "ss_sample_data.xlsx"
    output_path = project_root / "outputs" / "ss_sample_chart.xlsx"
    output_path.parent.mkdir(exist_ok=True)
    build_ss_sample(input_path=input_path, output_path=output_path)
