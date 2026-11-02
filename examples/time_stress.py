from pathlib import Path

from openpyxl import load_workbook

from openpyxl_chartkit import AxisSpec, ChartRenderer, ChartSpec, GroupingSpec, LayoutSpec, SeriesSpec


def build_time_stress_sample(input_path: Path, output_path: Path) -> None:
    workbook = load_workbook(input_path)

    spec = ChartSpec(
        source_sheet="data",
        output_sheet="charts",
        chart_type="scatter",
        x=AxisSpec(column="Time ms", title="Time (ms)"),
        series=(
            SeriesSpec(
                y=AxisSpec(column="Stress MPa", title="Stress (MPa)"),
                title="Stress MPa",
                style="line",
                grouping=GroupingSpec(columns=("DataNum",)),
            ),
        ),
        horizontal_grouping=GroupingSpec(columns=("SampleNum",)),
        vertical_grouping=GroupingSpec(columns=("Cycle",)),
        layout=LayoutSpec(width=20, height=12),
        title_template="Time-Stress Chart",
    )

    ChartRenderer().render(workbook, spec)
    workbook.save(output_path)


if __name__ == "__main__":
    project_root = Path(__file__).resolve().parents[1]
    input_path = project_root / "samples" / "time_stress_sample_data.xlsx"
    output_path = project_root / "outputs" / "time_stress_sample_chart.xlsx"

    output_path.parent.mkdir(parents=True, exist_ok=True)

    build_time_stress_sample(
        input_path=input_path,
        output_path=output_path,
    )
