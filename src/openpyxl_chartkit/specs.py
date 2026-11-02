from dataclasses import dataclass, field
from typing import Literal


ChartType = Literal["scatter", "bar"]
SeriesStyle = Literal["line", "circle"]
SplitAxis = Literal["row", "column"]


@dataclass(frozen=True)
class AxisSpec:
    """A worksheet column used as a chart axis."""

    column: str
    title: str | None = None
    log_scale: bool = False
    minimum: float | None = None
    maximum: float | None = None


@dataclass(frozen=True)
class GroupingSpec:
    """The columns used to split a chart or a series."""

    columns: tuple[str, ...]
    source: SplitAxis = "row"


@dataclass(frozen=True)
class SeriesSpec:
    """A single Y series in a chart."""

    y: AxisSpec
    title: str | None = None
    style: SeriesStyle = "line"
    grouping: GroupingSpec | None = None


@dataclass(frozen=True)
class LayoutSpec:
    """The dimensions and position of a chart on its output sheet."""

    width: float = 10
    height: float = 8
    horizontal_interval: int | None = None
    vertical_interval: int | None = None


@dataclass(frozen=True)
class ChartSpec:
    """The complete declarative definition of one chart."""

    source_sheet: str
    output_sheet: str
    chart_type: ChartType
    x: AxisSpec
    series: tuple[SeriesSpec, ...]
    horizontal_grouping: GroupingSpec | None = None
    vertical_grouping: GroupingSpec | None = None
    layout: LayoutSpec = field(default_factory=LayoutSpec)
    title_template: str | None = None
    y_axis_title_from: str | None = None
