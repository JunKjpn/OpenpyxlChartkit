import pytest

from openpyxl_chartkit.session import ChartLayoutSession


def test_allocate_starts_at_a1():
    session = ChartLayoutSession()

    assert session.allocate(
        rows=1,
        row_step=19,
        chart_height_rows=17,
    ) == (1, 1)


def test_allocate_moves_down_by_previous_group_height():
    session = ChartLayoutSession(vertical_gap=2)

    first = session.allocate(
        rows=2,
        row_step=19,
        chart_height_rows=17,
    )
    second = session.allocate(
        rows=1,
        row_step=19,
        chart_height_rows=17,
    )

    assert first == (1, 1)
    assert second == (39, 1)


def test_allocate_multiple_groups():
    session = ChartLayoutSession(vertical_gap=2)

    assert session.allocate(
        rows=2,
        row_step=19,
        chart_height_rows=17,
    ) == (1, 1)

    assert session.allocate(
        rows=1,
        row_step=19,
        chart_height_rows=17,
    ) == (39, 1)

    assert session.allocate(
        rows=3,
        row_step=19,
        chart_height_rows=17,
    ) == (58, 1)


@pytest.mark.parametrize("vertical_gap", [-1, -5])
def test_negative_vertical_gap_raises_error(vertical_gap):
    with pytest.raises(ValueError):
        ChartLayoutSession(vertical_gap=vertical_gap)


@pytest.mark.parametrize(
    ("rows", "row_step", "chart_height_rows"),
    [
        (0, 19, 17),
        (-1, 19, 17),
        (1, 0, 17),
        (1, -1, 17),
        (1, 19, 0),
        (1, 19, -1),
    ],
)
def test_invalid_dimensions_raise_error(
    rows,
    row_step,
    chart_height_rows,
):
    session = ChartLayoutSession()

    with pytest.raises(ValueError):
        session.allocate(
            rows=rows,
            row_step=row_step,
            chart_height_rows=chart_height_rows,
        )

def test_allocate_multiple_rows():
    session = ChartLayoutSession(vertical_gap=2)

    first = session.allocate(
        rows=2,
        row_step=19,
        chart_height_rows=17,
    )
    second = session.allocate(
        rows=1,
        row_step=17,
        chart_height_rows=17,
    )

    assert first == (1, 1)
    assert second == (39, 1)