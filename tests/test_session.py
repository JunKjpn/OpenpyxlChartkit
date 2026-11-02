import pytest

from openpyxl_chartkit.session import ChartLayoutSession


def test_allocate_starts_at_a1():
    session = ChartLayoutSession()

    assert session.allocate(rows=1, row_step=19) == (1, 1)


def test_allocate_moves_down_by_previous_group_height():
    session = ChartLayoutSession(vertical_gap=2)

    assert session.allocate(rows=2, row_step=19) == (1, 1)
    assert session.allocate(rows=1, row_step=19) == (41, 1)


def test_allocate_multiple_groups():
    session = ChartLayoutSession(vertical_gap=2)

    assert session.allocate(rows=2, row_step=19) == (1, 1)
    assert session.allocate(rows=1, row_step=19) == (41, 1)
    assert session.allocate(rows=3, row_step=19) == (62, 1)


@pytest.mark.parametrize("vertical_gap", [-1, -5])
def test_negative_vertical_gap_raises_error(vertical_gap):
    with pytest.raises(ValueError):
        ChartLayoutSession(vertical_gap=vertical_gap)


@pytest.mark.parametrize(
    ("rows", "row_step"),
    [
        (0, 19),
        (-1, 19),
        (1, 0),
        (1, -1),
    ],
)
def test_invalid_dimensions_raise_error(rows, row_step):
    session = ChartLayoutSession()

    with pytest.raises(ValueError):
        session.allocate(rows=rows, row_step=row_step)