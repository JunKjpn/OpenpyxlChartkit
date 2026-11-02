class ChartLayoutSession:
    """Manage chart group positions across multiple ChartSpec renders."""

    def __init__(self, vertical_gap: int = 2):
        if vertical_gap < 0:
            raise ValueError("vertical_gap must be non-negative.")

        self._vertical_gap = vertical_gap
        self._next_row = 1

    def allocate(self, rows: int, row_step: int, chart_height_rows: int) -> tuple[int, int]:
        if rows < 1:
            raise ValueError("rows must be at least 1.")
        if row_step < 1:
            raise ValueError("row_step must be at least 1.")
        if chart_height_rows < 1:
            raise ValueError("chart_height_rows must be at least 1.")

        origin = (self._next_row, 1)

        occupied_rows = (
            (rows - 1) * row_step
            + chart_height_rows
        )

        self._next_row += occupied_rows + self._vertical_gap

        return origin
