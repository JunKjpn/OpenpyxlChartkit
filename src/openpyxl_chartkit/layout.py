import math

from openpyxl.utils import get_column_letter

class ChartLayoutEngine:
    _WIDTH_TO_COLUMNS = 0.525
    _HEIGHT_TO_ROWS = 2.1
    DEFAULT_TOP_ROW = 2
    DEFAULT_LEFT_COLUMN = 1
    DEFAULT_MARGIN = 2

    def __init__(self, worksheet: Worksheet, layout: LayoutSpec):
        self._worksheet = worksheet
        self._layout = layout
        self._origin_row, self._origin_column = self._origin()
        self._row_step = (
                math.ceil(layout.height * self._HEIGHT_TO_ROWS) + (layout.vertical_interval or 0)
        )
        self._column_step = (
                math.ceil(layout.width * self._WIDTH_TO_COLUMNS) + (layout.horizontal_interval or 0)
        )

    def _origin(self) -> tuple[int, int]:
        if self._layout.placement == "left_down":
            return self._worksheet.max_row + self.DEFAULT_MARGIN, self.DEFAULT_LEFT_COLUMN
        elif self._layout.placement == "right_up":
            return self.DEFAULT_TOP_ROW, self._worksheet.max_column + self.DEFAULT_MARGIN
        else:
            return self.DEFAULT_TOP_ROW, self.DEFAULT_LEFT_COLUMN

    def anchor(self, grid_row: int, grid_column: int) -> str:
        row = self._origin_row + grid_row * self._row_step
        col = self._origin_column + grid_column * self._column_step
        return f"{get_column_letter(col)}{row}"