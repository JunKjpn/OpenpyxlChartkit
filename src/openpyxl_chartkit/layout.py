import math

from openpyxl.utils import get_column_letter

from .specs import LayoutSpec

class ChartLayoutEngine:
    _WIDTH_TO_COLUMNS = 0.525
    _HEIGHT_TO_ROWS = 2.1

    def __init__(self, layout: LayoutSpec):
        self._layout = layout
        self._row_step = math.ceil(layout.height * self._HEIGHT_TO_ROWS) + (layout.vertical_interval or 0)
        self._column_step = math.ceil(layout.width * self._WIDTH_TO_COLUMNS) + (layout.horizontal_interval or 0)

    def anchor(self, grid_row: int, grid_column: int, origin_row: int = 1, origin_column: int = 1) -> str:
        row = origin_row + grid_row * self._row_step
        col = origin_column + grid_column * self._column_step
        return f"{get_column_letter(col)}{row}"

    @property
    def row_step(self) -> int:
        return self._row_step
