import unittest
from ...canvas import Canvas, CellFormat


class TestCanvas(unittest.TestCase):
    def test_set_format(self):
        c = Canvas(3, 3)
        c.set_data(0, 0, [[0, 1, 2], [3, 4, 5], [6, 7, 8]])
        fmt_italic = CellFormat(italic=1)
        c.set_cell_format(0, 0, 2, 2, fmt_italic)
        for i in range(3):
            for j in range(3):
                if i < 2 and j < 2:
                    assert c.cells[i][j].cell_format.italic == 1

    def test_add_format(self):
        c = Canvas(3, 3)
        c.set_data(0, 0, [[0, 1, 2], [3, 4, 5], [6, 7, 8]])
        fmt_italic = CellFormat(italic=1)
        c.set_cell_format(0, 0, 2, 2, fmt_italic)
        fmt_bold = CellFormat(bold=1)
        c.add_cell_format(1, 1, 2, 2, fmt_bold)
        assert {'italic': 1} == c.cells[0][0].cell_format._asdict(include_none=False)
        assert {'bold': 1, 'italic': 1} == c.cells[1][1].cell_format._asdict(include_none=False)
        assert {'bold': 1} == c.cells[2][2].cell_format._asdict(include_none=False)

