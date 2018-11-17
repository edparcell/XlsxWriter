import unittest
from ...canvas import CellFormat


class TestCellFormat(unittest.TestCase):
    def test_add_cell_formats(self):
        fmt_bold = CellFormat(bold=1)
        fmt_italic = CellFormat(italic=1)
        fmt_bold_italic = fmt_bold + fmt_italic
        assert {'bold': 1, 'italic': 1} == fmt_bold_italic._asdict(include_none=False)

    def test_add_cell_formats_overrides(self):
        fmt_tnr = CellFormat(font_name='Times New Roman')
        fmt_tahoma = CellFormat(font_name='Tahoma')
        fmt_result = fmt_tnr + fmt_tahoma
        assert {'font_name': 'Tahoma'} == fmt_result._asdict(include_none=False)