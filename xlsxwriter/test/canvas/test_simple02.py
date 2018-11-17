from __future__ import with_statement
from ..excel_comparsion_test import ExcelComparisonTest
from ...workbook import Workbook
from ...canvas import Canvas, CellFormat


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename('simple02.xlsx')

    def test_create_file(self):
        """Test the creation of a simple workbook."""

        workbook = Workbook(self.got_filename)

        worksheet1 = workbook.add_worksheet()
        workbook.add_worksheet('Data Sheet')
        worksheet3 = workbook.add_worksheet()

        c = Canvas(2, 1)
        c.set_data(0, 0, 'Foo')
        c.set_data(1, 0, 123)
        c.write_to_sheet(workbook, worksheet1)

        bold = CellFormat(bold=1)
        c = Canvas(4, 3)
        c.set_data(1, 1, 'Foo')
        c.set_data(2, 1, 'Bar')
        c.set_data(3, 2, 234)
        c.set_cell_format(2, 1, 1, 1, bold)
        c.write_to_sheet(workbook, worksheet3)

        workbook.close()

        self.assertExcelEqual()
