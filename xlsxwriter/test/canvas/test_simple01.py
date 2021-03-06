from __future__ import with_statement
from ..excel_comparsion_test import ExcelComparisonTest
from ...workbook import Workbook
from ...canvas import Canvas


class TestCompareXLSXFiles(ExcelComparisonTest):
    """
    Test file created by XlsxWriter against a file created by Excel.

    """

    def setUp(self):
        self.set_filename('simple01.xlsx')

    def test_create_file(self):
        """Test the creation of a simple workbook."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        c = Canvas(2, 1)
        c.set_data(0, 0, 'Hello')
        c.set_data(1, 0, 123)

        c.write_to_sheet(workbook, worksheet)

        workbook.close()

        self.assertExcelEqual()

    def test_create_file_with_array(self):
        """Test the creation of a simple workbook."""

        workbook = Workbook(self.got_filename)
        worksheet = workbook.add_worksheet()

        c = Canvas(2, 1)
        c.set_data(0, 0, [['Hello'], [123]])

        c.write_to_sheet(workbook, worksheet)

        workbook.close()

        self.assertExcelEqual()
