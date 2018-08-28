from xlsxwriter.format import Format


def mk_data_class(typename, field_names):
    class C(object):
        __slots__ = field_names

        def __init__(self, **kwds):
            for x in field_names:
                v = kwds.pop(x, None)
                self.__setattr__(x, v)
            assert len(kwds) == 0

        def __repr__(self):
            field_vals = [(x, self.__getattribute__(x)) for x in field_names]
            field_vals = [(x, v) for x, v in field_vals if v is not None]
            s = ", ".join('{}={}'.format(x, repr(v)) for x, v in field_vals)
            return "{}({})".format(typename, s)

        def __eq__(self, other):
            if not isinstance(other, C):
                return False
            for x in field_names:
                if self.__getattribute__(x) != other.__getattribute__(x):
                    return False
            return True

        def __hash__(self):
            return hash(tuple(self.__getattribute__(x) for x in field_names))

        def _asdict(self, include_none=True):
            d = {x: self.__getattribute__(x) for x in field_names}
            if not include_none:
                d = {x: v for x, v in d.items() if v is not None}
            return d

    return C


CellFormat = mk_data_class('CellFormat', [
        'font_name', 'font_size', 'font_color', 'bold', 'italic', 'underline', 'font_strikeout', 'font_script',
        'num_format',
        'locked', 'hidden',
        'align', 'valign', 'rotation', 'text_wrap', 'reading_order', 'text_justlast', 'center_across', 'indent', 'shrink',
        'pattern', 'bg_color', 'fg_color'
    ])


BorderFormat = mk_data_class('BorderFormat', ['format', 'color'])


Cell = mk_data_class('Cell', ['data', 'cell_format', 'bottom_format', 'top_format', 'left_format', 'right_format'])


class Canvas(object):
    def __init__(self, height, width):
        self.width = width
        self.height = height
        self.cells = [[Cell() for _ in range(width)] for _ in range(height)]

    def set_data(self, row, column, data):
        if hasattr(data, '__len__') and not isinstance(data, str):
            if len(data) > 0 and hasattr(data[0], '__len__') and not isinstance(data[0], str):
                for i, row_data in enumerate(data):
                    for j, v in enumerate(row_data):
                        self.cells[row + i][column + j].data = v
            else:
                for j, v in enumerate(data):
                    self.cells[row][column + j].data = v
        else:
            self.cells[row][column].data = data

    def set_cell_format(self, row, column, height, width, format):
        assert 0 <= row < row + height < self.height
        assert 0 <= column < column + width < self.width
        for i in range(row, row + height):
            for j in range(column, column + width):
                self.cells[i][j].cell_format = format

    def set_border_format(self, row, column, height, width, border_format):
        for i in range(row, row + height):
            self.cells[i][column].left_format = border_format
            if width > 0:
                self.cells[i][column + width - 1].right_format = border_format
        for j in range(column, column + width):
            self.cells[row][j].top_format = border_format
            if height > 0:
                self.cells[row + height - 1][j].bottom_format = border_format

    def write_to_sheet(self, workbook, worksheet):
        formats = {}
        for i, row in enumerate(self.cells):
            for j, cell in enumerate(row):
                t_format = (cell.cell_format, cell.bottom_format, cell.top_format, cell.left_format, cell.right_format)
                cell_format = formats.get(t_format)
                if cell_format is None:
                    if cell.cell_format is None:
                        d = {}
                    else:
                        d = cell.cell_format._asdict(include_none=False)

                    if cell.bottom_format is not None:
                        d['bottom'] = cell.bottom_format.format
                        d['bottom_color'] = cell.bottom_format.color

                    if cell.top_format is not None:
                        d['top'] = cell.top_formati.format
                        d['top_color'] = cell.top_format.color

                    if cell.left_format is not None:
                        d['left'] = cell.left_format.format
                        d['left_color'] = cell.left_format.color

                    if cell.right_format is not None:
                        d['right'] = cell.right_format.format
                        d['right_color'] = cell.right_format.color
                    if d == {}:
                        cell_format = None
                    else:
                        cell_format = workbook.add_format(d)
                    formats[t_format] = cell_format
                worksheet.write(i, j, cell.data, cell_format)
