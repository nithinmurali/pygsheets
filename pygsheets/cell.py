# -*- coding: utf-8 -*-.

"""
pygsheets.cell
~~~~~~~~~~~~~~

This module contains cell model

"""

# import warnings
from .custom_types import *
from .exceptions import (IncorrectCellLabel, CellNotFound, InvalidArgumentValue)
from .utils import format_addr


class Cell(object):
    """An instance of this class represents a single cell
    in a :class:`worksheet <Worksheet>`.

    """

    def __init__(self, pos, val='', worksheet=None):
        self.worksheet = worksheet
        if type(pos) == str:
            pos = format_addr(pos, 'tuple')
        self._row, self._col = pos
        self._label = format_addr(pos, 'label')
        self._value = val  # formated vlaue
        self._unformated_value = val  # unformated vlaue
        self._formula = ''
        self._note = ''
        self._color = (1.0, 1.0, 1.0, 1.0)
        self._simplecell = True

        self.format = (FormatType.CUSTOM, '')
        """tuple specifying data format (format type, pattern) or just format"""
        self.parse_value = True
        """if set false, value will be shown as it is"""
        self.text_format = {"foregroundColor": {}, "fontFamily": '', "fontSize": 10, "bold": False, "italic": False,
                            "strikethrough": False, "underline": False}
        """the text format as json"""
        self.borders = {}
        """border properties as json"""

    @property
    def row(self):
        """Row number of the cell."""
        return self._row

    @row.setter
    def row(self, row):
        if self.worksheet:
            ncell = self.worksheet.cell((row, self.col))
            self.__dict__.update(ncell.__dict__)
        else:
            self._row = row

    @property
    def col(self):
        """Column number of the cell."""
        return self._col

    @col.setter
    def col(self, col):
        if self.worksheet:
            ncell = self.worksheet.cell((self._row, col))
            self.__dict__.update(ncell.__dict__)
        else:
            self._col = col

    @property
    def label(self):
        """Cell Label - Eg A1"""
        return self._label

    @label.setter
    def label(self, label):
        if self.worksheet:
            ncell = self.worksheet.cell(label)
            self.__dict__.update(ncell.__dict__)
        else:
            self._label = label

    @property
    def value(self):
        """formatted value of the cell"""
        return self._value

    @value.setter
    def value(self, value):
        if self.worksheet:
            self.worksheet.update_cell(self.label, value, self.parse_value)
            if not self._simplecell:
                self.fetch()
        else:
            self._value = value

    @property
    def value_unformatted(self):
        """ return unformatted value of the cell """
        return self._unformated_value

    @property
    def formula(self):
        """formula if any of the cell"""
        # self.fetch()
        return self._formula

    @formula.setter
    def formula(self, formula):
        if not formula.startswith('='):
            formula = "=" + formula
        tmp = self.parse_value
        self.parse_value = True
        self.value = formula
        self._formula = formula
        self.parse_value = tmp
        self.fetch()

    @property
    def note(self):
        """note on the cell"""
        return self._note

    @note.setter
    def note(self, note):
        self._note = note

    @property
    def color(self):
        """background color of the cell as (red, green, blue, alpha)"""
        return self._color

    @color.setter
    def color(self, value):
        if type(value) is tuple:
            if len(value) < 4:
                value = list(value) + [1.0]*(4-len(value))
        else:
            value = (value, 1.0, 1.0, 1.0)
        self._color = tuple(value)

    def unlink(self):
        """unlink the cell from worksheet"""
        self.worksheet = None

    def neighbour(self, position):
        """
        get a neighbouring cell of this cell

        :param position: a tuple of relative position of position as string as
                        right, left, top, bottom or combinatoin
        :return: Cell object of neighbouring cell
        """
        if not self.worksheet:
            return False
        addr = [self.row, self.col]
        if type(position) == tuple:
            addr = (addr[0] + position[0], addr[1] + position[1])
        elif type(position) == str:
            if "right" in position:
                addr[1] += 1
            if "left" in position:
                addr[1] -= 1
            if "top" in position:
                addr[0] -= 1
            if "bottom" in position:
                addr[0] += 1
        try:
            ncell = self.worksheet.cell(tuple(addr))
        except IncorrectCellLabel:
            raise CellNotFound
        return ncell

    def fetch(self):
        """ Update the value of the cell from sheet """
        self._simplecell = False
        if self.worksheet:
            self._value = self.worksheet.cell(self._label).value
            result = self.worksheet.client.sh_get_ssheet(self.worksheet.spreadsheet.id, fields='sheets/data/rowData',
                                                         include_data=True,
                                                         ranges=self.worksheet._get_range(self.label))
            result = result['sheets'][0]['data'][0]['rowData'][0]['values'][0]
            self._value = result.get('formattedValue','')
            try:
                self._unformated_value = list(result['effectiveValue'].values())[0]
            except KeyError:
                self._unformated_value = ''
            self._formula = result.get('userEnteredValue', {}).get('formulaValue', '')
            self._note = result.get('note', '')
            nformat = result.get('userEnteredFormat', {}).get('numberFormat', {})
            self.format = (nformat.get('type', FormatType.CUSTOM), nformat.get('pattern', ''))
            self.color = tuple(result.get('userEnteredFormat', {})
                                .get('backgroundColor', {'r':1.0,'g':1.0,'b':1.0,'a':1.0}).values())
            self.text_format = result.get('userEnteredFormat', {}).get('textFormat', {})
            self.borders = result.get('userEnteredFormat', {}).get('borders', {})
            return self
        else:
            return False

    def update(self):
        """update the sheet cell value with the attributes set """
        self._simplecell = False
        request = {
            "repeatCell": {
                "range": {
                    "sheetId": self.worksheet.id,
                    "startRowIndex": self.row - 1,
                    "endRowIndex": self.row,
                    "startColumnIndex": self.col - 1,
                    "endColumnIndex": self.col
                },
                "cell": self.get_json(),
                "fields": "userEnteredFormat, note"
            }
        }
        self.worksheet.client.sh_batch_update(self.worksheet.spreadsheet.id, request, None, False)
        self.value = self._value  # @TODO combine to above?

    def get_json(self):
        """get the json representation of the cell as per google api"""
        try:
            nformat, pattern = self.format
        except ValueError:
            nformat, pattern = self.format, ""
        json_repr = {
                    "userEnteredFormat": {
                        "numberFormat": {
                            "type": nformat.value,
                            "pattern": pattern
                        },
                        "backgroundColor": {
                            "red": self._color[0],
                            "green": self._color[1],
                            "blue": self._color[2],
                            "alpha": self._color[3],
                        },
                        "textFormat": self.text_format,
                        "borders": self.borders
                    },
                    "note": self._note,
                }
        return json_repr

    def __eq__(self, other):
        if self.worksheet is not None and other.worksheet is not None:
            if self.worksheet != other.worksheet:
                return False
        if self.label != other.label:
            return False
        return True

    def __repr__(self):
        return '<%s R%sC%s %s>' % (self.__class__.__name__,
                                   self.row,
                                   self.col,
                                   repr(self.value))

