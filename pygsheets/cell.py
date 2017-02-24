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
        self._simplecell = True

        self.format = (FormatType.CUSTOM, None)
        """tuple specifying data format (format type, pattern) or just format"""
        self.parse_value = True
        """if set false, value will be shown as it is"""
        self.bg_color = (1.0, 1.0, 1.0, 1.0)
        """background color of the cell as (red, green, blue, alpha)"""
        self.fg_color = (1.0, 1.0, 1.0, 1.0)
        """foreground color of the cell as (red, green, blue, alpha)"""
        self.text_format = {}
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
            try:
                self._value = result['formattedValue']
                self._unformated_value = list(result['effectiveValue'].values())[0]
            except KeyError:
                self._value = ''
                self._unformated_value = ''
            try:
                self._formula = result['userEnteredValue']['formulaValue']
            except KeyError:
                self._formula = ''
            try:
                self._note = result['note']
            except KeyError:
                self._note = ''

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
                "fields": "userEnteredFormat, note, userEnteredValue.stringValue"
            }
        }
        self.worksheet.client.sh_batch_update(self.worksheet.spreadsheet.id, request, None, False)

    def get_json(self):
        """get the json representation of the cell as per google api"""
        try:
            nformat, pattern = self.format
        except ValueError:
            nformat, pattern = self.format, ""
        json_repr = {"cell": {
                    "userEnteredFormat": {
                        "numberFormat": {
                            "type": nformat.value,
                            "pattern": pattern
                        },
                        "backgroundColor": {
                            "red": self.bg_color[0],
                            "green": self.bg_color[1],
                            "blue": self.bg_color[2],
                            "alpha": self.bg_color[3],
                        },
                        "textFormat": self.text_format,
                        "borders": self.borders
                    },
                    "note": self._note,
                },
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

