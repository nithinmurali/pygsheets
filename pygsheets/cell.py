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
        self._worksheet = worksheet
        if type(pos) == str:
            pos = format_addr(pos, 'tuple')
        self._row, self._col = pos
        self._label = format_addr(pos, 'label')
        self._value = val  # formated vlaue
        self._unformated_value = val  # unformated vlaue
        self._formula = ''
        self._format = FormatType.CUSTOM
        self._format_pattern = None
        self.parse_value = True  # if set false, value will be shown as it is
        self._note = ''
        self._simplecell = True  # if format, notes etc wont fe fetced on each update
        if self._worksheet is None:
            self._linked = False
        else:
            self._linked = True

    @property
    def row(self):
        """Row number of the cell."""
        return self._row

    @row.setter
    def row(self, row):
        if self._linked:
            ncell = self._worksheet.cell((row, self.col))
            self.__dict__.update(ncell.__dict__)
        else:
            self._row = row

    @property
    def col(self):
        """Column number of the cell."""
        return self._col

    @col.setter
    def col(self, col):
        if self._linked:
            ncell = self._worksheet.cell((self._row, col))
            self.__dict__.update(ncell.__dict__)
        else:
            self._col = col

    @property
    def label(self):
        """Cell Label - Eg A1"""
        return self._label

    @label.setter
    def label(self, label):
        if self._linked:
            ncell = self._worksheet.cell(label)
            self.__dict__.update(ncell.__dict__)
        else:
            self._label = label

    @property
    def value(self):
        """formatted value of the cell"""
        return self._value

    @value.setter
    def value(self, value):
        if self._linked:
            self._worksheet.update_cell(self.label, value, self.parse_value)
            if not self._simplecell:
                self.fetch()
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
        self.update()

    def unlink(self):
        """unlink the cell from worksheet"""
        self._linked = True
        return self

    def link(self, worksheet=None, update=False):
        """
        link cell with a worksheet

        :param worksheet: the worksheet to link to
        :param update: if the cell should be synces as after linking
        :return: :class: Cell
        """
        if worksheet is None and self._worksheet is None:
            raise InvalidArgumentValue("Worksheet not set for uplink")
        self._linked = False
        if worksheet:
            self._worksheet = worksheet
        if update:
            self.update()
        return self

    def set_format(self, format_type, pattern=None):
        """
        set cell format

        :param format_type: number format of the cell as enum FormatType
        :param pattern: Pattern string used for formatting.
        :return:
        """
        self._simplecell = False
        self._format = format_type
        self._format_pattern = pattern
        if not self._linked:
            return False
        if not isinstance(format_type, FormatType):
            raise InvalidArgumentValue("format_type")
        request = {
            "repeatCell": {
                "range": {
                    "sheetId": self._worksheet.id,
                    "startRowIndex": self.row - 1,
                    "endRowIndex": self.row,
                    "startColumnIndex": self.col - 1,
                    "endColumnIndex": self.col
                },
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {
                            "type": format_type.value,
                            "pattern": pattern
                        }
                    }
                },
                "fields": "userEnteredFormat.numberFormat"
            }
        }
        self._worksheet.client.sh_batch_update(self._worksheet.spreadsheet.id, request, None, False)
        self.fetch()
        return self

    def neighbour(self, position):
        """
        get a neighbouring cell of this cell

        :param position: a tuple of relative position of position as string as
                        right, left, top, bottom or combinatoin
        :return: Cell object of neighbouring cell
        """
        if not self._linked:
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
            ncell = self._worksheet.cell(tuple(addr))
        except IncorrectCellLabel:
            raise CellNotFound
        return ncell

    def fetch(self):
        """ Update the value of the cell from sheet """
        self._simplecell = False
        if self._linked:
            self._value = self._worksheet.cell(self._label).value
            result = self._worksheet.client.sh_get_ssheet(self._worksheet.spreadsheet.id, fields='sheets/data/rowData',
                                                          include_data=True,
                                                          ranges=self._worksheet._get_range(self.label))
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
        if not self._linked:
            return False
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
                "cell": {
                    "userEnteredFormat": {
                        "numberFormat": {
                            "type": self._format.value,
                            "pattern": self._format_pattern
                        }
                    },
                    "note": self._note,
                },
                "fields": "userEnteredFormat.numberFormat, note, userEnteredValue.stringValue"
            }
        }
        self._worksheet.client.sh_batch_update(self._worksheet.spreadsheet.id, request, None, False)

    def __eq__(self, other):
        if self._worksheet is not None and other._worksheet is not None:
            if self._worksheet != other._worksheet:
                return False
        if self.label != other.label:
            return False
        return True

    def __repr__(self):
        return '<%s R%sC%s %s>' % (self.__class__.__name__,
                                   self.row,
                                   self.col,
                                   repr(self.value))

