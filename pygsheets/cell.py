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

    :param pos: position of the cell adress
    :param val: value of the cell
    :param worksheet: worksheet this cell belongs to
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
        self._note = ''
        if self._worksheet is None:
            self._linked = False
        else:
            self._linked = True
        self._color = (1.0, 1.0, 1.0, 1.0)
        self._simplecell = True  # if format, notes etc wont be fetched on each update
        self.format = (FormatType.CUSTOM, '')
        """tuple specifying data format (format type, pattern) or just format"""
        self.parse_value = True
        """if set false, value will be shown as it is"""
        self.text_format = {}
        """the text format as json"""
        self.text_rotation = {}
        """the text rotation as json"""
        self.borders = {}
        """border properties as json"""

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
            self._label = format_addr((self._row, self._col), 'label')

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
            self._label = format_addr((self._row, self._col), 'label')

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
            self._row, self._col = format_addr(label, 'tuple')

    @property
    def value(self):
        """formatted value of the cell"""
        return self._value

    @value.setter
    def value(self, value):
        if self._linked:
            self._worksheet.update_cell(self.label, value, self.parse_value)
            if not self._simplecell:  # for unformated value and formula
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
        # self.fetch()
        return self._note

    @note.setter
    def note(self, note):
        if self._simplecell:
            self.fetch()
        self._note = note
        self.update()

    @property
    def color(self):
        """background color of the cell as (red, green, blue, alpha)"""
        return self._color

    @color.setter
    def color(self, value):
        if self._simplecell:
            self.fetch()
        if type(value) is tuple:
            if len(value) < 4:
                value = list(value) + [1.0]*(4-len(value))
        else:
            value = (value, 1.0, 1.0, 1.0)
        for c in value:
            if c < 0 or c > 1:
                raise InvalidArgumentValue("Color should be in range 0-1")
        self._color = tuple(value)
        self.update()

    def set_text_format(self, attribute, value):
        """
        set the text format
        :param attribute: one of the following "foregroundColor" "fontFamily", "fontSize", "bold", "italic",
                            "strikethrough", "underline"
        :param value: corresponding value for the attribute
        :return: :class: Cell
        """
        if self._simplecell:
            self.fetch()
        if attribute not in ["foregroundColor" "fontFamily", "fontSize", "bold", "italic",
                            "strikethrough", "underline"]:
            raise InvalidArgumentValue("not a valid argument, please see the docs")
        self.text_format[attribute] = value
        self.update()
        return self

    def set_text_rotation(self, attribute, value):
        """
        set the text rotation
        :param attribute: "angle" or "vertical"
        :param value: corresponding value for the attribute
        :return: :class: Cell
        """
        if self._simplecell:
            self.fetch()
        if attribute not in ["angle", "vertical"]:
            raise InvalidArgumentValue("not a valid argument, please see the docs")
        if attribute == "angle":
            if type(value) != int:
                raise InvalidArgumentValue("angle value must be of type int")
            if value not in range(-90, 91):
                raise InvalidArgumentValue("angle value range must be between -90 and 90")
        if attribute == "vertical":
            if type(value) != bool:
                raise InvalidArgumentValue("vertical value must be of type bool")

        self.text_rotation = {attribute: value}
        self.update()
        return self

    def unlink(self):
        """unlink the cell from worksheet"""
        self._linked = False
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
        self._linked = True
        if worksheet:
            self._worksheet = worksheet
        if update:
            self.update()
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

    def fetch(self, keep_simple=False):
        """ Update the value of the cell from sheet """
        if not keep_simple: self._simplecell = False
        if self._linked:
            self._value = self._worksheet.cell(self._label).value
            result = self._worksheet.client.sh_get_ssheet(self._worksheet.spreadsheet.id, fields='sheets/data/rowData',
                                                          include_data=True,
                                                          ranges=self._worksheet._get_range(self.label))
            result = result['sheets'][0]['data'][0]['rowData'][0]['values'][0]
            self._value = result.get('formattedValue', '')
            try:
                self._unformated_value = list(result['effectiveValue'].values())[0]
            except KeyError:
                self._unformated_value = ''
            self._formula = result.get('userEnteredValue', {}).get('formulaValue', '')
            self._note = result.get('note', '')
            nformat = result.get('userEnteredFormat', {}).get('numberFormat', {})
            self.format = (nformat.get('type', FormatType.CUSTOM), nformat.get('pattern', ''))
            color = result.get('userEnteredFormat', {})\
                .get('backgroundColor', {'red':1.0,'green':1.0,'blue':1.0,'alpha':1.0})
            self._color = (color.get('red', 0), color.get('green', 0), color.get('blue', 0), color.get('alpha', 0))
            self.text_format = result.get('userEnteredFormat', {}).get('textFormat', {})
            self.text_rotation = result.get('userEnteredFormat', {}).get('textRotation', {})
            self.borders = result.get('userEnteredFormat', {}).get('borders', {})
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
                    "sheetId": self._worksheet.id,
                    "startRowIndex": self.row - 1,
                    "endRowIndex": self.row,
                    "startColumnIndex": self.col - 1,
                    "endColumnIndex": self.col
                },
                "cell": self.get_json(),
                "fields": "userEnteredFormat, note"
            }
        }
        self._worksheet.client.sh_batch_update(self._worksheet.spreadsheet.id, request, None, False)
        self.value = self._value  # @TODO combine to above?

    def get_json(self):
        """get the json representation of the cell as per google api"""
        try:
            nformat, pattern = self.format
        except TypeError:
            nformat, pattern = self.format, ""
        return {"userEnteredFormat": {
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
                        "borders": self.borders,
                        "textRotation": self.text_rotation
                    },
                "note": self._note,
                }

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

