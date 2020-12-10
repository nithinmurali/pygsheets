# -*- coding: utf-8 -*-.

"""
pygsheets.cell
~~~~~~~~~~~~~~

This module represents a cell within the worksheet.

"""

from pygsheets.custom_types import *
from pygsheets.exceptions import (IncorrectCellLabel, CellNotFound, InvalidArgumentValue)
from pygsheets.utils import format_addr, is_number, format_color
from pygsheets.address import Address, GridRange


class Cell(object):
    """
    Represents a single cell of a sheet.

    Each cell is either a simple local value or directly linked to a specific cell of a sheet. When linked any
    changes to the cell will update the :class:`Worksheet <Worksheet>` immediately.

    :param pos:         Address of the cell as coordinate tuple or label.
    :param val:         Value stored inside of the cell.
    :param worksheet:   Worksheet this cell belongs to.
    :param cell_data:   This cells data stored in json, with the same structure as cellData of the Google Sheets API v4.
    """

    def __init__(self, pos, val='', worksheet=None, cell_data=None):
        self._worksheet = worksheet

        self._address = Address(pos, False)

        self._value = val  # formatted value
        self._unformated_value = val  # un-formatted value
        self._formula = ''
        self._note = None
        if self._worksheet is None:
            self._linked = False
        else:
            self._linked = True
        self._parent = None
        self._color = (None, None, None, None)
        self._simplecell = True  # if format, notes etc wont be fetched on each update
        self.format = (None, None)  # number format
        self.text_format = {}  # the text format as json
        self.text_rotation = None  # the text rotation as json

        self._horizontal_alignment = None
        self._vertical_alignment = None
        self.borders = None
        """Border Properties as dictionary. 
        Reference: `api object <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#borders>`__."""
        self.parse_value = True
        """Determines how values are interpreted by Google Sheets (True: USER_ENTERED; False: RAW).
        
        Reference: `sheets api <https://developers.google.com/sheets/api/reference/rest/v4/ValueInputOption>`__"""
        self._wrap_strategy = None
        self.is_dirty = True

        if cell_data is not None:
            self.set_json(cell_data)

    @property
    def row(self):
        """Row number of the cell."""
        return self.address.row

    @row.setter
    def row(self, row):
        self.address = (row, self.col)

    @property
    def col(self):
        """Column number of the cell."""
        return self.address.col

    @col.setter
    def col(self, col):
        self.address = (self.row, col)

    @property
    def label(self):
        """This cells label (e.g. 'A1')."""
        return self.address.label

    @label.setter
    def label(self, label):
        self.address = label

    @property
    def address(self):
        """ Address object representing the cell location. """
        return self._address

    @address.setter
    def address(self, value):
        if self._linked:
            ncell = self._worksheet.cell(value)
            self.__dict__.update(ncell.__dict__)
        else:
            self._address = Address(value)

    @property
    def value(self):
        """This cells formatted value."""
        return self._value

    @value.setter
    def value(self, value):
        self._value = value
        if self._linked:
            self._worksheet.update_value(self.label, value, self.parse_value)
            if not self._simplecell:  # for unformated value and formula
                self.fetch()
        else:
            self._formula = value if str(value).startswith('=') else ''
            self._unformated_value = ''

    @property
    def value_unformatted(self):
        """Unformatted value of this cell."""
        return self._unformated_value

    @property
    def formula(self):
        """Get/Set this cells formula if any."""
        if self._simplecell:
            self.fetch()
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
    def horizontal_alignment(self):
        """Horizontal alignment of the value in this cell.
           possible vlaues: :class:`HorizontalAlignment <pygsheets.custom_types.HorizontalAlignment>` """
        self.update()
        return self._horizontal_alignment

    @horizontal_alignment.setter
    def horizontal_alignment(self, value):
        if isinstance(value, HorizontalAlignment):
            self._horizontal_alignment = value
            self.update()
        else:
            raise InvalidArgumentValue('Use HorizontalAlignment object for setting the horizontal alignment.')

    @property
    def vertical_alignment(self):
        """Vertical alignment of the value in this cell.
            possible vlaues: :class:`VerticalAlignment <pygsheets.custom_types.VerticalAlignment>` """
        self.update()
        return self._vertical_alignment

    @vertical_alignment.setter
    def vertical_alignment(self, value):
        if isinstance(value, VerticalAlignment):
            self._vertical_alignment = value
            self.update()
        else:
            raise InvalidArgumentValue('Use VerticalAlignment for setting the vertical alignment.')

    @property
    def wrap_strategy(self):
        """
        How to wrap text in this cell.
        Possible wrap strategies: 'OVERFLOW_CELL', 'LEGACY_WRAP', 'CLIP', 'WRAP'.
        `Reference: api docs <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#wrapstrategy>`__
        """  
        return self._wrap_strategy

    @wrap_strategy.setter
    def wrap_strategy(self, wrap_strategy):
        self._wrap_strategy = wrap_strategy
        self.update()

    @property
    def note(self):
        """Get/Set note of this cell."""
        if self._simplecell:
            self.fetch()
        return self._note

    @note.setter
    def note(self, note):
        if self._simplecell:
            self.fetch()
        self._note = note
        self.update()

    @property
    def color(self):
        """Get/Set background color of this cell as a tuple (red, green, blue, alpha)."""
        if self._simplecell:
            self.fetch()
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

    @property
    def simple(self):
        """Simple cells only fetch the value itself. Set to false to fetch all cell properties."""
        return self._simplecell

    @simple.setter
    def simple(self, value):
        self._simplecell = value

    def set_text_format(self, attribute, value):
        """
        Set a text format property of this cell.

        Each format property must be set individually. Any format property which is not set will be considered
        unspecified.

        Attribute:
            - foregroundColor:    Sets the texts color. (tuple as (red, green, blue, alpha))
            - fontFamily:         Sets the texts font. (string)
            - fontSize:           Sets the text size. (integer)
            - bold:               Set/remove bold format. (boolean)
            - italic:             Set/remove italic format. (boolean)
            - strikethrough:      Set/remove strike through format. (boolean)
            - underline:          Set/remove underline format. (boolean)

        Reference: `api docs <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#textformat>`__

        :param attribute:   The format property to set.
        :param value:       The value the format property should be set to.

        :return: :class:`cell <Cell>`
        """
        if self._simplecell:
            self.fetch()
        if attribute not in ["foregroundColor", "fontFamily", "fontSize", "bold", "italic",
                             "strikethrough", "underline"]:
            raise InvalidArgumentValue("Not a valid attribute. Check documentation for more information.")
        if self.text_format:
            self.text_format[attribute] = value
        else:
            self.text_format = {attribute: value}
        self.update()
        return self

    def set_number_format(self, format_type, pattern=''):
        """
        Set number format of this cell.

        Reference: `api docs <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#NumberFormat>`__

        :param format_type: The type of the number format. Should be of type :class:`FormatType <FormatType>`.
        :param pattern: Pattern string used for formatting. If not set, a default pattern will be used.
                        See reference for supported patterns.
        :return: :class:`cell <Cell>`

        """
        if not isinstance(format_type, FormatType):
            raise InvalidArgumentValue("format_type should be of type pygsheets.FormatType")
        if self._simplecell:
            self.fetch()
        self.format = (format_type, pattern)
        self.update()
        return self

    def set_text_rotation(self, attribute, value):
        """
        The rotation applied to text in this cell.

        Can be defined as "angle" or as "vertical". May not define both!

        angle:
            [number] The angle between the standard orientation and the desired orientation.
            Measured in degrees. Valid values are between -90 and 90. Positive angles are angled upwards,
            negative are angled downwards.

            Note: For LTR text direction positive angles are in the counterclockwise direction,
            whereas for RTL they are in the clockwise direction.

        vertical:
            [boolean] If true, text reads top to bottom, but the orientation of individual characters is unchanged.

        Reference: `api_docs <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#textrotation>__`

        :param attribute:   "angle" or "vertical"
        :param value:       Corresponding value for the attribute. angle in (-90,90) for 'angle', boolean for 'vertical'
        :return: :class:`cell <Cell>`
        """
        if self._simplecell:
            self.fetch()
        if attribute not in ["angle", "vertical"]:
            raise InvalidArgumentValue("Text rotation can be set as 'angle' or 'vertical'. "
                                       "See documentation for details.")
        if attribute == "angle":
            if type(value) != int:
                raise InvalidArgumentValue("Property 'angle' must be an int.")
            if value not in range(-90, 91):
                raise InvalidArgumentValue("Property 'angle' must be in range -90 and 90.")
        if attribute == "vertical":
            if type(value) != bool:
                raise InvalidArgumentValue("Property 'vertical' must be set as boolean.")

        self.text_rotation = {attribute: value}
        self.update()
        return self

    def set_horizontal_alignment(self, value):
        """
        Set horizondal alignemnt of text in the cell

        :param value: Horizondal alignment value, instance of :class:`HorizontalAlignment <HorizontalAlignment>`
        :return: :class:`cell <Cell>`
        """
        if self._simplecell:
            self.fetch()
        self.horizontal_alignment = value
        return self

    def set_vertical_alignment(self, value):
        """
        Set vertical alignemnt of text in the cell

        :param value: Vertical alignment value, instance of :class:`VerticalAlignment <VerticalAlignment>`
        :return: :class:`cell <Cell>`
        """
        if self._simplecell:
            self.fetch()
        self.vertical_alignment = value
        return self

    def set_value(self, value):
        """
        Set value of the cell

        :param value: value to be set
        :return: :class:`cell <Cell>`
        """
        self.value = value
        return self

    def unlink(self):
        """Unlink this cell from its worksheet.

        Unlinked cells will no longer automatically update the sheet when changed. Use update or link to update the
        sheet."""
        self._linked = False
        self.is_dirty = False
        return self

    def link(self, worksheet=None, update=False):
        """
        Link cell with the specified worksheet.

        Linked cells will synchronize any changes with the sheet as they happen.

        :param worksheet:   The worksheet to link to. Can be None if the cell was linked to a worksheet previously.
        :param update:      Update the cell immediately after linking if the cell has changed
        :return: :class:`cell <Cell>`
        """
        if worksheet is None and self._worksheet is None:
            raise InvalidArgumentValue("No worksheet defined to link this cell to.")
        self._linked = True
        if worksheet:
            self._worksheet = worksheet
        if update and self.is_dirty:
            self.update()
        return self

    def neighbour(self, position):
        """
        Get a neighbouring cell of this cell.

        :param position:    This may be a string 'right', 'left', 'top', 'bottom' or a tuple of relative positions
                            (e.g. (1, 2) will return a cell one below and two to the right).
        :return: :class:`neighbouring cell <Cell>`
        """
        if not self._linked:
            return False
        addr = Address(self._address)
        if type(position) == tuple:
            addr = addr + position
        # TODO: this does not work if position is a list...
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
            ncell = self._worksheet.cell(addr)
        except IncorrectCellLabel:
            raise CellNotFound
        return ncell

    def fetch(self, keep_simple=False):
        """Update the value in this cell from the linked worksheet."""
        if not keep_simple: self._simplecell = False
        if self._linked:
            result = self._worksheet.client.sheet.get(self._worksheet.spreadsheet.id,
                                                      fields='sheets/data/rowData',
                                                      includeGridData=True,
                                                      ranges=self._worksheet._get_range(self.label))
            try:
                result = result['sheets'][0]['data'][0]['rowData'][0]['values'][0]
            except (KeyError, IndexError):
                result = dict()
            self.set_json(result)
            return self
        else:
            return False

    def refresh(self):
        """Refresh the value and properties in this cell from the linked worksheet.
           Same as fetch.
        """
        self.fetch(False)

    def update(self, force=False, get_request=False, worksheet_id=None):
        """
        Update the cell of the linked sheet or the worksheet given as parameter.

        :param force:           Force an update from the sheet, even if it is unlinked.
        :param get_request:     Return the request object instead of sending the request directly.
        :param worksheet_id:    Needed if the cell is not linked otherwise the cells worksheet is used.
        """
        if not (self._linked or force) and not get_request:
            return False
        self._simplecell = False
        worksheet_id = worksheet_id if worksheet_id is not None else self._worksheet.id
        request = {
            "repeatCell": {
                "range": GridRange(start=self._address, end=self._address, worksheet_id=worksheet_id).to_json(),
                "cell": self.get_json(),
                "fields": "userEnteredFormat, note, userEnteredValue"
            }
        }
        if get_request:
            return request
        self._worksheet.client.sheet.batch_update(self._worksheet.spreadsheet.id, request)

    def get_json(self):
        """Returns the cell as a dictionary structured like the Google Sheets API v4."""
        try:
            nformat, pattern = self.format
        except TypeError:
            nformat, pattern = self.format, ""

        if self._formula != '':
            value = self._formula
            value_key = 'formulaValue'
        elif is_number(self._value):
            value = self._value
            value_key = 'numberValue'
        elif type(self._value) is bool:
            value = self._value
            value_key = 'boolValue'
        elif type(self._value) is str or type(self._value) is unicode:
            value = self._value
            value_key = 'stringValue'
        else:   # @TODO errorValue key not handled
            value = self._value
            value_key = 'errorValue'

        ret_json = dict()
        ret_json["userEnteredFormat"] = dict()

        if self.format[0] is not None:
            ret_json["userEnteredFormat"]["numberFormat"] = {"type": getattr(nformat, 'value', nformat),
                                                             "pattern": pattern}
        if self._color[0] is not None:
            ret_json["userEnteredFormat"]["backgroundColor"] = {"red": self._color[0], "green": self._color[1],
                                                                "blue": self._color[2], "alpha": self._color[3]}
        if self.text_format is not None:
            ret_json["userEnteredFormat"]["textFormat"] = self.text_format.copy()
            fg = ret_json["userEnteredFormat"]["textFormat"].get('foregroundColor', None)
            ret_json["userEnteredFormat"]["textFormat"]['foregroundColor'] = format_color(fg, to='dict')

        if self.borders is not None:
            ret_json["userEnteredFormat"]["borders"] = self.borders
        if self._horizontal_alignment is not None:
            ret_json["userEnteredFormat"]["horizontalAlignment"] = self._horizontal_alignment.value
        if self._vertical_alignment is not None:
            ret_json["userEnteredFormat"]["verticalAlignment"] = self._vertical_alignment.value
        if self._wrap_strategy is not None:
            ret_json["userEnteredFormat"]["wrapStrategy"] = self._wrap_strategy
        if self.text_rotation is not None:
            ret_json["userEnteredFormat"]["textRotation"] = self.text_rotation

        if self._note is not None:
            ret_json["note"] = self._note
        ret_json["userEnteredValue"] = {value_key: value}

        return ret_json

    def set_json(self, cell_data):
        """
        Reads a json-dictionary returned by the Google Sheets API v4 and initialize all the properties from it.

        :param cell_data:   The cells data.
        """
        self._simplecell = False

        self._value = cell_data.get('formattedValue', '')
        try:
            self._unformated_value = list(cell_data['effectiveValue'].values())[0]
        except (KeyError, IndexError):
            self._unformated_value = ''
        self._formula = cell_data.get('userEnteredValue', {}).get('formulaValue', '')

        self._note = cell_data.get('note', None)
        nformat = cell_data.get('userEnteredFormat', {}).get('numberFormat', {})
        self.format = (nformat.get('type', None), nformat.get('pattern', ''))
        color = cell_data.get('userEnteredFormat', {}) \
            .get('backgroundColor', {'red': None, 'green': None, 'blue': None, 'alpha': None})

        self._color = (color.get('red', 0), color.get('green', 0), color.get('blue', 0), color.get('alpha', 0))
        self.text_format = cell_data.get('userEnteredFormat', {}).get('textFormat', None)
        if self.text_format and self.text_format.get('foregroundColor', None):
            self.text_format['foregroundColor'] = format_color(self.text_format['foregroundColor'], to='tuple')
        self.text_rotation = cell_data.get('userEnteredFormat', {}).get('textRotation', None)
        self.borders = cell_data.get('userEnteredFormat', {}).get('borders', None)
        self._wrap_strategy = cell_data.get('userEnteredFormat', {}).get('wrapStrategy', "WRAP_STRATEGY_UNSPECIFIED")

        nhorozondal_alignment = cell_data.get('userEnteredFormat', {}).get('horizontalAlignment', None)
        self._horizontal_alignment = \
            HorizontalAlignment[nhorozondal_alignment] if nhorozondal_alignment is not None else None
        nvertical_alignment = cell_data.get('userEnteredFormat', {}).get('verticalAlignment', None)
        self._vertical_alignment = \
            VerticalAlignment[nvertical_alignment] if nvertical_alignment is not None else None

        self.hyperlink = cell_data.get('hyperlink', '')
        
    def __setattr__(self, key, value):
        if key not in ['_linked', '_worksheet']:
            self.__dict__['is_dirty'] = True
        super(Cell, self).__setattr__(key, value)

    def __eq__(self, other):
        if self._worksheet is not None and other._worksheet is not None:
            if self._worksheet != other._worksheet:
                return False
        if self.label != other.label:
            return False
        return True

    def __repr__(self):
        return '<%s %s %s>' % (self.__class__.__name__, self.label, repr(self.value))
