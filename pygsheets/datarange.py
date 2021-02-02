# -*- coding: utf-8 -*-.

"""
pygsheets.datarange
~~~~~~~~~~~~~~~~~~~

This module contains DataRange class.

"""

import logging

from pygsheets.address import GridRange
from pygsheets.exceptions import InvalidArgumentValue, CellNotFound


class DataRange(object):
    """
    DataRange specifies a range of cells in the sheet. It can be unbounded on one or more axes.
    DataRange is for storing/manipulating a range of data in worksheet. This class can be used for
    group operations, e.g. changing format of all cells in a given range. This can also
    represent named ranges protected ranges, banned ranges etc.

    All the proteted range properties are stored in protected_properties.


    :param start: top left cell address. can be unbounded.
    :param end: bottom right cell address
    :param worksheet: worksheet where this range belongs
    :param name: name of the named range
    :param data: data of the range in as row major matrix
    :param name_id: id of named range
    :param namedjson: json representing the NamedRange from api

    >>> drange = Datarange(start='A1', end='B2', worksheet=wks)
    <Datarange Sheet1!A1:B2>
    >>> drange.name = "my_named_range" # make this datarange a named range
    <Datarange my_named_range Sheet1!A1:B2>
    >>> drange.protected = True # make the range protected
    <Datarange my_named_range Sheet1!A1:B2 protected>
    >>> drange.start_addr = 'B' # make the range unbounded on rows
    <Datarange my_named_range Sheet1!A:B protected>
    >>> drange.end_addr = None # make the range unbounded on both axes
    <Datarange my_named_range Sheet1 protected>

    """

    def __init__(self, start=None, end=None, worksheet=None, name='', data=None, name_id=None, namedjson=None,
                 protectedjson=None, grange=None):
        self._worksheet = worksheet
        self.logger = logging.getLogger(__name__)
        self.protected_properties = ProtectedRangeProperties()
        if grange:
            self.grid_range = grange
        else:
            self.grid_range = GridRange(worksheet=worksheet, start=start, end=end)

        if namedjson:
            self.grid_range.set_json(namedjson['range'])
            name_id = namedjson['namedRangeId']
            name = namedjson['name']
        if protectedjson:
            self.grid_range.set_json(protectedjson['range'])
            name_id = protectedjson.get('namedRangeId', '')  # @TODO get the name also
            self.protected_properties = ProtectedRangeProperties(protectedjson)

        if data:
            if len(data) == self.grid_range.height and len(data[0]) == self.grid_range.width:
                self._data = data
            else:
                self.fetch()
        else:
            self._data = [[]]

        self._linked = True
        self._name_id = name_id
        self._name = name

    @property
    def name(self):
        """name of the named range. setting a name will make this a range a named range
            setting this to empty string will delete the named range
        """
        return self._name

    @name.setter
    def name(self, name):
        if type(name) is not str:
            raise InvalidArgumentValue('name should be a string')
        if not name:
            self._worksheet.delete_named_range(range_id=self._name_id)
            self._name = ''
            self._name_id = ''
        else:
            if not self._name_id:
                # @TODO handle when not linked (create an range on link)
                if not self._linked:
                    self.logger.warn("unimplimented bahaviour")
                api_obj = self._worksheet.create_named_range(name, grange=self.grid_range, returnas='json')
                self._name = name
                self._name_id = api_obj['namedRangeId']
            else:
                self._name = name
                self.update_named_range()

    @property
    def name_id(self):
        """ if of the named range """
        return self._name_id

    # Protected properties TODO move to protected properties @next_version

    @property
    def protect_id(self):
        """ id of the protected range """
        return self.protected_properties.protected_id

    @property
    def protected(self):
        """get/set the range as protected
        setting this to False will make this range unprotected
        """
        return self.protected_properties.is_protected()

    @protected.setter
    def protected(self, value):
        if value:
            if not self.protected:
                resp = self._worksheet.create_protected_range(grange=self.grid_range, named_range_id=self._name_id,
                                                              returnas='json')
                self.protected_properties.set_json(resp)
        else:
            if self.protected:
                self._worksheet.remove_protected_range(self.protect_id)
                self.protected_properties.clear()

    @property
    def editors(self):
        """
        Lists the editors of the protected range
        can also set a list of editors, take a tuple ('users' or 'groups', [<editors>])
        can also set ('domainUsersCanEdit', Boolean)
        """
        return self.protected_properties.editors

    @editors.setter
    def editors(self, value):
        if type(value) is not tuple or value[0] not in ['users', 'groups', 'domainUsersCanEdit']:
            raise InvalidArgumentValue
        self.protected_properties.editors[value[0]] = value[1]
        self.update_protected_range(fields='editors')

    @property
    def requesting_user_can_edit(self):
        """ if the requesting user can edit protected range """
        return self.protected_properties.requestingUserCanEdit

    @requesting_user_can_edit.setter
    def requesting_user_can_edit(self, value):
        self.protected_properties.requestingUserCanEdit = value
        self.update_protected_range(fields='requestingUserCanEdit')

    @property
    def description(self):
        """ if the requesting user can edit protected range """
        return self.protected_properties.description

    @description.setter
    def description(self, value):
        self.protected_properties.description = value
        self.update_protected_range(fields='description')

    # End protected properties

    @property
    def start_addr(self):
        """top-left address of the range"""
        return self.grid_range.start.index

    @start_addr.setter
    def start_addr(self, addr):
        self.grid_range.start = addr
        self.update_named_range()
        self.update_protected_range()

    @property
    def end_addr(self):
        """bottom-right address of the range"""
        return self.grid_range.end.index

    @end_addr.setter
    def end_addr(self, addr):
        self.grid_range.end = addr
        self.update_named_range()
        self.update_protected_range()

    @property
    def range(self):
        """Range in format A1:C5"""
        return self.grid_range.start.label + ':' + self.grid_range.end.label

    @property
    def worksheet(self):
        """ linked worksheet """
        return self._worksheet

    @property
    def cells(self):
        """Get cells of this range"""
        if len(self._data[0]) == 0:
            self.fetch()
        return self._data

    def link(self, update=True):
        """link the datarange so that all properties are synced right after setting them

        :param update: if the range should be synced to cloud on link
        """
        if not self._worksheet:
            raise InvalidArgumentValue("No worksheet defined to link this range to.")
        self._linked = True

        [[y.link(worksheet=self._worksheet, update=update) for y in x] for x in self._data]
        if update:
            self.update_protected_range()
            self.update_named_range()

    def unlink(self):
        """unlink the sheet so that all properties are not synced as it is changed"""
        self._linked = False
        [[y.unlink() for y in x] for x in self._data]

    def fetch(self, only_data=True):
        """
        update the range data/properties from cloud

        .. warning::
                Currently only data is fetched not properties, so `only_data` wont work

        :param only_data: fetch only data

        """
        self._data = self._worksheet.get_values(grange=self.grid_range, returnas='cells',
                                                include_tailing_empty_rows=True, include_tailing_empty=True)
        if not only_data:
            logging.error("functionality not implimented")

    def apply_format(self, cell, fields=None, cell_json=None):
        """
        Change format of all cells in the range

        :param cell: a model :class: Cell whose format will be applied to all cells
        :param fields: comma seprated string of fields of cell to apply, refer to `google api docs <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells#CellData>`__
        :param cell_json: if not providing a cell object, provide a cell json. refer to  `google api docs <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells>`__
        """
        if not cell_json:
            cell_json = cell.get_json()
        request = {"repeatCell": {
            "range": self._get_gridrange(),
            "cell": cell_json,
            "fields": fields or "userEnteredFormat,hyperlink,note,textFormatRuns,dataValidation,pivotTable"
            }
        }
        self._worksheet.client.sheet.batch_update(self._worksheet.spreadsheet.id, request)

    def update_values(self, values=None):
        """
        Update the worksheet with values of the cells in this range

        :param values: values as matrix, which has same size as the range

        """
        if self._linked and values:
            self._worksheet.update_values(crange=self.range, values=values)
            self.fetch()
        if self._linked and not values:
            self._worksheet.update_values(cell_list=self._data)

    def sort(self, basecolumnindex=0, sortorder="ASCENDING"):
        """sort the values in the datarange

        :param basecolumnindex:     Index of the base column in which sorting is to be done (Integer).
                                    The index here is the index of the column in range (first columen is 0).
        :param sortorder:           either "ASCENDING" or "DESCENDING" (String)
        """
        self._worksheet.sort_range(grange=self.grid_range,
                                   basecolumnindex=basecolumnindex + self.grid_range.start[1]-1,
                                   sortorder=sortorder)

    def clear(self, fields="userEnteredValue"):
        """
        Clear values in this datarange.

        Reference:

            -  `FieldMask Api object <https://developers.google.com/protocol-buffers/docs/reference/google.protobuf#google.protobuf.FieldMask>`_

        :param fields: Comma separated list of field masks.

        """
        self._worksheet.clear(grange=self.grid_range, fields=fields)

    def update_named_range(self):
        """update the named range properties"""
        if not self._name_id or not self._linked:
            return False
        if self.protected:
            self.update_protected_range()
        request = {'updateNamedRange': {
          "namedRange": {
              "namedRangeId": self._name_id,
              "name": self._name,
              "range": self._get_gridrange(),
          },
          "fields": '*',
        }}
        self._worksheet.client.sheet.batch_update(self._worksheet.spreadsheet.id, request)

    def update_protected_range(self, fields='*'):
        """ update the protected range properties """
        if not self.protected or not self._linked:
            return False

        request = {'updateProtectedRange': {
          "protectedRange": self.protected_properties.to_json(),
          "fields": fields,
        }}
        if self._name_id:
            request['updateProtectedRange']['protectedRange']['namedRangeId'] = self._name_id
        else:
            request['updateProtectedRange']['protectedRange']['range'] = self._get_gridrange()
        self._worksheet.client.sheet.batch_update(self._worksheet.spreadsheet.id, request)

    def update_borders(self, top=False, right=False, bottom=False, left=False, inner_horizontal=False, inner_vertical=False,
                       style='NONE', width=1, red=0, green=0, blue=0):
        """
        update borders for range

        NB  use style='NONE' to erase borders
            default color is black

        :param top: make a top border
        :param right: make a right border
        :param bottom: make a bottom border
        :param left: make a left border
        :param style: either 'SOLID', 'DOTTED', 'DASHED', 'SOLID', 'SOLID_MEDIUM', 'SOLID_THICK', 'DOUBLE' or 'NONE' (String).
        :param width: border width (depreciated) (Integer).
        :param red: 0-255 (Integer).
        :param green: 0-255 (Integer).
        :param blue: 0-255 (Integer).
        """
        if not (top or right or bottom or left):
            return

        if style not in ['SOLID', 'DOTTED', 'DASHED', 'SOLID', 'SOLID_MEDIUM', 'SOLID_THICK', 'DOUBLE', 'NONE']:
            raise ValueError('specified value is not a valid border style')

        request = {"updateBorders": {"range": self._get_gridrange()}}

        border = {
            "style": style,
            "width": width,
            "color": {
                "red": red,
                "green": green,
                "blue": blue
            }}

        if top:
            request["updateBorders"]["top"] = border
        if bottom:
            request["updateBorders"]["bottom"] = border
        if left:
            request["updateBorders"]["left"] = border
        if right:
            request["updateBorders"]["right"] = border
        if inner_horizontal:
            request["updateBorders"]["innerHorizontal"] = border
        if inner_vertical:
            request["updateBorders"]["innerVertical"] = border

        self._worksheet.client.sheet.batch_update(self._worksheet.spreadsheet.id, request)

    def merge_cells(self, merge_type='MERGE_ALL'):
        """
        Merge cells in range

        ! You can't vertically merge cells that intersect an existing filter

        :param merge_type: either   'MERGE_ALL'
                                    ,'MERGE_COLUMNS'  ( = merge multiple rows (!) together to make column(s))
                                    ,'MERGE_ROWS' ( = merge multiple columns (!) together to make a row(s))
                                    ,'NONE' (unmerge)
        """
        self._worksheet.merge_cells(merge_type=merge_type, grange=self.grid_range)

    def _get_gridrange(self):
        return self.grid_range.to_json()

    def __getitem__(self, item):
        if len(self._data[0]) == 0 and self.grid_range.width > 0:
            self.fetch()
        if type(item) is int:
            try:
                return self._data[item]
            except IndexError:
                raise CellNotFound

    def __eq__(self, other):
        return self.start_addr == other.start_addr and self.end_addr == other.end_addr\
               and self.name == other.name and self.protect_id == other.protect_id

    def __repr__(self):
        range_str = self.range
        if self.worksheet:
            range_str = str(self.grid_range.label)
        protected_str = " protected" if self.protected else ""

        return '<%s %s %s%s>' % (self.__class__.__name__, str(self._name), range_str, protected_str)


class ProtectedRangeProperties(object):

    def __init__(self, api_obj=None):
        self.protected_id = None
        self.description = None
        self.warningOnly = None
        self.requestingUserCanEdit = None
        self.editors = None
        if api_obj:
            self.set_json(api_obj)

    def set_json(self, api_obj):
        if type(api_obj) is not dict:
            raise InvalidArgumentValue
        self.protected_id = api_obj['protectedRangeId']
        self.description = api_obj.get('description', '')
        self.editors = api_obj.get('editors', {})
        self.warningOnly = api_obj.get('warningOnly', False)
        self.requestingUserCanEdit = api_obj.get('requestingUserCanEdit', None)

    def to_json(self):
        api_obj = {
            "protectedRangeId": self.protected_id,
            "description": self.description,
            "warningOnly": self.warningOnly,
            "requestingUserCanEdit": self.requestingUserCanEdit,
            "editors": self.editors
        }
        return api_obj

    def is_protected(self):
        return self.protected_id is not None

    def clear(self):
        self.protected_id = None
        self.description = None
        self.warningOnly = None
        self.requestingUserCanEdit = None
        self.editors = None


