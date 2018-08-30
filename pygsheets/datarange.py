# -*- coding: utf-8 -*-.

"""
pygsheets.datarange
~~~~~~~~~~~~~~~~~~~

This module contains DataRange class for storing/manipulating a range of data in spreadsheet. This class can
be used for group operations, e.g. changing format of all cells in a given range. This can also represent named ranges
protected ranges, banned ranges etc.

"""

import logging

from pygsheets.utils import format_addr
from pygsheets.exceptions import InvalidArgumentValue, CellNotFound


class DataRange(object):
    """
    DataRange specifies a range of cells in the sheet

    :param start: top left cell address
    :param end: bottom right cell address
    :param worksheet: worksheet where this range belongs
    :param name: name of the named range
    :param data: data of the range in as row major matrix
    :param name_id: id of named range
    :param namedjson: json representing the NamedRange from api
    """

    def __init__(self, start=None, end=None, worksheet=None, name='', data=None, name_id=None, namedjson=None, protectedjson=None):
        self._worksheet = worksheet
        self.logger = logging.getLogger(__name__)
        self._protected_properties = ProtectedRangeProperties()

        if namedjson:
            start = (namedjson['range'].get('startRowIndex', 0)+1, namedjson['range'].get('startColumnIndex', 0)+1)
            # @TODO this won't scale if the sheet size is changed
            end = (namedjson['range'].get('endRowIndex', self._worksheet.cols),
                   namedjson['range'].get('endColumnIndex', self._worksheet.rows))
            name_id = namedjson['namedRangeId']
        if protectedjson:
            # TODO dosent consider backing named range
            start = (protectedjson['range'].get('startRowIndex', 0)+1, protectedjson['range'].get('startColumnIndex', 0)+1)
            # @TODO this won't scale if the sheet size is changed
            end = (protectedjson['range'].get('endRowIndex', self._worksheet.cols),
                   protectedjson['range'].get('endColumnIndex', self._worksheet.rows))
            name_id = protectedjson.get('namedRangeId', '')  # @TODO get the name also
            self._protected_properties = ProtectedRangeProperties(protectedjson)

        self._start_addr = format_addr(start, 'tuple')
        self._end_addr = format_addr(end, 'tuple')
        if data:
            if len(data) == self._end_addr[0] - self._start_addr[0] + 1 and \
                            len(data[0]) == self._end_addr[1] - self._start_addr[1] + 1:
                self._data = data
            else:
                self.fetch()
        else:
            self.fetch()

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
                api_obj = self._worksheet.create_named_range(name, start=self._start_addr,
                                                             end=self._end_addr, returnas='json')
                self._name = name
                self._name_id = api_obj['namedRangeId']
            else:
                self._name = name
                self.update_named_range()

    @property
    def name_id(self):
        return self._name_id

    @property
    def protect_id(self):
        return self._protected_properties.protected_id

    @property
    def protected(self):
        """get/set range protection"""
        return self._protected_properties.is_protected()

    @protected.setter
    def protected(self, value):
        if value:
            if not self.protected:
                resp = self._worksheet.create_protected_range(self._get_gridrange())
                self._protected_properties.set_json(resp['replies'][0]['addProtectedRange']['protectedRange'])
        else:
            if self.protected:
                self._worksheet.remove_protected_range(self.protect_id)
                self._protected_properties.clear()

    @property
    def editors(self):
        return self._protected_properties.editors

    @editors.setter
    def editors(self, value):
        """ set a list of editors, take a tuple ('users' or 'groups', [<editors>]) """
        if type(value) is not tuple or value[0] not in ['users', 'groups']:
            raise InvalidArgumentValue
        self._protected_properties.editors[value[0]] = value[1]
        self.update_protected_range(fields='editors')

    @property
    def requesting_user_can_edit(self):
        return self._protected_properties.requestingUserCanEdit

    @requesting_user_can_edit.setter
    def requesting_user_can_edit(self, value):
        self._protected_properties.requestingUserCanEdit = value
        self.update_protected_range(fields='requestingUserCanEdit')

    @property
    def start_addr(self):
        """top-left address of the range"""
        return self._start_addr

    @start_addr.setter
    def start_addr(self, addr):
        self._start_addr = format_addr(addr, 'tuple')
        self.update_named_range()

    @property
    def end_addr(self):
        """bottom-right address of the range"""
        return self._end_addr

    @end_addr.setter
    def end_addr(self, addr):
        self._end_addr = format_addr(addr, 'tuple')
        self.update_named_range()

    @property
    def range(self):
        """Range in format A1:C5"""
        return format_addr(self._start_addr) + ':' + format_addr(self._end_addr)

    @property
    def worksheet(self):
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

        .. warn::
                Currently only data is fetched not properties, so `only_data` wont work

        :param only_data: fetch only data

        """
        self._data = self._worksheet.get_values(self._start_addr, self._end_addr, returnas='cells',
                                                include_tailing_empty_rows=True, include_tailing_empty=True)
        if not only_data:
            logging.error("functionality not implimented")

    def apply_format(self, cell):
        """
        Change format of all cells in the range

        :param cell: a model :class: Cell whose format will be applied to all cells

        """
        request = {"repeatCell": {
            "range": self._get_gridrange(),
            "cell": cell.get_json(),
            "fields": "userEnteredFormat,hyperlink,note,textFormatRuns,dataValidation,pivotTable"
            }
        }
        self._worksheet.client.sheet.batch_update(self._worksheet.spreadsheet.id, request)

    def update_values(self, values=None):
        """
        Update the values of the cells in this range

        :param values: values as matrix

        """
        if values and self._linked:
            self._worksheet.update_values(crange=self.range, values=values)
            self.fetch()
        if self._linked and not values:
            self._worksheet.update_values(cell_list=self._data)

    def sort(self, basecolumnindex=0, sortorder="ASCENDING"):
        """sort the datarange

        :param basecolumnindex:     Index of the base column in which sorting is to be done (Integer).
                                    The index here is the index of the column in range (first columen is 0).
        :param sortorder:           either "ASCENDING" or "DESCENDING" (String)
        """
        self._worksheet.sort_range(self._start_addr, self._end_addr, basecolumnindex=basecolumnindex + format_addr(self._start_addr, 'tuple')[1]-1, sortorder=sortorder)

    def update_named_range(self):
        """update the named properties"""
        if not self._name_id or not self._linked:
            return False
        if self.protected:
            self.update_protected_range()
        request = {'updateNamedRange':{
          "namedRange": {
              "namedRangeId": self._name_id,
              "name": self._name,
              "range": self._get_gridrange(),
          },
          "fields": '*',
        }}
        self._worksheet.client.sheet.batch_update(self._worksheet.spreadsheet.id, request)

    def update_protected_range(self, fields='*'):
        if not self.protected or not self._linked:
            return False

        request = {'updateProtectedRange': {
          "protectedRange": self._protected_properties.to_json(),
          "fields": fields,
        }}
        request['updateProtectedRange']['protectedRange']['range'] = self._get_gridrange()
        self._worksheet.client.sheet.batch_update(self._worksheet.spreadsheet.id, request)

    def _get_gridrange(self):
        return {
            "sheetId": self._worksheet.id,
            "startRowIndex": self._start_addr[0]-1,
            "endRowIndex": self._end_addr[0],
            "startColumnIndex": self._start_addr[1]-1,
            "endColumnIndex": self._end_addr[1],
        }

    def __getitem__(self, item):
        if len(self._data[0]) == 0:
            self.fetch()
        if type(item) == int:
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
            range_str = str(self.range)
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


