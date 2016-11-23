# -*- coding: utf-8 -*-.

"""
pygsheets.models
~~~~~~~~~~~~~~~~

This module contains common spreadsheets' models

"""

import re
import warnings

from .exceptions import IncorrectCellLabel, WorksheetNotFound, CellNotFound, InvalidArgumentValue, InvalidUser
from .utils import finditem, numericise_all
from custom_types import *


class Spreadsheet(object):

    """ A class for a spreadsheet object."""

    def __init__(self, client, jsonsheet=None, id=None):
        """ spreadsheet init.

        :param client: the client object which links to this ssheet
        :param jsonsheet: the json sheet which has properties of this ssheet
        :param id: id of the spreadsheet
        """
        if type(jsonsheet) != dict and type(jsonsheet):
            raise InvalidArgumentValue
        self.client = client
        self._sheet_list = []
        self._jsonsheet = jsonsheet
        self._id = id
        self._update_properties(jsonsheet)
        self._permissions = dict()
        self.batch_mode = False

    def __repr__(self):
        return '<%s %s Sheets:%s>' % (self.__class__.__name__,
                                      repr(self.title), len(self._sheet_list))

    @property
    def id(self):
        """ id of the spreadsheet """
        return self._id

    @property
    def title(self):
        """ title of the spreadsheet """
        return self._title

    @property
    def defaultformat(self):
        """ deafault cell format"""
        return self._defaultFormat

    @property
    def sheet1(self):
        """Shortcut property for getting the first worksheet."""
        return self.worksheet()

    def _update_properties(self, jsonsheet=None):
        """ Update all sheet properies.

        :param jsonsheet: json var to update values form \
                if not specified, will fetch it and update

        """
        if not jsonsheet and len(self.id) > 1:
            self._jsonsheet = self.client.open_by_key(self.id, 'json')
        elif not jsonsheet and len(self.id) == 0:
            raise InvalidArgumentValue
        # print self._jsonsheet
        self._id = self._jsonsheet['spreadsheetId']
        self._fetch_sheets(self._jsonsheet)
        self._title = self._jsonsheet['properties']['title']
        self._defaultFormat = self._jsonsheet['properties']['defaultFormat']
        self.client.spreadsheetId = self._id

    def _fetch_sheets(self, jsonsheet=None):
        """update sheets list"""
        if not jsonsheet:
            jsonsheet = self.client.open_by_key(self.id, returnas='json')
        for sheet in jsonsheet.get('sheets'):
            self._sheet_list.append(Worksheet(self, sheet))

    def worksheets(self, sheet_property=None, value=None):
        """
        Get all worksheets filtered by a property.

        :param sheet_property: proptery to filter - 'title', 'index', 'id'
        :param value: value of property to match

        :returns: list of all :class:`worksheets <Worksheet>`
        """
        if not sheet_property and not value:
            return self._sheet_list

        if sheet_property not in ['title', 'index', 'id']:
            raise InvalidArgumentValue
        elif sheet_property in ['index', 'id']:
            value = int(value)

        sheets = [x for x in self._sheet_list if getattr(x, sheet_property) == value]
        if not len(sheets) > 0:
            self._fetch_sheets()
            sheets = [x for x in self._sheet_list if getattr(x, sheet_property) == value]
            if not len(sheets) > 0:
                raise WorksheetNotFound()
        return sheets

    def worksheet(self, property='id', value=0):
        """Returns a worksheet with specified property.

        :param property: A property of a worksheet. If there're multiple worksheets \
                        with the same title, first one will be returned.
        :param value: value of given property

        :type property: 'title','index','id'

        :returns: instance of :class:`Worksheet`

        Example. Getting worksheet named 'Annual bonuses'

        >>> sht = client.open('Sample one')
        >>> worksheet = sht.worksheet('title','Annual bonuses')

        """
        return self.worksheets(property, value)[0]

    def worksheet_by_title(self, title):
        """
        returns worksheet by title

        :param title: title of the sheet

        :returns: Spresheet instance
        """
        return self.worksheet('title', title)

    def add_worksheet(self, title, rows, cols):
        """Adds a new worksheet to a spreadsheet.

        :param title: A title of a new worksheet.
        :param rows: Number of rows.
        :param cols: Number of columns.

        :returns: a newly created :class:`worksheets <Worksheet>`.
        """
        request = {"addSheet": {"properties": {'title': title, "gridProperties": {"rowCount": rows, "columnCount": cols}}}}
        result = self.client.sh_batch_update(self.id, request, 'replies/addSheet', False)
        jsheet = dict()
        jsheet['properties'] = result['replies'][0]['addSheet']['properties']
        wks = Worksheet(self, jsheet)
        self._sheet_list.append(wks)
        return wks

    def del_worksheet(self, worksheet):
        """Deletes a worksheet from a spreadsheet.

        :param worksheet: The :class:`worksheets <Worksheet>` to be deleted.

        """
        if worksheet not in self.worksheets():
            raise WorksheetNotFound
        request = {"deleteSheet": {'sheetId': worksheet.id}}
        self.client.sh_batch_update(self.id, request, '', False)
        self._sheet_list.remove(worksheet)

    def share(self, addr, role='reader', expirationTime=None, is_group=False):
        """
        create/update permission for user/group/domain

        :param addr: this is the email for user/group and domain adress for domains
        :param role: permission to be applied ('owner','writer','commenter','reader')
        :param expirationTime: (Not Implimented) time until this permission should last (datetime)
        :param is_group: boolean , Is this a use/group used only when email provided

        """
        return self.client.add_permission(self.id, addr, role=role, is_group=False)

    def list_permissions(self):
        """
        list all the permissions of the spreadsheet

        :returns: list of permissions as json object

        """
        permissions = self.client.list_permissions(self.id)
        self._permissions = permissions['permissions']
        return self._permissions

    def remove_permissions(self, addr):
        """
        Removes all permissions of the user provided

        :param addr: email/domain of the user

        """
        try:
            result = self.client.remove_permissions(self.id, addr, self._permissions)
        except InvalidUser:
            result = self.client.remove_permissions(self.id, addr)
        return result

    def batch_start(self):
        """
        Start batch mode, where all updates to sheet values will be batched

        """
        self.batch_mode = True
        warnings.warn('Batching is only for Update operations')

    def batch_stop(self, discard=False):
        """
        Stop batch Mode

        :param discard: discard all changes done in batch mode
        """
        self.batch_mode = False
        if not discard:
            self.client.send_batch(self.id)

    # @TODO
    def link(self, syncToColoud=False):
        """ Link the spread sheet with colud, so all local changes \
            will be updated instantly, so does all data fetches

            :param  syncToColoud: update the cloud with local changes if set to true
                          update the local copy with cloud if set to false
        """
        # just link all child sheets
        pass

    # @TODO
    def unlink(self):
        """ Unlink the spread sheet with colud, so all local changes
            will be made on local copy fetched
        """
        # just unlink all sheets
        pass

    def __iter__(self):
        for sheet in self.worksheets():
            yield(sheet)

    def __getitem__(self, item):
        if type(item) == int:
            return self._sheet_list[item]


class Worksheet(object):

    """A class for worksheet object."""

    def __init__(self, spreadsheet, jsonSheet):
        self.spreadsheet = spreadsheet
        self.client = spreadsheet.client
        self._linked = True
        self.jsonSheet = jsonSheet
        self.data_grid = ''  # for storing sheet data while unlinked

    def __repr__(self):
        return '<%s %s index:%s>' % (self.__class__.__name__,
                                  repr(self.title),
                                  self.index)

    @property
    def id(self):
        """Id of a worksheet."""
        return self.jsonSheet['properties']['sheetId']

    @property
    def index(self):
        """Index of worksheet"""
        return self.jsonSheet['properties']['index']

    @property
    def title(self):
        """Title of a worksheet."""
        return self.jsonSheet['properties']['title']

    @title.setter
    def title(self, title):
        self.jsonSheet['properties']['title'] = title
        if self._linked:
            self.client.update_sheet_properties(self.jsonSheet['properties'], 'title')

    @property
    def rows(self):
        """Number of rows"""
        return int(self.jsonSheet['properties']['gridProperties']['rowCount'])

    @rows.setter
    def rows(self, row_count):
        self.jsonSheet['properties']['gridProperties']['rowCount'] = int(row_count)
        if self._linked:
            self.client.update_sheet_properties(self.spreadsheet.id, self.jsonSheet['properties'],
                                                'gridProperties/rowCount')

    @property
    def cols(self):
        """Number of columns"""
        return int(self.jsonSheet['properties']['gridProperties']['columnCount'])

    @cols.setter
    def cols(self, col_count):
        self.jsonSheet['properties']['gridProperties']['columnCount'] = int(col_count)
        if self._linked:
            self.client.update_sheet_properties(self.spreadsheet.id, self.jsonSheet['properties'],
                                                'gridProperties/columnCount')

    # @TODO
    @property
    def updated(self):
        """Updated time in RFC 3339 format(use drive api)"""
        return None

    # @TODO
    def link(self, syncToColoud=True):
        """ Link the spread sheet with colud, so all local changes
            will be updated instantly, so does all data fetches

            :param  syncToColoud: update the cloud with local changes if set to true
                          update the local copy with cloud if set to false
        """
        # warnings.warn("Complete functionality not implimented")
        if syncToColoud:
            self.client.update_sheet_properties(self.spreadsheet.id, self.jsonSheet['properties'])
        else:
            wks = self.spreadsheet.worksheet(self, property='id', value=self.id)
            self.jsonSheet = wks.jsonSheet
        self._linked = True

    # @TODO
    def unlink(self):
        """ Unlink the spread sheet with colud, so all local changes
            will be made on local copy fetched
        """
        warnings.warn("Complete functionality not implimented")
        self._linked = False

    @staticmethod
    def get_addr(addr, output='flip'):
        """
        function to convert adress format of cells from one to another

        :param addr: adress as tuple or label
        :param output: -'label' will output label
                      - 'tuple' will output tuple
                      - 'flip' will convert to other type
        :returns: tuple or label
        """
        _MAGIC_NUMBER = 64
        if type(addr) == tuple:
            if output == 'label' or output == 'flip':
                # return self.get_addr_int(*addr)
                row = int(addr[0])
                col = int(addr[1])
                if row < 1 or col < 1:
                    raise IncorrectCellLabel('(%s, %s)' % (row, col))
                div = col
                column_label = ''
                while div:
                    (div, mod) = divmod(div, 26)
                    if mod == 0:
                        mod = 26
                        div -= 1
                    column_label = chr(mod + _MAGIC_NUMBER) + column_label
                label = '%s%s' % (column_label, row)
                return label

            elif output == 'tuple':
                return addr

        elif type(addr) == str:
            if output == 'tuple' or output == 'flip':
                # return self.get_int_addr(addr)
                _cell_addr_re = re.compile(r'([A-Za-z]+)(\d+)')
                m = _cell_addr_re.match(addr)
                if m:
                    column_label = m.group(1).upper()
                    row, col = int(m.group(2)), 0
                    for i, c in enumerate(reversed(column_label)):
                        col += (ord(c) - _MAGIC_NUMBER) * (26 ** i)
                else:
                    raise IncorrectCellLabel(addr)
                return row, col
            elif output == 'label':
                return addr
        else:
            raise InvalidArgumentValue

    def _get_range(self, start_label, end_label):
        """get range in A1 notation, given start and end labels

        """
        return self.title + '!' + ('%s:%s' % (start_label, end_label))

    def cell(self, addr):
        """
        Returns  cell object at given address.

        :param addr: cell adress as either tuple (row, col) or cell label 'A1'

        :returns: an instance of a :class:`Cell`

        Example:

        >>> wks.cell((1,1))
        <Cell R1C1 "I'm cell A1">
        >>> wks.cell('A1')
        <Cell R1C1 "I'm cell A1">

        """
        try:
            if type(addr) is str:
                val = self.client.get_range(self.spreadsheet.id, self._get_range(addr, addr), 'ROWS')[0][0]
            elif type(addr) is tuple:
                label = Worksheet.get_addr(addr, 'label')
                val = self.client.get_range(self.spreadsheet.id, self._get_range(label, label), 'ROWS')[0][0]
            else:
                raise CellNotFound
        except Exception as e:
            if str(e).find('exceeds grid limits') != -1:
                raise CellNotFound
            else:
                raise e

        return Cell(addr, val, self)

    def range(self, crange):
        """Returns a list of :class:`Cell` objects from specified range.

        :param crange: A string with range value in common format,
                         e.g. 'A1:A5'.
        """
        startcell = crange.split(':')[0]
        endcell = crange.split(':')[1]
        return self.values(startcell, endcell, returnas='cell')

    def values(self, start, end, returnas='matrix', majdim='ROWS'):
        """Returns value of cells given the topleft corner position
        and bottom right position

        :param start: topleft position as tuple or label
        :param end: bottomright position as tuple or label
        :param majdim: output as rowwise or columwise
                       takes - 'ROWS' or 'COLMUNS'
        :param returnas: return as list of strings of cell objects
                         takes - 'matrix' or 'cell'

        Example:

        >>> wks.values((1,1),(3,3))
        [[u'another look.', u'', u'est'],
         [u'EE 4212', u"it's down there "],
         [u'ee 4210', u'somewhere, let me take ']]

        """
        start_label = Worksheet.get_addr(start, 'label')
        end_label = Worksheet.get_addr(end, 'label')
        values = self.client.get_range(self.spreadsheet.id, self._get_range(start_label, end_label), majdim.upper())

        if returnas.lower() == 'matrix':
            return values
        elif returnas.lower() == 'cell' or returnas.lower() == 'cells':
            cells = []
            for k in range(0, len(values)):
                row = []
                for i in range(0, len(values[k])):
                    row.append(Cell((k+1, i+1), values[k][i], self))
                cells.append(row)
            return cells
        else:
            return None

    def all_values(self, returnas='matrix', majdim='ROWS'):
        """Returns a list of lists containing all cells' values as strings.

        :param majdim: output as row wise or columwise
        :param returnas: return as list of strings of cell objects

        :type majdim: 'ROWS', 'COLUMNS'
        :type returnas: 'matrix','cell'

        Example:

        >>> wks.all_values()
        [[u'another look.', u'', u'est'],
         [u'EE 4212', u"it's down there "],
         [u'ee 4210', u'somewhere, let me take ']]
        """
        return self.values((1, 1), (self.rows, self.cols), returnas=returnas, majdim=majdim)

    # @TODO improve empty2zero for other types also and clustring
    def get_all_records(self, empty2zero=False, head=1):
        """
        Returns a list of dictionaries, all of them having:
            - the contents of the spreadsheet's with the head row as keys, \
            And each of these dictionaries holding
            - the contents of subsequent rows of cells as values.

        Cell values are numericised (strings that can be read as ints
        or floats are converted).

        :param empty2zero: determines whether empty cells are converted to zeros.
        :param head: determines wich row to use as keys, starting from 1
            following the numeration of the spreadsheet.

        :returns: a dict dict with header column values as head and rows as list
        """
        idx = head - 1
        data = self.all_values()
        keys = data[idx]
        values = [numericise_all(row, empty2zero) for row in data[idx + 1:]]
        return [dict(zip(keys, row)) for row in values]

    def row(self, row, returnas='matrix'):
        """Returns a list of all values in a `row`.

        Empty cells in this list will be rendered as :const:` `.

        :param row: index of row
        :param returnas: ('matrix' or 'cell') return as cell objects or just 2d array

        """
        return self.values((row, 1), (row, self.cols), returnas=returnas)[0]

    def col(self, col, returnas='matrix'):
        """Returns a list of all values in column `col`.

        Empty cells in this list will be rendered as :const:` `.

        :param col: index of col
        :param returnas: ('matrix' or 'cell') return as cell objects or just values

        """
        return self.values((1, col), (self.rows, col), majdim='COLUMNS', returnas=returnas)[0]

    def update_cell(self, addr, val, parse=True):
        """Sets the new value to a cell.

        :param addr: cell adress as tuple (row,column) or label 'A1'.
        :param val: New value
        :param parse: if the values should be stored \
                        as is or should be as if the user typed them into the UI

        Example:

        >>> wks.update_cell('A1', '42') # this could be 'a1' as well
        <Cell R1C1 "42">
        >>> wks.update_cell('A3', '=A1+A2', True)
        <Cell R1C3 "57">
        """
        label = Worksheet.get_addr(addr, 'label')
        body = dict()
        body['range'] = self._get_range(label, label)
        body['majorDimension'] = 'ROWS'
        body['values'] = [[val]]
        self.client.sh_update_range(self.spreadsheet.id, body, self.spreadsheet.batch_mode, parse)

    def update_cells(self, cell_list=None, range=None, values=None, majordim='ROWS'):
        """Updates cells in batch, it can take either a cell list or a range and values

        :param cell_list: List of a :class:`Cell` objects to update with their values
        :param range: range in format A1:A2
        :param values: list of values if range given
        :param majordim: major dimension of given data

        """
        if cell_list:
            if not self.spreadsheet.batch_mode:
                self.spreadsheet.batch_start()
            for cell in cell_list:
                self.update_cell(cell.label, cell.value)
            self.spreadsheet.batch_stop()  # @TODO fix this
        elif range and values:
            body = dict()
            if range.find(':') == -1:
                start_r_tuple = Worksheet.get_addr(range, output='tuple')
                print (start_r_tuple)
                if majordim == 'ROWS':
                    end_r_tuple = (start_r_tuple[0]+len(values), start_r_tuple[1]+len(values[0]))
                    print (end_r_tuple)
                else:
                    end_r_tuple = (start_r_tuple[0] + len(values[0]), start_r_tuple[1] + len(values))
                body['range'] = self._get_range(range, Worksheet.get_addr(end_r_tuple))
                print(body['range'])
            else:
                body['range'] = self._get_range(*range.split(':'))
            body['majorDimension'] = majordim
            body['values'] = values
            self.client.sh_update_range(self.spreadsheet.id, body, self.spreadsheet.batch_mode)
        else:
            raise InvalidArgumentValue(cell_list)  # @TODO test

    def update_col(self, index, values):
        """update an existing colum with values

        """
        colrange = Worksheet.get_addr((1, index), 'label') + ":" + Worksheet.get_addr((len(values), index), "label")
        self.update_cells(range=colrange, values=[values], majordim='COLUMNS')

    def update_row(self, index, values):
        """update an existing row with values

        """
        colrange = self.get_addr((index, 1), 'label') + ':' + self.get_addr((index, len(values)), 'label')
        self.update_cells(range=colrange, values=[values], majordim='ROWS')

    def resize(self, rows=None, cols=None):
        """Resizes the worksheet.

        :param rows: New rows number.
        :param cols: New columns number.
        """
        self.unlink()
        self.rows = rows
        self.cols = cols
        self.link(True)

    def add_rows(self, rows):
        """Adds rows to worksheet.

        :param rows: Rows number to add.
        """
        self.resize(rows=self.rows + rows, cols=self.cols)

    def add_cols(self, cols):
        """Adds colums to worksheet.

        :param cols: Columns number to add.
        """
        self.resize(cols=self.cols + cols, rows=self.rows)

    # @TODO
    def delete_cols(self, cols):
        warnings.warn("Method not Implimented")

    # @TODO
    def delete_rows(self, rows):
        warnings.warn("Method not Implimented")

    def insert_cols(self, col, number=1, values=None):
        """
        Insert a colum after the colum <col> and fill with values <values>

        :param col: columer after which new colum should be inserted
        :param number: number of colums to be inserted
        :param values: values to filled in new colum

        """
        request = {'insertDimension': {'inheritFromBefore': False,
                                       'range': {'sheetId': self.id, 'dimension': 'COLUMNS',
                                                 'endIndex': (col+number), 'startIndex': col}
                                       }}
        self.client.sh_batch_update(self.spreadsheet.id, request, batch=self.spreadsheet.batch_mode)
        # self.client.insertdim(self.id, 'COLUMNS', col, (col+number), False)
        self.jsonSheet['properties']['gridProperties']['columnCount'] = self.cols+number
        if values:
            self.update_col(col+1, values)

    def insert_rows(self, row, number=1, values=None):
        """
        Insert a row after the row <row> and fill with values <values>

        :param row: row after which new colum should be inserted
        :param number: number of rows to be inserted
        :param values: values to be filled in new row

        """
        request = {'insertDimension': {'inheritFromBefore': False,
                                       'range': {'sheetId': self.id, 'dimension': 'ROWS',
                                                 'endIndex': (row+number), 'startIndex': row}}}
        self.client.sh_batch_update(self.spreadsheet.id, request, batch=self.spreadsheet.batch_mode)

        # self.client.insertdim(self.id, 'ROWS', row, (row+number), False)
        self.jsonSheet['properties']['gridProperties']['rowCount'] = self.rows + number
        # @TODO fore multiple rows inserted change
        if values:
            self.update_row(row+1, values)

    # @TODO
    def clear(self):
        """clear thw worksheet"""
        Warning("Not yet implimented")

    # @TODO
    def append_row(self, values):
        """Adds a row to the worksheet and populates it with values.
        Widens the worksheet if there are more values than columns.

        :param values: List of values for the new row.
        """
        warnings.warn("Method not Implimented")

    # @TODO
    def _finder(self, func, query):
        warnings.warn("Method not Implimented")

    # @TODO
    def find(self, query):
        """Finds first cell matching query.

        :param query: A text string or compiled regular expression.
        """
        pass

    # @TODO
    def findall(self, query):
        """Finds all cells matching query.

        :param query: A text string or compiled regular expression.
        """
        warnings.warn("Method not Implimented")

    # @TODO
    def export(self, format='csv'):
        """Export the worksheet in specified format.

        :param format: A format of the output.
        """
        warnings.warn("Method not Implimented")

    def __iter__(self):
        rows = self.all_values(majdim='ROWS')
        for row in rows:
            yield(row + (self.cols - len(row))*[''])

    def __getitem__(self, item):
        if type(item) == int:
            row = self.all_values()[item]
            return row + (self.cols - len(row))*['']


class Cell(object):

    """An instance of this class represents a single cell
    in a :class:`worksheet <Worksheet>`.

    """

    def __init__(self, pos, val='', worksheet=None):
        self.worksheet = worksheet
        if type(pos) == str:
            pos = Worksheet.get_addr(pos, 'tuple')
        self._row, self._col = pos
        self._label = Worksheet.get_addr(pos, 'label')
        self._value = val  # formated vlaue
        self._unformated_value = val    # @TODO
        self._formula = ''  # @TODO use the render_as
        self.format = FormatType.CUSTOM  # @TODO
        self.render_as = ValueRenderOption.FORMATTED  # @TODO use this
        self.parse_value = True
        self._comment = ''  # @TODO
    
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
        """formated value of the cell"""
        return self._value

    @value.setter
    def value(self, value):
        if self.worksheet:
            self.worksheet.update_cell(self.label, value, self.parse_value)
            self._value = value
        else:
            self._value = value

    @property
    def formula(self):
        """formula if any of the cell"""
        # fetch formula
        return self._formula

    @formula.setter
    def formula(self, formula):
        tmp, self.parse_value = self.parse_value, False
        self.value = formula
        self.parse_value = True

    def fetch(self):
        """ Update the value of the cell from sheet """
        if self.worksheet:
            self._value = self.worksheet.cell(self._label)

    def __repr__(self):
        return '<%s R%sC%s %s>' % (self.__class__.__name__,
                                   self.row,
                                   self.col,
                                   repr(self.value))

