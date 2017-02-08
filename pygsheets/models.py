# -*- coding: utf-8 -*-.

"""
pygsheets.models
~~~~~~~~~~~~~~~~

This module contains common spreadsheets' models

"""

import re
import warnings
import datetime

from .exceptions import (IncorrectCellLabel, WorksheetNotFound, RequestError,
                         CellNotFound, InvalidArgumentValue, InvalidUser)
from .utils import finditem, numericise_all
from .custom_types import *
try:
    from pandas import DataFrame
except ImportError:
    DataFrame = None


class Spreadsheet(object):

    """ A class for a spreadsheet object."""

    def __init__(self, client, jsonsheet=None, id=None):
        """ spreadsheet init.

        :param client: the client object which links to this ssheet
        :param jsonsheet: the json sheet which has properties of this ssheet
        :param id: id of the spreadsheet
        """
        if type(jsonsheet) != dict and jsonsheet is not None:
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
        self._sheet_list = []
        if not jsonsheet:
            jsonsheet = self.client.open_by_key(self.id, returnas='json')
        for sheet in jsonsheet.get('sheets'):
            self._sheet_list.append(Worksheet(self, sheet))

    def worksheets(self, sheet_property=None, value=None, force_fetch=False):
        """
        Get all worksheets filtered by a property.

        :param sheet_property: proptery to filter - 'title', 'index', 'id'
        :param value: value of property to match
        :param force_fetch: update the sheets, from cloud

        :returns: list of all :class:`worksheets <Worksheet>`
        """
        if not sheet_property and not value:
            return self._sheet_list

        if sheet_property not in ['title', 'index', 'id']:
            raise InvalidArgumentValue
        elif sheet_property in ['index', 'id']:
            value = int(value)

        sheets = [x for x in self._sheet_list if getattr(x, sheet_property) == value]
        if not len(sheets) > 0 or force_fetch:
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

    def add_worksheet(self, title, rows=100, cols=26, src_tuple=None, src_worksheet=None):
        """Adds a new worksheet to a spreadsheet.

        :param title: A title of a new worksheet.
        :param rows: Number of rows.
        :param cols: Number of columns.
        :param src_tuple: a tuple (spreadsheet id, worksheet id) specifying a worksheet to copy
        :param src_worksheet: source worksheet to copy values from

        :returns: a newly created :class:`worksheets <Worksheet>`.
        """
        if self.batch_mode:
            raise RequestError("not supported in batch Mode")

        jsheet = dict()
        if src_tuple:
            jsheet['properties'] = self.client.sh_copy_worksheet(src_tuple[0], src_tuple[1], self.id)
            wks = Worksheet(self, jsheet)
            wks.title = title
        elif src_worksheet:
            if type(src_worksheet) != Worksheet:
                raise InvalidArgumentValue("src_worksheet")
            jsheet['properties'] = self.client.sh_copy_worksheet(src_worksheet.spreadsheet.id, src_worksheet.id, self.id)
            wks = Worksheet(self, jsheet)
            wks.title = title
        else:
            request = {"addSheet": {"properties": {'title': title, "gridProperties": {"rowCount": rows, "columnCount": cols}}}}
            result = self.client.sh_batch_update(self.id, request, 'replies/addSheet', False)
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

    def find(self, string, replace=None, regex=True, match_case=False, include_formulas=False,
             srange=None, sheet=True):
        """
        Find and replace cells in spreadsheet

        :param string: string to search for
        :param replace: string to replace with
        :param regex: is the search string regex
        :param match_case: match case in search
        :param include_formulas: include seach in formula
        :param srange: range to search in A1 format
        :param sheet: if True - search all sheets, else search specified sheet

        """
        if not replace:
            found_list = []
            for wks in self.worksheets():
                found_list.extend(wks.find(string))
            return found_list
        body = {
            "find": string,
            "replacement": replace,
            "matchCase": match_case,
            "matchEntireCell": False,
            "searchByRegex": regex,
            "includeFormulas": include_formulas,
        }
        if srange:
            body['range'] = srange
        elif type(sheet) == bool:
            body['allSheets'] = True
        elif type(sheet) == int:
            body['sheetId'] = sheet
        body = {'findReplace': body}
        response = self.client.sh_batch_update(self.id, request=body, batch=self.batch_mode)
        return response['replies'][0]['findReplace']
    
    # @TODO impliment expiration time
    def share(self, addr, role='reader', expirationTime=None, is_group=False):
        """
        create/update permission for user/group/domain

        :param addr: this is the email for user/group and domain address for domains
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
        Stop batch Mode and Update the changes

        :param discard: discard all changes done in batch mode
        """
        self.batch_mode = False
        if not discard:
            self.client.send_batch(self.id)

    # @TODO
    def link(self, syncToColoud=False):
        """ Link the spreadsheet with colud, so all local changes \
            will be updated instantly, so does all data fetches

            :param  syncToColoud: true ->  update the cloud with local changes
                                  false -> update the local copy with cloud
        """
        # just link all child sheets
        warnings.warn("method not implimented")

    # @TODO
    def unlink(self):
        """ Unlink the spread sheet with colud, so all local changes
            will be made on local copy fetched
        """
        # just unlink all sheets
        warnings.warn("method not implimented")

    def export(self, fformat=ExportType.CSV):
        """Export all the worksheet of the worksheet in specified format.

        :param fformat: A format of the output as Enum ExportType
        """
        if fformat is ExportType.CSV:
            for wks in self._sheet_list:
                wks.export(ExportType.CSV)
        elif isinstance(fformat, ExportType):
            self._sheet_list[0].export(fformat=fformat)

    @property
    def updated(self):
        """Last time the spreadsheet was modified, in RFC 3339 format"""
        request = self.client.driveService.files().get(fileId=self.id, fields='modifiedTime')
        response = self.client._execute_request(self.id, request, False)
        return response['modifiedTime']

    def __iter__(self):
        for sheet in self.worksheets():
            yield(sheet)

    def __getitem__(self, item):
        if type(item) == int:
            return self.worksheet('index', item)


class Worksheet(object):

    """A class for worksheet object."""

    def __init__(self, spreadsheet, jsonSheet):
        self.spreadsheet = spreadsheet
        self.client = spreadsheet.client
        self._linked = True
        self.jsonSheet = jsonSheet
        self.data_grid = None  # for storing sheet data while unlinked
        self.grid_update_time = None

    def __repr__(self):
        return '<%s %s index:%s>' % (self.__class__.__name__,
                                     repr(self.title), self.index)

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
            self.client.update_sheet_properties(self.spreadsheet.id, self.jsonSheet['properties'], 'title')

    @property
    def rows(self):
        """Number of rows"""
        return int(self.jsonSheet['properties']['gridProperties']['rowCount'])

    @rows.setter
    def rows(self, row_count):
        if row_count == self.rows:
            return
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
        if col_count == self.cols:
            return
        self.jsonSheet['properties']['gridProperties']['columnCount'] = int(col_count)
        if self._linked:
            self.client.update_sheet_properties(self.spreadsheet.id, self.jsonSheet['properties'],
                                                'gridProperties/columnCount')

    def _update_grid(self, force=False):
        """
        update the data grid with values from sheeet
        :param force: force update data grid

        """
        if not self.data_grid or force:
            self.data_grid = self.all_values(returnas='cells', include_empty=False)
        elif not force:  # @TODO the update is not instantaious
            updated = datetime.datetime.strptime(self.spreadsheet.updated, '%Y-%m-%dT%H:%M:%S.%fZ')
            if updated > self.grid_update_time:
                self.data_grid = self.all_values(returnas='cells', include_empty=False)
        self.grid_update_time = datetime.datetime.utcnow()

    # @TODO update values too (currently only sync worksheet properties)
    def link(self, syncToColoud=True):
        """ Link the spread sheet with colud, so all local changes
            will be updated instantly, so does all data fetches

            :param  syncToColoud: update the cloud with local changes if set to true
                          update the local copy with cloud if set to false
        """
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

    # @TODO Rename
    @staticmethod
    def get_addr(addr, output='flip'):
        """
        function to convert address format of cells from one to another

        :param addr: address as tuple or label
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
                _cell_addr_re = re.compile(r'([A-Za-z]+)(\d+)')
                m = _cell_addr_re.match(addr)
                if m:
                    column_label = m.group(1).upper()
                    row, col = int(m.group(2)), 0
                    for i, c in enumerate(reversed(column_label)):
                        col += (ord(c) - _MAGIC_NUMBER) * (26 ** i)
                else:
                    raise IncorrectCellLabel(addr)
                return int(row), int(col)
            elif output == 'label':
                return addr
        else:
            raise InvalidArgumentValue

    def _get_range(self, start_label, end_label):
        """get range in A1 notation, given start and end labels

        """
        return self.title + '!' + ('%s:%s' % (Worksheet.get_addr(start_label, 'label'),
                                              Worksheet.get_addr(end_label, 'label')))

    def cell(self, addr):
        """
        Returns cell object at given address.

        :param addr: cell address as either tuple (row, col) or cell label 'A1'

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
                raise

        return Cell(addr, val, self)

    def range(self, crange):
        """Returns a list of :class:`Cell` objects from specified range.

        :param crange: A string with range value in common format,
                         e.g. 'A1:A5'.
        """
        startcell = crange.split(':')[0]
        endcell = crange.split(':')[1]
        return self.values(startcell, endcell, returnas='cell')

    def value(self, addr):
        """
        value of a cell at given address

        :param addr: cell address as either tuple or label

        """
        addr = self.get_addr(addr, 'tuple')
        try:
            return self.values(addr, addr, include_empty=False)[0][0]
        except KeyError:
            raise CellNotFound

    def values(self, start, end, returnas='matrix', majdim='ROWS', include_empty=True):
        """Returns value of cells given the topleft corner position
        and bottom right position

        :param start: topleft position as tuple or label
        :param end: bottomright position as tuple or label
        :param majdim: output as rowwise or columwise
                       takes - 'ROWS' or 'COLMUNS'
        :param returnas: return as list of strings of cell objects
                         takes - 'matrix' or 'cell'
        :param include_empty: include empty trailing cells/values until last non-zero value

        Example:

        >>> wks.values((1,1),(3,3))
        [[u'another look.', u'', u'est'],
         [u'EE 4212', u"it's down there "],
         [u'ee 4210', u'somewhere, let me take ']]

        """
        values = self.client.get_range(self.spreadsheet.id, self._get_range(start, end), majdim.upper())
        start = self.get_addr(start, 'tuple')
        if not include_empty:
            matrix = values
        else:
            max_cols = len(max(values, key=len))
            matrix = [list(x + ['']*(max_cols-len(x))) for x in values]

        if returnas == 'matrix':
            return matrix
        else:
            cells = []
            for k in range(len(matrix)):
                row = []
                for i in range(len(matrix[k])):
                    if majdim == 'COLUMNS':
                        row.append(Cell((start[0]+i, start[1]+k), matrix[k][i], self))
                    elif majdim == 'ROWS':
                        row.append(Cell((start[0]+k, start[1]+i), matrix[k][i], self))
                    else:
                        raise InvalidArgumentValue('majdim')

                cells.append(row)
            return cells

    def all_values(self, returnas='matrix', majdim='ROWS', include_empty=True):
        """Returns a list of lists containing all cells' values as strings.

        :param majdim: output as row wise or columwise
        :param returnas: return as list of strings of cell objects
        :param include_empty: whether to include empty values
        :type returnas: 'matrix','cell'

        Example:

        >>> wks.all_values()
        [[u'another look.', u'', u'est'],
         [u'EE 4212', u"it's down there "],
         [u'ee 4210', u'somewhere, let me take ']]
        """
        return self.values((1, 1), (self.rows, self.cols), returnas=returnas,
                           majdim=majdim, include_empty=include_empty)

    # @TODO add clustring (use append?)
    def get_all_records(self, empty_value='', head=1):
        """
        Returns a list of dictionaries, all of them having:
            - the contents of the spreadsheet's with the head row as keys, \
            And each of these dictionaries holding
            - the contents of subsequent rows of cells as values.

        Cell values are numericised (strings that can be read as ints
        or floats are converted).

        :param empty_value: determines empty cell's value
        :param head: determines wich row to use as keys, starting from 1
            following the numeration of the spreadsheet.

        :returns: a list of dict with header column values as head and rows as list
        """
        idx = head - 1
        data = self.all_values(returnas='matrix', include_empty=False)
        keys = data[idx]
        values = [numericise_all(row, empty_value) for row in data[idx + 1:]]
        return [dict(zip(keys, row)) for row in values]

    def row(self, row, returnas='matrix', include_empty=True):
        """Returns a list of all values in a `row`.

        Empty cells in this list will be rendered as :const:` `.

        :param include_empty: whether to include empty values
        :param row: index of row
        :param returnas: ('matrix' or 'cell') return as cell objects or just 2d array

        """
        return self.values((row, 1), (row, self.cols),
                           returnas=returnas, include_empty=include_empty)[0]

    def col(self, col, returnas='matrix', include_empty=True):
        """Returns a list of all values in column `col`.

        Empty cells in this list will be rendered as :const:` `.

        :param include_empty: whether to include empty values
        :param col: index of col
        :param returnas: ('matrix' or 'cell') return as cell objects or just values

        """
        return self.values((1, col), (self.rows, col), majdim='COLUMNS',
                           returnas=returnas, include_empty=include_empty)[0]

    def update_cell(self, addr, val, parse=True):
        """Sets the new value to a cell.

        :param addr: cell address as tuple (row,column) or label 'A1'.
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

    def update_cells(self, crange=None, values=None, cell_list=None, majordim='ROWS'):
        """Updates cells in batch, it can take either a cell list or a range and values

        :param cell_list: List of a :class:`Cell` objects to update with their values
        :param crange: range in format A1:A2 or just 'A1' or even (1,2) end cell will be infered from values
        :param values: list of values if range given, if value is None its unchanged
        :param majordim: major dimension of given data

        """
        if cell_list:
            crange = 'A1:' + str(self.get_addr((self.rows, self.cols)))
            # @TODO fit the minimum rectangle than whole array
            values = [[None for x in range(self.cols)] for y in range(self.rows)]
            for cell in cell_list:
                values[cell.col-1][cell.row-1] = cell.value
        body = dict()
        estimate_size = False
        if type(crange) == str:
            if crange.find(':') == -1:
                estimate_size = True
        elif type(crange) == tuple:
            estimate_size = True
        else:
            raise InvalidArgumentValue

        if estimate_size:
            start_r_tuple = Worksheet.get_addr(crange, output='tuple')
            if majordim == 'ROWS':
                end_r_tuple = (start_r_tuple[0]+len(values), start_r_tuple[1]+len(values[0]))
            else:
                end_r_tuple = (start_r_tuple[0] + len(values[0]), start_r_tuple[1] + len(values))
            body['range'] = self._get_range(crange, Worksheet.get_addr(end_r_tuple))
        else:
            body['range'] = self._get_range(*crange.split(':'))
        body['majorDimension'] = majordim
        body['values'] = values
        self.client.sh_update_range(self.spreadsheet.id, body, self.spreadsheet.batch_mode)

    def update_col(self, index, values):
        """update an existing colum with values

        """
        colrange = Worksheet.get_addr((1, index), 'label') + ":" + Worksheet.get_addr((len(values), index), "label")
        self.update_cells(crange=colrange, values=[values], majordim='COLUMNS')

    def update_row(self, index, values):
        """update an existing row with values

        """
        colrange = self.get_addr((index, 1), 'label') + ':' + self.get_addr((index, len(values)), 'label')
        self.update_cells(crange=colrange, values=[values], majordim='ROWS')

    def resize(self, rows=None, cols=None):
        """Resizes the worksheet.

        :param rows: New rows number.
        :param cols: New columns number.
        """
        self.unlink()
        self.rows = rows
        self.cols = cols
        self.link()

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

    def delete_cols(self, index, number=1):
        """
        delete a number of colums stating from index

        :param index: indenx of first col to delete
        :param number: number of cols to delete

        """
        index = index - 1
        if number < 1:
            raise InvalidArgumentValue
        request = {'deleteDimension': {'range': {'sheetId': self.id, 'dimension': 'COLUMNS',
                                                 'endIndex': (index+number), 'startIndex': index}}}
        self.client.sh_batch_update(self.spreadsheet.id, request, batch=self.spreadsheet.batch_mode)
        self.jsonSheet['properties']['gridProperties']['columnCount'] = self.cols-number

    def delete_rows(self, index, number=1):
        """
        delete a number of rows stating from index

        :param index: index of first row to delete
        :param number: number of rows to delete

        """
        index = index-1
        if number < 1:
            raise InvalidArgumentValue
        request = {'deleteDimension': {'range': {'sheetId': self.id, 'dimension': 'ROWS',
                                                 'endIndex': (index+number), 'startIndex': index}}}
        self.client.sh_batch_update(self.spreadsheet.id, request, batch=self.spreadsheet.batch_mode)
        self.jsonSheet['properties']['gridProperties']['rowCount'] = self.rows-number

    # @TODO remove???
    def fill_cells(self, start, end, value=''):
        """
        fill a range of cells with given value

        :param start: start cell address as tuple or label
        :param end: end cell address
        :param empty_value: empty value to replace with
        """
        start = self.get_addr(start, "tuple")
        end = self.get_addr(end, "tuple")
        values = [[value]*(end[1]-start[1]+1)]*(end[0]-start[0]+1)
        self.update_cells(crange=start, values=values)

    def insert_cols(self, col, number=1, values=None):
        """
        Insert a colum after the colum <col> and fill with values <values>
        Widens the worksheet if there are more values than columns.

        :param col: columer after which new colum should be inserted
        :param number: number of colums to be inserted
        :param values: values matrix to filled in new colum

        """
        request = {'insertDimension': {'inheritFromBefore': False,
                                       'range': {'sheetId': self.id, 'dimension': 'COLUMNS',
                                                 'endIndex': (col+number), 'startIndex': col}
                                       }}
        self.client.sh_batch_update(self.spreadsheet.id, request, batch=self.spreadsheet.batch_mode)
        self.jsonSheet['properties']['gridProperties']['columnCount'] = self.cols+number
        if values and number == 1:
            if len(values) > self.rows:
                self.rows = len(values)
            self.update_col(col+1, values)

    def insert_rows(self, row, number=1, values=None):
        """
        Insert a row after the row <row> and fill with values <values>
        Widens the worksheet if there are more values than columns.

        :param row: row after which new colum should be inserted
        :param number: number of rows to be inserted
        :param values: values matrix to be filled in new row

        """
        request = {'insertDimension': {'inheritFromBefore': False,
                                       'range': {'sheetId': self.id, 'dimension': 'ROWS',
                                                 'endIndex': (row+number), 'startIndex': row}}}
        self.client.sh_batch_update(self.spreadsheet.id, request, batch=self.spreadsheet.batch_mode)
        self.jsonSheet['properties']['gridProperties']['rowCount'] = self.rows + number
        # @TODO fore multiple rows inserted change
        if values and number == 1:
            if len(values) > self.cols:
                self.cols = len(values)
            self.update_row(row+1, values)

    def clear(self, start='A1', end=None):
        """
        clears the worksheet by default, if range given then clears range

        :param start: topright cell address
        :param end: bottom left cell of range

        """
        if not end:
            end = (self.rows, self.cols)
        body = {'ranges': [self._get_range(start, end)]}
        self.client.sh_batch_clear(self.spreadsheet.id, body)

    def append_row(self, start='A1', end=None, values=None):
        """Search for a table in the given range and will
         append it with values

        :param start: start cell of range
        :param end: end cell of range
        :param values: List of values for the new row.

        """
        if type(values[0]) != list:
            values = [values]
        if not end:
            end = (self.rows, self.cols)
        body = {"values": values}
        self.client.sh_append(self.spreadsheet.id, body=body, rranage=self._get_range(start, end))

    def find(self, query, replace=None, force_fetch=True):
        """Finds first cell matching query.

        :param query: A text string or compiled regular expression.
        :param replace: string to replace
        :param force_fetch: if local datagrid should be updated before searching, even if file is not modified
        """
        self._update_grid(force_fetch)
        found_list = []
        if isinstance(query, type(re.compile(""))):
            match = lambda x: query.search(x.value)
        else:
            match = lambda x: x.value == query
        for row in self.data_grid:
            found_list.extend(filter(match, row))
        if replace:
            for cell in found_list:
                cell.value = replace
        return found_list

    def set_dataframe(self, df, start, copy_index=False, copy_head=True, fit=False, escape_formulae=False):
        """
        set the values of a pandas dataframe at cell <start>

        :param df: pandas dataframe
        :param start: top right cell address from where values are inserted
        :param copy_index: if index should be copied
        :param copy_head: if headers should be copied
        :param fit: should the worksheet should be resized to fit the dataframe
        :param escape_formulae: If any value starts with an equals sign =, it will be
               prefixed with a apostrophe ', to avoid being interpreted as a formula.

        """
        start = self.get_addr(start, 'tuple')
        values = df.values.tolist()
        if copy_index:
            for i in range(values):
                values[i].insert(0, i)
        if copy_head:
            head = df.columns.tolist()
            if copy_index:
                head.insert(0, '')
            values.insert(0, head)
        if fit:
            self.rows = start[0] - 1 + len(values[0])
            self.cols = start[1] - 1 + len(values)
        if escape_formulae:
            warnings.warn("Functionality not implimented")
        self.update_cells(crange=start, values=values)

    def get_as_df(self, head=1, numerize=True, empty_value=''):
        """
        get value of wprksheet as a pandas dataframe

        :param head: colum head for df
        :param numerize: if values should be numerized
        :param empty_value: valued  used to indicate empty cell value

        :returns: pandas.Dataframe

        """
        if not DataFrame:
            raise ImportError("pandas")
        idx = head - 1
        values = self.all_values(returnas='matrix', include_empty=True)
        keys = list(''.join(values[idx]))
        if numerize:
            values = [numericise_all(row[:len(keys)], empty_value) for row in values[idx + 1:]]
        else:
            values = [row[:len(keys)] for row in values[idx + 1:]]
        return DataFrame(values, columns=keys)

    def export(self, fformat=ExportType.CSV):
        """Export the worksheet in specified format.

        :param fformat: A format of the output as Enum ExportType
        """
        # warnings.warn("Method not Implimented")
        if fformat is ExportType.CSV:
            import csv
            filename = 'worksheet'+str(self.id)+'.csv'
            with open(filename, 'wt') as f:
                writer = csv.writer(f, lineterminator="\n")
                writer.writerows(self.all_values())
        elif isinstance(fformat, ExportType):
            self.client.export(self.spreadsheet.id, fformat)

    def copy_to(self, spreadsheet_id):
        """
        copy the worksheet to specified spreadsheet
        :param spreadsheet_id: id to the spreadsheet to copy

        """
        jsheet = dict()
        self.client.sh_copy_worksheet(self.spreadsheet.id, self.id, spreadsheet_id)

    def __iter__(self):
        rows = self.all_values(majdim='ROWS')
        for row in rows:
            yield(row + (self.cols - len(row))*[''])

    def __getitem__(self, item):
        if type(item) == int:
            if item >= self.cols:
                raise CellNotFound
            try:
                row = self.all_values()[item]
            except IndexError:
                row = ['']*self.cols
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
        self._formula = ''
        # self.format = FormatType.CUSTOM  # @TODO
        self.parse_value = True  # if the value will be parsed to corsp types
        self._parse_formula = False # should value parse formula
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
        """formatted value of the cell"""
        return self._value

    @value.setter
    def value(self, value):
        if type(value) == str and not self._parse_formula:  # remove formulae
            if value.startswith('='):
                value = "'"+str(value)
        if self.worksheet:
            self.worksheet.update_cell(self.label, value, self.parse_value)
            self._value = value
        else:
            self._value = value

    @property
    def formula(self):
        """formula if any of the cell"""
        if self.worksheet:
            self._formula = self.worksheet.client.get_range(self.worksheet.spreadsheet.id,
                                                            self.worksheet._get_range(self.label, self.label),
                                                            majordim='ROWS', value_render=ValueRenderOption.FORMULA)[0][0]
        if not self._formula.startswith('='):
            self._formula = ""
        return self._formula

    @formula.setter
    def formula(self, formula):
        self._parse_formula = True
        self.value = formula
        self._parse_formula = False

    def fetch(self):
        """ Update the value of the cell from sheet """
        if self.worksheet:
            self._value = self.worksheet.cell(self._label).value

    def __repr__(self):
        return '<%s R%sC%s %s>' % (self.__class__.__name__,
                                   self.row,
                                   self.col,
                                   repr(self.value))

