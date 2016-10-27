# -*- coding: utf-8 -*-

"""
pygsheets.models
~~~~~~~~~~~~~~

This module contains common spreadsheets' models

"""

import re

from .exceptions import IncorrectCellLabel, WorksheetNotFound, CellNotFound, InvalidArgumentValue, InvalidUser
from .utils import finditem, numericise_all


class Spreadsheet(object):

    """ A class for a spreadsheet object."""

    def __init__(self, client, jsonsheet=None, id=None):
        self.client = client
        self._sheet_list = []
        self._jsonsheet = jsonsheet
        self._id = id
        self._update_properties(jsonsheet)
        self._permissions = dict()

    def __repr__(self):
        return '<%s %s Sheets:%s>' % (self.__class__.__name__,
                                      repr(self.title), len(self._sheet_list))

    @property
    def id(self):
        """ id of the sheet """
        return self._id

    # @TODO - link changes with colud
    @property
    def title(self):
        """ title of the sheet """
        return self._title

    # @TODO - link changes with colud
    @property
    def defaultformat(self):
        """ deafault cell format"""
        return self._defaultFormat

    @property
    def sheet1(self):
        """Shortcut property for getting the first worksheet."""
        return self.worksheet()

    def get_id_fields(self):
        return {'spreadsheet_id': self.id}

    def _update_properties(self, jsonsheet=None):
        """ update all sheet properies

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

    def _fetch_sheets(self, jsonsheet):
        """update sheets list
        """
        if not jsonsheet:
            jsonsheet = self.client.open_by_key(self.id, 'json')
        for sheet in jsonsheet.get('sheets'):
            self._sheet_list.append(Worksheet(self, sheet))

    def link(self, syncToColoud=False):
        """ Link the spread sheet with colud, so all local changes
            will be updated instantly, so does all data fetches

            :param  syncToColoud: update the cloud with local changes if set to true
                          update the local copy with cloud if set to false
        """
        pass

    def unlink(self):
        """ Unlink the spread sheet with colud, so all local changes
            will be made on local copy fetched
        """
        pass

    def share(self, addr, role='reader', expirationTime=None, is_group=False):
        """
        create/update permission for user/group/domain
        :param addr: this is the email for user/group and domain adress for domains
        :param role: permission to be applied
        :param expirationTime: (Not Implimented) time until this permission should last
        :param is_group: boolean , Is this a use/group used only when email provided

        :type addr : email
        :type role: 'owner','writer','commenter','reader'
        :type expirationTime: datetime
        :type is_group: bool
        :return:
        """
        return self.client.add_permission(self.id, addr, role=role, is_group=False)

    def list_permissions(self):
        """
        list all the permissions of the spreadsheet
        :return:
        """
        permissions = self.client.list_permissions(self.id)
        self._permissions = permissions['permissions']
        return self._permissions

    def remove_permissions(self, addr):
        """
        removes all permissions of the user provided
        :param addr: email/domain of the user
        :return:
        """
        try:
            result = self.client.remove_permissions(self.id, addr, self._permissions)
        except InvalidUser:
            result = self.client.remove_permissions(self.id, addr)
        return result

    def add_worksheet(self, title, rows, cols):
        """Adds a new worksheet to a spreadsheet.

        :param title: A title of a new worksheet.
        :param rows: Number of rows.
        :param cols: Number of columns.

        @TODO Returns a newly created :class:`worksheets <Worksheet>`.
        """
        jsheet = dict()
        jsheet['properties'] = self.client.add_worksheet(title, rows, cols)
        self._sheet_list.append(Worksheet(self, jsheet))

    def del_worksheet(self, worksheet):
        """Deletes a worksheet from a spreadsheet.

        :param worksheet: The worksheet to be deleted.

        """
        if worksheet not in self.worksheets():
            raise WorksheetNotFound
        self.client.del_worksheet(worksheet.id)
        self._sheet_list.remove(worksheet)

    def worksheets(self, property=None, value=None):
        """Returns a list of all :class:`worksheets <Worksheet>`
        in a spreadsheet.

        """
        if not property and not value:
            return self._sheet_list
        
        sheets = [x for x in self._sheet_list if getattr(x,property)==value]
        if not len(sheets) > 0:
            self._fetch_sheets()
            sheets = [x for x in self._sheet_list if getattr(x,property)==value]
            if not len(sheets) > 0:
                raise WorksheetNotFound()
        return sheets

    def worksheet(self, property='id', value=0):
        """Returns a worksheet with specified `title`.

        The returning object is an instance of :class:`Worksheet`.

        :param property: A property of a worksheet. If there're multiple
                      worksheets with the same title, first one will
                      be returned.
        :param value: value of given property

        Example. Getting worksheet named 'Annual bonuses'

        >>> sht = client.open('Sample one')
        >>> worksheet = sht.worksheet('title','Annual bonuses')

        """
        return self.worksheets(property, value)[0]

    def __iter__(self):
        for sheet in self.worksheets():
            yield(sheet)


class Worksheet(object):

    """A class for worksheet object."""

    def __init__(self, spreadsheet, jsonSheet):
        self.spreadsheet = spreadsheet
        self.client = spreadsheet.client
        self._linked = True
        self.jsonSheet = jsonSheet
        self.data_grid = ''  # for storing sheet data while unlinked

    def __repr__(self):
        return '<%s %s id:%s>' % (self.__class__.__name__,
                                  repr(self.title),
                                  self.id)

    @property
    def id(self):
        """Id of a worksheet."""
        return self.jsonSheet['properties']['sheetId']

    @property
    def index(self):
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
    def row_count(self):
        """Number of rows"""
        return int(self.jsonSheet['properties']['gridProperties']['rowCount'])

    @row_count.setter
    def row_count(self, row_count):
        self.jsonSheet['properties']['gridProperties']['rowCount'] = int(row_count)
        if self._linked:
            self.client.update_sheet_properties(self.jsonSheet['properties'], 'gridProperties/rowCount')

    @property
    def col_count(self):
        """Number of columns"""
        return int(self.jsonSheet['properties']['gridProperties']['columnCount'])

    @col_count.setter
    def col_count(self, col_count):
        self.jsonSheet['properties']['gridProperties']['columnCount'] = int(col_count)
        if self._linked:
            self.client.update_sheet_properties(self.jsonSheet['properties'], 'gridProperties/columnCount')

    @property
    def updated(self):
        """ @TODO Updated time in RFC 3339 format(use drive api)"""
        return None

    def link(self, syncToColoud=True):
        """ Link the spread sheet with colud, so all local changes
            will be updated instantly, so does all data fetches

            :param  syncToColoud: update the cloud with local changes if set to true
                          update the local copy with cloud if set to false
        """
        if syncToColoud:
            self.client.update_sheet_properties(self.jsonSheet['properties'])
        else:
            wks = self.spreadsheet.worksheet(self, property='id', value=self.id)
            self.jsonSheet = wks.jsonSheet
        self._linked = True

    def unlink(self):
        """ Unlink the spread sheet with colud, so all local changes
            will be made on local copy fetched
        """
        self._linked = False

    def get_id_fields(self):
        return {'spreadsheet_id': self.spreadsheet.id,
                'worksheet_id': self.id}

    @staticmethod
    def get_int_addr(label):
        """Translates cell's label address to a tuple of integers.

        The result is a tuple containing `row` and `column` numbers.

        :param label: String with cell label in common format, e.g. 'B1'.
                      Letter case is ignored.

        Example:

        """
        _MAGIC_NUMBER = 64
        _cell_addr_re = re.compile(r'([A-Za-z]+)(\d+)')

        m = _cell_addr_re.match(label)
        if m:
            column_label = m.group(1).upper()
            row = int(m.group(2))

            col = 0
            for i, c in enumerate(reversed(column_label)):
                col += (ord(c) - _MAGIC_NUMBER) * (26 ** i)
        else:
            raise IncorrectCellLabel(label)

        return row, col

    @staticmethod
    def get_addr_int(row, col):
        """Translates cell's tuple of integers to a cell label.

        The result is a string containing the cell's coordinates in label form.

        :param row: The row of the cell to be converted.
                    Rows start at index 1.

        :param col: The column of the cell to be converted.
                    Columns start at index 1.

        Example:

        A1

        """
        _MAGIC_NUMBER = 64

        row = int(row)
        col = int(col)

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

    @staticmethod
    def get_addr(addr, output='flip'):
        """
        fcuntion to change the adress of cells

        :param addr: adress as tuple or label
        :param output: operation - 'label' will output label
                     'tuple' will output tuple
                     'flip' will convert to other type
        :return: tuple or label
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
        """Returns an instance of a :class:`Cell` positioned in `row`
           and `col` column.

        :param addr cell adress as either tuple - (row, col) or cell label 'A1'

        Example:

        >>> wks.cell((1,1))
        <Cell R1C1 "I'm cell A1">

        >>> wks.cell('A1')
        <Cell R1C1 "I'm cell A1">

        """
        if type(addr) is str:
            val = self.client.get_range(self._get_range(addr, addr), 'ROWS')[0][0]
        elif type(addr) is tuple:
            label = Worksheet.get_addr(addr, 'label')
            val = self.client.get_range(self._get_range(label, label), 'ROWS')[0][0]
        else:
            raise CellNotFound
        return Cell(addr, val, self)

    def range(self, alphanum):
        """Returns a list of :class:`Cell` objects from specified range.

        :param alphanum: A string with range value in common format,
                         e.g. 'A1:A5'.

        """
        startcell = Cell(alphanum.split(':')[0])
        endcell = Cell(alphanum.split(':')[1])
        return self.get_values(startcell, endcell, returnas='cell')

    def get_values(self, start, end, majdim='ROWS', returnas='value'):
        """Returns value of cells given the topleft corner position
        and bottom right position

        :param start: topleft position as tuple or label
        :param end: bottomright position as tuple or label
        :param majdim: output as rowwise or columwise
                       takes - 'ROWS' or 'COLMUNS'
        :param returnas: return as list of strings of cell objects
                         takes - 'value' or 'cell'

        Example:

        >>> wks.get_values((1,1),(3,3))
        [[u'another look.', u'', u'est'],
         [u'EE 4212', u"it's down there "],
         [u'ee 4210', u'somewhere, let me take ']]

        """
        start_label = Worksheet.get_addr(start, 'label')
        end_label = Worksheet.get_addr(end, 'label')
        values = self.client.get_range(self._get_range(start_label, end_label), majdim.upper())

        if returnas.lower() == 'value':
            return values
        elif returnas.lower() == 'cell':
            cells = []
            for k in xrange(0, len(values)):
                row = []
                for i in xrange(0, len(values[k])):
                    row.append(Cell((k+1, i+1), values[k][i], self))
                cells.append(row)
            return cells
        else:
            return None

    def get_all_values(self, majdim='ROWS', returnas='value'):
        """Returns a list of lists containing all cells' values as strings.

        :param majdim: output as rowwise or columwise
        :param returnas: return as list of strings of cell objects

        Example:

        >>> wks.get_all_values()
        [[u'another look.', u'', u'est'],
         [u'EE 4212', u"it's down there "],
         [u'ee 4210', u'somewhere, let me take ']]
        """
        return self.get_values((1, 1), (self.row_count, self.col_count), majdim, returnas)

    # @TODO improve empty2zero for other types also
    def get_all_records(self, empty2zero=False, head=1):
        """Returns a list of dictionaries, all of them having:
            - the contents of the spreadsheet's with the head row as keys,
            And each of these dictionaries holding
            - the contents of subsequent rows of cells as values.


        Cell values are numericised (strings that can be read as ints
        or floats are converted).

        :param empty2zero: determines whether empty cells are converted to zeros.
        :param head: determines wich row to use as keys, starting from 1
            following the numeration of the spreadsheet."""
        idx = head - 1
        data = self.get_all_values()
        keys = data[idx]
        values = [numericise_all(row, empty2zero) for row in data[idx + 1:]]
        return [dict(zip(keys, row)) for row in values]

    def row_values(self, row, returnas='value'):
        """Returns a list of all values in a `row`.
            :param row - index of row
            :param returnas - ('value' or 'cell') return as cell objects or just values

        Empty cells in this list will be rendered as :const:``.

        """
        return self.get_values((row, 1), (row, self.col_count), returnas=returnas)[0]

    def col_values(self, col, returnas='value'):
        """Returns a list of all values in column `col`.
            :param col - index of col
            :param returnas - ('value' or 'cell') return as cell objects or just values

        Empty cells in this list will be rendered as :const:``.

        """
        return self.get_values((1, col), (self.row_count, col), majdim='COLUMNS', returnas=returnas)[0]

    def update_cell(self, addr, val):
        """Sets the new value to a cell.

        :param addr: cell adress as tuple (row,column) or label 'A1'.
        :param val: New value.

        Example:

        >>> wks.update_cell('A1', '42') # this could be 'a1' as well
        <Cell R1C1 "I'm cell A1">

        """
        label = Worksheet.get_addr(addr, 'label')
        self.client.update_range(self._get_range(label, label), [[unicode(val)]])

    def update_cells(self, cell_list=None, range=None, values=None, majordim='ROWS'):
        """Updates cells in batch, it can take either a cell list or a range and values

        :param cell_list: List of a :class:`Cell` objects to update with their values
        :param range: range in format A1:A2
        :param values: list of values if range given
        :param majordim: major dimension of given data

        """
        if cell_list:
            self.client.start_batch()
            for cell in cell_list:
                self.update_cell(cell.label, cell.value)
            self.client.stop_batch()
        elif range and values:
            self.client.update_range(self._get_range(*range.split(':')), values, majordim)
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
        self.row_count = rows
        self.col_count = cols
        self.link(True)

    def add_rows(self, rows):
        """Adds rows to worksheet.

        :param rows: Rows number to add.
        """
        self.resize(rows=self.row_count + rows, cols=self.col_count)

    def add_cols(self, cols):
        """Adds colums to worksheet.

        :param cols: Columns number to add.
        """
        self.resize(cols=self.col_count + cols, rows=self.row_count)

    def insert_cols(self, col, number=1, values=None):
        """insert a colum after the colum <col> and fill with values <values>

        """
        self.client.insertdim(self.id, 'COLUMNS', col, (col+number), False)
        if values:
            self.update_col(col+1, values)

    def insert_rows(self, row, number=1, values=None):
        """ insert a row after the row <row> and fill with values <values>

        """
        self.client.insertdim(self.id, 'ROWS', row, (row+number), False)
        if values:
            self.update_row(row+1, values)

    # @TODO
    def append_row(self, values):
        """Adds a row to the worksheet and populates it with values.
        Widens the worksheet if there are more values than columns.

        :param values: List of values for the new row.
        """
        pass

    # @TODO
    def _finder(self, func, query):
        pass

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
        pass

    # @TODO
    def export(self, format='csv'):
        """Export the worksheet in specified format.

        :param format: A format of the output.
        """
        pass


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
        self._value = val
        self.format = None
    
    @property
    def row(self):
        """Row number of the cell."""
        return self._row

    @row.setter
    def row(self, row):
        if self.worksheet:
            ncell = self.worksheet.cell(row)
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
        return self._value

    @value.setter
    def value(self, value):
        if self.worksheet:
            self.worksheet.update_cell(self.label, value)
            self._value = value
        else:
            self._value = value

    def fetch(self):
        if self.worksheet:
            self._value = self.worksheet.cell(self._label)

    def __repr__(self):
        return '<%s R%sC%s %s>' % (self.__class__.__name__,
                                   self.row,
                                   self.col,
                                   repr(self.value))
