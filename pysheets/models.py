# -*- coding: utf-8 -*-

"""
gspread.models
~~~~~~~~~~~~~~

This module contains common spreadsheets' models

"""

import re
from collections import defaultdict
from itertools import chain
import json

from .utils import finditem, numericise_all
from .exceptions import IncorrectCellLabel, WorksheetNotFound, CellNotFound


class Spreadsheet(object):

    """ A class for a spreadsheet object."""

    def __init__(self, client, jsonSsheet):
        self.client = client
        self._sheet_list = []
        self._jsonSsheet = jsonSsheet
        #print jsonSsheet['spreadsheetId']
        self._update_properties(jsonSsheet)
        

    @property
    def id(self):
        return self._id

    @property
    def title(self):
        return self._title

    def get_id_fields(self):
        return {'spreadsheet_id': self.id}

    def _update_properties(self, jsonSsheet=None):
        ''' update all sheet properies

        '''
        if not jsonSsheet and len(self.id)>1:
            jsonSsheet = self.client.fetch_worksheet(self.id)

        self._id = jsonSsheet['spreadsheetId']
        self._fetch_sheets(jsonSsheet)
        self._title = jsonSsheet['properties']['title']

    def _fetch_sheets(self,jsonSsheet):
        ''' update sheets list

        '''
        if not jsonSsheet:
            jsonSsheet = self.client.fetch_worksheet(self.id)
        for sheet in jsonSsheet.get('sheets'):
            self._sheet_list.append(Worksheet(self,sheet))

    def add_worksheet(self, title, rows, cols):
        """Adds a new worksheet to a spreadsheet.

        :param title: A title of a new worksheet.
        :param rows: Number of rows.
        :param cols: Number of columns.

        @TODO Returns a newly created :class:`worksheets <Worksheet>`.
        """
        client.add_worksheet(title, rows, cols)
        self._fetch_sheets()

    def del_worksheet(self, worksheet):
        """Deletes a worksheet from a spreadsheet.

        :param worksheet: The worksheet to be deleted.

        """
        self.client.del_worksheet(worksheet)
        self._sheet_list.remove(worksheet)

    #@TODO
    def worksheets(self, property=None, value=None):
        """Returns a list of all :class:`worksheets <Worksheet>`
        in a spreadsheet.

        """
        pass

    def worksheet(self, property='id', value=0):
        """Returns a worksheet with specified `title`.

        The returning object is an instance of :class:`Worksheet`.

        :param title: A title of a worksheet. If there're multiple
                      worksheets with the same title, first one will
                      be returned.

        Example. Getting worksheet named 'Annual bonuses'

        >>> sht = client.open('Sample one')
        >>> worksheet = sht.worksheet('Annual bonuses')

        """
        try:
            return [x for x in self._sheet_list if getattr(x,property)][0]
        except IndexError:
            self._fetch_sheets()
            try:
                return [x for x in self._sheet_list if getattr(x,property)][0]
            except IndexError:
                raise WorksheetNotFound(title)

    @property
    def sheet1(self):
        """Shortcut property for getting the first worksheet."""
        return self.worksheet()

    @property
    def title(self):
        return self.title

    def __iter__(self):
        for sheet in self.worksheets():
            yield(sheet)


class Worksheet(object):

    """A class for worksheet object."""

    def __init__(self, spreadsheet, jsonSheet):
        self.spreadsheet = spreadsheet
        self.client = spreadsheet.client
        self._id = ''
        self._title = ''
        self._index = ''
        self.jsonSheet = jsonSheet
        self._update_properties(jsonSheet)

    def __repr__(self):
        return '<%s %s id:%s>' % (self.__class__.__name__,
                                  repr(self.title),
                                  self.id)

    @property
    def id(self):
        """Id of a worksheet."""
        return self._id

    @property
    def index(self):
        return self._index

    @property
    def title(self):
        """Title of a worksheet."""
        return self._title

    @property
    def row_count(self):
        """Number of rows"""
        return int(self.jsonSheet['gridProperties']['rowCount'])

    @property
    def col_count(self):
        """Number of columns"""
        return int(self.jsonSheet['gridProperties']['columnCount'])

    @property
    def updated(self):
        """ @TODO Updated time in RFC 3339 format"""
        return None

    def _update_properties(self, jsonSheet):
        self._id = jsonSheet['properties']['sheetId']
        self._type = jsonSheet['properties']['sheetType']
        self._title = jsonSheet['properties']['title']
        self._index = jsonSheet['properties']['index']
        self.rowCount = jsonSheet['properties']['gridProperties']['rowCount']
        self.colCount = jsonSheet['properties']['gridProperties']['columnCount']

    def get_id_fields(self):
        return {'spreadsheet_id': self.spreadsheet.id,
                'worksheet_id': self.id}

    def _cell_addr(self, row, col):
        return 'R%sC%s' % (row, col)

    @staticmethod
    def get_int_addr(label):
        """Translates cell's label address to a tuple of integers.

        The result is a tuple containing `row` and `column` numbers.

        :param label: String with cell label in common format, e.g. 'B1'.
                      Letter case is ignored.

        Example:

        >>> Worksheet.get_int_addr('A1')
        (1, 1)

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

        return (row, col)

    @staticmethod
    def get_addr_int(row, col):
        """Translates cell's tuple of integers to a cell label.

        The result is a string containing the cell's coordinates in label form.

        :param row: The row of the cell to be converted.
                    Rows start at index 1.

        :param col: The column of the cell to be converted.
                    Columns start at index 1.

        Example:

        >>> Worksheet.get_addr_int(1, 1)
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

    def acell(self, label):
        """Returns an instance of a :class:`Cell`.

        :param label: String with cell label in common format, e.g. 'B1'.
                      Letter case is ignored.

        Example:

        >>> wks.acell('A1') # this could be 'a1' as well
        <Cell R1C1 "I'm cell A1">

        """
        val = self.client.get_range(self._get_range(label,label),  'ROWS')[0][0]
        return Cell(label, self,val)

    def cell(self, row, col):
        """Returns an instance of a :class:`Cell` positioned in `row`
           and `col` column.

        :param row: Integer row number.
        :param col: Integer column number.

        Example:

        >>> wks.cell(1, 1)
        <Cell R1C1 "I'm cell A1">

        """
        return self.acell(Worksheet.get_addr_int(row,col))

    def range(self, alphanum):
        """Returns a list of :class:`Cell` objects from specified range.

        :param alphanum: A string with range value in common format,
                         e.g. 'A1:A5'.

        """
        startcell = Cell( alphanum.split(':')[0] )
        endcell = Cell( alphanum.split(':')[1] )
        values = self.client.get_range(self._get_range(startcell.label,endcell.label),  'ROWS')
        cells = []
        for i in range(startcell.col,endcell.col):
            rcells = []
            for j in xrange(startcell.row,endcell.row):
                rcells = [rcells, Cell(Worksheet.get_addr_int((j+1),i), values[i]) ]
            cells = [cells, rcells]
        return cells

    #@TODO
    def get_all_values(self):
        """Returns a list of lists containing all cells' values as strings."""
        cells = self._fetch_cells()

        # defaultdicts fill in gaps for empty rows/cells not returned by gdocs
        rows = defaultdict(lambda: defaultdict(str))
        for cell in cells:
            row = rows.setdefault(int(cell.row), defaultdict(str))
            row[cell.col] = cell.value

        # we return a whole rectangular region worth of cells, including
        # empties
        if not rows:
            return []

        all_row_keys = chain.from_iterable(row.keys() for row in rows.values())
        rect_cols = range(1, max(all_row_keys) + 1)
        rect_rows = range(1, max(rows.keys()) + 1)

        return [[rows[i][j] for j in rect_cols] for i in rect_rows]

    #@TODO
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

    def row_values(self, row):
        """Returns a list of all values in a `row`.

        Empty cells in this list will be rendered as :const:`None`.

        """
        vlaues = self.client.get_range(self._get_range(Worksheet.get_addr_int(row,1),Worksheet.get_addr_int(self.colCount, col)),  'ROWS')
        cells = []
        for i in range(0,self.colCount):
            cells = [cells, Cell(Worksheet.get_addr_int(row,(i+1)), values[i] ) ]
        return cells

    def col_values(self, col):
        """Returns a list of all values in column `col`.

        Empty cells in this list will be rendered as :const:`None`.

        """
        values = self.client.get_range(self._get_range(Worksheet.get_addr_int(1, col),Worksheet.get_addr_int(self.rowCount, col)),  'COLUMNS')
        cells = []
        for i in range(0,self.rowCount):
            cells = [cells, Cell(Worksheet.get_addr_int((i+1),col), values[i] ) ]
        return cells

    def update_acell(self, label, val):
        """Sets the new value to a cell.

        :param label: String with cell label in common format, e.g. 'B1'.
                      Letter case is ignored.
        :param val: New value.

        Example:

        >>> wks.update_acell('A1', '42') # this could be 'a1' as well
        <Cell R1C1 "I'm cell A1">

        """
        self.client.update_range(_get_range(label,label),[[ unicode(val) ]])
        
    def _get_range(self, start_cell,end_cell):
        '''get range in A1 notatin given start and end labels

        '''
        return self.title + '!' + ('%s:%s' % (start_cell, end_cell))

    def update_cell(self, row, col, val):
        """Sets the new value to a cell.

        :param row: Row number.
        :param col: Column number.
        :param val: New value.

        """
        self.update_acell( Worksheet.get_addr_int(row,col), val=unicode(val))

    def update_cells(self, cell_list):
        """Updates cells in batch.

        :param cell_list: List of a :class:`Cell` objects to update.

        """
        self.client.start_batch()
        for cell in cell_list:
            update_acell(cell.label,cell.value)
        self.client.stop_batch()

    #@TODO
    def resize(self, rows=None, cols=None):
        """Resizes the worksheet.

        :param rows: New rows number.
        :param cols: New columns number.
        """
        pass

    def add_rows(self, rows):
        """Adds rows to worksheet.

        :param rows: Rows number to add.
        """
        self.resize(rows=self.row_count + rows)

    def add_cols(self, cols):
        """Adds colums to worksheet.

        :param cols: Columns number to add.
        """
        self.resize(cols=self.col_count + cols)

    #@TODO
    def append_row(self, values):
        """Adds a row to the worksheet and populates it with values.
        Widens the worksheet if there are more values than columns.

        Note that a new Google Sheet has 100 or 1000 rows by default. You
        may need to scroll down to find the new row.

        :param values: List of values for the new row.
        """
        self.add_rows(1)
        new_row = self.row_count
        data_width = len(values)
        if self.col_count < data_width:
            self.resize(cols=data_width)

        cell_list = []
        for i, value in enumerate(values, start=1):
            cell = self.cell(new_row, i)
            cell.value = value
            cell_list.append(cell)

        self.update_cells(cell_list)

    #@TODO
    def insert_row(self, values, index=1):
        """"Adds a row to the worksheet at the specified index and populates it with values.
        Widens the worksheet if there are more values than columns.

        :param values: List of values for the new row.
        """
        if index == self.row_count + 1:
            return self.append_row(values)
        elif index > self.row_count + 1:
            raise IndexError('Row index out of range')

        
        self.update_cells(cells_after_insert)

    #@TODO
    def _finder(self, func, query):
        cells = self._fetch_cells()

        if isinstance(query, basestring):
            match = lambda x: x.value == query
        else:
            match = lambda x: query.search(x.value)

        return func(match, cells)

    #@TODO
    def find(self, query):
        """Finds first cell matching query.

        :param query: A text string or compiled regular expression.
        """
        try:
            return self._finder(finditem, query)
        except StopIteration:
            raise CellNotFound(query)

    #@TODO
    def findall(self, query):
        """Finds all cells matching query.

        :param query: A text string or compiled regular expression.
        """
        return list(self._finder(filter, query))

    #@TODO
    def export(self, format='csv'):
        """Export the worksheet in specified format.

        :param format: A format of the output.
        """
        pass


class Cell(object):

    """An instance of this class represents a single cell
    in a :class:`worksheet <Worksheet>`.

    """

    def __init__(self, label, worksheet = None, val = ''):
        self.worksheet = worksheet
        self._row = int(Worksheet.get_int_addr(label)[0])
        self._col = int(Worksheet.get_int_addr(label)[1])
        self._label = Worksheet.get_addr_int(self._row,self._col)
        self._value = val
        self.format = None
    
    @property
    def row(self):
        """Row number of the cell."""
        return self._row

    @property
    def col(self):
        """Column number of the cell."""
        return self._col

    @property
    def label(self):
        return self._label
    
    @property
    def value(self):
        return self._value

    @value.setter
    def update_val(self, value):
        if worksheet:
            self.worksheet.update_acell(self.label,value)
        else:
            self._value = value;

    def __repr__(self):
        return '<%s R%sC%s %s>' % (self.__class__.__name__,
                                   self.row,
                                   self.col,
                                   repr(self.value))
