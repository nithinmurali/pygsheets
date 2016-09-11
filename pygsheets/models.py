# -*- coding: utf-8 -*-

"""
pygsheets.models
~~~~~~~~~~~~~~

This module contains common spreadsheets' models

"""

import re
from collections import defaultdict
from itertools import chain
import json

from .exceptions import IncorrectCellLabel, WorksheetNotFound, CellNotFound, InvalidArgumentValue


class Spreadsheet(object):

    """ A class for a spreadsheet object."""

    def __init__(self, client, jsonSsheet=None,id=None):
        self.client = client
        self._sheet_list = []
        self._jsonSsheet = jsonSsheet
        self._id = id
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

            :param jssonSsheet json var to update values form \
                if not specified, will fetch it and update

        '''
        if not jsonSsheet and len(self.id)>1:
            self._jsonSsheet = self.client.open_by_key(self.id, 'json')
        elif not jsonSsheet and len(self.id)==0:
            raise InvalidArgumentValue

        self._id = self._jsonSsheet['spreadsheetId']
        self._fetch_sheets(self._jsonSsheet)
        self._title = self._jsonSsheet['properties']['title']
        self.client.spreadsheetId = self._id

    def _fetch_sheets(self,jsonSsheet):
        ''' update sheets list

        '''
        if not jsonSsheet:
            jsonSsheet = self.client.open_by_key(self.id, 'json')
        for sheet in jsonSsheet.get('sheets'):
            self._sheet_list.append(Worksheet(self,sheet))

    # @TODO
    def add_worksheet(self, title, rows, cols):
        """Adds a new worksheet to a spreadsheet.

        :param title: A title of a new worksheet.
        :param rows: Number of rows.
        :param cols: Number of columns.

        @TODO Returns a newly created :class:`worksheets <Worksheet>`.
        """
        self.client.add_worksheet(title, rows, cols)
        self._fetch_sheets()

    # @TODO
    def del_worksheet(self, worksheet):
        """Deletes a worksheet from a spreadsheet.

        :param worksheet: The worksheet to be deleted.

        """
        self.client.del_worksheet(worksheet)
        self._sheet_list.remove(worksheet)

    def worksheets(self, property=None, value=None):
        """Returns a list of all :class:`worksheets <Worksheet>`
        in a spreadsheet.

        """
        if not property and not value:
            return self._sheet_list
        
        sheets = [x for x in self._sheet_list if getattr(x,property)==value]
        if not len(sheets)>0:
            self._fetch_sheets()
            sheets = [x for x in self._sheet_list if getattr(x,property)==value]
            if not len(sheets)>0:
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
        return self.worksheets(property,value)[0]

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
        return int(self.jsonSheet['properties']['gridProperties']['rowCount'])

    @property
    def col_count(self):
        """Number of columns"""
        return int(self.jsonSheet['properties']['gridProperties']['columnCount'])

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

    def _get_range(self, start_label,end_label):
        '''get range in A1 notation, given start and end labels

        '''
        return self.title + '!' + ('%s:%s' % (start_label, end_label))

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
                rcells = [rcells, Cell((j+1,i), values[i]) ]
            cells = [cells, rcells]
        return cells

    def get_values(self,start,end,majDim='ROWS',returnas='value'):
        values = self.client.get_range( self._get_range(Worksheet.get_addr_int(*start),Worksheet.get_addr_int(*end)), majDim)
        if returnas == 'value':
            return values
        elif returnas == 'cell':
            cells = []
            for k in xrange(0,len(values)):
                row = []
                for i in xrange(0,len(values[k])):
                    row.append( Cell((k+1,i+1), self, values[k][i]) )
                cells.append(row)
            return cells
        else:
            return None

    #@TODO
    def get_all_values(self,majDim='ROWS',returnas='value'):
        """Returns a list of lists containing all cells' values as strings."""
        return self.get_values((1,1),(self.rowCount,self.colCount),majDim,returnas)

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

        pass

    def row_values(self, row, returnas='value'):
        """Returns a list of all values in a `row`.
            :param row - index of row
            :param returnas - ('value' or 'cell') return as cell objects or just values

        Empty cells in this list will be rendered as :const:``.

        """
        return self.get_values((row,1),(row,self.colCount),returnas=returnas)[0]


    def col_values(self, col, returnas='value'):
        """Returns a list of all values in column `col`.
            :param col - index of col
            :param returnas - ('value' or 'cell') return as cell objects or just values

        Empty cells in this list will be rendered as :const:``.

        """
        return self.get_values((1,col),(self.rowCount,col),majDim='COLUMNS',returnas=returnas)[0]

    def update_acell(self, label, val):
        """Sets the new value to a cell.

        :param label: String with cell label in common format, e.g. 'B1'.
                      Letter case is ignored.
        :param val: New value.

        Example:

        >>> wks.update_acell('A1', '42') # this could be 'a1' as well
        <Cell R1C1 "I'm cell A1">

        """
        self.client.update_range(self._get_range(label,label),[[ unicode(val) ]])
    
    def update_cell(self, row, col, val):
        """Sets the new value to a cell.

        :param row: Row number.
        :param col: Column number.
        :param val: New value.

        """
        self.update_acell( Worksheet.get_addr_int(row,col), val=unicode(val))

    def update_cells(self, cell_list=None, range=None, values=None, majorDim='ROWS'):
        """Updates cells in batch.

        :param cell_list: List of a :class:`Cell` objects to update.
        :param range: range in format A1:A2

        """
        print range
        if cell_list:
            self.client.start_batch()
            for cell in cell_list:
                update_acell(cell.label,cell.value)
            self.client.stop_batch()
        elif range and values:
            self.client.update_range(self._get_range(*range.split(':') ),values,majorDim)
        else:
            raise InvalidArgumentValue(cell_list) #@TODO test

    def update_col(self,index, values):
        '''update an existing colum with values

        '''
        range = Worksheet.get_addr_int(1,index) +":"+Worksheet.get_addr_int(len(values),index)
        self.update_cells(range=range,values=[values],majorDim='COLUMNS')

    def update_row(self,index, values):
        '''update an existing row with values

        '''
        range= self.get_addr_int(index,1) + ':' +self.get_addr_int(index,len(values))
        self.update_cells(range=range,values=[values],majorDim='ROWS')

    #@TODO
    def resize(self, rows=None, cols=None):
        """Resizes the worksheet.

        :param rows: New rows number.
        :param cols: New columns number.
        """
        pass

    #@TODO
    def add_rows(self, rows):
        """Adds rows to worksheet.

        :param rows: Rows number to add.
        """
        self.resize(rows=self.row_count + rows)

    #@TODO
    def add_cols(self, cols):
        """Adds colums to worksheet.

        :param cols: Columns number to add.
        """
        self.resize(cols=self.col_count + cols)

    def insert_cols(self, col, number=1, values = None):
        ''' insert a colum after the colum <col> and fill with values <values>

        '''
        self.client.insertdim(self.id,'COLUMNS',col, (col+number), False)
        if values:
            self.update_col(col+1,values)
    def insert_rows(self, row, number=1, values = None):
        ''' insert a row after the row <row> and fill with values <values>

        '''
        self.client.insertdim(self.id,'ROWS',row, (row+number), False)
        if values:
            self.update_row(row+1, values)
    #@TODO
    def append_row(self, values):
        """Adds a row to the worksheet and populates it with values.
        Widens the worksheet if there are more values than columns.

        Note that a new Google Sheet has 100 or 1000 rows by default. You
        may need to scroll down to find the new row.

        :param values: List of values for the new row.
        """
        pass

    #@TODO
    def _finder(self, func, query):
        pass

    #@TODO
    def find(self, query):
        """Finds first cell matching query.

        :param query: A text string or compiled regular expression.
        """
        pass

    #@TODO
    def findall(self, query):
        """Finds all cells matching query.

        :param query: A text string or compiled regular expression.
        """
        pass

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

    def __init__(self, pos, worksheet = None, val = ''):
        self.worksheet = worksheet
        if type(pos) == str:
            pos = Worksheet.get_int_addr(pos)
        self._row = int(pos[0])
        self._col = int(pos[1])
        self._label = Worksheet.get_addr_int(self._row,self._col)
        self._value = val
        self.format = None
    
    @property
    def row(self):
        """Row number of the cell."""
        return self._row

    @row.setter
    def row(self,row):
        if self.worksheet:
            ncell = self.worksheet.cell(row,self._col)
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
            ncell = self.worksheet.cell(self._row, col)
            self.__dict__.update(ncell.__dict__)
        else:
            self._col = col
    
    @property
    def label(self):
        return self._label

    @label.setter
    def label(self,label):
        if self.worksheet:
            ncell = self.worksheet.acell(label)
            self.__dict__.update(ncell.__dict__)
        else:
            self._label =label
    
    @property
    def value(self):
        return self._value

    @value.setter
    def value(self, value):
        if self.worksheet:
            self.worksheet.update_acell(self.label,value)
            self._value = value;
        else:
            self._value = value;

    def fetch(self):
        if worksheet:
            self._value = self.worksheet.acell(label)

    def __repr__(self):
        return '<%s R%sC%s %s>' % (self.__class__.__name__,
                                   self.row,
                                   self.col,
                                   repr(self.value))
