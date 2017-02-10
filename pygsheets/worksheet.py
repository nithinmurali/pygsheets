# -*- coding: utf-8 -*-.

"""
pygsheets.worksheet
~~~~~~~~~~~~~~~~~~~

This module contains worksheet model

"""

import datetime
import re
import warnings

from .cell import Cell
from .exceptions import (IncorrectCellLabel, CellNotFound, InvalidArgumentValue)
from .utils import numericise_all, format_addr
from .custom_types import *
try:
    from pandas import DataFrame
except ImportError:
    DataFrame = None


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

    # @TODO the update is not instantaious
    def _update_grid(self, force=False):
        """
        update the data grid with values from sheeet
        :param force: force update data grid

        """
        if not self.data_grid or force:
            self.data_grid = self.all_values(returnas='cells', include_empty=False)
        elif not force:
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

    def _get_range(self, start_label, end_label=None):
        """get range in A1 notation, given start and end labels

        """
        if not end_label:
            end_label = start_label
        return self.title + '!' + ('%s:%s' % (format_addr(start_label, 'label'),
                                              format_addr(end_label, 'label')))

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
                label = format_addr(addr, 'label')
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
        return self.get_values(startcell, endcell, returnas='cell')

    def get_value(self, addr):
        """
        value of a cell at given address

        :param addr: cell address as either tuple or label

        """
        addr = format_addr(addr, 'tuple')
        try:
            return self.get_values(addr, addr, include_empty=False)[0][0]
        except KeyError:
            raise CellNotFound

    def get_values(self, start, end, returnas='matrix', majdim='ROWS', include_empty=True):
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
        start = format_addr(start, 'tuple')
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
        return self.get_values((1, 1), (self.rows, self.cols), returnas=returnas,
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

    def get_row(self, row, returnas='matrix', include_empty=True):
        """Returns a list of all values in a `row`.

        Empty cells in this list will be rendered as :const:` `.

        :param include_empty: whether to include empty values
        :param row: index of row
        :param returnas: ('matrix' or 'cell') return as cell objects or just 2d array

        """
        return self.get_values((row, 1), (row, self.cols),
                               returnas=returnas, include_empty=include_empty)[0]

    def get_col(self, col, returnas='matrix', include_empty=True):
        """Returns a list of all values in column `col`.

        Empty cells in this list will be rendered as :const:` `.

        :param include_empty: whether to include empty values
        :param col: index of col
        :param returnas: ('matrix' or 'cell') return as cell objects or just values

        """
        return self.get_values((1, col), (self.rows, col), majdim='COLUMNS',
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
        label = format_addr(addr, 'label')
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
            crange = 'A1:' + str(format_addr((self.rows, self.cols)))
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
            start_r_tuple = format_addr(crange, output='tuple')
            if majordim == 'ROWS':
                end_r_tuple = (start_r_tuple[0]+len(values), start_r_tuple[1]+len(values[0]))
            else:
                end_r_tuple = (start_r_tuple[0] + len(values[0]), start_r_tuple[1] + len(values))
            body['range'] = self._get_range(crange, format_addr(end_r_tuple))
        else:
            body['range'] = self._get_range(*crange.split(':'))
        body['majorDimension'] = majordim
        body['values'] = values
        self.client.sh_update_range(self.spreadsheet.id, body, self.spreadsheet.batch_mode)

    def update_col(self, index, values):
        """update an existing colum with values

        """
        colrange = format_addr((1, index), 'label') + ":" + format_addr((len(values), index), "label")
        self.update_cells(crange=colrange, values=[values], majordim='COLUMNS')

    def update_row(self, index, values):
        """update an existing row with values

        """
        colrange = format_addr((index, 1), 'label') + ':' + format_addr((index, len(values)), 'label')
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
        index -= 1
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
        index -= 1
        if number < 1:
            raise InvalidArgumentValue
        request = {'deleteDimension': {'range': {'sheetId': self.id, 'dimension': 'ROWS',
                                                 'endIndex': (index+number), 'startIndex': index}}}
        self.client.sh_batch_update(self.spreadsheet.id, request, batch=self.spreadsheet.batch_mode)
        self.jsonSheet['properties']['gridProperties']['rowCount'] = self.rows-number

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
        start = format_addr(start, 'tuple')
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

    # @TODO optimize (use datagrid)
    def __iter__(self):
        rows = self.all_values(majdim='ROWS')
        for row in rows:
            yield(row + (self.cols - len(row))*[''])

    # @TODO optimize (use datagrid)
    def __getitem__(self, item):
        if type(item) == int:
            if item >= self.cols:
                raise CellNotFound
            try:
                row = self.all_values()[item]
            except IndexError:
                row = ['']*self.cols
            return row + (self.cols - len(row))*['']
