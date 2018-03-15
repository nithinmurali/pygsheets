# -*- coding: utf-8 -*-.

"""
pygsheets.worksheet
~~~~~~~~~~~~~~~~~~~

This module represents a worksheet within a spreadsheet.

"""

import datetime
import re
from io import open

from .cell import Cell
from .datarange import DataRange
from .exceptions import (IncorrectCellLabel, CellNotFound, InvalidArgumentValue, RangeNotFound)
from .utils import numericise_all, format_addr
from .custom_types import *
try:
    import pandas as pd
except ImportError:
    pd = None


class Worksheet(object):
    """
    A worksheet.

    :param spreadsheet:     Spreadsheet object to which this worksheet belongs to
    :param jsonSheet:       Contains properties to initialize this worksheet.

                      Ref to api details for more info
    """

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
        """The ID of this worksheet."""
        return self.jsonSheet['properties']['sheetId']

    @property
    def index(self):
        """The index of this worksheet"""
        return self.jsonSheet['properties']['index']

    @index.setter
    def index(self, index):
        self.jsonSheet['properties']['index'] = index
        if self._linked:
            self.client.update_sheet_properties(self.spreadsheet.id, self.jsonSheet['properties'], 'index')

    @property
    def title(self):
        """The title of this worksheet."""
        return self.jsonSheet['properties']['title']

    @title.setter
    def title(self, title):
        self.jsonSheet['properties']['title'] = title
        if self._linked:
            self.client.update_sheet_properties(self.spreadsheet.id, self.jsonSheet['properties'], 'title')

    @property
    def url(self):
        """The url of this worksheet."""
        return self.spreadsheet.url+"/edit#gid="+str(self.id)

    @property
    def rows(self):
        """Number of rows active within the sheet. A new sheet contains 1000 rows."""
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
        """Number of columns active within the sheet."""
        return int(self.jsonSheet['properties']['gridProperties']['columnCount'])

    @cols.setter
    def cols(self, col_count):
        if col_count == self.cols:
            return
        self.jsonSheet['properties']['gridProperties']['columnCount'] = int(col_count)
        if self._linked:
            self.client.update_sheet_properties(self.spreadsheet.id, self.jsonSheet['properties'],
                                                'gridProperties/columnCount')

    @property
    def frozen_rows(self):
        """Number of frozen rows."""
        return self.jsonSheet['properties']['gridProperties'].get('frozenRowCount', 0)

    @frozen_rows.setter
    def frozen_rows(self, row_count):
        self.jsonSheet['properties']['gridProperties']['frozenRowCount'] = int(row_count)
        if self._linked:
            self.client.update_sheet_properties(self.spreadsheet.id, self.jsonSheet['properties'],
                                                'gridProperties/frozenRowCount')

    @property
    def frozen_cols(self):
        """Number of frozen columns."""
        return self.jsonSheet['properties']['gridProperties'].get('frozenColumnCount', 0)

    @frozen_cols.setter
    def frozen_cols(self, col_count):
        self.jsonSheet['properties']['gridProperties']['frozenColumnCount'] = int(col_count)
        if self._linked:
            self.client.update_sheet_properties(self.spreadsheet.id, self.jsonSheet['properties'],
                                                'gridProperties/frozenColumnCount')

    def refresh(self, update_grid=False):
        """refresh worksheet data"""
        jsonsheet = self.client.open_by_key(self.spreadsheet.id, returnas='json')
        for sheet in jsonsheet.get('sheets'):
            if sheet['properties']['sheetId'] == self.id:
                self.jsonSheet = sheet
        if update_grid:
            self._update_grid()

    # @TODO the update is not instantaious
    def _update_grid(self, force=False):
        """
        update the data grid (offline) with values from sheeet
        :param force: force update data grid

        """
        if not self.data_grid or force:
            self.data_grid = self.get_all_values(returnas='cells', include_empty=False)
        elif not force:
            updated = datetime.datetime.strptime(self.spreadsheet.updated, '%Y-%m-%dT%H:%M:%S.%fZ')
            if updated > self.grid_update_time:
                self.data_grid = self.get_all_values(returnas='cells', include_empty=False)
        self.grid_update_time = datetime.datetime.utcnow()

    # @TODO update values too (currently only sync worksheet properties)
    def link(self, syncToCloud=True):
        """ Link the spread sheet with cloud, so all local changes
            will be updated instantly, so does all data fetches

            :param  syncToCloud: update the cloud with local changes if set to true
                          update the local copy with cloud if set to false
        """
        if syncToCloud:
            self.client.update_sheet_properties(self.spreadsheet.id, self.jsonSheet['properties'])
        else:
            wks = self.spreadsheet.worksheet(property='id', value=self.id)
            self.jsonSheet = wks.jsonSheet
        self._linked = True

    # @TODO
    def unlink(self):
        """ Unlink the spread sheet with cloud, so all local changes
            will be made on local copy fetched
        """
        self._linked = False

    def sync(self):
        """sync the worksheet to cloud"""
        self.link(True)

    def _get_range(self, start_label, end_label=None, rformat='A1'):
        """get range in A1 notation, given start and end labels

        :param start_label: range start label
        :param end_label: range end label
        :param rformat: can be A1 or GridRange

        """
        if not end_label:
            end_label = start_label
        if rformat == "A1":
            return self.title + '!' + ('%s:%s' % (format_addr(start_label, 'label'),
                                                  format_addr(end_label, 'label')))
        else:
            start_tuple = format_addr(start_label, "tuple")
            end_tuple = format_addr(end_label, "tuple")
            return {"sheetId": self.id, "startRowIndex": start_tuple[0]-1, "endRowIndex": end_tuple[0],
                    "startColumnIndex": start_tuple[1]-1, "endColumnIndex": end_tuple[1]}

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

    def range(self, crange, returnas='cells'):
        """Returns a list of :class:`Cell` objects from specified range.

        :param crange: A string with range value in common format,
                         e.g. 'A1:A5'.
        :param returnas: cann be 'matrix', 'cell', 'range' the correponding type will be returned
        """
        startcell = crange.split(':')[0]
        endcell = crange.split(':')[1]
        return self.get_values(startcell, endcell, returnas=returnas, include_all=True)

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

    def get_values(self, start, end, returnas='matrix', majdim='ROWS', include_empty=True, include_all=False,
                   value_render=ValueRenderOption.FORMATTED):
        """Returns a defined range of values.

        Returns a range of values from start Cell to end Cell. It will fetch these values from remote and then
        processes them. Will return either a simple list of lists, a list of Cell objects or a DataRange object with
        all the cells inside.

        >>> wks.get_values((1,1),(3,3))
        [[u'another look.', u'', u'est'],
         [u'EE 4212', u"it's down there "],
         [u'ee 4210', u'somewhere, let me take ']]

        :param start:           Top left cell as coordinates or label.
        :param end:             Bottom right cell as coordinates or label.
        :param majdim:          The major dimension of the matrix. ('ROWS' or 'COLUMNS')
        :param returnas:        The type to return the fetched values as. ('matrix', 'cell', 'range')
        :param include_empty:   Include empty cells up to 'self.cols'.
        :param include_all:     Include empty rows to return an exact rectangle.
        :param value_render:    The format of the return values.


        :returns 'range':   :class:`DataRange <DataRange>`
                 'cell':    [:class:`Cell <Cell>`]
                 'matrix':  [[ ... ], [ ... ], ...]
        """
        if returnas == 'matrix':
            values = self.client.get_range(self.spreadsheet.id, self._get_range(start, end), majdim.upper(),
                                           value_render=value_render)
            empty_value = ''
        else:
            values = self.client.sh_get_ssheet(self.spreadsheet.id, fields='sheets/data/rowData', include_data=True,
                                               ranges=self._get_range(start, end))
            values = values['sheets'][0]['data'][0].get('rowData', [])
            if include_all:
                values = [x.get('values', []) for x in values]
            else:
                values = [x.get('values', []) for x in values]
                values = list(filter(lambda x: any('effectiveValue' in item for item in x), values))  # skip empty rows
            empty_value = dict()

        if values == [['']] or values == []: values = [[]]

        start = format_addr(start, 'tuple')
        end = format_addr(end, 'tuple')
        if include_all or returnas == 'range':
            max_cols = end[1] - start[1] + 1
            max_rows = end[0] - start[0] + 1
            if majdim == "COLUMNS": max_cols, max_rows = max_rows, max_cols
            matrix = [list(x + [empty_value] * (max_cols - len(x))) for x in values]
            if max_rows > len(matrix):
                matrix.extend([[empty_value]*max_cols]*(max_rows - len(matrix)))
        elif include_empty and len(values) > 0 and values != [[]]:
            if returnas != "matrix":
                matrix = list(filter(lambda x: any('effectiveValue' in item for item in x), values))  # skip empty rows
                # @TODO issue here
            else:
                max_cols = end[1] - start[1] + 1 if majdim == "ROWS" else end[0] - start[0] + 1
                matrix = [list(x + [empty_value] * (max_cols - len(x))) for x in values]
        else:
            if returnas != "matrix":
                matrix = list(filter(lambda x: any('effectiveValue' in item for item in x), values))  # skip empty rows
                for i, row in enumerate(matrix):
                    for j, cell in reversed(list(enumerate(row))):
                        if 'effectiveValue' not in cell:
                            del matrix[i][j]
                        else:
                            break
            else:
                matrix = values

        if matrix == [[]]: return matrix

        if returnas == 'matrix':
            return matrix
        else:
            if majdim == "COLUMNS":
                cells = [[] for x in range(len(matrix[0]))]
                for k in range(len(matrix)):
                    for i in range(len(matrix[k])):
                        cells[i].append(Cell(pos=(start[0]+k, start[1]+i), worksheet=self, cell_data=matrix[k][i]))
            elif majdim == 'ROWS':
                cells = [[] for x in range(len(matrix))]
                for k in range(len(matrix)):
                    for i in range(len(matrix[k])):
                        cells[k].append(Cell(pos=(start[0]+k, start[1]+i), worksheet=self, cell_data=matrix[k][i]))
            else:
                raise InvalidArgumentValue('majdim')

            if returnas.startswith('cell'):
                return cells
            elif returnas == 'range':
                return DataRange(start, format_addr(end, 'label'), worksheet=self, data=cells)

    def get_all_values(self, returnas='matrix', majdim='ROWS', include_empty=True):
        """Returns a list of lists containing all cells' values as strings.

        :param majdim: output as row wise or columwise
        :param returnas: return as list of strings of cell objects
        :param include_empty: whether to include empty values
        :type returnas: 'matrix','cell'

        Example:

        >>> wks.get_all_values()
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
        data = self.get_all_values(returnas='matrix', include_empty=False)
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

    def get_gridrange(self, start, end):
        """
        get a range in gridrange format

        :param start: start adress
        :param end: end adress
        """
        return self._get_range(start, end, "gridrange")

    def update_value(self, addr, val, parse=None):
        """Sets the new value to a cell.

        :param addr: cell address as tuple (row,column) or label 'A1'.
        :param val: New value
        :param parse: if False, values will be stored \
                        as is else as if the user typed them into the UI default is spreadsheet.default_parse

        Example:

        >>> wks.update_value('A1', '42') # this could be 'a1' as well
        <Cell R1C1 "42">
        >>> wks.update_value('A3', '=A1+A2', True)
        <Cell R1C3 "57">
        """
        label = format_addr(addr, 'label')
        body = dict()
        body['range'] = self._get_range(label, label)
        body['majorDimension'] = 'ROWS'
        body['values'] = [[val]]
        parse = parse if parse is not None else self.spreadsheet.default_parse
        self.client.sh_update_range(self.spreadsheet.id, body, self.spreadsheet.batch_mode, parse)

    def update_values(self, crange=None, values=None, cell_list=None, extend=False, majordim='ROWS', parse=None):
        """Updates cell values in batch, it can take either a cell list or a range and values. cell list is only efficient
        for large lists. This will only update the cell values not other properties.

        :param cell_list: List of a :class:`Cell` objects to update with their values
        :param crange: range in format A1:A2 or just 'A1' or even (1,2) end cell will be infered from values
        :param values: matrix of values if range given, if a value is None its unchanged
        :param extend: add columns and rows to the workspace if needed (not for cell list)
        :param majordim: major dimension of given data
        :param parse: if the values should be as if the user typed them into the UI else its stored as is. default is
                      spreadsheet.default_parse
        """
        if cell_list:
            values = [[None for x in range(self.cols)] for y in range(self.rows)]
            min_tuple = [cell_list[0].row, cell_list[0].col]
            max_tuple = [0, 0]
            for cell in cell_list:
                min_tuple[0] = min(min_tuple[0], cell.row)
                min_tuple[1] = min(min_tuple[1], cell.col)
                max_tuple[0] = max(max_tuple[0], cell.row)
                max_tuple[1] = max(max_tuple[1], cell.col)
                try:
                    values[cell.row-1][cell.col-1] = cell.value
                except IndexError:
                        raise CellNotFound(cell)
            values = [row[min_tuple[1]-1:max_tuple[1]] for row in values[min_tuple[0]-1:max_tuple[0]]]
            crange = str(format_addr(tuple(min_tuple))) + ':' + str(format_addr(tuple(max_tuple)))
        elif crange and values:
            if not isinstance(values, list) or not isinstance(values[0], list):
                raise InvalidArgumentValue("values should be a matrix")
        else:
            raise InvalidArgumentValue("provide either cells or values, not both")

        body = dict()
        estimate_size = False
        if type(crange) == str:
            if crange.find(':') == -1:
                estimate_size = True
        elif type(crange) == tuple:
            estimate_size = True
        else:
            raise InvalidArgumentValue('crange')

        if estimate_size:
            start_r_tuple = format_addr(crange, output='tuple')
            max_2nd_dim = max(map(len, values))
            if majordim == 'ROWS':
                end_r_tuple = (start_r_tuple[0]+len(values), start_r_tuple[1]+max_2nd_dim)
            else:
                end_r_tuple = (start_r_tuple[0] + max_2nd_dim, start_r_tuple[1] + len(values))
            body['range'] = self._get_range(crange, format_addr(end_r_tuple))
        else:
            body['range'] = self._get_range(*crange.split(':'))

        if extend:
            self.refresh()
            end_r_tuple = format_addr(str(body['range']).split(':')[-1])
            if self.rows < end_r_tuple[0]:
                self.rows = end_r_tuple[0]-1
            if self.cols < end_r_tuple[1]:
                self.cols = end_r_tuple[1]-1
        body['majorDimension'] = majordim
        body['values'] = values
        parse = parse if parse is not None else self.spreadsheet.default_parse
        self.client.sh_update_range(self.spreadsheet.id, body, self.spreadsheet.batch_mode, parse=parse)

    def update_cells(self, cell_list, fields='*'):
        """
        update cell properties and data from a list of cell obejcts

        :param cell_list: list of cell objects
        :param fields: cell fields to update, in google FieldMask format(see api docs)

        """
        requests = []
        for cell in cell_list:
            request = cell.update(get_request=True, worksheet_id=self.id)
            request['repeatCell']['fields'] = fields
            requests.append(request)

        self.client.sh_batch_update(self.spreadsheet.id, requests, None, True)

    def update_col(self, index, values, row_offset=0):
        """
        update an existing colum with values

        :param index: index of the starting column form where value should be inserted
        :param values: values to be inserted as matrix, column major
        :param row_offset: rows to skip before inserting values

        """
        if type(values[0]) is not list:
            values = [values]
        colrange = format_addr((row_offset+1, index), 'label') + ":" + format_addr((row_offset+len(values[0]),
                                                                                   index+len(values)-1), "label")
        self.update_values(crange=colrange, values=values, majordim='COLUMNS')

    def update_row(self, index, values, col_offset=0):
        """Update an existing row with values

        :param index:       Index of the starting row form where value should be inserted
        :param values:      Values to be inserted as matrix
        :param col_offset:  Columns to skip before inserting values

        """
        if type(values[0]) is not list:
            values = [values]
        colrange = format_addr((index, col_offset+1), 'label') + ':' + format_addr((index+len(values)-1,
                                                                                    col_offset+len(values[0])), 'label')
        self.update_values(crange=colrange, values=values, majordim='ROWS')

    def resize(self, rows=None, cols=None):
        """Resizes the worksheet.

        :param rows: New number of rows.
        :param cols: New number of columns.
        """
        self.unlink()
        trows, tcols = self.rows, self.cols
        self.rows, self.cols = rows, cols
        try:
            self.link()
        except:
            self.rows, self.cols = trows, tcols

    def add_rows(self, rows):
        """Adds new rows to this worksheet.

        :param rows: How many rows to add (integer)
        """
        self.resize(rows=self.rows + rows, cols=self.cols)

    def add_cols(self, cols):
        """Add new colums to this worksheet.

        :param cols: How many columns to add (integer)
        """
        self.resize(cols=self.cols + cols, rows=self.rows)

    def delete_cols(self, index, number=1):
        """Delete 'number' of columns from index.

        :param index:   Index of first column to delete
        :param number:  Number of columns to delete

        """
        index -= 1
        if number < 1:
            raise InvalidArgumentValue('number')
        request = {'deleteDimension': {'range': {'sheetId': self.id, 'dimension': 'COLUMNS',
                                                 'endIndex': (index+number), 'startIndex': index}}}
        self.client.sh_batch_update(self.spreadsheet.id, request, batch=self.spreadsheet.batch_mode)
        self.jsonSheet['properties']['gridProperties']['columnCount'] = self.cols-number

    def delete_rows(self, index, number=1):
        """Delete 'number' of rows from index.

        :param index:   Index of first row to delete
        :param number:  Number of rows to delete
        """
        index -= 1
        if number < 1:
            raise InvalidArgumentValue
        request = {'deleteDimension': {'range': {'sheetId': self.id, 'dimension': 'ROWS',
                                                 'endIndex': (index+number), 'startIndex': index}}}
        self.client.sh_batch_update(self.spreadsheet.id, request, batch=self.spreadsheet.batch_mode)
        self.jsonSheet['properties']['gridProperties']['rowCount'] = self.rows-number

    def insert_cols(self, col, number=1, values=None, inherit=False):
        """Insert new columns after 'col' and initialize all cells with values.

        Increases the number of rows if there are more values in values than rows.

        Reference: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#insertdimensionrequest

        :param col:     Index of the col at which the values will be inserted.
        :param number:  Number of columns to be inserted.
        :param values:  Content to be inserted into new columns.
        :param inherit: New cells will inherit properties from the column to the left (True) or to the right (False).
        """
        request = {'insertDimension': {'inheritFromBefore': inherit,
                                       'range': {'sheetId': self.id, 'dimension': 'COLUMNS',
                                                 'endIndex': (col+number), 'startIndex': col}
                                       }}
        self.client.sh_batch_update(self.spreadsheet.id, request, batch=self.spreadsheet.batch_mode)
        self.jsonSheet['properties']['gridProperties']['columnCount'] = self.cols+number
        if values:
            self.update_col(col+1, values)

    def insert_rows(self, row, number=1, values=None, inherit=False):
        """Insert a new row after 'row' and initialize all cells with values.

        Widens the worksheet if there are more values than columns.

        Reference: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#insertdimensionrequest

        :param row:     Index of the row at which the values will be inserted.
        :param number:  Number of rows to be inserted.
        :param values:  Content to be inserted into new rows.
        :param inherit: New cells will inherit properties from the row above (True) or below (False).
        """
        request = {'insertDimension': {'inheritFromBefore': inherit,
                                       'range': {'sheetId': self.id, 'dimension': 'ROWS',
                                                 'endIndex': (row+number), 'startIndex': row}}}
        self.client.sh_batch_update(self.spreadsheet.id, request, batch=self.spreadsheet.batch_mode)
        self.jsonSheet['properties']['gridProperties']['rowCount'] = self.rows + number
        if values:
            self.update_row(row+1, values)

    def clear(self, start='A1', end=None, fields="userEnteredValue"):
        """Clear all values in worksheet.

        Can be limited to a specific range with start & end.

        Fields specifies which cell properties should be cleared. Use "*" to clear all fields.

        Reference CellData: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#CellData
        Reference FieldMask: https://developers.google.com/protocol-buffers/docs/reference/google.protobuf#google.protobuf.FieldMask

        :param start:   Top left cell label.
        :param end:     Bottom right cell label.
        :param fields:  Comma separated list of field masks.

        """
        if not end:
            end = (self.rows, self.cols)
        request = {"updateCells": {"range": self._get_range(start, end, "GridRange"), "fields": fields}}
        self.client.sh_batch_update(self.spreadsheet.id, request, batch=self.spreadsheet.batch_mode)

    def adjust_column_width(self, start, end=None, pixel_size=100):
        """Set the width of one or more columns.

        :param start:       Index of the first column to be widened.
        :param end:         Index of the last column to be widened.
        :param pixel_size:  New width in pixels.

        """
        if end is None or end <= start:
            end = start + 1

        request = {
          "updateDimensionProperties": {
            "range": {
              "sheetId": self.id,
              "dimension": "COLUMNS",
              "startIndex": start,
              "endIndex": end
            },
            "properties": {
              "pixelSize": pixel_size
            },
            "fields": "pixelSize"
          }
        },

        self.client.sh_batch_update(self.spreadsheet.id, request, batch=self.spreadsheet.batch_mode)

    def auto_resize_dimensions(self, dimension, start, end=None):
        """Automatically resizes one or more dimensions
        :param dimension: COLUMNS or ROWS
        :param start: index of dimension to be resized
        :param end: index of dimension that will be resized
        """
        if end is None or end <= start:
            end = start + 1

        request = {
          "autoResizeDimensions": {
            "dimensions": {
              "sheetId": self.id,
              "dimension": dimension,
              "startIndex": start,
              "endIndex": end
            }
          }
        },
        self.client.sh_batch_update(self.spreadsheet.id, request, batch=self.spreadsheet.batch_mode)
        return self

    def auto_resize_columns(self, start, end=None):
        """Automatically resizes one or more columns
        :param start: index of column to be resized
        :param end: index of column that will be resized
        """
        self.auto_resize_dimensions('COLUMNS', start, end)

    def auto_resize_rows(self, start, end=None):
        """Automatically resizes one or more rows
        :param start: index of row to be resized
        :param end: index of row that will be resized
        """
        self.auto_resize_dimensions('ROWS', start, end)

    def update_dimensions_visibility(self, start, end=None, dimension="ROWS", hidden=True):
        """Hide or show one or more rows or columns.

        :param start:       Index of the first row or column.
        :param end:         Index of the last row or column.
        :param dimension:   'ROWS' or 'COLUMNS'
        :param hidden:      Hide rows or columns
        """
        if end is None or end <= start:
            end = start + 1

        request = {
                      "updateDimensionProperties": {
                          "range": {
                              "sheetId": self.id,
                              "dimension": dimension,
                              "startIndex": start,
                              "endIndex": end
                          },
                          "properties": {
                              "hiddenByUser": hidden
                          },
                          "fields": "hiddenByUser"
                      }
                  },

        self.client.sh_batch_update(self.spreadsheet.id, request, batch=self.spreadsheet.batch_mode)

    def hide_dimensions(self, start, end=None, dimension="ROWS"):
        """Hide one ore more rows or columns.

        :param start:       Index of the first row or column.
        :param end:         Index of the first row or column.
        :param dimension:   'ROWS' or 'COLUMNS'
        """
        self.update_dimensions_visibility(start, end, dimension, hidden=True)

    def show_dimensions(self, start, end=None, dimension="ROWS"):
        """Show one ore more rows or columns.

        :param start:       Index of the first row or column.
        :param end:         Index of the first row or column.
        :param dimension:   'ROWS' or 'COLUMNS'
        """
        self.update_dimensions_visibility(start, end, dimension, hidden=False)

    def hide_rows(self, start, end=None):
        """Hide one or more rows.

        :param start:   Index of the first row.
        :param end:     Index of the last row.
        """
        self.hide_dimensions(start, end, "ROWS")

    def show_rows(self, start, end=None):
        """Show one or more rows.

        :param start:   Index of the first row.
        :param end:     Index of the last row.
        """

        self.show_dimensions(start, end, "ROWS")

    def hide_columns(self, start, end=None):
        """Hide one or more columns.

        :param start:   Index of the first column.
        :param end:     Index of the last column.
        """

        self.hide_dimensions(start, end, "COLUMNS")

    def show_columns(self, start, end=None):
        """Show one or more columns.

        :param start:   Index of the first column.
        :param end:     Index of the last column.
        """
        self.show_dimensions(start, end, "COLUMNS")

    def adjust_row_height(self, start, end=None, pixel_size=100):
        """Adjust the height of one or more rows.

        :param start:       Index of first row to be heightened.
        :param end:         Index of last row to be heightened.
        :param pixel_size:  New height in pixels.
        """
        if end is None or end <= start:
            end = start + 1

        request = {
          "updateDimensionProperties": {
            "range": {
              "sheetId": self.id,
              "dimension": "ROWS",
              "startIndex": start,
              "endIndex": end
            },
            "properties": {
              "pixelSize": pixel_size
            },
            "fields": "pixelSize"
          }
        }
        self.client.sh_batch_update(self.spreadsheet.id, request, batch=self.spreadsheet.batch_mode)

    def append_table(self, start='A1', end=None, values=None, dimension='ROWS', overwrite=False):
        """Append values to the sheet.

        Reference: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/append

        :param start:       Top left cell of the range (requires a label).
        :param end:         Bottom right cell of the range (requires a label).
        :param values:      List of values for the new row or column.
        :param dimension:   Dimension to which the values will be added ('ROWS' or 'COLUMNS')
        :param overwrite:   Overwrite existing values.
        """

        if type(values[0]) != list:
            values = [values]
        if not end:
            end = (self.rows, self.cols)
        body = {"values": values, "majorDimension": dimension}
        self.client.sh_append(self.spreadsheet.id, body=body, rranage=self._get_range(start, end), replace=overwrite)

    def find(self, query, replace=None, force_fetch=True):
        """Finds all cells matching the query.

        Can be fed with a simple string or a compiled regular expression. A string will only match as a full match,
        while the regular expression will be searched within the cell value.

        If replace is not None each matched value will be fully replaced.

        :param query:       A text string or compiled regular expression.
        :param replace:     Replacement string.
        :param force_fetch: Update worksheet from remote, even if unmodified.

        :returns    A list of all cells matching the query.
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

    # @TODO optimize with unlink
    def create_named_range(self, name, start, end):
        """Create a new named range in this worksheet.

        Reference: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#namedrange

        :param name:    Name of the range.
        :param start:   Top left cell address (label or coordinates)
        :param end:     Bottom right cell address (label or coordinates)
        :return :class:`DataRange`
        """
        start = format_addr(start, 'tuple')
        end = format_addr(end, 'tuple')
        request = {"addNamedRange": {
            "namedRange": {
                "name": name,
                "range": {
                    "sheetId": self.id,
                    "startRowIndex": start[0]-1,
                    "endRowIndex": end[0],
                    "startColumnIndex": start[1]-1,
                    "endColumnIndex": end[1],
                }
            }}}
        self.client.sh_batch_update(self.spreadsheet.id, request, batch=self.spreadsheet.batch_mode)
        return DataRange(start, end, self, name)

    def get_named_range(self, name):
        """Get a named range by name.

        Reference: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#namedrange

        :param name:    Name of the named range to be retrieved.
        :return: :class:`DataRange`

        :raises RangeNotFound, if no range matched the name given.
        """
        nrange = [x for x in self.spreadsheet.named_ranges if x.name == name and x.worksheet.id == self.id]
        if len(nrange) == 0:
            self.spreadsheet.update_properties()
            nrange = [x for x in self.spreadsheet.named_ranges if x.name == name and x.worksheet.id == self.id]
            if len(nrange) == 0:
                raise RangeNotFound(name)
        return nrange[0]

    def get_named_ranges(self, name=''):
        """Get named ranges from this worksheet.

        Reference: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#namedrange

        :param name:    Name of the named range to be retrieved, if omitted all ranges are retrieved.
        :return: :class:`DataRange`
        """
        if name == '':
            self.spreadsheet.update_properties()
            nrange = [x for x in self.spreadsheet.named_ranges if x.worksheet.id == self.id]
            return nrange
        else:
            return self.get_named_range(name)

    def delete_named_range(self, name, range_id=''):
        """Delete a named range.

        Reference: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#namedrange

        :param name:        Name of the range.
        :param range_id:    Id of the range (optional)

        """
        if not range_id:
            range_id = self.get_named_ranges(name=name).name_id
        request = {'deleteNamedRange': {
            "namedRangeId": range_id,
        }}
        self.client.sh_batch_update(self.spreadsheet.id, request, batch=self.spreadsheet.batch_mode)
        self.spreadsheet._named_ranges = [x for x in self.spreadsheet._named_ranges if x["namedRangeId"] != range_id]

    def create_protected_range(self, gridrange):
        """Create protected range.

        Reference: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#protectedrange

        :param gridrange:   Grid range of the cells to be protected.
        """
        request = {"addProtectedRange": {
            "protectedRange": {
                "range": gridrange
            },
        }}
        return self.client.sh_batch_update(self.spreadsheet.id, request, None, False)

    def remove_protected_range(self, range_id):
        """Remove protected range.

        Reference: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#protectedrange

        :param range_id:    ID of the protected range.
        """
        request = {"deleteProtectedRange": {
            "protectedRangeId": range_id
        }}
        return self.client.sh_batch_update(self.spreadsheet.id, request, None, False)

    def set_dataframe(self, df, start, copy_index=False, copy_head=True, fit=False, escape_formulae=False, nan='NaN'):
        """Load sheet from Pandas data frame.

        Will load all data contained within the Pandas data frame into this worksheet.
        It will begin filling the worksheet at cell start.

        :param df:              Pandas data frame.
        :param start:           Address of the top left corner where the data should be added.
        :param copy_index:      Copy data frame index (multi index supported).
        :param copy_head:       Copy header data into first row.
        :param fit:             Resize the worksheet to fit all data inside if necessary.
        :param escape_formulae: Any value starting with an equal sign (=), will be prefixed with an apostroph (') to
                                avoid value being interpreted as a formula.
        :param nan:             Value with which NaN values are replaced.
        """
        start = format_addr(start, 'tuple')
        df = df.replace(pd.np.nan, nan)
        values = df.values.tolist()
        (df_rows, df_cols) = df.shape

        if copy_index:
            if isinstance(df.index, pd.MultiIndex):
                for i, indexes in enumerate(df.index):
                    for index_item in reversed(indexes):
                        values[i].insert(0, index_item)
                df_cols += len(df.index[0])
            else:
                for i, val in enumerate(df.index):
                    values[i].insert(0, val)
                df_cols += 1

        if copy_head:
            # If multi index, copy indexes in each level to new row, colum/index names are not copied for now
            if isinstance(df.index, pd.MultiIndex):
                head = [""]*len(df.index[0]) if copy_index else []
                heads = [head[:] for x in df.columns[0]]
                for col_head in df.columns:
                    for i, col_item in enumerate(col_head):
                        heads[i].append(col_item)
                values = heads + values
                df_rows += len(df.columns[0])
            else:
                head = [""] if copy_index else []
                head.extend(df.columns.tolist())
                values.insert(0, head)
                df_rows += 1

        end = format_addr(tuple([start[0]+df_rows, start[1]+df_cols]))

        if fit:
            self.cols = start[1] - 1 + df_cols
            self.rows = start[0] - 1 + df_rows

        # @TODO optimize this
        if escape_formulae:
            for row in values:
                for i in range(len(row)):
                    if type(row[i]) == str and row[i].startswith('='):
                        row[i] = "'" + str(row[i])
        crange = format_addr(start) + ':' + end
        self.update_values(crange=crange, values=values)

    def get_as_df(self, has_header=True, index_colum=None, start=None, end=None, numerize=True, empty_value=''):
        """
        Get the content of this worksheet as a pandas data frame.

        :param has_header:      Interpret first row as data frame header.
        :param index_colum:     Column to use as data frame index (integer).
        :param numerize:        Numerize cell values.
        :param empty_value:     Placeholder value to represent empty cells.
        :param start:           Top left cell to load into data frame. (default: A1)
        :param end:             Bottom right cell to load into data frame. (default: (rows, cols))

        :returns: pandas.Dataframe
        """
        if not pd:
            raise ImportError("pandas")
        if start is not None or end is not None:
            if end is None:
                end = (self.rows, self.cols)
            values = self.get_values(start, end, include_empty=True)
        else:
            values = self.get_all_values(returnas='matrix', include_empty=True)

        if numerize:
            values = [numericise_all(row[:len(values[0])], empty_value) for row in values]

        if has_header:
            keys = values[0]
            values = [row[:len(values[0])] for row in values[1:]]
            df = pd.DataFrame(values, columns=keys)
        else:
            df = pd.DataFrame(values)

        if index_colum:
            if index_colum < 1 or index_colum > len(df.columns):
                raise ValueError("index_column %s not found" % index_colum)
            else:
                df.index = df[df.columns[index_colum - 1]]
                del df[df.columns[index_colum - 1]]
        return df

    def export(self, fformat=ExportType.CSV, filename=None):
        """Export this worksheet in the specified format.

        A worksheet can be exported as CSV, MS_Excel, Open_Office_sheet or PDF.

        :param fformat:     Format of the exported file.
        :param filename:    File name of the exported file (incl. appropriate file ending).
        """
        if fformat is ExportType.CSV:
            import csv
            ifilename = 'worksheet'+str(self.id)+'.csv' if filename is None else filename
            print (ifilename)
            with open(ifilename, 'wt', encoding="utf-8") as f:
                writer = csv.writer(f, lineterminator="\n")
                writer.writerows(self.get_all_values())
        elif isinstance(fformat, ExportType):
            self.client.export(self.spreadsheet.id, fformat, filename=filename)
        else:
            raise InvalidArgumentValue("Fformat needs to be a member of ExportType Enum")

    def copy_to(self, spreadsheet_id):
        """Copy this worksheet to the specified spreadsheet.

        Reference: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.sheets/copyTo

        :param spreadsheet_id:  Id of the spreadsheet this worksheet should be copied to.
        """
        self.client.sh_copy_worksheet(self.spreadsheet.id, self.id, spreadsheet_id)

    def __eq__(self, other):
        return self.id == other.id and self.spreadsheet == other.spreadsheet

    # @TODO optimize (use datagrid)
    def __iter__(self):
        rows = self.get_all_values(majdim='ROWS')
        for row in rows:
            yield(row + (self.cols - len(row))*[''])

    # @TODO optimize (use datagrid)
    def __getitem__(self, item):
        if type(item) == int:
            if item >= self.cols:
                raise CellNotFound
            try:
                row = self.get_all_values()[item]
            except IndexError:
                row = ['']*self.cols
            return row + (self.cols - len(row))*['']
