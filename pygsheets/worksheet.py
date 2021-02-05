# -*- coding: utf-8 -*-.

"""
pygsheets.worksheet
~~~~~~~~~~~~~~~~~~~

This module represents a worksheet within a spreadsheet.

"""

import datetime
import re
import warnings
import logging

from pygsheets.cell import Cell
from pygsheets.datarange import DataRange
from pygsheets.address import GridRange, Address
from pygsheets.exceptions import (CellNotFound, InvalidArgumentValue, RangeNotFound)
from pygsheets.utils import numericise_all, format_addr, fullmatch, batchable, allow_gridrange
from pygsheets.custom_types import *
from pygsheets.chart import Chart
from pygsheets.developer_metadata import DeveloperMetadataLookupDataFilter, DeveloperMetadata
try:
    import pandas as pd
except ImportError:
    pd = None


_warning_mesage = "this {} is deprecated. Use {} instead"
_deprecated_keyword_mapping = {
    'include_empty': 'include_tailing_empty',
    'include_all': 'include_tailing_empty_rows',
}


class Worksheet(object):
    """
    A worksheet.

    :param spreadsheet:     Spreadsheet object to which this worksheet belongs to
    :param jsonSheet:       Contains properties to initialize this worksheet.

                      Ref to api details for more info
    """

    def __init__(self, spreadsheet, jsonSheet):
        self.logger = logging.getLogger(__name__)
        self.spreadsheet = spreadsheet
        self.client = spreadsheet.client
        self._linked = True
        self.jsonSheet = jsonSheet
        self.data_grid = None  # for storing sheet data while unlinked
        self._func_calls = []
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
            self.client.sheet.update_sheet_properties_request(self.spreadsheet.id, self.jsonSheet['properties'], 'index')

    @property
    def title(self):
        """The title of this worksheet."""
        return self.jsonSheet['properties']['title']

    @title.setter
    def title(self, title):
        self.jsonSheet['properties']['title'] = title
        if self._linked:
            self.client.sheet.update_sheet_properties_request(self.spreadsheet.id, self.jsonSheet['properties'], 'title')

    @property
    def hidden(self):
        """Mark the worksheet as hidden."""
        return self.jsonSheet['properties'].get('hidden', False)

    @hidden.setter
    def hidden(self, hidden):
        self.jsonSheet['properties']['hidden'] = hidden
        if self._linked:
            self.client.sheet.update_sheet_properties_request(self.spreadsheet.id, self.jsonSheet['properties'], 'hidden')

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
            self.client.sheet.update_sheet_properties_request(self.spreadsheet.id, self.jsonSheet['properties'],
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
            self.client.sheet.update_sheet_properties_request(self.spreadsheet.id, self.jsonSheet['properties'],
                                                              'gridProperties/columnCount')

    @property
    def frozen_rows(self):
        """Number of frozen rows."""
        return self.jsonSheet['properties']['gridProperties'].get('frozenRowCount', 0)

    @frozen_rows.setter
    def frozen_rows(self, row_count):
        self.jsonSheet['properties']['gridProperties']['frozenRowCount'] = int(row_count)
        if self._linked:
            self.client.sheet.update_sheet_properties_request(self.spreadsheet.id, self.jsonSheet['properties'],
                                                              'gridProperties/frozenRowCount')

    @property
    def frozen_cols(self):
        """Number of frozen columns."""
        return self.jsonSheet['properties']['gridProperties'].get('frozenColumnCount', 0)

    @frozen_cols.setter
    def frozen_cols(self, col_count):
        self.jsonSheet['properties']['gridProperties']['frozenColumnCount'] = int(col_count)
        if self._linked:
            self.client.sheet.update_sheet_properties_request(self.spreadsheet.id, self.jsonSheet['properties'],
                                                              'gridProperties/frozenColumnCount')

    @property
    def linked(self):
        """If the sheet is linked."""
        return self._linked

    def refresh(self, update_grid=False):
        """refresh worksheet data"""
        jsonsheet = self.client.open_as_json(self.spreadsheet.id)
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
            self.data_grid = self.get_all_values(returnas='cells-unlinked', include_tailing_empty=True, include_tailing_empty_rows=True)
        elif not force:
            updated = datetime.datetime.strptime(self.spreadsheet.updated, '%Y-%m-%dT%H:%M:%S.%fZ')
            if updated > self.grid_update_time:
                self.data_grid = self.get_all_values(returnas='cells-unlinked', include_tailing_empty=True, include_tailing_empty_rows=True)
        self.grid_update_time = datetime.datetime.utcnow()

    def link(self, syncToCloud=True):
        """ Link the spreadsheet with cloud, so all local changes
            will be updated instantly. Data fetches will work only in linked sheet.
            All the data update calls made when sheet is unlinked will be replayed on linking.

            :param  syncToCloud: update the cloud with local changes (data_grid), cached update calls if set to true
                          update the local copy with cloud if set to false
        """
        self._linked = True
        if syncToCloud:
            self.client.sheet.update_sheet_properties_request(self.spreadsheet.id, self.jsonSheet['properties'], '*')
            if self.data_grid:
                tmp_data_grid = [item for sublist in self.data_grid for item in sublist]  # flatten the list
                self.update_cells(tmp_data_grid)

            # call all saved function calls
            for func, fargs in self._func_calls:
                func(*fargs[0], **fargs[1])
            self._func_calls = []
        else:
            wks = self.spreadsheet.worksheet(property='id', value=self.id)
            self.jsonSheet = wks.jsonSheet

    # TODO change to False @nextRelease
    def unlink(self, save_grid=True):
        """ Unlink the spreadsheet with cloud, so that any changes made wont be updated instantaneously.
            All the data update calls are cached and will be called once the sheet is linked again.

            .. warning::
             After unlinking, functions which return data won't work.

        """
        if save_grid:
            self._update_grid()
        self._linked = False
        self._func_calls = []

    def sync(self):
        """
        sync the worksheet (datagrid, and worksheet properties) to cloud

        """
        self.link(True)
        self.logger.warn("sync not implimented")

    def _get_range(self, start_label, end_label=None, rformat='A1'):
        """get range in A1 notation, given start and end labels

        :param start_label: range start label
        :param end_label: range end label
        :param rformat: can be A1 or GridRange

        """
        grange = GridRange(worksheet=self, start=start_label, end=end_label)
        if rformat == "A1":
            return grange.label
        else:
            return grange.to_json()

    def cell(self, addr):
        """
        Returns cell object at given address.

        :param addr: cell address as either tuple (row, col) or cell label 'A1' or Adress

        :returns: an instance of a :class:`Cell`

        Example:

        >>> wks.cell((1,1))
        <Cell R1C1 "I'm cell A1">
        >>> wks.cell('A1')
        <Cell R1C1 "I'm cell A1">

        """
        if not self._linked: return False
        addr = Address(addr)
        try:
            val = self.client.get_range(self.spreadsheet.id, self._get_range(addr, addr), 'ROWS')[0][0]
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
        :param returnas: can be 'matrix', 'cell', 'range' the corresponding type will be returned
        """
        startcell = crange.split(':')[0]
        endcell = crange.split(':')[1]
        return self.get_values(startcell, endcell, returnas=returnas, include_tailing_empty_rows=True)

    def get_value(self, addr, value_render=ValueRenderOption.FORMATTED_VALUE):
        """
        value of a cell at given address

        :param addr: cell address as either tuple or label
        :param value_render: how the output values should rendered. `api docs <https://developers.google.com/sheets/api/reference/rest/v4/ValueRenderOption>`__

        """
        addr = format_addr(addr, 'tuple')
        try:
            return self.get_values(addr, addr, returnas='matrix', include_tailing_empty=True,
                                   include_tailing_empty_rows=True, value_render=value_render)[0][0]
        except KeyError:
            raise CellNotFound

    @allow_gridrange
    def get_values(self, start, end, returnas='matrix', majdim='ROWS', include_tailing_empty=True,
                   include_tailing_empty_rows=False, value_render=ValueRenderOption.FORMATTED_VALUE,
                   date_time_render_option=DateTimeRenderOption.SERIAL_NUMBER, grange=None, **kwargs):
        """
        Returns a range of values from start cell to end cell. It will fetch these values from remote and then
        processes them. Will return either a simple list of lists, a list of Cell objects or a DataRange object with
        all the cells inside.

        :param start: Top left position as tuple or label
        :param end: Bottom right position as tuple or label
        :param grange: give range as grid range object, object of :class:`GridRange`
        :param majdim: The major dimension of the matrix. ('ROWS') ( 'COLUMNS' not implemented )
        :param returnas: The type to return the fetched values as. ('matrix', 'cell', 'range')
        :param include_tailing_empty: whether to include empty trailing cells/values after last non-zero value in a row
        :param include_tailing_empty_rows: whether to include tailing rows with no values; if include_tailing_empty is false,
                    will return unfilled list for each empty row, else will return rows filled with empty cells
        :param value_render: how the output values should rendered. `api docs <https://developers.google.com/sheets/api/reference/rest/v4/ValueRenderOption>`__
        :param date_time_render_option:     How dates, times, and durations should be represented in the output.
                                            This is ignored if `valueRenderOption` is `FORMATTED_VALUE`. The default
                                            dateTime render option is [`DateTimeRenderOption.SERIAL_NUMBER`].

        :returns: 'range':   :class:`DataRange <DataRange>`
                 'cell':    [:class:`Cell <Cell>`]
                 'matrix':  [[ ... ], [ ... ], ...]
                 append '-unlinked' to get unlinked objects
        """
        include_tailing_empty = kwargs.get('include_empty', include_tailing_empty)
        include_tailing_empty_rows = kwargs.get('include_all', include_tailing_empty_rows)

        return_unliked, return_worksheet = returnas.endswith('-unlinked'), self
        returnas = returnas.split('-')[0]
        if return_unliked:
            return_worksheet = None

        _deprecated_keywords = ['include_empty', 'include_all']
        for key in list(kwargs):
            if key in _deprecated_keywords:
                warnings.warn(
                    'The argument {} is deprecated. Use {} instead.'.format(key, _deprecated_keyword_mapping[key])
                    , category=DeprecationWarning)
                kwargs.pop(key, None)

        if not self._linked: return False

        majdim = majdim.upper()
        if majdim.startswith('COL'):
            majdim = "COLUMNS"
        prev_include_tailing_empty_rows, prev_include_tailing_empty = True, True

        if not grange:
            grange = GridRange(worksheet=self, start=start, end=end)
        grange.set_worksheet(self)

        # fetch the values
        if returnas == 'matrix':
            values = self.client.get_range(self.spreadsheet.id, grange.label, majdim,
                                           value_render_option=value_render,
                                           date_time_render_option=date_time_render_option, **kwargs)
            empty_value = ''
        else:
            values = self.client.sheet.get(self.spreadsheet.id, fields='sheets/data/rowData',
                                           includeGridData=True,
                                           ranges=grange.label)
            values = values['sheets'][0]['data'][0].get('rowData', [])
            values = [x.get('values', []) for x in values]
            empty_value = dict({"effectiveValue": {"stringValue": ""}})

            # Cells are always returned in row major form from api. lets keep them such that for now
            # So lets first make a complete rectangle and cleanup later
            if majdim == "COLUMNS":
                prev_include_tailing_empty = include_tailing_empty
                prev_include_tailing_empty_rows = include_tailing_empty_rows
                include_tailing_empty = True
                include_tailing_empty_rows = True

        if returnas == 'range':  # need perfect rectangle
            include_tailing_empty = True
            include_tailing_empty_rows = True

        if values == [['']] or values == []: values = [[]]

        # cleanup and re-structure the values
        start, end = [x.index for x in grange.get_bounded_indexes()]

        max_rows = end[0] - start[0] + 1
        max_cols = end[1] - start[1] + 1
        if majdim == "COLUMNS" and returnas == "matrix":
            max_cols = end[0] - start[0] + 1
            max_rows = end[1] - start[1] + 1

        # restructure values according to params
        if include_tailing_empty_rows and (max_rows-len(values)) > 0:  # append empty rows in end
            values.extend([[]]*(max_rows-len(values)))
        if include_tailing_empty:  # append tailing cells in rows
            values = [list(x + [empty_value] * (max_cols - len(x))) for x in values]
        elif returnas != 'matrix':
            for i, row in enumerate(values):
                for j, cell in reversed(list(enumerate(row))):
                    if 'effectiveValue' not in cell:
                        del values[i][j]
                    else:
                        break

        if values == [[]] or values == [['']]: return values

        if returnas == 'matrix':
            return values
        else:

            # Now the cells are complete rectangle, convert to columen major form and remove
            # the excess cells based on the params saved
            if majdim == "COLUMNS":
                values = list(map(list, zip(*values)))
                for i in range(len(values) - 1, -1, -1):
                    if not prev_include_tailing_empty_rows:
                        if not any((item.get("effectiveValue", {}).get("stringValue", "-1") != "" and "effectiveValue" in item) for item in values[i]):
                            del values[i]
                            continue
                    if not prev_include_tailing_empty:
                        for k in range(len(values[i])-1, -1, -1):
                            if values[i][k].get("effectiveValue", {}).get("stringValue", "-1") != "" and "effectiveValue" in values[i][k]:
                                break
                            else:
                                del values[i][k]
                max_cols = end[0] - start[0] + 1
                max_rows = end[1] - start[1] + 1

            cells = []
            for k in range(len(values)):
                cells.extend([[]])
                for i in range(len(values[k])):
                    if majdim == "ROWS":
                        cells[-1].append(Cell(pos=(start[0]+k, start[1]+i), worksheet=return_worksheet, cell_data=values[k][i]))
                    else:
                        cells[-1].append(Cell(pos=(start[0]+i, start[1]+k), worksheet=return_worksheet, cell_data=values[k][i]))

            if cells == []: cells = [[]]

            if returnas.startswith('cell'):
                return cells
            elif returnas == 'range':
                dd = DataRange(start, format_addr(end, 'label'), worksheet=return_worksheet, data=cells)
                return dd

    def get_values_batch(self, ranges, majdim='ROWS', value_render=ValueRenderOption.FORMATTED_VALUE,
                         date_time_render_option=DateTimeRenderOption.SERIAL_NUMBER,  **kwargs):
        """
        Returns a range of values from start cell to end cell. It will fetch these values from remote and then
        processes them. Will return either a simple list of lists, a list of Cell objects or a DataRange object with
        all the cells inside.

        :param ranges: list of ranges to get data as - objects of :class:`GridRange`, or tuple (with start and end)
                        or A1 notation string or dict in GridRange format
        :param majdim: The major dimension of the matrix. ('ROWS') ( 'COLMUNS' not implemented )
        :param value_render: refer get_values
        :param date_time_render_option: refer get_values

        Example:

        >>> wks.get_values_batch( ['A1:A2', 'C1:C2'] )
          [ [['3'], ['4']], [['c']]]
        >>> wks.get_values_batch( [('1', None), ('5', None)] )
        [ <values of row 1>, <values of row 5> ]
        >>> wks.get_values_batch( [('A1', 'B2'), ('5', None), 'Sheet1!D1:F10', 'A'])
        [ <values list of lists> ]
        """
        labels = [GridRange.create(x, self).label for x in ranges]
        values = self.client.get_range(self.spreadsheet.id, value_ranges=labels, major_dimension=majdim,
                                       value_render_option=value_render,
                                       date_time_render_option=date_time_render_option, **kwargs)

        return values

    def get_all_values(self, returnas='matrix', majdim='ROWS', include_tailing_empty=True,
                       include_tailing_empty_rows=True, **kwargs):
        """Returns a list of lists containing all cells' values as strings.

        :param majdim: output as row wise or columwise
        :param returnas: return as list of strings of cell objects
        :param include_tailing_empty: whether to include empty trailing cells/values after last non-zero value
        :param include_tailing_empty_rows: whether to include rows with no values; if include_tailing_empty is false,
                    will return unfilled list for each empty row, else will return rows filled with empty string
        :param kwargs: all parameters of :meth:`pygsheets.Worksheet.get_values`

        :type returnas: 'matrix','cell', 'range

        Example:

        >>> wks.get_all_values()
        [[u'another look.', u'', u'est'],
         [u'EE 4212', u"it's down there "],
         [u'ee 4210', u'somewhere, let me take ']]
        """
        return self.get_values(None, None, returnas=returnas, majdim=majdim,
                               include_tailing_empty=include_tailing_empty,
                               include_tailing_empty_rows=include_tailing_empty_rows, **kwargs)

    def get_all_records(self, empty_value='', head=1, majdim='ROWS', numericise_data=True, **kwargs):
        """
        Returns a list of dictionaries, all of them having

            - the contents of the spreadsheet's with the head row as keys, \
            And each of these dictionaries holding
            - the contents of subsequent rows of cells as values.

        Cell values are numericised (strings that can be read as ints
        or floats are converted).

        :param empty_value: determines empty cell's value
        :param head: determines wich row to use as keys, starting from 1
            following the numeration of the spreadsheet.
        :param majdim: ROW or COLUMN major form
        :param numericise_data: determines if data is converted to numbers or left as string
        :param kwargs: all parameters of :meth:`pygsheets.Worksheet.get_values`

        :returns: a list of dict with header column values as head and rows as list

        .. warning::
            Will work nicely only if there is a single table in the sheet

        """
        if not self._linked: return False

        idx = head - 1
        data = self.get_all_values(returnas='matrix', include_tailing_empty=False, include_tailing_empty_rows=False,
                                   majdim=majdim, **kwargs)
        keys = data[idx]
        num_keys = len(keys)
        values = []
        for row in data[idx+1:]:
            if len(row) < num_keys:
                row.extend([""]*(num_keys-len(row)))
            elif len(row) > num_keys:
                row = row[:num_keys]
            if numericise_data:
                values.append(numericise_all(row, empty_value))
            else:
                values.append(row)

        return [dict(zip(keys, row)) for row in values]

    def get_row(self, row, returnas='matrix', include_tailing_empty=True, **kwargs):
        """Returns a list of all values in a `row`.

        Empty cells in this list will be rendered as empty strings .

        :param include_tailing_empty: whether to include empty trailing cells/values after last non-zero value
        :param row: index of row
        :param kwargs: all parameters of :meth:`pygsheets.Worksheet.get_values`

        :param returnas: ('matrix', 'cell', 'range') return as cell objects or just 2d array or range object

        """
        row = self.get_values((row, 1), (row, None), returnas=returnas,
                              include_tailing_empty=include_tailing_empty, include_tailing_empty_rows=True, **kwargs)
        if returnas == 'range':
            return row
        else:
            return row[0]

    def get_col(self, col, returnas='matrix', include_tailing_empty=True, **kwargs):
        """Returns a list of all values in column `col`.

        Empty cells in this list will be rendered as :const:` ` .

        :param include_tailing_empty: whether to include empty trailing cells/values after last non-zero value
        :param col: index of col
        :param kwargs: all parameters of :meth:`pygsheets.Worksheet.get_values`

        :param returnas: ('matrix' or 'cell' or 'range') return as cell objects or just values

        """
        col = self.get_values((1, col), (None, col), returnas=returnas, majdim='COLUMNS',
                              include_tailing_empty=include_tailing_empty, include_tailing_empty_rows=True, **kwargs)
        if returnas == 'range':
            return col
        else:
            return col[0]

    def get_gridrange(self, start, end):
        """
        get a range in gridrange format

        :param start: start address
        :param end: end address
        """
        return self._get_range(start, end, "gridrange")

    @batchable
    def update_cell(self, **kwargs):
        warnings.warn(_warning_mesage.format("method", "update_value"), category=DeprecationWarning)
        self.update_value(**kwargs)

    @batchable
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
        if not self._linked:
            self._func_calls.append(())
            return False

        label = format_addr(addr, 'label')
        body = dict()
        body['range'] = self._get_range(label, label)
        body['majorDimension'] = 'ROWS'
        body['values'] = [[val]]
        parse = parse if parse is not None else self.spreadsheet.default_parse
        self.client.sheet.values_batch_update(self.spreadsheet.id, body, parse)

    @batchable
    def update_values(self, crange=None, values=None, cell_list=None, extend=False, majordim='ROWS', parse=None):
        """Updates a range cell values, it can take either a cell list or a range and its values. cell list is only efficient
        for small lists. This will only update the cell values not other properties.

        :param cell_list: List of a :class:`Cell` objects to update with their values. If you pass a matrix to this,\
        then it is assumed that the matrix is continous (range), and will just update values based on label of top \
        left and bottom right cells.

        :param crange: range in format A1:A2 or just 'A1' or even (1,2) end cell will be inferred from values
        :param values: matrix of values if range given, if a value is None its unchanged
        :param extend: add columns and rows to the workspace if needed (not for cell list)
        :param majordim: major dimension of given data
        :param parse: if the values should be as if the user typed them into the UI else its stored as is. Default is
                      spreadsheet.default_parse
        """
        if not self._linked: return False

        if cell_list:
            if type(cell_list[0]) is list:
                values = []
                for row in cell_list:
                    tmp_row = []
                    for col in cell_list:
                        tmp_row.append(cell_list[row][col].value)
                    values.append(tmp_row)
                crange = cell_list[0][0].label + ':' + cell_list[-1][-1].label
            else:
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
        self.client.sheet.values_batch_update(self.spreadsheet.id, body, parse)

    @batchable
    def update_values_batch(self, ranges, values, majordim='ROWS', parse=None):
        """
        update multiple ranges of values in a single call.

        :param ranges: list of addresses of the range. can be GridRange, label, tuple, etc
        :param values: list of values corresponding to ranges, should be list of matrices
        :param majordim: major dimension of values provided. 'ROWS' or 'COLUMNS'
        :param parse: if the values should be as if the user typed them into the UI else its stored as is. Default is
                      spreadsheet.default_parse

        Example:
        >>> wks.update_values_batch(['A1:A2', 'B1:B2'], [[[1],[2]], [[3],[4]]])
        >>> wks.get_values_batch(['A1:A2', 'B1:B2'])
        [[['1'], ['2']], [['3'], ['4']]]

        >>> wks.update_values_batch([((1,1), (2,1)), 'B1:B2'], [[[1,2]], [[3,4]]], 'COLUMNS')
        >>> wks.get_values_batch(['A1:A2', 'B1:B2'])
        [[['1'], ['2']], [['3'], ['4']]]

        """
        ranges = [GridRange.create(x, self).label for x in ranges]
        if not isinstance(values, list):
            raise InvalidArgumentValue('values is not a list')
        if len(ranges) != len(values):
            raise InvalidArgumentValue('number of ranges and values should match')
        # TODO update to enable filters
        data = [
            {'dataFilter': {'a1Range': x[0]}, 'values': x[1], 'majorDimension': majordim} for x in zip(ranges, values)
        ]
        self.client.sheet.values_batch_update_by_data_filter(self.spreadsheet.id, data, parse)

    @batchable
    def update_cells_prop(self, **kwargs):
        warnings.warn(_warning_mesage.format('method', 'update_cells'), category=DeprecationWarning)
        self.update_cells(**kwargs)

    @batchable
    def update_cells(self, cell_list, fields='*'):
        """
        update cell properties and data from a list of cell objects

        :param cell_list: list of cell objects
        :param fields: cell fields to update, in google `FieldMask format <https://developers.google.com/protocol-buffers/docs/reference/google.protobuf#google.protobuf.FieldMask>`_

        """
        if not self._linked: return False

        if fields == 'userEnteredValue':
            pass  # TODO Create a grid and put values there and update

        requests = []
        for cell in cell_list:
            request = cell.update(get_request=True, worksheet_id=self.id)
            request['repeatCell']['fields'] = fields
            requests.append(request)

        self.client.sheet.batch_update(self.spreadsheet.id, requests)

    @batchable
    def update_col(self, index, values, row_offset=0):
        """
        update an existing colum with values

        :param index: index of the starting column form where value should be inserted
        :param values: values to be inserted as matrix, column major
        :param row_offset: rows to skip before inserting values

        """
        if not self._linked: return False

        if type(values[0]) is not list:
            values = [values]
        colrange = format_addr((row_offset+1, index), 'label') + ":" + format_addr((row_offset+len(values[0]),
                                                                                   index+len(values)-1), "label")
        self.update_values(crange=colrange, values=values, majordim='COLUMNS')

    @batchable
    def update_row(self, index, values, col_offset=0):
        """Update an existing row with values

        :param index:       Index of the starting row form where value should be inserted
        :param values:      Values to be inserted as matrix
        :param col_offset:  Columns to skip before inserting values

        """
        if not self._linked: return False

        if type(values[0]) is not list:
            values = [values]
        colrange = format_addr((index, col_offset+1), 'label') + ':' + format_addr((index+len(values)-1,
                                                                                    col_offset+len(values[0])), 'label')
        self.update_values(crange=colrange, values=values, majordim='ROWS')

    @batchable
    def resize(self, rows=None, cols=None):
        """Resizes the worksheet.

        :param rows: New number of rows.
        :param cols: New number of columns.
        """
        trows, tcols = self.rows, self.cols
        try:
            self.rows, self.cols = rows or trows, cols or tcols
        except:
            self.logger.error("couldnt resize the sheet to " + str(rows) + ',' + str(cols))
            self.rows, self.cols = trows, tcols
            raise

    @batchable
    def add_rows(self, rows):
        """Adds new rows to this worksheet.

        :param rows: How many rows to add (integer)
        """
        self.resize(rows=self.rows + rows, cols=self.cols)

    @batchable
    def add_cols(self, cols):
        """Add new columns to this worksheet.

        :param cols: How many columns to add (integer)
        """
        self.resize(cols=self.cols + cols, rows=self.rows)

    @batchable
    def delete_cols(self, index, number=1):
        """Delete 'number' of columns from index.

        :param index:   Index of first column to delete
        :param number:  Number of columns to delete

        """
        if not self._linked: return False

        index -= 1
        if number < 1:
            raise InvalidArgumentValue('number')
        request = {'deleteDimension': {'range': {'sheetId': self.id, 'dimension': 'COLUMNS',
                                                 'endIndex': (index+number), 'startIndex': index}}}
        self.client.sheet.batch_update(self.spreadsheet.id, request)
        self.jsonSheet['properties']['gridProperties']['columnCount'] = self.cols-number

    @batchable
    def delete_rows(self, index, number=1):
        """Delete 'number' of rows from index.

        :param index:   Index of first row to delete
        :param number:  Number of rows to delete
        """
        if not self._linked: return False

        index -= 1
        if number < 1:
            raise InvalidArgumentValue
        request = {'deleteDimension': {'range': {'sheetId': self.id, 'dimension': 'ROWS',
                                                 'endIndex': (index+number), 'startIndex': index}}}
        self.client.sheet.batch_update(self.spreadsheet.id, request)
        self.jsonSheet['properties']['gridProperties']['rowCount'] = self.rows-number

    @batchable
    def insert_cols(self, col, number=1, values=None, inherit=False):
        """Insert new columns after 'col' and initialize all cells with values. Increases the
        number of rows if there are more values in values than rows.

        Reference: `insert request <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#insertdimensionrequest>`_

        :param col:     Index of the col at which the values will be inserted.
        :param number:  Number of columns to be inserted.
        :param values:  Content to be inserted into new columns.
        :param inherit: New cells will inherit properties from the column to the left (True) or to the right (False).
        """
        if not self._linked: return False

        request = {'insertDimension': {'inheritFromBefore': inherit,
                                       'range': {'sheetId': self.id, 'dimension': 'COLUMNS',
                                                 'endIndex': (col+number), 'startIndex': col}
                                       }}
        self.client.sheet.batch_update(self.spreadsheet.id, request)
        self.jsonSheet['properties']['gridProperties']['columnCount'] = self.cols+number
        if values:
            self.update_col(col+1, values)

    @batchable
    def insert_rows(self, row, number=1, values=None, inherit=False):
        """Insert a new row after 'row' and initialize all cells with values.

        Widens the worksheet if there are more values than columns.

        Reference: `insert request`_

        :param row:     Index of the row at which the values will be inserted.
        :param number:  Number of rows to be inserted.
        :param values:  Content to be inserted into new rows.
        :param inherit: New cells will inherit properties from the row above (True) or below (False).
        """
        if not self._linked: return False

        request = {'insertDimension': {'inheritFromBefore': inherit,
                                       'range': {'sheetId': self.id, 'dimension': 'ROWS',
                                                 'endIndex': (row+number), 'startIndex': row}}}
        self.client.sheet.batch_update(self.spreadsheet.id, request)
        self.jsonSheet['properties']['gridProperties']['rowCount'] = self.rows + number
        if values:
            self.update_row(row+1, values)

    @batchable
    @allow_gridrange
    def clear(self, start='A1', end=None, fields="userEnteredValue"):
        """Clear all values in worksheet. Can be limited to a specific range with start & end.

        Fields specifies which cell properties should be cleared. Use "*" to clear all fields.

        Reference:

            - `CellData Api object <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#CellData>`_
            -  `FieldMask Api object <https://developers.google.com/protocol-buffers/docs/reference/google.protobuf#google.protobuf.FieldMask>`_

        :param start:   Top left cell label.
        :param end:     Bottom right cell label.
        :param fields:  Comma separated list of field masks.

        """
        if not self._linked: return False

        if not end:
            end = (self.rows, self.cols)
        grange = GridRange(worksheet=self, start=start, end=end)
        request = {"updateCells": {"range": grange.to_json(), "fields": fields}}
        self.client.sheet.batch_update(self.spreadsheet.id, request)

    @batchable
    def adjust_column_width(self, start, end=None, pixel_size=None):
        """Set the width of one or more columns.

        :param start:       Index of the first column to be widened.
        :param end:         Index of the last column to be widened.
        :param pixel_size:  New width in pixels or None to set width automatically based on the size of the column content.

        """
        if not self._linked: return False

        if end is None or end <= start:
            end = start

        start -= 1

        if pixel_size:
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
        else:
            request = {
              "autoResizeDimensions": {
                "dimensions": {
                  "sheetId": self.id,
                  "dimension": "COLUMNS",
                  "startIndex": start,
                  "endIndex": end
                }
              }
            },

        self.client.sheet.batch_update(self.spreadsheet.id, request)

    @batchable
    def update_dimensions_visibility(self, start, end=None, dimension="ROWS", hidden=True):
        """Hide or show one or more rows or columns.

        :param start:       Index of the first row or column.
        :param end:         Index of the last row or column.
        :param dimension:   'ROWS' or 'COLUMNS'
        :param hidden:      Hide rows or columns
        """
        if not self._linked: return False

        if end is None or end <= start:
            end = start
        start -= 1

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

        self.client.sheet.batch_update(self.spreadsheet.id, request)

    @batchable
    def hide_dimensions(self, start, end=None, dimension="ROWS"):
        """Hide one ore more rows or columns.

        :param start:       Index of the first row or column.
        :param end:         Index of the last row or column.
        :param dimension:   'ROWS' or 'COLUMNS'
        """
        self.update_dimensions_visibility(start, end, dimension, hidden=True)

    @batchable
    def show_dimensions(self, start, end=None, dimension="ROWS"):
        """Show one ore more rows or columns.

        :param start:       Index of the first row or column.
        :param end:         Index of the last row or column.
        :param dimension:   'ROWS' or 'COLUMNS'
        """
        self.update_dimensions_visibility(start, end, dimension, hidden=False)

    @batchable
    def adjust_row_height(self, start, end=None, pixel_size=None):
        """Adjust the height of one or more rows.

        :param start:       Index of first row to be heightened.
        :param end:         Index of last row to be heightened.
        :param pixel_size:  New height in pixels or None to set height automatically based on the size of the row content.
        """
        if not self._linked: return False

        if end is None or end <= start:
            end = start
        start -= 1

        if pixel_size:
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
        else:
            request = {
              "autoResizeDimensions": {
                "dimensions": {
                  "sheetId": self.id,
                  "dimension": "ROWS",
                  "startIndex": start,
                  "endIndex": end
                }
              }
            },

        self.client.sheet.batch_update(self.spreadsheet.id, request)

    @batchable
    def append_table(self, values, start='A1', end=None, dimension='ROWS', overwrite=False, **kwargs):
        """Append a row or column of values to an existing table in the sheet.
        The input range is used to search for existing data and find a "table" within that range.
        Values will be appended to the next row of the table, starting with the first column of the table.

        Reference: `request <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/append>`_

        :param values:      List of values for the new row or column.
        :param start:       Top left cell of the range (requires a label).
        :param end:         Bottom right cell of the range (requires a label).
        :param dimension:   Dimension to which the values will be added ('ROWS' or 'COLUMNS')
        :param overwrite:   If true will overwrite data present in the spreadsheet. Otherwise will create new
                            rows to insert the data into.
        """
        if not self._linked:
            return False

        if type(values[0]) != list:
            values = [values]
        if not end:
            end = (self.rows, self.cols)
        self.client.sheet.values_append(self.spreadsheet.id, values, dimension, range=self._get_range(start, end),
                                        insertDataOption='OVERWRITE' if overwrite else 'INSERT_ROWS', **kwargs)
        self.refresh(False)

    def replace(self, pattern, replacement=None, **kwargs):
        """Replace values in any cells matched by pattern in this worksheet. Keyword arguments
        not specified will use the default value.

        If the worksheet is

            - **Unlinked** : Uses `self.find(pattern, **kwargs)` to find the cells and then replace the values in each cell.
            - **Linked** : The replacement will be done by a findReplaceRequest as defined by the Google Sheets API.\
             After the request the local copy is updated.

        Reference: `request <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#findreplacerequest>`__

        :param pattern:             Match cell values.
        :param replacement:         Value used as replacement.
        :arg searchByRegex:         Consider pattern a regex pattern. (default False)
        :arg matchCase:             Match case sensitive. (default False)
        :arg matchEntireCell:       Only match on full match. (default False)
        :arg includeFormulas:       Match fields with formulas too. (default False)
        """
        if self._linked:
            find_replace = dict()
            find_replace['find'] = pattern
            find_replace['replacement'] = replacement
            for key in kwargs:
                find_replace[key] = kwargs[key]
            find_replace['sheetId'] = self.id
            body = {'findReplace': find_replace}
            self.client.sheet.batch_update(self.spreadsheet.id, body)
            # self._update_grid(True)
        else:
            found_cells = self.find(pattern, **kwargs)
            if replacement is None:
                replacement = ''

            for cell in found_cells:
                if 'matchEntireCell' in kwargs and kwargs['matchEntireCell']:
                    cell.value = replacement
                else:
                    cell.value = re.sub(pattern, replacement, cell.value)

    def find(self, pattern, searchByRegex=False, matchCase=False, matchEntireCell=False, includeFormulas=False,
             cols=None, rows=None, forceFetch=True):
        """Finds all cells matched by the pattern.

        Compare each cell within this sheet with pattern and return all matched cells. All cells are compared
        as strings. If replacement is set, the value in each cell is set to this value. Unless full_match is False in
        in which case only the matched part is replaced.

        .. note::
            - Formulas are searched as their calculated values and not the actual formula.
            - Find fetches all data and then run a linear search on then, so this will be slow if you have a large sheet

        :param pattern:             A string pattern.
        :param searchByRegex:       Compile pattern as regex. (default False)
        :param matchCase:           Comparison is case sensitive. (default False)
        :param matchEntireCell:     Only match a cell if the pattern matches the entire value. (default False)
        :param includeFormulas:     Match cells with formulas. (default False)
        :param rows:                Range of rows to search in as tuple,  example (2, 10)
        :param cols:                Range of columns to search in as tuple, example (3, 10)
        :param forceFetch:          If the offline data should be updated before search. (default False)

        :returns:    A list of :class:`Cells <Cell>`.
        """
        if self._linked and forceFetch:
            self._update_grid(True)

        # flatten and filter data grid.
        cells = self.data_grid
        if rows: cells = self.data_grid[rows[0]-1: rows[1]]
        found_cells = []
        for cells_row in cells:
            if cols: cells_row = cells_row[cols[0]-1: cols[1]]
            found_cells.extend(cells_row)

        if not includeFormulas:
            found_cells = filter(lambda x: x.formula == '', found_cells)

        if not matchCase:
            pattern = pattern.lower()

        if searchByRegex and matchEntireCell and matchCase:
            return list(filter(lambda x: fullmatch(pattern, x.value), found_cells))
        elif searchByRegex and matchEntireCell and not matchCase:
            return list(filter(lambda x: fullmatch(pattern.lower(), x.value.lower()), found_cells))
        elif searchByRegex and not matchEntireCell and matchCase:
            return list(filter(lambda x: re.search(pattern, x.value), found_cells))
        elif searchByRegex and not matchEntireCell and not matchCase:
            return list(filter(lambda x: re.search(pattern, x.value.lower()), found_cells))

        elif not searchByRegex and matchEntireCell and matchCase:
            return list(filter(lambda x: x.value == pattern, found_cells))
        elif not searchByRegex and matchEntireCell and not matchCase:
            return list(filter(lambda x: x.value.lower() == pattern, found_cells))
        elif not searchByRegex and not matchEntireCell and matchCase:
            return list(filter(lambda x: False if x.value.find(pattern) == -1 else True, found_cells))
        else:  # if not searchByRegex and not matchEntireCell and not matchCase
            return list(filter(lambda x: False if x.value.lower().find(pattern) == -1 else True, found_cells))

    @batchable
    def create_named_range(self, name, start, end, grange, returnas='range'):
        """Create a new named range in this worksheet. Provide either start and end or grange.

        Reference: `Named range Api object <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#namedrange>`_

        :param name:    Name of the range.
        :param start:   Top left cell address (label or coordinates)
        :param end:     Bottom right cell address (label or coordinates)
        :param grange:  grid range, object of :class:`GridRange`
        :returns:   :class:`DataRange`
        """
        if not self._linked: return False

        if not grange:
            grange = GridRange(worksheet=self, start=start, end=end)

        request = {"addNamedRange": {
            "namedRange": {
                "name": name,
                "range": grange.to_json()
            }}}
        res = self.client.sheet.batch_update(self.spreadsheet.id, request)['replies'][0]['addNamedRange']['namedRange']
        if returnas == 'json':
            return res
        else:
            return DataRange(worksheet=self, namedjson=res)

    def get_named_range(self, name):
        """Get a named range by name.

        Reference: `Named range Api object <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#namedrange>`_

        :param name:    Name of the named range to be retrieved.
        :returns:   :class:`DataRange`

        :raises RangeNotFound: if no range matched the name given.
        """
        if not self._linked: return False

        nrange = [x for x in self.spreadsheet.named_ranges if x.name == name and x.worksheet.id == self.id]
        if len(nrange) == 0:
            self.spreadsheet.fetch_properties()
            nrange = [x for x in self.spreadsheet.named_ranges if x.name == name and x.worksheet.id == self.id]
            if len(nrange) == 0:
                raise RangeNotFound(name)
        return nrange[0]

    def get_named_ranges(self, name=''):
        """Get named ranges from this worksheet.

        Reference: `Named range Api object <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#namedrange>`_

        :param name:    Name of the named range to be retrieved, if omitted all ranges are retrieved.
        :return: :class:`DataRange`
        """
        if not self._linked: return False

        if name == '':
            self.spreadsheet.fetch_properties()
            nrange = [x for x in self.spreadsheet.named_ranges if x.worksheet.id == self.id]
            return nrange
        else:
            return self.get_named_range(name)

    @batchable
    def delete_named_range(self, name, range_id=''):
        """Delete a named range.

        Reference: `Named range Api object`_

        :param name:        Name of the range.
        :param range_id:    Id of the range (optional)

        """
        if not self._linked: return False

        if not range_id:
            range_id = self.get_named_ranges(name=name).name_id
        request = {'deleteNamedRange': {
            "namedRangeId": range_id,
        }}
        self.client.sheet.batch_update(self.spreadsheet.id, request)
        self.spreadsheet._named_ranges = [x for x in self.spreadsheet._named_ranges if x["namedRangeId"] != range_id]

    @batchable
    def create_protected_range(self, start=None, end=None, grange=None, named_range_id=None, returnas='range'):
        """Create protected range. Provide either start and end or grange.

        Reference: `Protected range Api object <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#protectedrange>`_

        :param start: adress of the topleft cell
        :param end: adress of the bottomright cell
        :param grange: grid range to protect, object of :class:`GridRange`
        :param named_range_id: id of named range to protect
        :param returnas: 'json' or 'range'

        """
        if not self._linked: return False

        if not grange:
            grange = GridRange(worksheet=self, start=start, end=end)

        request = {"addProtectedRange": {
            "protectedRange": {},
        }}
        if named_range_id:
            request['addProtectedRange']['protectedRange']['namedRangeId'] = named_range_id
        else:
            request['addProtectedRange']['protectedRange']['range'] = grange.to_json()
        drange = self.client.sheet.batch_update(self.spreadsheet.id,
                                                request)['replies'][0]['addProtectedRange']['protectedRange']
        if returnas == 'json':
            return drange
        else:
            return DataRange(protectedjson=drange, worksheet=self)

    @batchable
    def remove_protected_range(self, range_id):
        """Remove protected range.

        Reference: `Protected range Api object`_

        :param range_id:    ID of the protected range.
        """
        if not self._linked: return False

        request = {"deleteProtectedRange": {
            "protectedRangeId": range_id
        }}
        return self.client.sheet.batch_update(self.spreadsheet.id, request)

    def get_protected_ranges(self):
        """
        returns protected ranges in this sheet

        :return: Protected range objects
        :rtype: :class:`Datarange`
        """
        if not self._linked: return False

        self.refresh(False)
        return [DataRange(protectedjson=x, worksheet=self) for x in self.jsonSheet.get('protectedRanges', {})]

    @batchable
    def set_dataframe(self, df, start, copy_index=False, copy_head=True, extend=False, fit=False, escape_formulae=False, **kwargs):
        """Load sheet from Pandas Dataframe.

        Will load all data contained within the Pandas data frame into this worksheet.
        It will begin filling the worksheet at cell start. Supports multi index and multi header
        datarames.

        :param df:              Pandas data frame.
        :param start:           Address of the top left corner where the data should be added.
        :param copy_index:      Copy data frame index (multi index supported).
        :param copy_head:       Copy header data into first row.
        :param extend:          Add columns and rows to the worksheet if necessary, but won't delete any rows or columns.
        :param fit:             Resize the worksheet to fit all data inside if necessary.
        :param escape_formulae: Any value starting with an equal or plus sign (=/+), will be prefixed with an apostroph (') to
                                avoid value being interpreted as a formula.
        :param nan:             Value with which NaN values are replaced.

        """

        if not self._linked:
            return False
        nan = kwargs.get('nan', "NaN")

        start = Address(start)
        for col in df.select_dtypes('Int64'):
            df[col] = df[col].astype('unicode').replace('<NA>', nan)
        df = df.fillna(nan)
        values = df.astype('unicode').values.tolist()
        (df_rows, df_cols) = df.shape
        num_indexes = 1

        if copy_index:
            if isinstance(df.index, pd.MultiIndex):
                num_indexes = len(df.index[0])
                for i, indexes in enumerate(df.index):
                    indexes = map(str, indexes)
                    for index_item in reversed(list(indexes)):
                        values[i].insert(0, index_item)
                df_cols += num_indexes
            else:
                for i, val in enumerate(df.index.astype(str)):
                    values[i].insert(0, val)
                df_cols += num_indexes

        if copy_head:
            # If multi index, copy indexes in each level to new row, colum/index names are not copied for now
            if isinstance(df.columns, pd.MultiIndex):
                head = [""]*num_indexes if copy_index else []  # skip index columns
                heads = [head[:] for x in df.columns[0]]
                for col_head in df.columns:
                    for i, col_item in enumerate(col_head):
                        heads[i].append(str(col_item))
                values = heads + values
                df_rows += len(df.columns[0])
            else:
                head = [""]*num_indexes if copy_index else []  # skip index columns
                map(str, head)
                head.extend(df.columns.tolist())
                values.insert(0, head)
                df_rows += 1

        end = start + (df_rows, df_cols)

        if fit == extend is not False:
            raise InvalidArgumentValue("fit should not be same with extend")

        if fit:
            self.cols = start[1] - 1 + df_cols
            self.rows = start[0] - 1 + df_rows
        elif extend:
            self.cols = max(self.cols, start[1] - 1 + df_cols)
            self.rows = max(self.rows, start[0] - 1 + df_rows)
        else:
            if fit == "column":
                self.cols = start[1] - 1 + df_cols
            if fit == "row":
                self.rows = start[0] - 1 + df_rows
            if extend == "column":
                self.cols = max(self.cols, start[1] - 1 + df_cols)
            if extend == "row":
                self.rows = max(self.rows, start[0] - 1 + df_rows)

        if escape_formulae:
            values = list(map(lambda row: list(map(lambda cell: "'" + cell if type(cell) == str
                          and (cell.startswith('=') or cell.startswith('+')) else cell, row)), values))

        crange = start.label + ':' + end.label
        self.update_values(crange=crange, values=values)

    def get_as_df(self, has_header=True, index_column=None, start=None, end=None, numerize=True,
                  empty_value='', value_render=ValueRenderOption.FORMATTED_VALUE, **kwargs):
        """
        Get the content of this worksheet as a pandas data frame.

        :param has_header:      Interpret first row as data frame header.
        :param index_column:     Column to use as data frame index (integer).
        :param numerize:        Numerize cell values.
        :param empty_value:     Placeholder value to represent empty cells when numerizing.
        :param start:           Top left cell to load into data frame. (default: A1)
        :param end:             Bottom right cell to load into data frame. (default: (rows, cols))
        :param value_render:    How the output values should returned, `api docs <https://developers.google.com/sheets/api/reference/rest/v4/ValueRenderOption>`__
                                By default, will convert everything to strings. Setting as UNFORMATTED_VALUE will do
                                numerizing, but values will be unformatted.
        :param include_tailing_empty: whether to include empty trailing cells/values after last non-zero value in a row
        :param include_tailing_empty_rows: whether to include tailing rows with no values; if include_tailing_empty is false,
                    will return unfilled list for each empty row, else will return rows filled with empty cells
        :returns: pandas.Dataframe
        """
        if not self._linked: return False

        include_tailing_empty = kwargs.get('include_tailing_empty', False)
        include_tailing_empty_rows = kwargs.get('include_tailing_empty_rows', False)
        index_column = index_column or kwargs.get('index_colum', None)

        if not pd:
            raise ImportError("pandas")
        if start is not None or end is not None:
            if end is None:
                end = (self.rows, self.cols)
            values = self.get_values(start, end, value_render=value_render,
                                     include_tailing_empty=include_tailing_empty,
                                     include_tailing_empty_rows=include_tailing_empty_rows)
        else:
            values = self.get_all_values(returnas='matrix', include_tailing_empty=include_tailing_empty,
                                         value_render=value_render, include_tailing_empty_rows=include_tailing_empty_rows)
            
        max_row = max(len(row) for row in values)
        values = [row + [empty_value] * (max_row - len(row)) for row in values]

        if numerize:
            values = [numericise_all(row, empty_value) for row in values]

        if has_header:
            keys = values[0]
            values = values[1:]
            if any(key == '' for key in keys):
                warnings.warn('At least one column name in the data frame is an empty string. If this is a concern, please specify include_tailing_empty=False and/or ensure that each column containing data has a name.')
            df = pd.DataFrame(values, columns=keys)
        else:
            df = pd.DataFrame(values)

        if index_column:
            if index_column < 1 or index_column > len(df.columns):
                raise ValueError("index_column %s not found" % index_column)
            else:
                df.index = df[df.columns[index_column - 1]]
                del df[df.columns[index_column - 1]]
        return df

    def export(self, file_format=ExportType.CSV, filename=None, path=''):
        """Export this worksheet to a file.

        .. note::
            - Only CSV & TSV exports support single sheet export. In all other cases the entire \
        spreadsheet will be exported.
            - This can at most export files with 10 MB in size!

        :param file_format:     Target file format (default: CSV), enum :class:<pygsheets.ExportType>
        :param filename:        Filename (default: spreadsheet id + worksheet index).
        :param path:            Directory the export will be stored in. (default: current working directory)

        """
        if not self._linked:
            return
        self.client.drive.export(self, file_format=file_format, filename=filename, path=path)

    @batchable
    def copy_to(self, spreadsheet_id):
        """Copy this worksheet to another spreadsheet.

        This will copy the entire sheet into another spreadsheet and then return the new worksheet.
        Can be slow for huge spreadsheets.

        Reference: `request <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.sheets/copyTo>`__

        :param spreadsheet_id:  The id this should be copied to.
        :returns:               Copy of the worksheet in the new spreadsheet.
        """
        # TODO: Implement a way to limit returned data. For large spreadsheets.

        if not self._linked: return False

        response = self.client.sheet.sheets_copy_to(self.spreadsheet.id, self.id, spreadsheet_id)
        new_spreadsheet = self.client.open_by_key(spreadsheet_id)
        return new_spreadsheet[response['index']]

    @allow_gridrange
    def sort_range(self, start, end, basecolumnindex=0, sortorder="ASCENDING"):
        """Sorts the data in rows based on the given column index.

        :param start:               Address of the starting cell of the grid.
        :param end:                 Address of the last cell of the grid to be considered.
        :param basecolumnindex:     Index of the base column in which sorting is to be done (Integer),
                                    default value is 0. The index here is the index of the column in worksheet.
        :param sortorder:           either "ASCENDING" or "DESCENDING" (String)

        Example:
        If the data contain 5 rows and 6 columns and sorting is to be done in 4th column.
        In this case the values in other columns also change to maintain the same relative values.
        """

        if not self._linked: return False
        start = format_addr(start, 'tuple')
        end = format_addr(end, 'tuple')

        request = {"sortRange": {
            "range": {

                "sheetId": self.id,
                "startRowIndex": start[0]-1,
                "endRowIndex": end[0],
                "startColumnIndex": start[1]-1,
                "endColumnIndex": end[1],
            },
            "sortSpecs": [
                 {
                     "dimensionIndex": basecolumnindex,
                     "sortOrder": sortorder
                 }
             ],
        }}
        self.client.sheet.batch_update(self.spreadsheet.id, request)

    def add_chart(self, domain, ranges, title=None, chart_type=ChartType.COLUMN, anchor_cell=None):
        """
        Creates a chart in the sheet and retuns a chart object.

        :param domain:          Cell range of the desired chart domain in the form of tuple of adresses
        :param ranges:          Cell ranges of the desired ranges in the form of list of tuples of adresses
        :param title:           Title of the chart
        :param chart_type:      Basic chart type (default: COLUMN)
        :param anchor_cell:     position of the left corner of the chart in the form of cell address or cell object

        :return: :class:`Chart`

        Example:

        To plot a chart with x values from 'A1' to 'A6' and y values from 'B1' to 'B6'

        >>> wks.add_chart(('A1', 'A6'), [('B1', 'B6')], 'TestChart')
        <Chart 'COLUMN' 'TestChart'>

        """
        return Chart(self, domain, ranges, chart_type, title, anchor_cell)

    def get_charts(self, title=None):
        """Returns a list of chart objects, can be filtered by title.

        :param title:   title to be matched.

        :return: list of :class:`Chart`
        """
        matched_charts = []
        chart_data = self.client.sheet.get(self.spreadsheet.id,fields='sheets(charts,properties/sheetId)')
        sheet_list = chart_data.get('sheets')
        sheet = [x for x in sheet_list if x.get('properties', {}).get('sheetId') == self.id][0]
        chart_list = sheet.get('charts', [])
        for chart in chart_list:
            if not title or chart.get('spec', {}).get('title', '') == title:
                matched_charts.append(Chart(worksheet=self, json_obj=chart))
        return matched_charts

    @batchable
    def set_data_validation(self, start=None, end=None, condition_type=None, condition_values=None,
                            grange=None, **kwargs):
        """
        Sets a data validation rule to every cell in the range. To clear validation in a range,
        call this with no condition_type specified.

        refer to `api docs <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/other#conditiontype>`__ for possible inputs.

        :param start: start address
        :param end: end address
        :param grange: address as grid range
        :param condition_type: validation condition type: `possible values <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/other#conditiontype>`__
        :param condition_values: list of values for supporting condition type. For example ,
                when condition_type is NUMBER_BETWEEN, value should be two numbers indicationg lower
                and upper bound. See api docs for more info.
        :param kwargs: other options of rule.
                possible values: inputMessage, strict, showCustomUi
                `ref <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells#datavalidationrule>`__
        """
        if not grange:
            grange = GridRange(worksheet=self, start=start, end=end)
        grange.set_worksheet(self)

        condition_values = list() if not condition_values else condition_values
        json_values = []
        for value in condition_values:
            if condition_type in \
                    ['DATE_BEFORE', 'DATE_AFTER', 'DATE_ON_OR_BEFORE', 'DATE_ON_OR_AFTER']:
                json_values.append({'relativeDate': str(value)})
            else:
                json_values.append({'userEnteredValue': str(value)})

        request = {"setDataValidation": {
            "range": grange.to_json()
            }
        }
        if condition_type:
            rule = {'condition': {
                'type': condition_type,
                'values': json_values
                }
            }
            for kwarg in kwargs:
                rule[kwarg] = kwargs[kwarg]
            request['setDataValidation']['rule'] = rule
        self.client.sheet.batch_update(self.spreadsheet.id, request)

    def add_conditional_formatting(self, start, end, condition_type, format, condition_values=None, grange=None):
        """
        Adds a new conditional format rule.

        :param start: start address
        :param end: end address
        :param grange: address as grid range
        :param condition_type: validation condition type: `possible values <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/other#ConditionType>`__
        :param condition_values: list of values for supporting condition type. For example ,
                when condition_type is NUMBER_BETWEEN, value should be two numbers indicationg lower
                and upper bound. It also can be `this enum. <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/other#RelativeDate>`__ See api docs for more info.
        :param format: cell format json to apply if condition succeedes. `refer. <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/cells#CellFormat>`
        """
        if not grange:
            grange = GridRange(worksheet=self, start=start, end=end)
        grange.set_worksheet(self)
        condition_values = list() if not condition_values else condition_values

        condition_json = []
        for value in condition_values:
            if value in ['RELATIVE_DATE_UNSPECIFIED', 'PAST_YEAR', 'PAST_MONTH', 'PAST_WEEK',
                         'YESTERDAY', 'TODAY', 'TOMORROW']:
                condition_json.append({'relativeDate': value})
            else:
                condition_json.append({'userEnteredValue': value})
        request = {
            "addConditionalFormatRule": {
                "rule": {
                    "ranges": [grange.to_json()],
                    "booleanRule": {"condition": {"type": condition_type, "values": condition_json}, "format": format},
                },
                "index": 0
            }
        }
        self.client.sheet.batch_update(self.spreadsheet.id, request)

    @batchable
    def merge_cells(self, start=None, end=None, merge_type='MERGE_ALL', grange=None):
        """
        Merge cells in range

        ! You can't vertically merge cells that intersect an existing filter

        :param merge_type: either   'MERGE_ALL'
                                    ,'MERGE_COLUMNS'  ( = merge multiple rows (!) together to make column(s))
                                    ,'MERGE_ROWS' ( = merge multiple columns (!) together to make a row(s))
                                    ,'NONE' (unmerge)

        :param start: start Address
        :param end: end Address
        """
        if merge_type not in ['MERGE_ALL', 'MERGE_COLUMNS', 'MERGE_ROWS', 'NONE']:
            raise ValueError("merge_type should be one of the following : "
                             "'MERGE_ALL' 'MERGE_COLUMNS' 'MERGE_ROWS' 'NONE'")
        if not grange:
            grange = GridRange(worksheet=self, start=start, end=end)
        grange.set_worksheet(self)

        if merge_type == 'NONE':
            request = {'unmergeCells': {'range': grange.to_json()}}
        else:
            request = {'mergeCells': {'range': grange.to_json(), 'mergeType': merge_type}}
        self.client.sheet.batch_update(self.spreadsheet.id, request)

    def get_developer_metadata(self, key=None):
        """
        Fetch developer metadata associated with this worksheet

        :param key:  the key of the metadata to fetch. If unspecified, all metadata will be returned
        """
        data_filter = DeveloperMetadataLookupDataFilter(self.spreadsheet.id, self.id, meta_key=key)
        results = self.client.sheet.developer_metadata_search(self.spreadsheet.id, data_filter.to_json())
        metadata = []
        if results:
            for result in results["matchedDeveloperMetadata"]:
                meta_id = result["developerMetadata"]["metadataId"]
                key = result["developerMetadata"]["metadataKey"]
                value = result["developerMetadata"]["metadataValue"]
                metadata.append(DeveloperMetadata(meta_id, key, value, self.client, self.spreadsheet.id, self.id))
        return metadata

    def create_developer_metadata(self, key, value=None):
        """
        Create a new developer metadata associated with this worksheet

        Will return None when in batch mode, otherwise will return a DeveloperMetadata object

        :param key:    the key of the metadata to be created
        :param value:  the value of the metadata to be created (optional)
        """
        return DeveloperMetadata.new(key, value, self.client, self.spreadsheet.id, self.id)

    def __eq__(self, other):
        return self.id == other.id and self.spreadsheet == other.spreadsheet

    # @TODO optimize (use datagrid)
    def __iter__(self):
        rows = self.get_all_values(majdim='ROWS', include_tailing_empty=False, include_tailing_empty_rows=False)
        for row in rows:
            yield row + (self.cols - len(row))*['']

    # @TODO optimize (use datagrid)
    def __getitem__(self, item):
        if type(item) == int:
            if item < 1:
                raise CellNotFound("index start at 1")
            if item >= self.rows:
                raise CellNotFound
            try:
                row = self.get_row(item, include_tailing_empty=True)
            except IndexError:
                raise CellNotFound
            return row
