import sys
import re
from os import path
import pytest

sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))
import pygsheets
from pygsheets import Cell
from pygsheets.custom_types import HorizontalAlignment, VerticalAlignment

try:
    import ConfigParser
except ImportError:
    import configparser as ConfigParser

CONFIG_FILENAME = path.join(path.dirname(__file__), 'data/tests.config')
CREDS_FILENAME = path.join(path.dirname(__file__), 'data/creds.json')


def read_config(filename):
    config = ConfigParser.ConfigParser()
    config.readfp(open(filename))
    return config

config = None
gc = None


def setup_module(module):
    global config, gc
    try:
        config = read_config(CONFIG_FILENAME)
    except IOError as e:
        msg = "Can't find %s for reading test configuration. "
        raise Exception(msg % e.filename)

    try:
        gc = pygsheets.authorize(CREDS_FILENAME)
    except IOError as e:
        msg = "Can't find %s for reading credentials. "
        raise Exception(msg % e.filename)

    config_title = config.get('Spreadsheet', 'title')
    sheets = [x for x in gc.list_ssheets() if x["name"] == config_title]
    for sheet in sheets:
        gc.delete(sheet['name'])


def teardown_module(module):
    config_title = config.get('Spreadsheet', 'title')
    sheets = [x for x in gc.list_ssheets() if x["name"] == config_title]
    for sheet in sheets:
        gc.delete(sheet['name'])


# @pytest.mark.skip()
class TestPyGsheets(object):

    @pytest.mark.order1
    def test_gc(self):
        assert(isinstance(gc, pygsheets.Client))

    @pytest.mark.order2
    def test_create(self):
        spreadsheet = gc.create(title=config.get('Spreadsheet', 'title'))
        assert(isinstance(spreadsheet, pygsheets.Spreadsheet))

    @pytest.mark.order3
    def test_delete(self):
        config_title = config.get('Spreadsheet', 'title')
        gc.delete(title=config_title)
        with pytest.raises(IndexError):
            dummy = [x for x in gc._spreadsheeets if x["name"] == config_title][0]

    @pytest.mark.order4
    def test_create_delete_by_id(self):
        config_title = config.get('Spreadsheet', 'title')
        spreadsheet = gc.create(title=config_title)
        gc.delete(spreadsheet_id=spreadsheet.id)
        with pytest.raises(IndexError):
            dummy = [x for x in gc._spreadsheeets if x["id"] == spreadsheet.id][0]


# @pytest.mark.skip()
class TestClient(object):
    def setup_class(self):
        title = config.get('Spreadsheet', 'title')
        self.spreadsheet = gc.create(title)

    def teardown_class(self):
        title = config.get('Spreadsheet', 'title')
        gc.delete(title=title)

    def test_open_title(self):
        title = config.get('Spreadsheet', 'title')
        spreadsheet = gc.open(title)
        assert(isinstance(spreadsheet, pygsheets.Spreadsheet))
        assert spreadsheet.title == spreadsheet.title

    def test_open_key(self):
        title = config.get('Spreadsheet', 'title')
        spreadsheet = gc.open_by_key(self.spreadsheet.id)
        assert(isinstance(spreadsheet, pygsheets.Spreadsheet))
        assert spreadsheet.id == self.spreadsheet.id
        assert spreadsheet.title == title

    def test_open_url(self):
        url = "https://docs.google.com/spreadsheets/d/"+self.spreadsheet.id
        spreadsheet = gc.open_by_url(url)
        assert(isinstance(spreadsheet, pygsheets.Spreadsheet))
        assert spreadsheet.id == self.spreadsheet.id


# @pytest.mark.skip()
class TestSpreadSheet(object):
    def setup_class(self):
        title = config.get('Spreadsheet', 'title')
        self.spreadsheet = gc.create(title)

    def teardown_class(self):
        title = config.get('Spreadsheet', 'title')
        gc.delete(title=title)

    def test_properties(self):
        json_sheet = self.spreadsheet._jsonsheet

        assert self.spreadsheet.id == json_sheet['spreadsheetId']
        assert self.spreadsheet.title == json_sheet['properties']['title']
        assert self.spreadsheet.defaultformat == json_sheet['properties']['defaultFormat']
        assert isinstance(self.spreadsheet.sheet1, pygsheets.Worksheet)

    def test_permissions(self):
        old_per = self.spreadsheet.list_permissions()
        assert isinstance(old_per, list)

        self.spreadsheet.share('comp.tech.nm@gmail.com')
        assert len(self.spreadsheet.list_permissions()) == (len(old_per)+1)

        self.spreadsheet.remove_permissions('comp.tech.nm@gmail.com')
        assert len(self.spreadsheet.list_permissions()) == len(old_per)

    def test_workssheet_add_del(self):
        self.spreadsheet.add_worksheet("testSheetx", 50, 60)
        try:
            wks = self.spreadsheet.worksheet_by_title("testSheetx")
        except pygsheets.WorksheetNotFound:
            pytest.fail()
        assert wks.rows == 50
        assert wks.cols == 60

        self.spreadsheet.del_worksheet(wks)
        with pytest.raises(pygsheets.WorksheetNotFound):
            self.spreadsheet.worksheet_by_title("testSheetx")

    def test_worksheet_opening(self):
        wkss = self.spreadsheet.worksheets()
        assert isinstance(wkss, list)
        assert isinstance(wkss[0], pygsheets.Worksheet)

        assert isinstance(self.spreadsheet.worksheet(), pygsheets.Worksheet)

    def add_worksheet(self):
        self.spreadsheet.add_worksheet("dummy_temp_wks", 100, 50)
        wks = self.spreadsheet.worksheet_by_title("dummy_temp_wks")
        assert isinstance(wks, pygsheets.Worksheet)
        assert wks.rows == 100
        assert wks.cols == 50

    def delete_worksheet(self):
        wks = self.spreadsheet.worksheet_by_title("dummy_temp_wks")
        self.spreadsheet.del_worksheet(wks)
        with pytest.raises(pygsheets.WorksheetNotFound):
            self.spreadsheet.worksheet_by_title("dummy_temp_wks")

    def test_share(self):
        self.spreadsheet.share("comp.tech.nm@gmail.com")
        plist = self.spreadsheet.list_permissions()
        permission = [x for x in plist if x['emailAddress'] == 'comp.tech.nm@gmail.com']
        assert len(permission) == 1
        assert isinstance(permission[0], dict)
        assert permission[0]['role'] == 'reader'

    def test_remove_permission(self):
        self.spreadsheet.remove_permissions("comp.tech.nm@gmail.com")
        plist = self.spreadsheet.list_permissions()
        permission = [x for x in plist if x['emailAddress'] == 'comp.tech.nm@gmail.com']
        assert len(permission) == 0
        assert not isinstance(permission, dict)

    def test_updated(self):
        RFC_3339 = (r'\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?'
                    r'(Z|[+-]\d{2}:\d{2})')
        has_match = re.match(RFC_3339, self.spreadsheet.updated) is not None
        assert has_match

    def test_find(self):
        self.spreadsheet.add_worksheet('testFind1', 10, 10)
        self.spreadsheet.add_worksheet('testFind2', 10, 10)
        self.spreadsheet.worksheet('title', 'testFind1').update_row(1, ['test', 'test'])
        self.spreadsheet.worksheet('title', 'testFind2').update_row(1, ['test', 'test'])

        cells = self.spreadsheet.find('test')

        assert isinstance(cells, list)
        assert len(cells) == 3
        assert len(cells[0]) == 0
        assert len(cells[1]) == 2

        self.spreadsheet.del_worksheet(self.spreadsheet.worksheet('title', 'testFind1'))
        self.spreadsheet.del_worksheet(self.spreadsheet.worksheet('title', 'testFind2'))


# @pytest.mark.skip()
class TestWorkSheet(object):
    def setup_class(self):
        title = config.get('Spreadsheet', 'title')
        self.spreadsheet = gc.create(title)
        self.worksheet = self.spreadsheet.worksheet()

    def teardown_class(self):
        title = config.get('Spreadsheet', 'title')
        gc.delete(title=title)

    def test_properties(self):
        json_sheet = self.worksheet.jsonSheet

        assert self.worksheet.id == json_sheet['properties']['sheetId']
        assert self.worksheet.title == json_sheet['properties']['title']
        assert self.worksheet.index == json_sheet['properties']['index']

    def test_resize(self):
        rows = self.worksheet.rows
        cols = self.worksheet.cols

        self.worksheet.cols = cols+1
        assert self.worksheet.cols == cols+1

        self.worksheet.add_cols(1)
        assert self.worksheet.cols == cols + 2

        self.worksheet.rows = rows + 1
        assert self.worksheet.rows == rows + 1

        self.worksheet.add_rows(1)
        assert self.worksheet.rows == rows + 2

        self.worksheet.resize(rows, cols)
        assert self.worksheet.cols == cols
        assert self.worksheet.rows == rows

    def test_frozen_rows(self):
        ws = self.worksheet
        assert ws.frozen_rows == 0
        ws.frozen_rows = 1
        ws.refresh()
        assert ws.frozen_rows == 1
        ws.frozen_rows = 0
        ws.refresh()

    def test_frozen_cols(self):
        ws = self.worksheet
        assert ws.frozen_cols == 0
        ws.frozen_cols = 2
        ws.refresh()
        assert ws.frozen_cols == 2
        ws.frozen_cols = 0
        ws.refresh()

    def test_addr_reformat(self):
        addr = pygsheets.format_addr((1, 1))
        assert addr == 'A1'

        addr = pygsheets.format_addr('A1')
        assert addr == (1, 1)

    def test_cell(self):
        assert isinstance(self.worksheet.cell('A1'), pygsheets.Cell)
        assert isinstance(self.worksheet.cell((1, 1)), pygsheets.Cell)

        with pytest.raises(pygsheets.CellNotFound):
            self.worksheet.cell((self.worksheet.rows + 5, self.worksheet.cols + 5))

    def test_insert_cols_rows(self):
        cols = self.worksheet.cols
        self.worksheet.insert_cols(1, 2)
        assert self.worksheet.cols == (cols+2)

        rows = self.worksheet.rows
        self.worksheet.insert_rows(1, 2)
        assert self.worksheet.rows == (rows + 2)

        with pytest.raises(pygsheets.InvalidArgumentValue):
            pygsheets.format_addr([1, 1])

    def test_values(self):
        self.worksheet.update_value('A1', 'test val')
        vals = self.worksheet.get_values('A1', 'B4')
        assert isinstance(vals, list)
        assert vals[0][0] == 'test val'

        vals = self.worksheet.get_values('A1', (2, 2), 'cells')
        assert isinstance(vals, list)
        assert isinstance(vals[0][0], pygsheets.Cell)
        assert vals[0][0].value == 'test val'

    def test_update_cells(self):
        self.worksheet.update_values(crange='A1:B2', values=[[1, 2], [3, 4]])
        assert self.worksheet.cell((1, 1)).value == str(1)
        self.worksheet.resize(1, 1)
        self.worksheet.update_values(crange='A1', values=[[1, 2, 5], [3, 4, 6], [3, 4, 61]], extend=True)
        assert self.worksheet.cols == 3
        assert self.worksheet.rows == 3
        assert self.worksheet.cell((3, 3)).value == '61'

        self.worksheet.resize(30, 30)
        cells = [pygsheets.Cell('A1', 10), pygsheets.Cell('A2', 12)]
        self.worksheet.update_values(cell_list=cells)
        assert self.worksheet.cell((1, 1)).value == str(cells[0].value)

    def test_update_col(self):
        self.worksheet.resize(30, 30)
        self.worksheet.update_col(5, [1,2,3,4,5])
        cols = self.worksheet.get_col(5)
        assert isinstance(cols, list)
        assert cols[3] == str(4)

    def test_update_row(self):
        self.worksheet.resize(30, 30)
        self.worksheet.update_row(5,[1,2,3,4,5])
        rows = self.worksheet.get_row(5)
        assert isinstance(rows, list)
        assert rows[3] == str(4)

    def test_range(self):
        assert isinstance(self.worksheet.range('A1:A5'), list)

    def test_value_set(self):
        self.worksheet.update_value('A1', 'xxx')
        assert self.worksheet.cell('A1').value == 'xxx'

    def test_iter(self):
        self.worksheet.update_row(1, [1, 2, 3, 4, 5])
        self.worksheet.update_row(2, [2, 3, 4, 5, 6])
        wks_iter = iter(self.worksheet)
        assert next(wks_iter)[:5] == ['1', '2', '3', '4', '5']
        assert next(wks_iter)[:5] == ['2', '3', '4', '5', '6']

    def test_getitem(self):
        self.worksheet.update_row(1, [1, 2, 3, 4, 5])
        row = self.worksheet[0]
        assert len(row) == self.worksheet.cols
        assert row[0][0] == str(1)

    def test_clear(self):
        self.worksheet.update_value('S10', 100)
        self.worksheet.clear()
        assert self.worksheet.get_all_values() == [[]]

    def test_delete_dimension(self):
        rows = self.worksheet.rows
        self.worksheet.update_row(10, [1, 2, 3, 4, 5])
        self.worksheet.delete_rows(10)
        with pytest.raises(IndexError):
            assert self.worksheet.get_value((9, 2)) != 2
        assert self.worksheet.rows == rows - 1

        cols = self.worksheet.cols
        self.worksheet.update_col(10, [1, 2, 3, 4, 5])
        self.worksheet.delete_cols(10)
        with pytest.raises(IndexError):
            assert self.worksheet.get_value((10, 2)) != 2
        assert self.worksheet.cols == cols - 1

    # @TODO
    def test_append_row(self):
        assert True

    def test_set_dataframe(self):
        import pandas as pd
        df = pd.DataFrame({'a': [1, 2, 3, 'g'], 'x': [4, 5, 6, 'h']})
        self.worksheet.set_dataframe(df, 'B2', copy_head=True, fit=True, copy_index=True)
        assert self.worksheet.get_value('D5') == '6'
        assert self.worksheet.get_value('C5') == '3'
        assert self.worksheet.get_value('D2') == 'x'
        assert self.worksheet.cols == 4
        assert self.worksheet.rows == 6

        self.worksheet.set_dataframe(df, 'B2', copy_head=True, fit=True, copy_index=False)
        assert self.worksheet.get_value('C2') == 'x'

        self.worksheet.set_dataframe(df, 'B2', copy_head=False, fit=True, copy_index=False)
        assert self.worksheet.get_value('B2') == '1'
        assert self.worksheet.get_value('C2') == '4'

        # Test MultiIndex
        import numpy as np
        arrays = [np.array(['bar', 'bar', 'baz', 'baz', 'foo', 'foo', 'qux', 'qux']),
                  np.array(['one', 'two', 'one', 'two', 'one', 'two', 'one', 'two'])]
        tuples = list(zip(*arrays))
        index = pd.MultiIndex.from_tuples(tuples, names=['first', 'second'])
        df = pd.DataFrame(np.random.randn(8, 8), index=index, columns=index)
        self.worksheet.set_dataframe(df, 'A1', copy_index=True)
        assert self.worksheet.get_value('C1') == 'bar'
        assert self.worksheet.get_value('C2') == 'one'
        assert self.worksheet.get_value('F1') == 'baz'
        assert self.worksheet.get_value('F2') == 'two'
        self.worksheet.clear()

    # @TODO
    def test_get_as_df(self):
        assert True

    # @TODO
    def test_export(self):
        assert True

    def test_get_values(self):
        self.worksheet.resize(10, 10)
        self.worksheet.clear()
        self.worksheet.update_values('A1:C2', [[1, 2, ''], [2, 3, 4]])
        assert self.worksheet.get_values('A1', 'E5') == [[u'1', u'2', '', '', ''], [u'2', u'3', u'4', '', '']]

        # @TODO not working
        # assert self.worksheet.get_values('A1','D3', returnas="cells") == [[Cell('A1', '1'), Cell('B1','2'), Cell('C1',''), Cell('D1','')],
        #                                                                   [Cell('A2','2'), Cell('B2','3'), Cell('C2','4'), Cell('D2','')]]

        assert self.worksheet.get_values('A1','D3', returnas="cells", include_empty=False) == [[Cell('A1', '1'), Cell('B1','2') ],
                                                                          [Cell('A2','2'), Cell('B2','3'), Cell('C2','4')]]
        assert self.worksheet.get_values('D1', 'D3', returnas="cells", include_all=True) == [[Cell('D1', '')], [Cell('D2', '')], [Cell('D3', '')]]

    def test_hide_rows(self):
        self.worksheet.hide_rows(0, 2)
        json =self.spreadsheet.client.sh_get_ssheet(self.spreadsheet.id, fields="sheets/data/rowMetadata/hiddenByUser")
        assert json['sheets'][0]['data'][0]['rowMetadata'][0]['hiddenByUser'] == True
        assert json['sheets'][0]['data'][0]['rowMetadata'][1]['hiddenByUser'] == True
        self.worksheet.show_rows(0, 2)
        json =self.spreadsheet.client.sh_get_ssheet(self.spreadsheet.id, fields="sheets/data/rowMetadata/hiddenByUser")
        assert json['sheets'][0]['data'][0]['rowMetadata'][0].get('hiddenByUser', False) == False
        assert json['sheets'][0]['data'][0]['rowMetadata'][1].get('hiddenByUser', False) == False

    def test_hide_columns(self):
        self.worksheet.hide_columns(0, 2)
        json = self.spreadsheet.client.sh_get_ssheet(self.spreadsheet.id, fields="sheets/data/columnMetadata/hiddenByUser")
        assert json['sheets'][0]['data'][0]['columnMetadata'][0]['hiddenByUser'] == True
        assert json['sheets'][0]['data'][0]['columnMetadata'][1]['hiddenByUser'] == True
        self.worksheet.show_columns(0, 2)
        json =self.spreadsheet.client.sh_get_ssheet(self.spreadsheet.id, fields="sheets/data/columnMetadata/hiddenByUser")
        assert json['sheets'][0]['data'][0]['columnMetadata'][0].get('hiddenByUser', False) == False
        assert json['sheets'][0]['data'][0]['columnMetadata'][1].get('hiddenByUser', False) == False

    def test_find(self):
        cells = self.worksheet.find('test')
        assert isinstance(cells, list)
        assert 0 == len(cells)
        self.worksheet.update_row(1, ['test', 'test', 100, 'TEST', 'testtest', 'test', 'test', '=SUM(C:C)'])
        cells = self.worksheet.find('test')
        assert 4 == len(cells)
        self.worksheet.unlink()
        cells = self.worksheet.find('test', matchCase=False)
        assert 5 == len(cells)
        cells = self.worksheet.find('100')
        assert 1 == len(cells)
        cells = self.worksheet.find('test', matchEntireCell=False)
        assert 5 == len(cells)
        cells = self.worksheet.find('100', matchEntireCell=False, includeFormulas=True)
        assert 2 == len(cells)
        cells = self.worksheet.find('\w+')
        assert 7 == len(cells)
        cells = self.worksheet.find('test', 'value')
        assert 4 == len(cells)
        assert 'value' == cells[0].value
        cells = self.worksheet.find('test', 'value', matchEntireCell=False)
        assert 1 == len(cells)
        assert 'valuevalue' == cells[0].value

        self.worksheet.link(syncToCloud=False)
        row = self.worksheet.get_row(1)
        assert row[0] == 'value'
        assert row[3] == 'TEST'
        assert row[4] == 'valuevalue'
        self.worksheet.clear('A1', 'H1')

# @pytest.mark.skip()
class TestDataRange(object):
    def setup_class(self):
        title = config.get('Spreadsheet', 'title')
        self.spreadsheet = gc.create(title)
        self.worksheet = self.spreadsheet.worksheet()
        self.range = self.worksheet.range("A1:A2", returnas="range")

    def teardown_class(self):
        title = config.get('Spreadsheet', 'title')
        gc.delete(title=title)

    def test_protected_range(self):
        self.range.protected = True
        assert self.range.protected
        assert self.range.protect_id is not None
        assert self.range is not None
        assert len(self.spreadsheet.protected_ranges) == 1
        self.range.protected = False
        assert not self.range.protected
        assert self.range.protect_id is None
        assert len(self.spreadsheet.protected_ranges) == 0


# @pytest.mark.skip()
class TestCell(object):
    def setup_class(self):
        title = config.get('Spreadsheet', 'title')
        self.spreadsheet = gc.create(title)
        self.worksheet = self.spreadsheet.worksheet()
        self.cell = self.worksheet.cell('A1')
        self.cell.value = 'test_value'

    def teardown_class(self):
        title = config.get('Spreadsheet', 'title')
        gc.delete(title=title)

    def test_properties(self):
        assert self.cell.row == 1
        assert self.cell.col == 1
        assert self.cell.value == 'test_value'
        assert self.cell.label == 'A1'
        assert self.cell.horizontal_alignment == HorizontalAlignment.NONE
        assert self.cell.vertical_alignment == VerticalAlignment.NONE

    def test_alignment(self):
        self.cell.horizontal_alignment = HorizontalAlignment.RIGHT
        self.cell.vertical_alignment = VerticalAlignment.MIDDLE
        assert self.cell.horizontal_alignment == HorizontalAlignment.RIGHT
        assert self.cell.vertical_alignment == VerticalAlignment.MIDDLE

    def test_link(self):
        self.worksheet.update_value('B2', 'new_val')
        self.cell.row += 1
        self.cell.col += 1

        assert self.cell.row == 2
        assert self.cell.col == 2
        assert self.cell.value == 'new_val'
        assert self.cell.label == 'B2'

    def test_formula(self):
        self.worksheet.update_value('B1', 3)
        self.worksheet.update_value('C1', 4)
        cell = self.worksheet.cell('A1')
        cell.formula = '=B1+C1'
        assert cell.value == '7'
        assert cell.value_unformatted == 7

    def test_neighbour(self):
        self.worksheet.update_value('B1', 7)
        self.worksheet.update_value('C1', 8)
        cell = self.worksheet.cell('A1')

        assert cell.neighbour('right').value == '7'
        assert cell.neighbour((0, 1)).value == '7'
        assert cell.neighbour((0, 2)).value == '8'

    def test_link_unlink(self):
        self.worksheet.update_value('A1', 5)
        cell = self.worksheet.cell('A1')
        cell.unlink()
        cell.value = 10
        assert self.worksheet.get_value('A1') == '5'
        cell.link(update=True)
        assert self.worksheet.get_value('A1') == '10'

        cell.unlink()
        cell.value = 20
        assert self.worksheet.get_value('A1') == '10'
        cell.link()
        cell.update()
        assert self.worksheet.get_value('A1') == '20'

    def test_wrap_strategy(self):
        cell = self.worksheet.get_values('A1', 'A1', returnas="range")[0][0]
        assert cell.wrap_strategy == "WRAP_STRATEGY_UNSPECIFIED"
        cell.wrap_strategy = "WRAP"
        cell = self.worksheet.get_values('A1', 'A1', returnas="range")[0][0]
        assert cell.wrap_strategy == "WRAP"
