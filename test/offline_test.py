"""A test suite that doesn't query the Google API.

Avoiding direct network access is beneficial in that it markedly speeds up
testing, avoids error-prone credential setup, and enables validation even if
internet access is unavailable.
"""

import json
try:
    import ConfigParser
except ImportError:
    import configparser as ConfigParser

import mock
import pytest
import sys
from os import path

sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))
import pygsheets

DATA_DIR = path.join(path.dirname(__file__), 'data')
CONFIG_FILENAME = path.join(DATA_DIR, 'tests.config')


def read_config(filename):
    config = ConfigParser.ConfigParser()
    config.readfp(open(filename))
    return config


test_config = None
mock_gc = None


def setup_module(module):
    global test_config, mock_gc
    config = read_config(CONFIG_FILENAME)
    gc = mock.create_autospec(pygsheets.Client)

    sh_title = config.get('Spreadsheet', 'title')
    sh_id = config.get('Spreadsheet', 'id')

    with open(path.join(DATA_DIR, 'spreadsheet.json')) as data_file:
        spreadsheet_json = json.load(data_file)
    with open(path.join(DATA_DIR, 'worksheet.json')) as data_file:
        worksheet_json = json.load(data_file)
    spreadsheet_json["sheets"].append(worksheet_json)

    # path the object
    # gc._fetch_sheets.return_value = [{u'id': sh_id, u'name': sh_title}]
    gc.create.return_value = pygsheets.Spreadsheet(gc, spreadsheet_json)
    gc.open_by_key.return_value = pygsheets.Spreadsheet(gc, spreadsheet_json)
    gc.sh_batch_update.return_value = True

    mock_gc = gc


class TestSpreadsheet(object):

    @classmethod
    def setup_class(cls):
        sh_id = test_config.get('Spreadsheet', 'id')
        cls.spreadsheet = mock_gc.open_by_key(sh_id)

    def setup_method(self, method):
        sh_id = test_config.get('Spreadsheet', 'id')
        self.spreadsheet = mock_gc.open_by_key(sh_id)

    # @pytest.mark.skip()
    def test_properties(self):
        sh_id = test_config.get('Spreadsheet', 'id')
        sh_title = test_config.get('Spreadsheet', 'title')
        assert self.spreadsheet.id == sh_id
        assert self.spreadsheet.title == sh_title
        wks = self.spreadsheet.sheet1
        assert(isinstance(wks, pygsheets.Worksheet))

    # @pytest.mark.skip()
    @mock.patch('pygsheets.Spreadsheet._fetch_sheets', return_value=True)
    def test_worksheet_open(self, mock_fn):
        wks_id = int(test_config.get('Worksheet', 'id'))
        wks_title = test_config.get('Worksheet', 'title')

        wks = self.spreadsheet.worksheets()
        assert (isinstance(wks, list))

        wks = self.spreadsheet.worksheets('title', wks_title)[0]
        assert (isinstance(wks, pygsheets.Worksheet))
        with pytest.raises(pygsheets.WorksheetNotFound):
            wks = self.spreadsheet.worksheets('title', 'someinvalidkey')
            mock_gc.open_by_key.assert_caleld_with(self.spreadsheet.id)

        wks = self.spreadsheet.worksheet(property='id', value=wks_id)
        assert(isinstance(wks, pygsheets.Worksheet))

        wks = self.spreadsheet.worksheet_by_title(wks_title)
        assert(isinstance(wks, pygsheets.Worksheet))

    # def test_worksheet_remove(self):
    #     wks_id = config.get('Worksheet', 'id')
    #     wks = self.spreadsheet.worksheet('id', wks_id)
    #     self.spreadsheet.del_worksheet(wks)
    #     mock_gc.sh_batch_update.assert_called()
    #     with pytest.raises(pygsheets.WorksheetNotFound):
    #         self.spreadsheet.del_worksheet(mock.Mock(pygsheets.Worksheet))


class TestWorksheet(object):
    pass