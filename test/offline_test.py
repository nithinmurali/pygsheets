"""A test suite that doesn't query the Google API.

Avoiding direct network access is benefitial in that it markedly speeds up
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

config = None
mock_gc = None


def setup_module(module):
    global config, mock_gc
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

    mock_gc = gc


@pytest.mark.skip()
class TestClient(object):

    @classmethod
    def setup_class(cls):
        # spreadsheet = mock_gc
        pass

    def test_open(self):
        sh_title = config.get('Spreadsheet', 'title')
        sh_id = config.get('Spreadsheet', 'id')

        ret_sh = mock_gc.open(sh_title)
        assert (isinstance(ret_sh, pygsheets.Spreadsheet))
        mock_gc.open_by_key.assert_called_once_with(key=sh_id)

        sh_url = config.get('Spreadsheet', 'url')
        ret_sh = mock_gc.open_by_url(sh_url)
        assert (isinstance(ret_sh, pygsheets.Spreadsheet))
        mock_gc.open_by_key.assert_called_once_with(key=sh_id)

        ret_sh = mock_gc.open_all(sh_title)
        assert (isinstance(ret_sh, list))
        mock_gc.open_by_key.assert_called_once_with(key=sh_id)

    # @mock.patch('service.spreadsheets().get')
    # def test_open(self, mock_get):
    #     mock_response = mock.Mock()
    #     mock_response.execute.return_value = self.spreadsheet
    #     mock_get.return_value = mock_response
    #
    #     spreadsheet = self.gc.open('testssheettitle')
    #
    #     mock_get.assert_called_once_with(spreadsheetId=self.config.get('Spreadsheet', 'key'),
    #                                      fields='properties,sheets/properties,spreadsheetId')
    #     assert(isinstance(spreadsheet, pygsheets.Spreadsheet))


class TestSpreadsheet(object):

    @classmethod
    def setup_class(cls):
        sh_id = config.get('Spreadsheet', 'id')
        cls.spreadsheet = mock_gc.open_by_key(sh_id)

    def setup_method(self, method):
        sh_id = config.get('Spreadsheet', 'id')
        self.spreadsheet = mock_gc.open_by_key(sh_id)

    # @pytest.mark.skip()
    def test_properties(self):
        sh_id = config.get('Spreadsheet', 'id')
        sh_title = config.get('Spreadsheet', 'title')
        assert self.spreadsheet.id == sh_id
        assert self.spreadsheet.title == sh_title
        wks = self.spreadsheet.sheet1
        assert(isinstance(wks, pygsheets.Worksheet))

    # @pytest.mark.skip()
    @mock.patch('pygsheets.Spreadsheet._fetch_sheets', return_value=True)
    def test_worksheet_open(self, mock_fn):
        wks_id = int(config.get('Worksheet', 'id'))
        wks_title = config.get('Worksheet', 'title')

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

    def test_worksheet_remove(self):
        wks_id = config.get('Worksheet', 'id')
        wks = self.spreadsheet.worksheet('id', wks_id)
        self.spreadsheet.del_worksheet(wks)
        mock_gc.sh_batch_update.assert_called_once()
        with pytest.raises(pygsheets.WorksheetNotFound):
            self.spreadsheet.del_worksheet(mock.Mock(pygsheets.Worksheet))

    @pytest.mark.skip()
    def test_worksheet_add(self):
        wks_id = config.get('Worksheet', 'id')
        jsheet = {'replies': [{'addSheet': {'properties':
                                            self.spreadsheet.worksheet('id', wks_id).jsonSheet}}]}
        mock_gc.sh_batch_update.return_value = jsheet
        wks = self.spreadsheet.add_worksheet('wks1', 100, 100)
        mock_gc.sh_batch_update.assert_called_once()

