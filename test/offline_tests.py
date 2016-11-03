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

CONFIG_FILENAME = path.join(path.dirname(__file__), 'tests.config')


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
        gc = pygsheets.authorize(CREDS_FILENAME)
    except IOError as e:
        msg = "Can't find %s for reading test configuration. "
        raise Exception(msg % e.filename)

    config_title = config.get('Spreadsheet', 'title')
    sheets = [x for x in gc._spreadsheeets if x["name"] == config_title]
    for sheet in sheets:
        gc.delete(sheet['name'])


class TestSpreadsheet(object):

    @classmethod
    def setup_class(cls):
        with open('spreadsheet.json') as data_file:
            spreadsheet_json = json.load(data_file)
        with open('worksheet.json') as data_file:
            worksheet_json = json.load(data_file)
        spreadsheet_json["Sheets"].append(worksheet_json)

        cls.config = read_config(CONFIG_FILENAME)
        # cls.gc = mock.create_autospec(pygsheets.Client)
        # cls.gc.open_by_key.return_value = spreadsheet_json
        # cls.gc._fetch_sheets = [{'testssheetid': 'testssheettitle'}]
        # cls.spreadsheet = pygsheets.Spreadsheet(cls.gc)

        cls.gc.service.spreadsheets().get = mock.create_autospec(cls.gc.service.spreadsheets().get)
        cls.gc.open_by_key.return_value = spreadsheet_json
        cls.gc._fetch_sheets = [{'testssheetid': 'testssheettitle'}]
        cls.spreadsheet = pygsheets.Spreadsheet(cls.gc)

    @mock.patch('service.spreadsheets().get')
    def test_open(self, mock_get):

        mock_response = mock.Mock()
        mock_response.execute.return_value = self.spreadsheet
        mock_get.return_value = mock_response

        spreadsheet = self.gc.open('testssheettitle')

        mock_get.assert_called_once_with(spreadsheetId=self.config.get('Spreadsheet', 'key'),
                                         fields='properties,sheets/properties,spreadsheetId')
        assert(isinstance(spreadsheet, pygsheets.Spreadsheet))

    @classmethod
    def teardown_class(cls):
        pass
