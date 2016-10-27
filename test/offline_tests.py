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
import pytest, unittest
import sys
from os import path

sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))
import pygsheets

CONFIG_FILENAME = path.join(path.dirname(__file__), 'tests.config')


def read_config(filename):
    config = ConfigParser.ConfigParser()
    config.readfp(open(filename))
    return config


class TestSpreadsheet(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        with open('spreadsheet.json') as data_file:
            spreadsheet_json = json.load(data_file)
        with open('worksheet.json') as data_file:
            worksheet_json = json.load(data_file)
        spreadsheet_json["Sheets"].append(worksheet_json)

        cls.config = read_config(CONFIG_FILENAME)
        cls.gc = mock.create_autospec(pygsheets.Client)
        cls.gc.open_by_key.return_value = spreadsheet_json
        cls.gc._fetch_sheets = [{'testssheetid': 'testssheettitle'}]
        cls.spreadsheet = pygsheets.Spreadsheet(cls.gc)

    def setUp(self):
        print("setUp")

    def test_open(self):
        spreadsheet = self.gc.open('testssheettitle')
        self.assertTrue(isinstance(spreadsheet, pygsheets.Spreadsheet))

    def tearDown(self):
        pass

    @classmethod
    def tearDownClass(cls):
        pass

if __name__ == '__main__':
    unittest.main()
