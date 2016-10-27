import sys
from os import path
import pytest

sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))
import pygsheets

try:
    import ConfigParser
except ImportError:
    import configparser as ConfigParser

CONFIG_FILENAME = path.join(path.dirname(__file__), 'tests.config')
CREDS_FILENAME = path.join(path.dirname(__file__), 'creds.json')


def read_config(filename):
    config = ConfigParser.ConfigParser()
    config.readfp(open(filename))
    return config


class TestPyGsheets(object):

    @classmethod
    def setUpClass(cls):
        print "this called"
        try:
            cls.config = read_config(CONFIG_FILENAME)
            cls.gc = pygsheets.authorize(CREDS_FILENAME)
        except IOError as e:
            msg = "Can't find %s for reading test configuration. "
            raise Exception(msg % e.filename)

    def setUp(self):
        assert(isinstance(self.gc, pygsheets.Client))

    @pytest.mark.order1
    def test_create(self):
        spreadsheet = self.gc.create("this is dummy test ssheet for pygsheets")
        assert(isinstance(spreadsheet, pygsheets.Spreadsheet))

    @pytest.mark.order2
    def test_delete(self):
        self.gc.delete("this is dummy test ssheet for pygsheets")

#
# class TestClient(TestPyGsheets):
#     def setup_class(self):
#         title = self.config.get('Spreadsheet', 'title')
#         self.gc.create(title)
#
#     def teardown_class(self):
#         title = self.config.get('Spreadsheet', 'title')
#         self.gc.delete(title=title)
#
#     def test_create(self):
#         title = self.config.get('Spreadsheet', 'title')
#         spreadsheet = self.gc.create(title)
#         assert(isinstance(spreadsheet, pygsheets.Spreadsheet))
#
#     def test_open(self):
#         title = self.config.get('Spreadsheet', 'title')
#         spreadsheet = self.gc.open(title)
#         assert(isinstance(spreadsheet, pygsheets.Spreadsheet))
#
#
# class TestSpreadSheet(TestPyGsheets):
#     def setup_class(self):
#         title = self.config.get('Spreadsheet', 'title')
#         self.gc.create(title)
#
#     def teardown_class(self):
#         title = self.config.get('Spreadsheet', 'title')
#         self.gc.delete(title=title)
