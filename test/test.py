import sys
from os import path
import pytest

sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))
import pygsheets

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
        gc = pygsheets.authorize(CREDS_FILENAME)
    except IOError as e:
        msg = "Can't find %s for reading test configuration. "
        raise Exception(msg % e.filename)

    config_title = config.get('Spreadsheet', 'title')
    sheets = [x for x in gc._spreadsheeets if x["name"] == config_title]
    for sheet in sheets:
        gc.delete(sheet['name'])


def teardown_module(module):
    config_title = config.get('Spreadsheet', 'title')
    sheets = [x for x in gc._spreadsheeets if x["name"] == config_title]
    for sheet in sheets:
        gc.delete(sheet['name'])


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

#
# class TestSpreadSheet(object):
#     def setup_class(self):
#         title = self.config.get('Spreadsheet', 'title')
#         self.gc.create(title)
#
#     def teardown_class(self):
#         title = self.config.get('Spreadsheet', 'title')
#         self.gc.delete(title=title)


