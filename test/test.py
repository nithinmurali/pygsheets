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


class TestPyGsheets(object):

    @classmethod
    def setup_class(cls):
        try:
            cls.config = read_config(CONFIG_FILENAME)
            cls.gc = pygsheets.authorize(CREDS_FILENAME)
        except IOError as e:
            msg = "Can't find %s for reading test configuration. "
            raise Exception(msg % e.filename)

    @pytest.mark.order1
    def test_gc(self):
        # print "Called 2"
        assert(isinstance(self.gc, pygsheets.Client))

    @pytest.mark.order2
    def test_create(self):
        # print "Called 2"
        spreadsheet = self.gc.create(title=self.config.get('Spreadsheet', 'title'))
        self.id = spreadsheet.id
        assert(isinstance(spreadsheet, pygsheets.Spreadsheet))

    @pytest.mark.order3
    def test_delete(self):
        # print "Called 3"
        self.gc.delete(title=self.config.get('Spreadsheet', 'title'))
        # with pytest.raises(KeyError):
        #     dummy = [x for x in self.gc._spreadsheeets if x["id"] == self.id][0]


class TestClient(TestPyGsheets):
    def setup_class(self):
        title = self.config.get('Spreadsheet', 'title')
        self.gc.create(title)
        self.spreadsheet = self.gc.open(title)

    def teardown_class(self):
        title = self.config.get('Spreadsheet', 'title')
        self.gc.delete(title=title)

    def test_open_title(self):
        title = self.config.get('Spreadsheet', 'title')
        spreadsheet = self.gc.open(title)
        assert(isinstance(spreadsheet, pygsheets.Spreadsheet))
        assert spreadsheet.title == title

    def test_open_key(self):
        spreadsheet = self.gc.open_by_key(self.spreadsheet.id)
        assert(isinstance(spreadsheet, pygsheets.Spreadsheet))
        assert spreadsheet.id == self.spreadsheet.id

    # def test_open_url(self):
    #     url = "https://docs.google.com/spreadsheets/d/"+self.spreadsheet.id
    #     print url
    #     spreadsheet = self.gc.open_by_url(url)
    #     assert(isinstance(spreadsheet, pygsheets.Spreadsheet))
    #     assert spreadsheet.id == self.spreadsheet.id


class TestSpreadSheet(TestPyGsheets):
    def setup_class(self):
        title = self.config.get('Spreadsheet', 'title')
        self.gc.create(title)

    def teardown_class(self):
        title = self.config.get('Spreadsheet', 'title')
        self.gc.delete(title=title)


