import os
import sys

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pygsheets
from pygsheets.client import Client

from googleapiclient.http import HttpError


class TestAuthorization(object):

    def setup_class(self):
        self.base_path = os.path.join(os.path.dirname(__file__), 'auth_test_data')

    def test_service_account_authorization(self):
        c = pygsheets.authorize(service_account_file=self.base_path + '/pygsheettest_service_account.json')
        assert isinstance(c, Client)

        self.sheet = c.create('test_sheet')
        self.sheet.share('pygsheettest@gmail.com')
        self.sheet.delete()

    def test_user_credentials_loading(self):
        c = pygsheets.authorize(client_secret=self.base_path + '/client_secret.json',
                                credentials_directory=self.base_path)
        assert isinstance(c, Client)

        self.sheet = c.create('test_sheet')
        self.sheet.share('pygsheettest@gmail.com')
        self.sheet.delete()

    def test_deprecated_kwargs_removal(self):
        c = pygsheets.authorize(service_file=self.base_path + '/pygsheettest_service_account.json')
        assert isinstance(c, Client)