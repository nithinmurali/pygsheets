import os

import pytest


import pygsheets
from pygsheets.client import Client


class TestAuthorization(object):

    def setup_class(self):
        self.base_path = os.path.join(os.path.dirname(__file__), 'auth_test_data')

    def test_service_account_authorization(self):
        c = pygsheets.authorize(service_account_file=self.base_path + '/pygsheettest_service_account.json')
        assert isinstance(c, Client)

        c.create('test_sheet')

    def test_user_credentials_flow(self):
        pass

    def test_user_credentials_loading(self):
        pass

    def teardown_class(self):
        pass
