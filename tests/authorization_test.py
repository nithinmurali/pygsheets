from google.auth.exceptions import RefreshError
from google.oauth2.credentials import Credentials
import os
import sys
from unittest.mock import patch

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
import pygsheets
from pygsheets.authorization import _SCOPES
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

    def test_kwargs_passed_to_client(self):
        c = pygsheets.authorize(service_file=self.base_path + '/pygsheettest_service_account.json', retries=3)
        assert isinstance(c, Client)
        assert c.sheet.retries == 3

    def test_should_reload_client_secret_on_refresh_error(self):
        # First connection
        initial_c = pygsheets.authorize(
            client_secret=self.base_path + "/client_secret.json",
            credentials_directory=self.base_path,
        )
        credentials_filepath = self.base_path + "/sheets.googleapis.com-python.json"
        assert os.path.exists(credentials_filepath)

        # After a while, the refresh token is not working and raises RefreshError
        refresh_c = None
        with patch(
            "pygsheets.authorization._get_initial_user_authentication_credentials"
        ) as mock_initial_credentials:
            real_credentials = Credentials.from_authorized_user_file(
                credentials_filepath, scopes=_SCOPES
            )
            mock_initial_credentials.return_value = real_credentials

            with patch("pygsheets.authorization.Credentials") as mock_credentials:
                mock_credentials.from_authorized_user_file.return_value.refresh.side_effect = RefreshError(
                    "Error using refresh token"
                )
                mock_initial_credentials
                refresh_c = pygsheets.authorize(
                    client_secret=self.base_path + "/client_secret.json",
                    credentials_directory=self.base_path,
                )

            mock_initial_credentials.assert_called_once()
        assert isinstance(refresh_c, Client)
