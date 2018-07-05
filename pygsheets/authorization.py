import logging
import os
import json

from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow

from pygsheets.client import Client


def _get_user_authentication_credentials(client_secret_file, scopes, credential_directory=None):
    """"""
    if credential_directory is None:
        credential_directory = os.getcwd()
    elif credential_directory == 'global':
        home_dir = os.path.expanduser('~')
        credential_directory = os.path.join(home_dir, '.credentials')
        if not os.path.exists(credential_directory):
            os.makedirs(credential_directory)
    else:
        pass

    credentials_path = os.path.join(credential_directory, 'sheets.googleapis.com-python.json')

    if os.path.exists(credentials_path):
        # expect these to be valid. may expire at some point, but should be refreshed by google api client...
        return Credentials.from_authorized_user_file(credentials_path, scopes=scopes)

    flow = Flow.from_client_secrets_file(client_secret_file, scopes=scopes, redirect_uri='urn:ietf:wg:oauth:2.0:oob')

    auth_url, _ = flow.authorization_url(prompt='consent')

    print('Please go to this URL and finish the authentication flow: {}'.format(auth_url))
    code = input('Enter the authorization code: ')
    flow.fetch_token(code=code)

    credentials = flow.credentials

    credentials_as_dict = {
        'token': credentials.token,
        'refresh_token': credentials.refresh_token,
        'id_token': credentials.id_token,
        'token_uri': credentials.token_uri,
        'client_id': credentials.client_id,
        'client_secret': credentials.client_secret
    }

    with open(credentials_path, 'w') as file:
        file.write(json.dumps(credentials_as_dict))

    return credentials


def authorize(client_secret='client_secret.json', service_account_file=None, credentials_directory='',
              scopes=('https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive')):
    """

    :param client_secret:           Path to client secret file for OAuth. Ignored if service_account_file is given.
    :param service_account_file:    Path of the service account file.
    :param credentials_directory:   Where credential-tokens are stored (only relevant for OAuth authentication).
    :param scopes:
    :return:
    """
    if service_account_file is not None:
        credentials = service_account.Credentials.from_service_account_file(service_account_file, scopes=scopes)
    else:
        credentials = _get_user_authentication_credentials(client_secret, scopes, credentials_directory)

    return Client(credentials)