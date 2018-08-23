# -*- coding: utf-8 -*-.
import os
import json
import warnings

from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow

from pygsheets.client import Client


def _get_user_authentication_credentials(client_secret_file, scopes, credential_directory=None):
    """Returns user credentials."""
    if credential_directory is None:
        credential_directory = os.getcwd()
    elif credential_directory == 'global':
        home_dir = os.path.expanduser('~')
        credential_directory = os.path.join(home_dir, '.credentials')
        if not os.path.exists(credential_directory):
            os.makedirs(credential_directory)
    else:
        pass

    credentials_path = os.path.join(credential_directory, 'sheets.googleapis.com-python.json')  # TODO Change hardcoded name?

    if os.path.exists(credentials_path):
        # expect these to be valid. may expire at some point, but should be refreshed by google api client...
        return Credentials.from_authorized_user_file(credentials_path, scopes=scopes)

    flow = Flow.from_client_secrets_file(client_secret_file, scopes=scopes,
                                         redirect_uri='urn:ietf:wg:oauth:2.0:oob')

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


_SCOPES = ('https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive')

_deprecated_keyword_mapping = {
    'outh_file': 'client_secret',
    'outh_creds_store': 'credentials_directory',
    'service_file': 'service_account_file',
    'credentials': 'custom_credentials'
}


def authorize(client_secret='client_secret.json',
              service_account_file=None,
              credentials_directory='',
              scopes=_SCOPES,
              custom_credentials=None,
              **kwargs):

    """Authenticate this application with a google account.

    See general authorization documentation for details on how to attain the necessary files.

    :param client_secret:           Location of the oauth2 credentials file.
    :param service_account_file:    Location of a service account file.
    :param credentials_directory:   Location of the token file created by the OAuth2 process. Use 'global' to store in
                                    global location, which is OS dependent. Default None will store token file in
                                    current working directory. Please note that this is override your client secret.
    :param custom_credentials:      A custom or pre-made credentials object. Will ignore all other params.
    :param scopes:                  The scopes for which the authentication applies.
    :param kwargs:                  Parameters to be handed into the client constructor.
    :returns:                       :class:`Client`
    """
    if 'outh_file' in kwargs:
        warnings.warn('The argument {} is deprecated. Use {} instead.'.
                      format('outh_file', _deprecated_keyword_mapping['outh_file']), category=DeprecationWarning)
        client_secret = kwargs['outh_file']
        del kwargs['outh_file']
    if 'outh_creds_store' in kwargs:
        warnings.warn('The argument {} is deprecated. Use {} instead.'.
                      format('outh_creds_store', _deprecated_keyword_mapping['outh_creds_store']), category=DeprecationWarning)
        credentials_directory = kwargs['outh_creds_store']
        del kwargs['outh_creds_store']
    if 'service_file' in kwargs:
        warnings.warn('The argument {} is deprecated. Use {} instead.'.
                      format('service_file', _deprecated_keyword_mapping['service_file']), category=DeprecationWarning)
        service_account_file = kwargs['service_file']
        del kwargs['service_file']
    if 'credentials' in kwargs:
        warnings.warn('The argument {} is deprecated. Use {} instead.'.
                      format('credentials', _deprecated_keyword_mapping['credentials']), category=DeprecationWarning)
        custom_credentials = kwargs['credentials']
        del kwargs['credentials']
            
    if custom_credentials is not None:
        credentials = custom_credentials
    elif service_account_file is not None:
        credentials = service_account.Credentials.from_service_account_file(service_account_file, scopes=scopes)
    else:
        credentials = _get_user_authentication_credentials(client_secret, scopes, credentials_directory)

    return Client(credentials)
