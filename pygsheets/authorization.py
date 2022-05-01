# -*- coding: utf-8 -*-.
import os
import json
import warnings

from google.oauth2 import service_account
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import Flow, InstalledAppFlow
from google.auth.transport.requests import Request

from pygsheets.client import Client

try:
    input = raw_input
except NameError:
    pass


def _get_user_authentication_credentials(client_secret_file, scopes, credential_directory=None, local=False):
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

    credentials = None
    if os.path.exists(credentials_path):
        # expect these to be valid. may expire at some point, but should be refreshed by google api client...
        credentials = Credentials.from_authorized_user_file(credentials_path, scopes=scopes)

    if credentials:
        if credentials.expired and credentials.refresh_token:
            credentials.refresh(Request())
    else:
        if local:
            flow = InstalledAppFlow.from_client_secrets_file(client_secret_file, scopes)
            credentials = flow.run_local_server()
        else:
            flow = Flow.from_client_secrets_file(client_secret_file, scopes=scopes,
                                                 redirect_uri='urn:ietf:wg:oauth:2.0:oob')
            auth_url, _ = flow.authorization_url(prompt='consent')

            print('Please go to this URL and finish the authentication flow: {}'.format(auth_url))
            code = input('Enter the authorization code: ')
            flow.fetch_token(code=code)
            credentials = flow.credentials

    # Save the credentials for the next run
    credentials_as_dict = {
        'token': credentials.token,
        'refresh_token': credentials.refresh_token,
        'id_token': credentials.id_token,
        'token_uri': credentials.token_uri,
        'client_id': credentials.client_id,
        'client_secret': credentials.client_secret
    }
    try:
        with open(credentials_path, 'w') as file:
            file.write(json.dumps(credentials_as_dict))
    except OSError:
        print("Unable to save the credentials to file-system")

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
              service_account_env_var=None,
              service_account_json=None,
              credentials_directory='',
              scopes=_SCOPES,
              custom_credentials=None,
              local=False,
              **kwargs):

    """Authenticate this application with a google account.

    See general authorization documentation for details on how to attain the necessary files.

    :param client_secret:           Location of the oauth2 credentials file.
    :param service_account_file:    Location of a service account file.
    :param service_account_env_var: Use an environment variable to provide service account credentials.
    :param service_account_json:    pass in json string directly; could use aws secret manager or azure key vault to
                                    store value
    :param credentials_directory:   Location of the token file created by the OAuth2 process. Use 'global' to store in
                                    global location, which is OS dependent. Default None will store token file in
                                    current working directory. Please note that this is override your client secret.
    :param custom_credentials:      A custom or pre-made credentials object. Will ignore all other params.
    :param scopes:                  The scopes for which the authentication applies.
    :param local:                   If local then a browser will be opened to autheticate
    :param kwargs:                  Parameters to be handed into the client constructor.
    :returns:                       :class:`Client`

    .. warning::
        The `credentials_directory` overrides `client_secret`. So you might be accidently using a different credential
        than intended, if you are using global `credentials_directory` in more than one script.

    """

    for key in kwargs:
        if key in ['outh_file', 'outh_creds_store', 'service_file', 'credentials']:
            warnings.warn('The argument {} is deprecated. Use {} instead.'.format(key, _deprecated_keyword_mapping[key])
                          , category=DeprecationWarning)
    client_secret = kwargs.pop('outh_file', client_secret)
    service_account_file = kwargs.pop('service_file', service_account_file)
    credentials_directory = kwargs.pop('outh_creds_store', credentials_directory)
    custom_credentials = kwargs.pop('credentials', custom_credentials)

    if custom_credentials is not None:
        credentials = custom_credentials
    elif service_account_env_var is not None:
        service_account_info = json.loads(os.environ[service_account_env_var])
        credentials = service_account.Credentials.from_service_account_info(
            service_account_info, scopes=scopes)
    elif service_account_json is not None:
        service_account_info = json.loads(service_account_json)
        credentials = service_account.Credentials.from_service_account_info(
            service_account_info, scopes=scopes)
    elif service_account_file is not None:
        credentials = service_account.Credentials.from_service_account_file(service_account_file, scopes=scopes)
    else:
        credentials = _get_user_authentication_credentials(client_secret, scopes, credentials_directory, local)

    return Client(credentials, **kwargs)
