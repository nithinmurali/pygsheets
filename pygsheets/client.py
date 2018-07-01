# -*- coding: utf-8 -*-.

"""
pygsheets.client
~~~~~~~~~~~~~~~~

This module contains Client class responsible for communicating with
Google SpreadSheet API.

"""

import re
import warnings
import os
import tempfile
import uuid
import logging


from pygsheets.drive import DriveAPIWrapper
from pygsheets.sheet import SheetAPIWrapper
from pygsheets.spreadsheet import Spreadsheet
from pygsheets.exceptions import (AuthenticationError, SpreadsheetNotFound,
                                  NoValidUrlKeyFound, RequestError,
                                  InvalidArgumentValue)
from pygsheets.custom_types import *
from pygsheets.utils import format_addr

import httplib2
from json import load as jload
from googleapiclient import discovery
from oauth2client.file import Storage
from oauth2client import client
from oauth2client import tools
from oauth2client.service_account import ServiceAccountCredentials
try:
    import argparse
    flags = tools.argparser.parse_args([])
except ImportError:
    flags = None

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']
GOOGLE_SHEET_CELL_UPDATES_LIMIT = 50000

_url_key_re_v1 = re.compile(r'key=([^&#]+)')
_url_key_re_v2 = re.compile(r"/spreadsheets/d/([a-zA-Z0-9-_]+)")
_email_patttern = re.compile(r"\"?([-a-zA-Z0-9.`?{}]+@[-a-zA-Z0-9.]+\.\w+)\"?")
# _domain_pattern = re.compile("(?!-)[A-Z\d-]{1,63}(?<!-)$", re.IGNORECASE)


class Client(object):
    """An instance of this class communicates with Google API.

    :param oauth:                   An OAuth2 credential object. Credential objects are those created by the
                                    oauth2client library. `https://github.com/google/oauth2client`_
    :param http_client: (optional)  A object capable of making HTTP requests
    :param retries: (optional)      The number of times connection will be
                                    tried before raising a timeout error.

    >>> c = authorize()

    """

    spreadsheet_cls = Spreadsheet

    def __init__(self, oauth, http_client=None, retries=1, no_cache=False):
        if no_cache:
            cache = None
        else:
            cache = os.path.join(tempfile.gettempdir(), str(uuid.uuid4()))
        if os.name == "nt":
            cache = "\\\\?\\" + cache

        self.oauth = oauth
        self.logger = logging.getLogger(__name__)
        http_client = http_client or httplib2.Http(cache=cache, timeout=20)
        http = self.oauth.authorize(http_client)
        data_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

        self.sheet = SheetAPIWrapper(http, data_path)
        with open(os.path.join(data_path, "sheets_discovery.json")) as jd:
            self.service = discovery.build_from_document(jload(jd), http=http)
        self.drive = DriveAPIWrapper(http, data_path)
        self.batch_requests = dict()
        self.retries = retries

    @property
    def teamDriveId(self):
        """ Enable team drive support

            Deprecated: use client.drive.enable_team_drive(team_drive_id=?)
        """
        return self.drive.team_drive_id

    @teamDriveId.setter
    def teamDriveId(self, value):
        warnings.warn("Depricated  please use drive.enable_team_drive")
        self.drive.enable_team_drive(value)

    def spreadsheet_ids(self, query=None):
        """A list of all the ids of spreadsheets present in the users drive or TeamDrive."""
        return [x['id'] for x in self.drive.spreadsheet_metadata(query)]

    def spreadsheet_titles(self, query=None):
        """A list of all the titles of spreadsheets present in the users drive or TeamDrive."""
        return [x['name'] for x in self.drive.spreadsheet_metadata(query)]

    def create(self, title, template=None, folder=None, **kwargs):
        """Create a new spreadsheet.

        The title will always be set to the given value (even overwriting the templates title). The template
        can either be a `spreadsheet resource <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#resource-spreadsheet>`
        or an instance of `~pygsheets.Spreadsheet`. In both cases undefined values will be ignored.

        :param title:       Title of the new spreadsheet.
        :param template:    A template to create the new spreadsheet from.
        :param folder:      The Id of the folder this sheet will be stored in.
        :param kwargs:      Standard parameters (see reference for details).
        :return: `~pygsheets.Spreadsheet`
        """
        result = self.sheet.create(title, template=template, **kwargs)
        if folder:
            self.drive.move_file(result['spreadsheetId'],
                                 old_folder=self.drive.spreadsheet_metadata(query="name = '" + title + "'")[0]['parents'][0],
                                 new_folder=folder)
        return self.spreadsheet_cls(self, jsonsheet=result)

    def open(self, title):
        """Open a spreadsheet by title.

        In a case where there are several sheets with the same title, the first one is returned.

        >>> import pygsheets
        >>> c = pygsheets.authorize()
        >>> c.open('TestSheet')

        :param title:                           A title of a spreadsheet.
        :returns                                :class:`~pygsheets.Spreadsheet`.
        :raises pygsheets.SpreadsheetNotFound:  No spreadsheet with the given title was found.
        """
        try:
            spreadsheet = list(filter(lambda x: x['name'] == title, self.drive.spreadsheet_metadata()))[0]
            return self.open_by_key(spreadsheet['id'])
        except KeyError:
            raise SpreadsheetNotFound('Could not find a spreadsheet with title %s.' % title)

    def open_by_key(self, key):
        """Open a spreadsheet by key.

        >>> import pygsheets
        >>> c = pygsheets.authorize()
        >>> c.open_by_key('0BmgG6nO_6dprdS1MN3d3MkdPa142WFRrdnRRUWl1UFE')

        :param key:                             The key of a spreadsheet. (can be found in the sheet URL)
        :returns                                :class:`~pygsheets.Spreadsheet`
        :raises pygsheets.SpreadsheetNotFound:  No spreadsheet with the given key was found.
        """
        response = self.sheet.get(key,
                                  fields='properties,sheets/properties,spreadsheetId,namedRanges',
                                  includeGridData=False)
        return self.spreadsheet_cls(self, response)

    def open_by_url(self, url):
        """Open a spreadsheet by URL.

        >>> import pygsheets
        >>> c = pygsheets.authorize()
        >>> c.open_by_url('https://docs.google.com/spreadsheet/ccc?key=0Bm...FE&hl')

        :param url:                             URL of a spreadsheet as it appears in a browser.
        :returns:                               :class:`~pygsheets.Spreadsheet`
        :raises pygsheets.SpreadsheetNotFound:  No spreadsheet was found with the given URL.
        """
        m1 = _url_key_re_v1.search(url)
        if m1:
            return self.open_by_key(m1.group(1))

        else:
            m2 = _url_key_re_v2.search(url)
            if m2:
                return self.open_by_key(m2.group(1))
            else:
                raise NoValidUrlKeyFound

    def open_all(self, query=''):
        """Opens all available spreadsheets.

        Result can be filtered when specifying the query parameter. On the details on how to form the query:

        `Reference <https://developers.google.com/drive/v3/web/search-parameters>`_

        :param query:   (Optional) Can be used to filter the returned metadata.
        :returns:       A list of :class:`~pygsheets.Spreadsheet`.
        """
        return [self.open_by_key(key) for key in self.spreadsheet_ids(query=query)]

    def open_as_json(self, key):
        """Returns the json response from a spreadsheet.

        See API Reference on how it is constructed.
        """
        return self.sheet.get(key, fields='properties,sheets/properties,spreadsheetId,namedRanges',
                              includeGridData=False)

    def get_range(self, spreadsheet_id, value_range, major_dimension='ROWS',
                  value_render_option=ValueRenderOption.FORMATTED_VALUE):
        """Returns a range of values from a spreadsheet. The caller must specify the spreadsheet ID and a range.

        `Reference <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/get>`_

        :param spreadsheet_id:              The ID of the spreadsheet to retrieve data from.
        :param value_range:                 The A1 notation of the values to retrieve.
        :param major_dimension:             The major dimension that results should use.
                                            For example, if the spreadsheet data is: A1=1,B1=2,A2=3,B2=4, then
                                            requesting range=A1:B2,majorDimension=ROWS will return [[1,2],[3,4]],
                                            whereas requesting range=A1:B2,majorDimension=COLUMNS will return
                                            [[1,3],[2,4]].
        :param value_render_option:         How values should be represented in the output. The default
                                            render option is ValueRenderOption.FORMATTED_VALUE.
        :param date_time_render_option:     How dates, times, and durations should be represented in the output.
                                            This is ignored if valueRenderOption is FORMATTED_VALUE. The default
                                            dateTime render option is [DateTimeRenderOption.SERIAL_NUMBER].
        :return:                            An array of arrays with the values fetched. Returns an empty array if no
                                            values were fetched.
        """
        result = self.sheet.values_get(spreadsheet_id, value_range, major_dimension, value_render_option)
        try:
            return result['values']
        except KeyError:
            self.logger.warning('No values were fetched from the specified range: %s.', value_range)
            return [['']]

    # @TODO combine adj batch requests into 1
    def send_batch(self, spreadsheet_id):
        """Send all batched requests
        :param spreadsheet_id: id of ssheet batch requests to send
        :return: False if no batched requests
        """
        if spreadsheet_id not in self.batch_requests or self.batch_requests == []:
            return False

        def callback(request_id, response, exception):
            if exception:
                print(exception)
            else:
                self.logger.debug("batch request #" + request_id + " completed")
                pass
        i = 0
        batch_req = self.service.new_batch_http_request(callback=callback)
        for req in self.batch_requests[spreadsheet_id]:
            batch_req.add(req)
            i += 1
            if i % 100 == 0:  # as there is an limit of 100 requests
                i = 0
                batch_req.execute()
                batch_req = self.service.new_batch_http_request(callback=callback)
        batch_req.execute()
        self.batch_requests[spreadsheet_id] = []
        return True


def get_outh_credentials(client_secret_file, credential_dir=None, outh_nonlocal=False):
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    :param client_secret_file: path to outh2 client secret file
    :param credential_dir: path to directory where tokens should be stored
                           'global' if you want to store in system-wide location
                           None if you want to store in current script directory
    :param outh_nonlocal: if the authorization should be done in another computer,
                     this will provide a url which when run will ask for credentials

    :return
        Credentials, the obtained credential.
    """
    lflags = flags
    if credential_dir == 'global':
        home_dir = os.path.expanduser('~')
        credential_dir = os.path.join(home_dir, '.credentials')
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
    elif not credential_dir:
        credential_dir = os.getcwd()
    else:
        pass

    # verify credentials directory
    if not os.path.isdir(credential_dir):
        raise IOError(2, "Credential directory does not exist.", credential_dir)
    credential_path = os.path.join(credential_dir, 'sheets.googleapis.com-python.json')

    # check if refresh token file is passed
    with warnings.catch_warnings():
        warnings.simplefilter("ignore")
        try:
            store = Storage(client_secret_file)
            credentials = store.get()
        except KeyError:
            credentials = None

    # else try to get credentials from storage
    if not credentials or credentials.invalid:
        try:
            with warnings.catch_warnings():
                warnings.simplefilter("ignore")
                store = Storage(credential_path)
                credentials = store.get()
        except KeyError:
            credentials = None

    # else get the credentials from flow
    if not credentials or credentials.invalid:
        # verify client secret file
        if not os.path.isfile(client_secret_file):
            raise IOError(2, "Client secret file does not exist.", client_secret_file)
        # execute flow
        flow = client.flow_from_clientsecrets(client_secret_file, SCOPES)
        flow.user_agent = 'pygsheets'
        if lflags:
            lflags.noauth_local_webserver = outh_nonlocal
            credentials = tools.run_flow(flow, store, lflags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def authorize(outh_file='client_secret.json', outh_creds_store=None, outh_nonlocal=False, service_file=None,
              credentials=None, **client_kwargs):
    """Login to Google API using OAuth2 credentials.

    This function instantiates :class:`Client` and performs authentication.

    :param outh_file: path to outh2 credentials file, or tokens file
    :param outh_creds_store: path to directory where tokens should be stored
                           'global' if you want to store in system-wide location
                           None if you want to store in current script directory
    :param outh_nonlocal: if the authorization should be done in another computer,
                         this will provide a url which when run will ask for credentials
    :param service_file: path to service credentials file
    :param credentials: outh2 credentials object

    :param no_cache: (http client arg) do not ask http client to use a cache in tmp dir, useful for environments where
                     filesystem access prohibited
                     default: False

    :returns: :class:`Client` instance.

    """
    # @TODO handle exceptions
    if not credentials:
        if service_file:
            with open(service_file) as data_file:
                data = jload(data_file)
            credentials = ServiceAccountCredentials.from_json_keyfile_name(service_file, SCOPES)
        elif outh_file:
            credentials = get_outh_credentials(client_secret_file=outh_file, credential_dir=outh_creds_store,
                                               outh_nonlocal=outh_nonlocal)
        else:
            raise AuthenticationError
    rclient = Client(oauth=credentials, **client_kwargs)
    return rclient


# @TODO
def public():
    """
    return a :class:`Client` which can acess only publically shared sheets

    """
    pass
