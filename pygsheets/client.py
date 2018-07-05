# -*- coding: utf-8 -*-.
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


from google.auth.transport.requests import AuthorizedSession
from json import load as jload
from googleapiclient import discovery

GOOGLE_SHEET_CELL_UPDATES_LIMIT = 50000

_url_key_re_v1 = re.compile(r'key=([^&#]+)')
_url_key_re_v2 = re.compile(r"/spreadsheets/d/([a-zA-Z0-9-_]+)")
_email_patttern = re.compile(r"\"?([-a-zA-Z0-9.`?{}]+@[-a-zA-Z0-9.]+\.\w+)\"?")
# _domain_pattern = re.compile("(?!-)[A-Z\d-]{1,63}(?<!-)$", re.IGNORECASE)


class Client(object):
    """Create or access Google speadsheets.

    Exposes members to create new spreadsheets or open existing ones. Use `authorize` to instantiate an instance of this
    class.

    >>> import pygsheets
    >>> c = pygsheets.authorize()

    The sheet API service object is stored in the sheet property and the drive API service object in the drive property.

    >>> c.sheet.get('<SPREADSHEET ID>')
    >>> c.drive.delete('<FILE ID>')

    :param oauth:                   An credentials object created by the `oauth2client library <https://github.com/google/oauth2client>`_.
    :param http_client:             (Optional) The object responsible to handle HTTP requests. Defaults to the
                                    googleapiclient http-object.
    :param retries:                 (Optional) Number of times to retry a connection before raising a TimeOut error.
    """

    spreadsheet_cls = Spreadsheet

    def __init__(self, credentials, retries=3, no_cache=False):
        if no_cache:
            cache = None
        else:
            cache = os.path.join(tempfile.gettempdir(), str(uuid.uuid4()))
        if os.name == "nt":
            cache = "\\\\?\\" + cache
        self.logger = logging.getLogger(__name__)

        http = AuthorizedSession(credentials)
        data_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

        self.sheet = SheetAPIWrapper(http, data_path, retries=retries)
        self.drive = DriveAPIWrapper(http, data_path)

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
        """Get a list of all spreadsheet ids present in the Google Drive or TeamDrive accessed."""
        return [x['id'] for x in self.drive.spreadsheet_metadata(query)]

    def spreadsheet_titles(self, query=None):
        """Get a list of all spreadsheet titles present in the Google Drive or TeamDrive accessed."""
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

        In a case where there are several sheets with the same title, the first one found is returned.

        >>> import pygsheets
        >>> c = pygsheets.authorize()
        >>> c.open('TestSheet')

        :param title:                           A title of a spreadsheet.

        :returns:                               :class:`~pygsheets.Spreadsheet`
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
        :returns:                               :class:`~pygsheets.Spreadsheet`
        :raises pygsheets.SpreadsheetNotFound:  The given spreadsheet ID was not found.
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
        """Return a json representation of the spreadsheet.

        `See Reference for details <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#Spreadsheet>`_.
        """
        return self.sheet.get(key, fields='properties,sheets/properties,spreadsheetId,namedRanges',
                              includeGridData=False)

    def get_range(self, spreadsheet_id,
                  value_range,
                  major_dimension='ROWS',
                  value_render_option=ValueRenderOption.FORMATTED_VALUE,
                  date_time_render_option=DateTimeRenderOption.FORMATTED_STRING):
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
                                            values were fetched. Values are dynamically typed as int, float or string.
        """
        result = self.sheet.values_get(spreadsheet_id, value_range, major_dimension, value_render_option,
                                       date_time_render_option)
        try:
            return result['values']
        except KeyError:
            self.logger.warning('No values were fetched from the specified range: %s.', value_range)
            return [['']]

    def update_sheet_properties(self, spreadsheet_id, propertyObj, fields_to_update=
                                'title,hidden,gridProperties,tabColor,rightToLeft', batch=False):
        """wrapper for updating sheet properties"""
        request = {"updateSheetProperties": {"properties": propertyObj, "fields": fields_to_update}}
        return self.sh_batch_update(spreadsheet_id, request, None, batch)

    def sh_get_ssheet(self, spreadsheet_id, fields, ranges=None, include_data=True):
        request = self.service.spreadsheets().get(spreadsheetId=spreadsheet_id,
                                                  fields=fields, ranges=ranges, includeGridData=include_data)
        return self._execute_request(spreadsheet_id, request, False)

    def sh_update_range(self, spreadsheet_id, body, batch, parse=True):
        cformat = 'USER_ENTERED' if parse else 'RAW'
        batch_limit = GOOGLE_SHEET_CELL_UPDATES_LIMIT
        if body['majorDimension'] == 'ROWS':
            batch_length = int(batch_limit / len(body['values'][0]))  # num of rows to include in a batch
            num_rows = len(body['values'])
        else:
            batch_length = int(batch_limit / len(body['values']))  # num of rows to include in a batch
            num_rows = len(body['values'][0])
        if len(body['values'])*len(body['values'][0]) <= batch_limit:
            final_request = self.service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, range=body['range'],
                                                                        valueInputOption=cformat, body=body)
            self._execute_request(spreadsheet_id, final_request, batch)
        else:
            if batch_length == 0:
                raise AssertionError("num_columns < "+str(GOOGLE_SHEET_CELL_UPDATES_LIMIT))
            values = body['values']
            title, value_range = body['range'].split('!')
            value_range_start, value_range_end = value_range.split(':')
            value_range_end = list(format_addr(str(value_range_end), output='tuple'))
            value_range_start = list(format_addr(str(value_range_start), output='tuple'))
            max_rows = value_range_end[0]
            start_row = value_range_start[0]
            for batch_start in range(0, num_rows, batch_length):
                if body['majorDimension'] == 'ROWS':
                    body['values'] = values[batch_start:batch_start+batch_length]
                else:
                    body['values'] = [col[batch_start:batch_start + batch_length] for col in values]
                value_range_start[0] = batch_start + start_row
                value_range_end[0] = min(batch_start+batch_length, max_rows) + start_row
                body['range'] = title+'!'+format_addr(tuple(value_range_start), output='label')+':' + \
                                format_addr(tuple(value_range_end), output='label')
                final_request = self.service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, body=body,
                                                                            range=body['range'], valueInputOption=cformat)
                self._execute_request(spreadsheet_id, final_request, batch)

    def sh_batch_clear(self, spreadsheet_id, body, batch=False):
        """wrapper around batch clear"""
        final_request = self.service.spreadsheets().values().batchClear(spreadsheetId=spreadsheet_id, body=body)
        self._execute_request(spreadsheet_id, final_request, batch)

    def sh_copy_worksheet(self, src_ssheet, src_worksheet, dst_ssheet):
        """wrapper of sheets copyTo"""
        final_request = self.service.spreadsheets().sheets().copyTo(spreadsheetId=src_ssheet, sheetId=src_worksheet,
                                                                    body={"destinationSpreadsheetId": dst_ssheet})
        return self._execute_request(dst_ssheet, final_request, False)

    def sh_append(self, spreadsheet_id, body, rranage, replace=False, batch=False):
        """wrapper around batch append"""
        if replace:
            inoption = "OVERWRITE"
        else:
            inoption = "INSERT_ROWS"
        final_request = self.service.spreadsheets().values().append(spreadsheetId=spreadsheet_id, range=rranage, body=body,
                                                                    insertDataOption=inoption, includeValuesInResponse=False,
                                                                    valueInputOption="USER_ENTERED")
        self._execute_request(spreadsheet_id, final_request, batch)

    # @TODO use batch update more efficiently
    def sh_batch_update(self, spreadsheet_id, request, fields=None, batch=False):
        if type(request) == list:
            body = {'requests': request}
        else:
            body = {'requests': [request]}
        final_request = self.service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id, body=body,
                                                                fields=fields)
        return self._execute_request(spreadsheet_id, final_request, batch)

    def _execute_request(self, spreadsheet_id, request, batch):
        """Execute the request"""
        if batch:
            try:
                self.batch_requests[spreadsheet_id].append(request)
            except KeyError:
                self.batch_requests[spreadsheet_id] = [request]
            self.logger.debug("batch request added")
        else:
            self.logger.debug("request : " + request.uri)
            for i in range(self.retries):
                try:
                    response = request.execute()
                except Exception as e:
                    if repr(e).find('timed out') == -1:
                        raise
                    if i == self.retries-1:
                        self.logger.exception("Timeout")
                        raise RequestError("Timeout : " + repr(e))
                    self.logger.debug("Cant connect, retrying - #" + str(i))
                else:
                    return response

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
