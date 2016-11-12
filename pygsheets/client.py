# -*- coding: utf-8 -*-.

"""
pygsheets.client
~~~~~~~~~~~~~~~~

This module contains Client class responsible for communicating with
Google SpreadSheet API.

"""
import re
import warnings

from .models import Spreadsheet
from .exceptions import (AuthenticationError, SpreadsheetNotFound,
                         NoValidUrlKeyFound,
                         InvalidArgumentValue, InvalidUser)
# from custom_types import *

import httplib2
import os
from json import load as jload

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools
from oauth2client.service_account import ServiceAccountCredentials
try:
    import argparse
    flags = tools.argparser.parse_args([])
except ImportError:
    flags = None

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

_url_key_re_v1 = re.compile(r'key=([^&#]+)')
_url_key_re_v2 = re.compile(r"/spreadsheets/d/([a-zA-Z0-9-_]+)")


class Client(object):

    """An instance of this class communicates with Google API.

    :param auth: An OAuth2 credential object. Credential objects are those created by the
                 oauth2client library. https://github.com/google/oauth2client

    >>> c = pygsheets.Client(auth=OAuthCredentialObject)


    """
    def __init__(self, auth):
        self.auth = auth
        http = auth.authorize(httplib2.Http(cache="/tmp/.pygsheets_cache", timeout=10))
        discoveryurl = ('https://sheets.googleapis.com/$discovery/rest?'
                        'version=v4')
        self.service = discovery.build('sheets', 'v4', http=http,
                                       discoveryServiceUrl=discoveryurl)
        self.driveService = discovery.build('drive', 'v3', http=http)
        self._spreadsheeets = []
        self._fetch_sheets()
        self.sendBatch = False
        self.spreadsheetId = None  # @TODO remove this
    
    def _fetch_sheets(self):
        """
        fetch all the sheets info from user's gdrive

        :returns: None
        """
        results = self.driveService.files().list(corpus='user', pageSize=500,
                                                 q="mimeType='application/vnd.google-apps.spreadsheet'",
                                                 fields="files(id, name)").execute()
        try:
            results = results['files']
        except KeyError:
            results = []
        self._spreadsheeets = results

    def create(self, title):
        """Creates a spreadsheet, returning a :class:`~pygsheets.Spreadsheet` instance.

        :param title: A title of a spreadsheet.
        """
        body = {'properties': {'title': title}}
        result = self.service.spreadsheets().create(body=body).execute()
        self._spreadsheeets.append({'name': title, "id": result['spreadsheetId']})
        return Spreadsheet(self, jsonsheet=result)

    def delete(self, title=None, id=None):
        """Deletes, a spreadsheet by title or id.

        :param title: title of a spreadsheet.
        :param id: id of a spreadsheet this takes precedence if both given.

        :raise pygsheets.SpreadsheetNotFound: if no spreadsheet is found.
        """
        if not id and not title:
            raise SpreadsheetNotFound
        if id:
            if len([x for x in self._spreadsheeets if x["id"] == id]) == 0:
                raise SpreadsheetNotFound
        try:
            if title and not id:
                id = [x["id"] for x in self._spreadsheeets if x["name"] == title][0]
        except IndexError:
            raise SpreadsheetNotFound

        self.driveService.files().delete(fileId=id).execute()
        self._spreadsheeets.remove([x for x in self._spreadsheeets if x["name"] == title][0])

    def open(self, title):
        """Opens a spreadsheet, returning a :class:`~pygsheets.Spreadsheet` instance.

        :param title: A title of a spreadsheet.

        If there's more than one spreadsheet with same title the first one
        will be opened.

        :raises pygsheets.SpreadsheetNotFound: if no spreadsheet with
                                             specified `title` is found.

        >>> c = pygsheets.Client(auth=('user@example.com', 'qwertypassword'))
        >>> c.login()
        >>> c.open('My fancy spreadsheet')

        """
        try:
            ssheet_id = [x['id'] for x in self._spreadsheeets if x["name"] == title][0]
            return self.open_by_key(ssheet_id)
        except IndexError:
            self._fetch_sheets()
            try:
                return [Spreadsheet(self, id=x['id']) for x in self._spreadsheeets if x["name"] == title][0]
            except IndexError:
                raise SpreadsheetNotFound(title)

    def open_by_key(self, key, returnas='spreadsheet'):
        """Opens a spreadsheet specified by `key`, returning a :class:`~pygsheets.Spreadsheet` instance.

        :param key: A key of a spreadsheet as it appears in a URL in a browser.
        :param returnas: return as spreadhseet of json object
        :raises pygsheets.SpreadsheetNotFound: if no spreadsheet with
                                             specified `key` is found.

        >>> c = pygsheets.authorize()
        >>> c.open_by_key('0BmgG6nO_6dprdS1MN3d3MkdPa142WFRrdnRRUWl1UFE')

        """
        result = ''
        try:
            result = self.service.spreadsheets().get(spreadsheetId=key,
                                                     fields='properties,sheets/properties,spreadsheetId')\
                                                    .execute()
        except Exception as e:
            raise e
        if returnas == 'spreadsheet':
            return Spreadsheet(self, result)
        elif returnas == 'json':
            return result
        else:
            raise InvalidArgumentValue(returnas)
    
    def open_by_url(self, url):
        """Opens a spreadsheet specified by `url`,

        :param url: URL of a spreadsheet as it appears in a browser.

        :raises pygsheets.SpreadsheetNotFound: if no spreadsheet with
                                             specified `url` is found.
        :returns: a `~pygsheets.Spreadsheet` instance.

        >>> c = pygsheets.authorize()
        >>> c.open_by_url('https://docs.google.com/spreadsheet/ccc?key=0Bm...FE&hl')

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

    def open_all(self, title=None):
        """
        Opens all available spreadsheets,

        :param title: (optional) If specified can be used to filter spreadsheets by title.

        :returns: list of a :class:`~pygsheets.Spreadsheet` instances
        """
        return [Spreadsheet(self, id=x['id']) for x in self._spreadsheeets if ((title is None) or (x['name'] == title))]

    def list_ssheets(self):
        return self._spreadsheeets

    # @TODO
    def start_batch(self):
        self.sendBatch = True

    # @TODO
    def stop_batch(self):
        self.sendBatch = False

    def update_range(self, range, values, majorDim='ROWS', parse=True):
        """

        :param range: range in A1 format
        :param values: values as 2d array
        :param majorDim: major dimesion
        :param parse: should the vlaues be parsed or trated as strings eg. formulas
        :returns:
        """
        if self.sendBatch:
            pass
        else:
            body = dict()
            body['range'] = range
            body['majorDimension'] = str(majorDim)
            body['values'] = values
            cformat = 'USER_ENTERED' if parse else 'RAW'
            result = self.service.spreadsheets().values().update(spreadsheetId=self.spreadsheetId, range=body['range'],
                                                                 valueInputOption=cformat, body=body).execute()

    def get_range(self, range, majorDim='ROWS', value_render='FORMATTED_VALUE'):
        """
         fetches  values from sheet.
        :param range: range in A! format
        :param majorDim: if the major dimension is rows or cols
        :param value_render: format of output values

        :type majorDim: 'ROWS' or 'COLUMNS'
        :type value_render: 'FORMATTED_VALUE', 'UNFORMATTED_VALUE', 'FORMULA'

        :returns: 2d array
        """
        if value_render not in ['FORMATTED_VALUE', 'UNFORMATTED_VALUE', 'FORMULA']:
            raise InvalidArgumentValue
        if not self.spreadsheetId:
            return None

        if self.sendBatch:
            pass
        else:
            result = self.service.spreadsheets().values().get(spreadsheetId=self.spreadsheetId, range=range,
                        majorDimension=majorDim, valueRenderOption=value_render, dateTimeRenderOption=None).execute()
            try:
                return result['values']
            except KeyError:
                return [['']]

    def insertdim(self, sheetId, majorDim, startindex, endIndex, inheritbefore=False):
        if self.sendBatch:
            pass
        else:
            body = {'requests': [{'insertDimension': {'inheritFromBefore': False,
                    'range': {'sheetId': sheetId, 'dimension': majorDim, 'endIndex': endIndex, 'startIndex': startindex}
                    }}]}
            self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=body).execute()

    def update_sheet_properties(self, propertyObj, fieldsToUpdate='title,hidden,gridProperties,tabColor,rightToLeft'):
        requests = {"updateSheetProperties": {"properties": propertyObj, "fields": fieldsToUpdate}}
        body = {'requests': [requests]}
        self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=body).execute()

    def add_worksheet(self, title, rows=1000, cols=26):
        requests = {"addSheet": {"properties": {'title': title, "gridProperties": {"rowCount": rows, "columnCount": cols}}}}
        body = {'requests': [requests]}
        result = self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=body, fields='replies/addSheet').execute()
        return result['replies'][0]['addSheet']['properties']

    def del_worksheet(self, sheetId):
        requests = {"deleteSheet": {'sheetId': sheetId}}
        body = {'requests': [requests]}
        result = self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=body).execute()
        return result

    # @TODO impliment expirationTime
    def add_permission(self, file_id, addr, role='reader', is_group=False, expirationTime=None):
        """
        create/update permission for user/group/domain

        :param file_id: id of the file whose permissions to manupulate
        :param addr: this is the email for user/group and domain adress for domains
        :param role: permission to be applied
        :param expirationTime: (Not Implimented) time until this permission should last
        :param is_group: boolean , Is this addr a group; used only when email provided

        """
        if re.match('^[_a-z0-9-]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,4})$', addr):
            permission = {
                'type': 'user',
                'role': role,
                'emailAddress': addr
            }
            if is_group:
                permission['type'] = 'group'
        elif re.match('[a-zA-Z\d-]{,63}(\.[a-zA-Z\d-]{,63})*', addr):
            permission = {
                'type': 'domain',
                'role': role,
                'domain': addr
            }
        else:
            print ("invalid adress: %s" % addr)
            return False
        self.driveService.permissions().create(
            fileId=file_id,
            body=permission,
            fields='id',
        ).execute()

    def list_permissions(self, file_id):
        """
        list permissions of a file
        :param file_id: file id
        :returns:
        """
        result = self.driveService.permissions().list(fileId=file_id,
                                                      fields='permissions(domain,emailAddress,expirationTime,id,role,type)'
                                                      ).execute()
        return result

    def remove_permissions(self, file_id, addr, permisssions_in=None):
        """
        remove a users permission

        :param file_id: id of drive file
        :param addr: user email/domain name
        :param permisssions_in: permissions of the sheet if not provided its fetched

        :returns:
        """
        if not permisssions_in:
            permissions = self.list_permissions(file_id)['permissions']
        else:
            permissions = permisssions_in

        if re.match('^[_a-z0-9-]+(\.[_a-z0-9-]+)*@[a-z0-9-]+(\.[a-z0-9-]+)*(\.[a-z]{2,4})$', addr):
            permission_id = [x['id'] for x in permissions if x['emailAddress'] == addr]
        elif re.match('[a-zA-Z\d-]{,63}(\.[a-zA-Z\d-]{,63})*', addr):
            permission_id = [x['id'] for x in permissions if x['domain'] == addr]
        else:
            raise InvalidArgumentValue
        if len(permission_id) == 0:
            # @TODO maybe raise an userInvalid exception
            raise InvalidUser
        result = self.driveService.permissions().delete(fileId=file_id, permissionId=permission_id[0]).execute()
        return result


def get_outh_credentials(client_secret_file, application_name='PyGsheets', credential_dir=None):
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    :param client_secret_file: path to outh2 client secret file
    :param application_name: name of application
    :param credential_dir: path to directory where tokens should be stored
                           'global' if you want to store in system-wide location
                           None if you want to store in current script directory
    :return
        Credentials, the obtained credential.
    """
    if credential_dir == 'global':
        home_dir = os.path.expanduser('~')
        credential_dir = os.path.join(home_dir, '.credentials')
        if not os.path.exists(credential_dir):
            os.makedirs(credential_dir)
    elif not credential_dir:
        credential_dir = os.getcwd()
    else:
        pass

    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(client_secret_file, SCOPES)
        flow.user_agent = application_name
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def authorize(outh_file='client_secret.json', service_file=None, credentials=None):
    """Login to Google API using OAuth2 credentials.

    This is a shortcut function which instantiates :class:`Client`
    and performs auhtication.

    :param outh_file: path to outh2 credentials file
    :param service_file: name of the application
    :param credentials: outh2 credentials object,

    :returns: :class:`Client` instance.

    """
    # @TODO handle exceptions
    if not credentials:
        if service_file:
            with open(service_file) as data_file:
                data = jload(data_file)
                print('service_email : '+str(data['client_email']))
            credentials = ServiceAccountCredentials.from_json_keyfile_name(service_file, SCOPES)
        elif outh_file:
            credentials = get_outh_credentials(client_secret_file=outh_file)
        else:
            raise AuthenticationError
    rclient = Client(auth=credentials)
    return rclient
