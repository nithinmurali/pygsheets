# -*- coding: utf-8 -*-

"""
pygsheets.client
~~~~~~~~~~~~~~

This module contains Client class responsible for communicating with
Google Data API.

"""
import re
import warnings

from . import __version__
from .models import Spreadsheet
from .exceptions import (AuthenticationError, SpreadsheetNotFound,
                         NoValidUrlKeyFound, UpdateCellError,
                         RequestError,InvalidArgumentValue)

import httplib2
import os

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools
import json
try:
    import argparse
    flags = tools.argparser.parse_args([])
    # print(flags)
except ImportError:
    flags = None


SCOPES = ['https://www.googleapis.com/auth/spreadsheets','https://www.googleapis.com/auth/drive']
APPLICATION_NAME = 'PyGsheets'

_url_key_re_v1 = re.compile(r'key=([^&#]+)')
_url_key_re_v2 = re.compile(r'spreadsheets/d/([^&#]+)/edit')


class Client(object):

    """An instance of this class communicates with Google Data API.

    :param auth: An OAuth2 credential object. Credential objects are those created by the
                 oauth2client library. https://github.com/google/oauth2client

    >>> c = pygsheets.Client(auth=OAuthCredentialObject)


    """
    def __init__(self, auth):
        self.auth = auth
        http = auth.authorize(httplib2.Http())
        discoveryurl = ('https://sheets.googleapis.com/$discovery/rest?'
                        'version=v4')
        self.service = discovery.build('sheets', 'v4', http=http,
                                       discoveryServiceUrl=discoveryurl)
        self.driveService = discovery.build('drive', 'v3', http=http)
        self._spreadsheeets = []
        self._fetch_sheets()
        self.sendBatch = False
        self.spreadsheetId = None
    
    def _fetch_sheets(self):
        """
        fetch all the sheets info from user's gdrive

        :return: None
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

    def delete(self, id=None, title=None):
        """Deletes a spreadsheet by title or id.

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
            return [Spreadsheet(self, id=x['id']) for x in self._spreadsheeets if x["name"] == title][0]
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
            result = self.service.spreadsheets().get(spreadsheetId=key).execute()
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
           returning a :class:`~pygsheets.Spreadsheet` instance.

        :param url: URL of a spreadsheet as it appears in a browser.

        :raises pygsheets.SpreadsheetNotFound: if no spreadsheet with
                                             specified `url` is found.

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

    def openall(self, title=None):
        """Opens all available spreadsheets,
           returning a list of a :class:`~pygsheets.Spreadsheet` instances.

        :param title: (optional) If specified can be used to filter
                      spreadsheets by title.

        """
        return [Spreadsheet(self, id=x['id']) for x in self._spreadsheeets if ((title is None) or (x['name'] == title))]
    
    def start_batch(self):
        self.sendBatch = True

    #@TODO
    def stop_batch(self):
        self.sendBatch = False

    def update_range(self, range, values, majorDim='ROWS', format=False):
        """
        @TODO group requests based on value input option
        
        """
        if self.sendBatch:
            pass
        else:
            body = dict()
            body['range'] = range
            body['majorDimension'] = str(majorDim)
            body['values'] = values
            cformat = 'RAW' if format else 'USER_ENTERED'
            result = self.service.spreadsheets().values().update(spreadsheetId=self.spreadsheetId, range=body['range'],
                                                                 valueInputOption=cformat, body=body).execute()

    def get_range(self, range, majorDim='ROWS'):
        """
        @TODO group requests based on value input option

        """
        if not self.spreadsheetId:
            return None

        if self.sendBatch:
            pass
        else:
            result = self.service.spreadsheets().values().get(spreadsheetId=self.spreadsheetId, range=range,\
                        majorDimension=majorDim, valueRenderOption=None, dateTimeRenderOption=None).execute()
            try:
                return result['values']
            except KeyError:
                return [['']]

    def insertdim(self, sheetId, majorDim, startindex, endIndex, inheritbefore=False):
        if self.sendBatch:
            pass
        else:
            body = {'requests':[{'insertDimension':{'inheritFromBefore':False, \
                    'range':{'sheetId':sheetId,'dimension':majorDim,'endIndex':endIndex,'startIndex':startindex} \
                    } }] }
            print body
            result = self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId,body=body).execute()

    def update_sheet_properties(self, propertyObj, fieldsToUpdate='title,hidden,gridProperties,tabColor,rightToLeft'):
        requests = {"updateSheetProperties": propertyObj, "fields": fieldsToUpdate}
        body = {'requests': [requests]}
        result = self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=body).execute()

    def add_worksheet(self, title, rows, cols):
        requests = {"addSheet": {"properties": {'title': title, "gridProperties": {"rowCount": rows, "columnCount": cols}}}}
        body = {'requests': [requests]}
        result = self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=body).execute()
        return result

    def del_worksheet(self, sheetId):
        requests = {"deleteSheet":{ 'sheetId':sheetId } }
        body = {'requests': [requests]}
        result = self.service.spreadsheets().batchUpdate(spreadsheetId=self.spreadsheetId, body=body).execute()
        return result


def get_credentials(client_secret_file):
    """Gets valid user credentials from storage.

    If nothing has been stored, or if the stored credentials are invalid,
    the OAuth2 flow is completed to obtain the new credentials.

    Returns:
        Credentials, the obtained credential.
    """
    home_dir = os.path.expanduser('~')
    credential_dir = os.path.join(home_dir, '.credentials')
    if not os.path.exists(credential_dir):
        os.makedirs(credential_dir)
    credential_path = os.path.join(credential_dir,
                                   'sheets.googleapis.com-python.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(client_secret_file, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else:  # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def authorize(sfile='client_secret.json', credentials=None):
    """Login to Google API using OAuth2 credentials.

    This is a shortcut function which instantiates :class:`Client`
    and performs login right away.

    :returns: :class:`Client` instance.

    """
    if not credentials:
        credentials = get_credentials(sfile)
    rclient = Client(auth=credentials)
    return rclient
