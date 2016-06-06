# -*- coding: utf-8 -*-

"""
pysheets.client
~~~~~~~~~~~~~~

This module contains Client class responsible for communicating with
Google Data API.

"""
import re
import warnings

from xml.etree import ElementTree

from . import __version__
# from . import urlencode
# from .ns import _ns
from .models import Spreadsheet
from .exceptions import (AuthenticationError, SpreadsheetNotFound,
                         NoValidUrlKeyFound, UpdateCellError,
                         RequestError)


#from __future__ import print_function
import httplib2
import os

from apiclient import discovery
import oauth2client
from oauth2client import client
from oauth2client import tools
import json


SCOPES = 'https://www.googleapis.com/auth/spreadsheets'
APPLICATION_NAME = 'pySheets'

_url_key_re_v1 = re.compile(r'key=([^&#]+)')
_url_key_re_v2 = re.compile(r'spreadsheets/d/([^&#]+)/edit')

class Client(object):

    """An instance of this class communicates with Google Data API.

    :param auth: A tuple containing an *email* and a *password* used for ClientLogin
                 authentication or an OAuth2 credential object. Credential objects are those created by the
                 oauth2client library. https://github.com/google/oauth2client
    :param http_session: (optional) A session object capable of making HTTP requests while persisting headers.
                                    Defaults to :class:`~gspread.httpsession.HTTPSession`.

    >>> c = gspread.Client(auth=('user@example.com', 'qwertypassword'))

    or

    >>> c = gspread.Client(auth=OAuthCredentialObject)


    """
    def __init__(self, auth):
        self.auth = auth
        http = auth.authorize(httplib2.Http())
        discoveryUrl = ('https://sheets.googleapis.com/$discovery/rest?'
                'version=v4')
        self.service = discovery.build('sheets', 'v4', http=http,
                          discoveryServiceUrl=discoveryUrl)
        self.sendBatch = False
        self.spreadsheetId = None

    #@TODO
    def open(self, title):
        """Opens a spreadsheet, returning a :class:`~gspread.Spreadsheet` instance.

        :param title: A title of a spreadsheet.

        If there's more than one spreadsheet with same title the first one
        will be opened.

        :raises gspread.SpreadsheetNotFound: if no spreadsheet with
                                             specified `title` is found.

        >>> c = gspread.Client(auth=('user@example.com', 'qwertypassword'))
        >>> c.login()
        >>> c.open('My fancy spreadsheet')

        """
        pass

    def open_by_key(self, key):
        """Opens a spreadsheet specified by `key`, returning a :class:`~gspread.Spreadsheet` instance.

        :param key: A key of a spreadsheet as it appears in a URL in a browser.

        :raises gspread.SpreadsheetNotFound: if no spreadsheet with
                                             specified `key` is found.

        >>> c = gspread.Client(auth=('user@example.com', 'qwertypassword'))
        >>> c.login()
        >>> c.open_by_key('0BmgG6nO_6dprdS1MN3d3MkdPa142WFRrdnRRUWl1UFE')

        """
        try:
            result = self.service.spreadsheets().get(spreadsheetId=key).execute()
        except Exception as e:
            if (json.loads(result)['error']['status'] == 'NOT_FOUND'):
                raise SpreadsheetNotFound
            else:
                raise e
        self.spreadsheetId = key
        return Spreadsheet(self,result)
            

    def open_by_url(self, url):
        """Opens a spreadsheet specified by `url`,
           returning a :class:`~gspread.Spreadsheet` instance.

        :param url: URL of a spreadsheet as it appears in a browser.

        :raises gspread.SpreadsheetNotFound: if no spreadsheet with
                                             specified `url` is found.

        >>> c = gspread.Client(auth=('user@example.com', 'qwertypassword'))
        >>> c.login()
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

    #@TODO
    def openall(self, title=None):
        """Opens all available spreadsheets,
           returning a list of a :class:`~gspread.Spreadsheet` instances.

        :param title: (optional) If specified can be used to filter
                      spreadsheets by title.

        """
        pass

    #@TODO
    def start_batch(self):
        self.sendBatch = True

    #@TODO
    def stop_batch(self):
        self.sendBatch = False

    #@TODO
    def update_range(self,range,values,majorDim='ROWS',format=False):
        '''
        @TODO group requests based on value input option
        
        '''
        if self.sendBatch:
            pass
        else:
            body = {}
            body['range'] = range
            body['majorDimension'] = str(majorDim)
            body['values'] = values
            if format: format = 'RAW';
            else: format = 'USER_ENTERED';
            result = self.service.spreadsheets().values().update(spreadsheetId=spreadsheetId,range=body['range'],valueInputOption=format,body=body).execute()


    #@TODO
    def get_range(self,range,majorDim='ROWS'):
        '''
        @TODO group requests based on value input option

        '''
        if not self.spreadsheetId:
            return None

        if self.sendBatch:
            pass
        else:
            result = self.service.spreadsheets().values().get(spreadsheetId=self.spreadsheetId, range = range,\
                        majorDimension=majorDim,valueRenderOption=None,dateTimeRenderOption=None).execute()
            return result['values']



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
                                   'sheets.googleapis.com-python-quickstart.json')

    store = oauth2client.file.Storage(credential_path)
    credentials = store.get()
    if not credentials or credentials.invalid:
        flow = client.flow_from_clientsecrets(client_secret_file, SCOPES)
        flow.user_agent = APPLICATION_NAME
        if flags:
            credentials = tools.run_flow(flow, store, flags)
        else: # Needed only for compatibility with Python 2.6
            credentials = tools.run(flow, store)
        print('Storing credentials to ' + credential_path)
    return credentials


def authorize(file = 'client_secret.json',credentials = None):
    """Login to Google API using OAuth2 credentials.

    This is a shortcut function which instantiates :class:`Client`
    and performs login right away.

    :returns: :class:`Client` instance.

    """
    if not credentials:
        credentials =  get_credentials(file)
    #print 'cred: ',credential
    client = Client(auth=credentials)
    return client
