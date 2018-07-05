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
    """An instance of this class communicates with Google API.

    :param oauth: An OAuth2 credential object. Credential objects are those created by the
                 oauth2client library. https://github.com/google/oauth2client
    :param http_client: (optional) A object capable of making HTTP requests
    :param retries: (optional) The number of times connection will be
                tried before raising a timeout error.

    >>> c = pygsheets.Client(oauth=OAuthCredentialObject)

    """

    spreadsheet_cls = Spreadsheet

    def __init__(self, credentials, retries=1, no_cache=False):
        if no_cache:
            cache = None
        else:
            cache = os.path.join(tempfile.gettempdir(), str(uuid.uuid4()))
        if os.name == "nt":
            cache = "\\\\?\\" + cache
        self.logger = logging.getLogger(__name__)

        http = AuthorizedSession(credentials)
        data_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
        with open(os.path.join(data_path, "sheets_discovery.json")) as jd:
            self.service = discovery.build_from_document(jload(jd), http=http)
        self.drive = DriveAPIWrapper(http, data_path)
        self._spreadsheeets = []
        self.batch_requests = dict()
        self.retries = retries

    @property
    def teamDriveId(self):
        """ Enable team drive support
            Depricated  please use drive.enable_team_drive
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

    def create(self, title, folder=None, template=None):
        """Creates a spreadsheet, returning a :class:`~pygsheets.Spreadsheet` instance.

        :param folder: id of the parent folder, where the spreadsheet is to be created
        :param title: A title of a spreadsheet.
        :param template: Id of the template spreadsheet to create this sheet from.

        """
        if template:
            result = self.drive.copy_file(template, title, folder)
            spreadsheet_id = result['id']
            self._spreadsheeets.append({'name': title, "id": spreadsheet_id})
            return self.open_by_key(spreadsheet_id)
        else:
            body = {'properties': {'title': title}}
            request = self.service.spreadsheets().create(body=body)
            result = self._execute_request(None, request, False)
            spreadsheet_id = result['spreadsheetId']
            if folder:
                self.drive.move_file(spreadsheet_id, new_folder=folder,
                                     old_folder=self.drive.spreadsheet_metadata(query="name = '" +
                                                                                      title + "'")[0]['parents'][0])
            self._spreadsheeets.append({'name': title, "id": spreadsheet_id})
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
        response = self.sh_get_ssheet(key, 'properties,sheets/properties,spreadsheetId,namedRanges', include_data=False)
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

    def open_all(self, query=None):
        """Opens all available spreadsheets.

        Result can be filtered when specifying the query parameter.

        `Reference <https://developers.google.com/drive/v3/web/search-parameters>`

        :param query:   Can be used to filter the returned metadata.
        :returns:       A list of :class:`~pygsheets.Spreadsheet`.
        """
        return [self.open_by_key(key) for key in self.spreadsheet_ids(query=query)]

    def open_as_json(self, key):
        """Returns the json response from a spreadsheet."""
        return self.sh_get_ssheet(key, 'properties,sheets/properties,spreadsheetId,namedRanges', include_data=False)

    def get_range(self, spreadsheet_id, vrange, majordim='ROWS', value_render=ValueRenderOption.FORMATTED):
        """
         fetches  values from sheet.

        :param spreadsheet_id:  spreadsheet id
        :param vrange: range in A! format
        :param majordim: if the major dimension is rows or cols 'ROWS' or 'COLUMNS'
        :param value_render: format of output values

        :returns: 2d array
        """

        if isinstance(value_render, ValueRenderOption):
            value_render = value_render.value

        if not type(value_render) == str:
            raise InvalidArgumentValue("value_render")

        request = self.service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=vrange,
                                                           majorDimension=majordim, valueRenderOption=value_render,
                                                           dateTimeRenderOption=None)
        result = self._execute_request(spreadsheet_id, request, False)
        try:
            return result['values']
        except KeyError:
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