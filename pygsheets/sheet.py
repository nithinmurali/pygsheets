from pygsheets.spreadsheet import Spreadsheet
from pygsheets.worksheet import Worksheet
from pygsheets.custom_types import ExportType
from pygsheets.exceptions import InvalidArgumentValue, CannotRemoveOwnerError, RequestError

from googleapiclient import discovery
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.errors import HttpError

import logging
import json
import os
import time
import re


class SheetAPIWrapper(object):

    def __init__(self, http, data_path, quota=100, seconds_per_quota=100, retries=1, logger=logging.getLogger(__name__)):
        """

        :param http:
        :param data_path:
        :param logger:
        """
        self.logger = logger
        with open(os.path.join(data_path, "sheets_discovery.json")) as jd:
            self.service = discovery.build_from_document(json.load(jd), http=http)

        self.retries = retries

        self.collect_batch_updates = False
        self.batch_requests = dict()

        self.number_of_calls = 0
        self.start_time = time.time()
        self.quota = quota
        self.seconds_per_quota = seconds_per_quota

    def batch_update(self, spreadsheet_id, requests, **kwargs):
        """
        Applies one or more updates to the spreadsheet.

        Each request is validated before being applied. If any request is not valid then the entire request will
        fail and nothing will be applied.

        Some requests have replies to give you some information about how they are applied. The replies will mirror
        the requests. For example, if you applied 4 updates and the 3rd one had a reply, then the response will have
        2 empty replies, the actual reply, and another empty reply, in that order.

        Due to the collaborative nature of spreadsheets, it is not guaranteed that the spreadsheet will reflect exactly
        your changes after this completes, however it is guaranteed that the updates in the request will be applied
        together atomically. Your changes may be altered with respect to collaborator changes. If there are no
        collaborators, the spreadsheet should reflect your changes.

        +-----------------------------------+-----------------------------------------------------+
        | Request body params               | Description                                         |
        +===================================+=====================================================+
        | includeSpreadsheetInResponse      | | Determines if the update response should include  |
        |                                   | | the spreadsheet resource. (default: False)        |
        +-----------------------------------+-----------------------------------------------------+
        | responseRanges[]                  | | Limits the ranges included in the response        |
        |                                   | | spreadsheet. Only applied if the first param is   |
        |                                   | | True.                                             |
        +-----------------------------------+-----------------------------------------------------+
        | responseIncludeGridData           | | True if grid data should be returned. Meaningful  |
        |                                   | | only if if includeSpreadsheetInResponse is 'true'.|
        |                                   | | This parameter is ignored if a field mask was set |
        |                                   | | in the request.                                   |
        +-----------------------------------+-----------------------------------------------------+

        :param spreadsheet_id:  The spreadsheet to apply the updates to.
        :param requests:        A list of updates to apply to the spreadsheet. Requests will be applied in the order
                                they are specified. If any request is not valid, no requests will be applied.
        :param kwargs:          Request body params & standard parameters (see reference for details).
        :return:
        """
        if isinstance(requests, list):
            body = {'requests': requests}
        else:
            body = {'requests': [requests]}

        for param in ['includeSpreadsheetInResponse', 'responseRanges', 'responseIncludeGridData']:
            if param in kwargs:
                body['requests'][param] = kwargs[param]
                del kwargs[param]

        request = self.service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id,
                                                          body=body, **kwargs)
        return self._execute_requests(spreadsheet_id, request)

    def create(self, title, template=None, **kwargs):
        """Create a spreadsheet.

        Can be created with just a title. All other values will be set to default.

        A template can be either a JSON representation of a Spreadsheet Resource as defined by the
        Google Sheets API or an instance of the Spreadsheet class. Missing fields will be set to default.

        :param title:       Title of the new spreadsheet.
        :param template:    Template used to create the new spreadsheet.
        :param kwargs:      Standard parameters (see reference for details).
        :return:            A Spreadsheet Resource.
        """
        if template is None:
            body = {'properties': {'title': title}}
        else:
            if isinstance(template, dict):
                if 'properties' in template:
                    template['properties']['title'] = title
                else:
                    template['properties'] = {'title': title}
                body = template
            elif isinstance(template, Spreadsheet):
                body = template.to_json()
                body['properties']['title'] = title
            else:
                raise InvalidArgumentValue('Need a dictionary or spreadsheet for a template.')
        return self._execute_requests(self.service.spreadsheets().create(body=body, **kwargs))

    def get(self, spreadsheet_id, **kwargs):
        """Returns a full spreadsheet with the entire data.

        The data returned can be limited with parameters.

        :param spreadsheet_id:  The Id of the spreadsheet to return.
        :param kwargs:          Standard parameters (see reference for details).
        :return:                Return a SheetResource.
        """
        if 'fields' not in kwargs:
            kwargs['fields'] = '*'
        if 'includeGridData' not in kwargs:
            kwargs['includeGridData'] = True
        return self._execute_requests(self.service.spreadsheets().get(spreadsheetId=spreadsheet_id, **kwargs))

    def get_by_data_filter(self):
        pass

    def developer_metadata_get(self):
        pass

    def developer_metadata_search(self):
        pass

    def sheets_copy_to(self, source_spreadsheet_id, worksheet_id, destination_spreadsheet_id, **kwargs):
        """Copies a worksheet from one spreadsheet to another.

        `Reference <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.sheets/copyTo>`_

        :param source_spreadsheet_id:       The ID of the spreadsheet containing the sheet to copy.
        :param worksheet_id:                The ID of the sheet to copy.
        :param destination_spreadsheet_id:  The ID of the spreadsheet to copy the sheet to.
        :param kwargs:                      Standard parameters (see reference for details).
        :return:  SheetProperties
        """
        if 'fields' not in kwargs:
            kwargs['fields'] = '*'

        body = {"destinationSpreadsheetId": destination_spreadsheet_id}
        request = self.service.spreadsheets().sheets().copyTo(spreadsheetId=source_spreadsheet_id,
                                                              sheetId=worksheet_id,
                                                              body=body,
                                                              **kwargs)
        return self._execute_requests(request)

    def values_append(self, spreadsheet_id, values, major_dimension, range, replace):
        """wrapper around batch append"""
        body = {
            'values': values,
            'majorDimension': major_dimension
        }

        if replace:
            inoption = "OVERWRITE"
        else:
            inoption = "INSERT_ROWS"
        request = self.service.spreadsheets().values().append(spreadsheetId=spreadsheet_id, range=range,
                                                                    body=body,
                                                                    insertDataOption=inoption,
                                                                    includeValuesInResponse=False,
                                                                    valueInputOption="USER_ENTERED")
        self._execute_requests(request)

    def values_batch_clear(self):
        pass

    def values_batch_clear_by_data_filter(self):
        pass

    def values_batch_get(self):
        pass

    def values_batch_get_by_data_filter(self):
        pass

    def values_batch_update(self):
        pass

    def values_batch_update_by_data_filter(self):
        pass

    def values_clear(self):
        pass

    def values_get(self):
        pass

    def values_update(self):
        pass

    def _execute_requests(self, request, spreadsheet_id=None):
        """

        :param request:
        :param spreadsheet_id:
        :return:
        """
        if self.collect_batch_updates:
            try:
                self.batch_requests[spreadsheet_id].append(request)
            except KeyError:
                self.batch_requests[spreadsheet_id] = [request]
        else:
            now = time.time()
            # if more than seconds per quota elapsed since the first call the counter is reset.
            if self.start_time < now - self.seconds_per_quota:
                self.start_time = now
                self.number_of_calls = 0

            self.number_of_calls += 1
            # if the number of calls would exceed the quota wait until the quota is reset.
            if self.number_of_calls > self.quota:
                time.sleep((self.start_time + 100) - now)
                self.number_of_calls = 0
                self.start_time = time.time()
            response = request.execute(num_retries=self.retries)

            return response
