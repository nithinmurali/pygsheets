from pygsheets.spreadsheet import Spreadsheet
from pygsheets.utils import format_addr
from pygsheets.exceptions import InvalidArgumentValue

from googleapiclient import discovery
from googleapiclient.errors import HttpError

import logging
import json
import os
import time

GOOGLE_SHEET_CELL_UPDATES_LIMIT = 50000


class SheetAPIWrapper(object):

    def __init__(self, http, data_path, seconds_per_quota=100, retries=1, logger=logging.getLogger(__name__)):
        """A wrapper class for the Google Sheets API v4.

        All calls to the the API are made in this class. This ensures that the quota is never hit.

        The default quota for the API is 100 requests per 100 seconds. Each request is made immediately and counted.
        When 100 seconds have passed the counter is reset. Should the counter reach 101 the request is delayed until 100
        seconds since the first request pass.

        :param http:                The http object used to execute the requests.
        :param data_path:           Where the discovery json file is stored.
        :param quota:               Default value is 100
        :param seconds_per_quota:   Default value is 100 seconds
        :param retries:             How often the requests will be repeated if the connection times out. (Default 1)
        :param logger:
        """
        self.logger = logger
        with open(os.path.join(data_path, "sheets_discovery.json")) as jd:
            self.service = discovery.build_from_document(json.load(jd), http=http)

        self.retries = retries

        self.collect_batch_updates = False
        self.batch_requests = dict()

        self.seconds_per_quota = seconds_per_quota

    # TODO: Implement feature to actually combine update requests.
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

        if 'fields' not in kwargs:
            kwargs['fields'] = '*'

        request = self.service.spreadsheets().batchUpdate(spreadsheetId=spreadsheet_id,
                                                          body=body, **kwargs)
        return self._execute_requests(request)

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

    #################################
    #     BATCH UPDATE REQUESTS     #
    #################################

    def update_sheet_properties_request(self, spreadsheet_id, properties, fields):
        """Updates the properties of the specified sheet.

        Properties must be an instance of `SheetProperties <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#SheetProperties>`_.

        :param spreadsheet_id:  The id of the spreadsheet to be updated.
        :param properties:      The properties to be updated.
        :param fields:          Specifies the fields which should be updated.
        :return: SheetProperties
        """
        request = {
            'updateSheetProperties':
                {
                    'properties': properties,
                    'fields': fields
                }
        }
        return self.batch_update(spreadsheet_id, request)

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
            insert_data_option = "OVERWRITE"
        else:
            insert_data_option = "INSERT_ROWS"
        request = self.service.spreadsheets().values().append(spreadsheetId=spreadsheet_id, range=range,
                                                              body=body,
                                                              insertDataOption=insert_data_option,
                                                              includeValuesInResponse=False,
                                                              valueInputOption="USER_ENTERED")
        self._execute_requests(request)

    def values_batch_clear(self, spreadsheet_id, ranges):
        """Clear values from sheet.

        Clears one or more ranges of values from a spreadsheet. The caller must specify the spreadsheet ID and one or
        more ranges. Only values are cleared -- all other properties of the cell (such as formatting, data validation,
        etc..) are kept.

        `Reference <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/batchClear>`_

        :param spreadsheet_id:  The ID of the spreadsheet to update.
        :param ranges:          A list of ranges to clear in A1 notation.
        """
        body = {'ranges': ranges}
        request = self.service.spreadsheets().values().batchClear(spreadsheetId=spreadsheet_id, body=body)
        self._execute_requests(request)

    def values_batch_clear_by_data_filter(self):
        pass

    def values_batch_get(self):
        pass

    def values_batch_get_by_data_filter(self):
        pass

    def values_batch_update(self, spreadsheet_id, body, parse=True):
        """

        :param spreadsheet_id:
        :param body:
        :param parse:
        :return:
        """
        cformat = 'USER_ENTERED' if parse else 'RAW'
        batch_limit = GOOGLE_SHEET_CELL_UPDATES_LIMIT
        if body['majorDimension'] == 'ROWS':
            batch_length = int(batch_limit / len(body['values'][0]))  # num of rows to include in a batch
            num_rows = len(body['values'])
        else:
            batch_length = int(batch_limit / len(body['values']))  # num of rows to include in a batch
            num_rows = len(body['values'][0])
        if len(body['values']) * len(body['values'][0]) <= batch_limit:
            request = self.service.spreadsheets().values().update(spreadsheetId=spreadsheet_id,
                                                                  range=body['range'],
                                                                  valueInputOption=cformat, body=body)
            self._execute_requests(request)
        else:
            if batch_length == 0:
                raise AssertionError("num_columns < " + str(GOOGLE_SHEET_CELL_UPDATES_LIMIT))
            values = body['values']
            title, value_range = body['range'].split('!')
            value_range_start, value_range_end = value_range.split(':')
            value_range_end = list(format_addr(str(value_range_end), output='tuple'))
            value_range_start = list(format_addr(str(value_range_start), output='tuple'))
            max_rows = value_range_end[0]
            start_row = value_range_start[0]
            for batch_start in range(0, num_rows, batch_length):
                if body['majorDimension'] == 'ROWS':
                    body['values'] = values[batch_start:batch_start + batch_length]
                else:
                    body['values'] = [col[batch_start:batch_start + batch_length] for col in values]
                value_range_start[0] = batch_start + start_row
                value_range_end[0] = min(batch_start + batch_length, max_rows) + start_row
                body['range'] = title + '!' + format_addr(tuple(value_range_start), output='label') + ':' + \
                                format_addr(tuple(value_range_end), output='label')
                request = self.service.spreadsheets().values().update(spreadsheetId=spreadsheet_id, body=body,
                                                                      range=body['range'],
                                                                      valueInputOption=cformat)
                self._execute_requests(request)

    def values_batch_update_by_data_filter(self):
        pass

    def values_clear(self):
        pass

    def values_get(self):
        pass

    def values_update(self):
        pass

    def _execute_requests(self, request):
        """

        :param request:
        :param spreadsheet_id:
        :return:
        """
        try:
            response = request.execute(num_retries=self.retries)
        except HttpError as error:
            if error.resp['status'] == '429':
                time.sleep(self.seconds_per_quota)
                response = request.execute(num_retries=self.retries)
            else:
                raise
        return response
