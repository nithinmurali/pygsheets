from pygsheets.spreadsheet import Spreadsheet
from pygsheets.utils import format_addr
from pygsheets.exceptions import InvalidArgumentValue
from pygsheets.custom_types import ValueRenderOption, DateTimeRenderOption

from googleapiclient import discovery
from googleapiclient.errors import HttpError

import logging
import json
import os
import time

GOOGLE_SHEET_CELL_UPDATES_LIMIT = 50000


class SheetAPIWrapper(object):

    def __init__(self, http, data_path, seconds_per_quota=100, retries=1, logger=logging.getLogger(__name__),
                 check=True):
        """A wrapper class for the Google Sheets API v4.

        All calls to the the API are made in this class. This ensures that the quota is never hit.

        The default quota for the API is 100 requests per 100 seconds. Each request is made immediately and counted.
        When 100 seconds have passed the counter is reset. Should the counter reach 101 the request is delayed until seconds_per_quota
        seconds since the first request pass.

        :param http:                The http object used to execute the requests.
        :param data_path:           Where the discovery json file is stored.
        :param seconds_per_quota:   Default value is 100 seconds
        :param retries:             How often the requests will be repeated if the connection times out. (Default 1)
        :param check:               Check for quota error and apply rate limiting.
        :param logger:
        """

        self.logger = logger
        try:
            with open(os.path.join(data_path, "sheets_discovery.json")) as jd:
                self.service = discovery.build_from_document(json.load(jd), http=http)
        except:
            self.service = discovery.build('sheets', 'v4', http=http)
        self.retries = retries
        self.seconds_per_quota = seconds_per_quota
        self.check = check

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
        body.pop('spreadsheetId', None)
        return self._execute_requests(self.service.spreadsheets().create(body=body, **kwargs))

    def get(self, spreadsheet_id, **kwargs):
        """Returns a full spreadsheet with the entire data.

        The data returned can be limited with parameters. See `reference <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/get>`__  for details .

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

        Properties must be an instance of `SheetProperties <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#SheetProperties>`__.

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

    # def get_by_data_filter(self):
    #    pass

    # def developer_metadata_get(self):
    #    pass

    # def developer_metadata_search(self):
    #    pass

    def sheets_copy_to(self, source_spreadsheet_id, worksheet_id, destination_spreadsheet_id, **kwargs):
        """Copies a worksheet from one spreadsheet to another.

        Reference: `request <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.sheets/copyTo>`_

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

    def values_append(self, spreadsheet_id, values, major_dimension, range, **kwargs):
        """Appends values to a spreadsheet.

        The input range is used to search for existing data and find a "table" within that range. Values will be
        appended to the next row of the table, starting with the first column of the table. See the guide and
        sample code for specific details of how tables are detected and data is appended.

        The caller must specify the spreadsheet ID, range, and a valueInputOption. The valueInputOption only
        controls how the input data will be added to the sheet (column-wise or row-wise),
        it does not influence what cell the data starts being written to.

        Reference: `request <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/append>`__

        :param spreadsheet_id:      The ID of the spreadsheet to update.
        :param values:              The values to be appended in the body.
        :param major_dimension:     The major dimension of the values provided (e.g. row or column first?)
        :param range:               The A1 notation of a range to search for a logical table of data.
                                    Values will be appended after the last row of the table.
        :param kwargs:              Query & standard parameters (see reference for details).
        """
        body = {
            'values': values,
            'majorDimension': major_dimension
        }
        request = self.service.spreadsheets().values().append(spreadsheetId=spreadsheet_id,
                                                              range=range,
                                                              body=body,
                                                              valueInputOption=kwargs.get('valueInputOption', 'USER_ENTERED'),
                                                              **kwargs)
        return self._execute_requests(request)

    def values_batch_clear(self, spreadsheet_id, ranges):
        """Clear values from sheet.

        Clears one or more ranges of values from a spreadsheet. The caller must specify the spreadsheet ID and one or
        more ranges. Only values are cleared -- all other properties of the cell (such as formatting, data validation,
        etc..) are kept.

        Reference: `request <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/batchClear>`__

        :param spreadsheet_id:  The ID of the spreadsheet to update.
        :param ranges:          A list of ranges to clear in A1 notation.
        """
        body = {'ranges': ranges}
        request = self.service.spreadsheets().values().batchClear(spreadsheetId=spreadsheet_id, body=body)
        self._execute_requests(request)

    # def values_batch_clear_by_data_filter(self):
    #    pass

    # def values_batch_get(self):
    #    pass

    # def values_batch_get_by_data_filter(self):
    #    pass

    # TODO: actually implement batch update. Only uses one or several update requests.
    def values_batch_update(self, spreadsheet_id, body, parse=True):
        """
        Impliments batch update

        :param spreadsheet_id: id of spreadsheet
        :param body: body of request
        :param parse:
        """
        cformat = 'USER_ENTERED' if parse else 'RAW'
        batch_limit = GOOGLE_SHEET_CELL_UPDATES_LIMIT
        lengths = [len(x) for x in body['values']]
        avg_row_length = (min(lengths) + max(lengths))/2
        avg_row_length = 1 if avg_row_length == 0 else avg_row_length
        if body['majorDimension'] == 'ROWS':
            batch_length = int(batch_limit / avg_row_length)  # num of rows to include in a batch
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

    # def values_batch_update_by_data_filter(self):
    #    pass

    # def values_clear(self):
    #    pass

    def values_get(self, spreadsheet_id, value_range, major_dimension='ROWS',
                   value_render_option=ValueRenderOption.FORMATTED_VALUE,
                   date_time_render_option=DateTimeRenderOption.SERIAL_NUMBER):
        """Returns a range of values from a spreadsheet. The caller must specify the spreadsheet ID and a range.

        Reference: `request <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values/get>`__
        
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
        :return:                            `ValueRange <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets.values#ValueRange>`_
        """
        if isinstance(value_render_option, ValueRenderOption):
            value_render_option = value_render_option.value

        if isinstance(date_time_render_option, DateTimeRenderOption):
            date_time_render_option = date_time_render_option.value

        request = self.service.spreadsheets().values().get(spreadsheetId=spreadsheet_id,
                                                           range=value_range,
                                                           majorDimension=major_dimension,
                                                           valueRenderOption=value_render_option,
                                                           dateTimeRenderOption=date_time_render_option)
        return self._execute_requests(request)

    # TODO: implement as base for batch update.
    # def values_update(self):
    #    pass

    def _execute_requests(self, request):
        """Execute a request to the Google Sheets API v4.

        When the API returns a 429 Error will sleep for the specified time and try again.

        :param request:     The request to be made.
        :return:            Response
        """
        try:
            response = request.execute(num_retries=self.retries)
        except HttpError as error:
            if error.resp['status'] == '429' and self.check:
                time.sleep(self.seconds_per_quota)  # TODO use asyncio
                response = request.execute(num_retries=self.retries)
            else:
                raise
        return response
