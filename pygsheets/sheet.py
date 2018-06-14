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
import re


class SheetAPIWrapper(object):

    def __init__(self, http, data_path, retries=1, logger=logging.getLogger(__name__)):
        """

        :param http:
        :param data_path:
        :param logger:
        """
        self.logger = logger
        with open(os.path.join(data_path, "sheets_discovery.json")) as jd:
            self.service = discovery.build_from_document(json.load(jd), http=http)

        self.retries = retries

    def batch_update(self, sheet_id, requests):
        pass

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

        Can be limited with parameters.

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

    def values_append(self):
        pass

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

    def _execute_requests(self, request, collect=False, spreadsheet_id=None):
        """

        :param request:
        :param collect:         Collect all the requests and combine them into a singe request.
        :param spreadsheet_id:
        :return:
        """
        if collect:
            try:
                self.batch_requests[spreadsheet_id].append(request)
            except KeyError:
                self.batch_requests[spreadsheet_id] = [request]
        else:
            for i in range(self.retries):
                try:
                    response = request.execute()
                except Exception as e:
                    if repr(e).find('timed out') == -1:
                        raise
                    if i == self.retries - 1:
                        raise RequestError("Timeout : " + repr(e))
                    # print ("Cant connect, retrying ... " + str(i))
                else:
                    return response
