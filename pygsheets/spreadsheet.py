# -*- coding: utf-8 -*-.

"""
pygsheets.spreadsheet
~~~~~~~~~~~~~~~~~~~~~

This module represents an entire spreadsheet. Which can have several worksheets.

"""

import logging

from .worksheet import Worksheet
from .datarange import DataRange
from .exceptions import (WorksheetNotFound, RequestError,
                         InvalidArgumentValue, InvalidUser)
from .custom_types import *


class Spreadsheet(object):
    """ A class for a spreadsheet object."""

    worksheet_cls = Worksheet

    def __init__(self, client, jsonsheet=None, id=None):
        """The spreadsheet is used to store and manipulate metadata and load specific sheets.

        :param client:      The client which is responsible to connect the sheet with the remote.
        :param jsonsheet:   The json-dict representation of the spreadsheet as returned by Google Sheets API v4.
        :param id:          Id of this spreadsheet
        """
        if type(jsonsheet) != dict and jsonsheet is not None:
            raise InvalidArgumentValue("jsonsheet")
        self.logger = logging.getLogger(__name__)
        self.client = client
        self._sheet_list = []
        self._jsonsheet = jsonsheet
        self._id = id
        self._title = ''
        self._named_ranges = []
        self.update_properties(jsonsheet)
        self._permissions = dict()
        self.batch_mode = False
        self.default_parse = True

    @property
    def id(self):
        """Id of the spreadsheet."""
        return self._id

    @property
    def title(self):
        """Title of the spreadsheet."""
        return self._title

    @property
    def sheet1(self):
        """Direct access to the first worksheet."""
        return self.worksheet()

    @property
    def url(self):
        """Url of the spreadsheet."""
        return "https://docs.google.com/spreadsheets/d/"+self.id

    @property
    def named_ranges(self):
        """All named ranges in this spreadsheet."""
        return [DataRange(namedjson=x, name=x['name'], worksheet=self.worksheet('id', x['range'].get('sheetId', 0)))
                for x in self._named_ranges]

    @property
    def protected_ranges(self):
        """All protected ranges in this spreadsheet."""
        request = self.client.service.spreadsheets().get(spreadsheetId=self.id, fields="sheets/(properties/sheetId,protectedRanges)", includeGridData=True)
        response = self.client._execute_request(self.id, request, False)
        return [DataRange(protectedjson=x, worksheet=self.worksheet('id', sheet['properties']['sheetId']))
                for sheet in response['sheets']
                for x in sheet.get('protectedRanges', [])]

    @property
    def defaultformat(self):
        """Default cell format used."""
        return self._defaultFormat

    @property
    def updated(self):
        """Last time the spreadsheet was modified using RFC 3339 format."""
        request = self.client.driveService.files().get(fileId=self.id, fields='modifiedTime',
                                                       supportsTeamDrives=self.client.enableTeamDriveSupport)
        response = self.client._execute_request(self.id, request, False)
        return response['modifiedTime']

    def update_properties(self, jsonsheet=None, fetch_sheets=True):
        """Update all properties of this spreadsheet with the remote.

        The provided json representation must be the same as the Google Sheets v4 Response. If no sheet is given this
        will simply fetch all data from remote and update the local representation.

        Reference: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets

        :param jsonsheet:       Used to update the spreadsheet.
        :param fetch_sheets:    Fetch sheets from remote.

        """
        if not jsonsheet and len(self.id) > 1:
            self._jsonsheet = self.client.open_by_key(self.id, 'json')
        elif not jsonsheet and len(self.id) == 0:
            raise InvalidArgumentValue('jsonsheet')
        # print self._jsonsheet
        self._id = self._jsonsheet['spreadsheetId']
        if fetch_sheets:
            self._fetch_sheets(self._jsonsheet)
        self._title = self._jsonsheet['properties']['title']
        self._defaultFormat = self._jsonsheet['properties']['defaultFormat']
        self.client.spreadsheetId = self._id
        self._named_ranges = self._jsonsheet.get('namedRanges', [])

    def _fetch_sheets(self, jsonsheet=None):
        """Update the sheets stored in this spreadsheet."""
        self._sheet_list = []
        if not jsonsheet:
            jsonsheet = self.client.open_by_key(self.id, returnas='json')
        for sheet in jsonsheet.get('sheets'):
            self._sheet_list.append(self.worksheet_cls(self, sheet))

    def worksheets(self, sheet_property=None, value=None, force_fetch=False):
        """Get worksheets matching the specified property.

        :param sheet_property:  Property used to filter ('title', 'index', 'id').
        :param value:           Value of the property.
        :param force_fetch:     Fetch data from remote.

        :returns: List of :class:`Worksheets <Worksheet>`.
        """
        if not sheet_property and not value:
            return self._sheet_list

        if sheet_property not in ['title', 'index', 'id']:
            raise InvalidArgumentValue('sheet_property')
        elif sheet_property in ['index', 'id']:
            value = int(value)

        sheets = [x for x in self._sheet_list if getattr(x, sheet_property) == value]
        if not len(sheets) > 0 or force_fetch:
            self._fetch_sheets()
            sheets = [x for x in self._sheet_list if getattr(x, sheet_property) == value]
            if not len(sheets) > 0:
                raise WorksheetNotFound()
        return sheets

    def worksheet(self, property='index', value=0):
        """Returns the worksheet with the specified index, title or id.

        If several worksheets with the same property are found the first is returned. This may not be the same
        worksheet every time.

        Example: Getting a worksheet named 'Annual bonuses'

        >>> sht = client.open('Sample one')
        >>> worksheet = sht.worksheet('title','Annual bonuses')

        :param property:    The searched property.
        :param value:       Value of the property.

        :returns: :class:`Worksheets <Worksheet>`.
        """
        return self.worksheets(property, value)[0]

    def worksheet_by_title(self, title):
        """Returns worksheet by title.

        :param title:   Title of the sheet

        :returns: :class:`Worksheets <Worksheet>`.
        """
        return self.worksheet('title', title)

    def add_worksheet(self, title, rows=100, cols=26, src_tuple=None, src_worksheet=None, index=None):
        """Creates or copies a worksheet and adds it to this spreadsheet.

        When creating only a title is needed. Rows & columns can be adjusted to match your needs.
        Index can be specified to set position of the sheet.

        When copying another worksheet supply the spreadsheet id & worksheet id and the worksheet wrapped in a Worksheet
        class.

        :param title:           Title of the worksheet.
        :param rows:            Number of rows which should be initialized (default 100)
        :param cols:            Number of columns which should be initialized (default 26)
        :param src_tuple:       Tuple of (spreadsheet id, worksheet id) specifying the worksheet to copy.
        :param src_worksheet:   The source worksheet.
        :param index:           Tab index of the worksheet.

        :returns: :class:`Worksheets <Worksheet>`.
        """
        if self.batch_mode:
            raise Exception("not supported in batch Mode")

        jsheet = dict()
        if src_tuple:
            jsheet['properties'] = self.client.sh_copy_worksheet(src_tuple[0], src_tuple[1], self.id)
            wks = self.worksheet_cls(self, jsheet)
            wks.title = title
        elif src_worksheet:
            if type(src_worksheet) != Worksheet:
                raise InvalidArgumentValue("src_worksheet")
            jsheet['properties'] = self.client.sh_copy_worksheet(src_worksheet.spreadsheet.id, src_worksheet.id, self.id)
            wks = self.worksheet_cls(self, jsheet)
            wks.title = title
        else:
            request = {"addSheet": {"properties": {'title': title, "gridProperties": {"rowCount": rows, "columnCount": cols}}}}
            if index is not None:
                request["addSheet"]["properties"]["index"] = index
            result = self.client.sh_batch_update(self.id, request, 'replies/addSheet', False)
            jsheet['properties'] = result['replies'][0]['addSheet']['properties']
            wks = self.worksheet_cls(self, jsheet)
        self._sheet_list.append(wks)
        return wks

    def del_worksheet(self, worksheet):
        """Deletes the worksheet from this spreadsheet.

        :param worksheet: The :class:`worksheets <Worksheet>` to be deleted.
        """
        if worksheet not in self.worksheets():
            raise WorksheetNotFound
        request = {"deleteSheet": {'sheetId': worksheet.id}}
        self.client.sh_batch_update(self.id, request, '', False)
        self._sheet_list.remove(worksheet)

    def replace(self, pattern, replacement=None, **kwargs):
        """Replace values in any cells matched by pattern in all worksheets.

        Keyword arguments not specified will use the default value.

        Unlinked:
            Uses self.find(pattern, **kwargs) to find the cells and then replace the values in each cell.

        Linked:
            The replacement will be done by a findReplaceRequest as defined by the Google Sheets API. After the request
            the local copy is updated.

        Request: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#findreplacerequest

        :param pattern:             Match cell values.
        :param replacement:         Value used as replacement.
        :key searchByRegex:         Consider pattern a regex pattern. (default False)
        :key matchCase:             Match case sensitive. (default False)
        :key matchEntireCell:       Only match on full match. (default False)
        :key includeFormulas:       Match fields with formulas too. (default False)
        """
        for wks in self.worksheets():
            wks.replace(pattern, replacement=replacement, **kwargs)

    def find(self, pattern, **kwargs):
        """Searches through all worksheets.

        Search all worksheets with the options given. If an option is not given, the default will be used.
        Will return a list of cells for each worksheet packed into a list. If a worksheet has no cell which
        matches pattern an empty list is added.

        :param pattern:             The value to search.
        :key searchByRegex:         Consider pattern a regex pattern. (default False)
        :key matchCase:             Match case sensitive. (default False)
        :key matchEntireCell:       Only match on full match. (default False)
        :key includeFormulas:       Match fields with formulas too. (default False)

        :returns A list of lists of :class:`Cells <Cell>`
        """
        found_cells = []
        for sheet in self.worksheets():
            found_cells.append(sheet.find(pattern, **kwargs))
        return found_cells

    # @TODO impliment expiration time
    def share(self, addr, role='reader', expirationTime=None, is_group=False):
        """Create/update permission for user/group/domain/anyone

        :param addr: email for user/group, domain address for domains or 'anyone'
        :param role: permission to be applied ('owner','writer','commenter','reader')
        :param expirationTime: (Not Implimented) time until this permission should last (datetime)
        :param is_group: boolean , Is this a use/group used only when email provided
        """
        return self.client.add_permission(self.id, addr, role=role, is_group=False)

    def list_permissions(self):
        """List all permissions for this spreadsheet.

        :returns: Permissions (list)
        """
        permissions = self.client.list_permissions(self.id)
        self._permissions = permissions['permissions']
        return self._permissions

    def remove_permissions(self, addr):
        """Remove user from permissions list.

        :param addr:    User email.
        """
        try:
            self.client.remove_permissions(self.id, addr, self._permissions)
        except InvalidUser:
            self.client.remove_permissions(self.id, addr)

    def batch_start(self):
        """Start batch mode.

        This will begin batch mode. All requests made to the sheet will instead be collected and
        executed once done. This should speed up processing of local file and reduce the number of
        API calls.
        """
        self.batch_mode = True
        self.logger.warn('Batching is only for Update operations')

    def batch_stop(self, discard=False):
        """Stop batch mode.

        This will end batch mode and all changes made during batch mode will be either synched with
        the remote spreadsheet or discarded.

        :param discard: Discard all changes made during batch mode.
        """
        self.batch_mode = False
        if not discard:
            self.client.send_batch(self.id)

    # @TODO
    def link(self, syncToCloud=False):
        """Link spreadsheet with remote.

        Linked spreadsheets will upload each change to the remote. This ensures that the local copy will always be up
        to date. This will link all sheets and cause an update. Either local or remote data will be overwritten.

        :param  syncToCloud:    True  -> Overwrite remote with local changes.
                                False -> Overwrite local with remote changes.
        """
        # just link all child sheets
        warnings.warn("method not implimented")

    # @TODO
    def unlink(self):
        """Unlink spreadsheet from remote.

        Unlinked spreadsheets will no longer update the remote. All changes will only apply to the local copy.
        Use link() to re-link this spreadsheet with remote.
        """
        # just unlink all sheets
        warnings.warn("method not implimented")

    def export(self, fformat=ExportType.CSV, filename=None):
        """Export all the worksheets to the file.

        The filename must have an appropriate file extension. Each sheet will be exported into a separate file.
        The filename is extended (before the extension) with the index number of the worksheet to not overwrite
        each file.

        :param fformat:     ExportType.<?>
        :param filename:    File name with path. Otherwise file will be stored in working directory.
        """
        filename = filename.split('.')
        if fformat is ExportType.CSV:
            for wks in self._sheet_list:
                wks.export(ExportType.CSV, filename=filename[0]+str(wks.index)+"."+filename[1])
        elif isinstance(fformat, ExportType):
            for wks in self._sheet_list:
                wks.export(fformat=fformat, filename=filename[0]+str(wks.index)+"."+filename[1])
        else:
            raise InvalidArgumentValue("fformat should be of ExportType Enum")

    def custom_request(self, request, fields):
        """
        Send a custom batch update request to this spreadsheet.

        These requests have to be properly constructed. All possible requests are documented in the reference.

        Reference: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request

        :param request: One or several requests as dictionaries.
        :param fields:  Fields which should be included in the response.
        :return:   json response -> https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/response
        """
        return self.client.sh_batch_update(self.id, request, fields=fields, batch=False)

    def __repr__(self):
        return '<%s %s Sheets:%s>' % (self.__class__.__name__,
                                      repr(self.title), len(self._sheet_list))

    def __eq__(self, other):
        return self.id == other.id

    def __iter__(self):
        for sheet in self.worksheets():
            yield(sheet)

    def __getitem__(self, item):
        if type(item) == int:
            return self.worksheet('index', item)
