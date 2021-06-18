# -*- coding: utf-8 -*-.

"""
pygsheets.spreadsheet
~~~~~~~~~~~~~~~~~~~~~

This module represents an entire spreadsheet. Which can have several worksheets.

"""

import logging
import warnings

from pygsheets.worksheet import Worksheet
from pygsheets.datarange import DataRange
from pygsheets.exceptions import (WorksheetNotFound, RequestError,
                         InvalidArgumentValue, InvalidUser)
from pygsheets.custom_types import *
from pygsheets.developer_metadata import DeveloperMetadataLookupDataFilter, DeveloperMetadata


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
        self.fetch_properties(jsonsheet)
        self.default_parse = True

    @property
    def id(self):
        """Id of the spreadsheet."""
        return self._id

    @property
    def title(self):
        """Title of the spreadsheet."""
        return self._title

    @title.setter
    def title(self, value):
        self._title = value
        self._jsonsheet['properties']['title'] = value
        self.update_properties()
    
    @property
    def locale(self):
        """Locale of the spreadsheet."""
        return self._jsonsheet['properties']['locale']

    @locale.setter
    def locale(self, value):
        self._locale = value
        self._jsonsheet['properties']['locale'] = value
        self.update_properties()

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
        response = self.client.sheet.get(spreadsheet_id=self.id, fields='sheets(properties.sheetId,protectedRanges)')
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
        return self.client.drive.get_update_time(self.id)

    def update_properties(self):
        """
        Update the sheet properties in cloud
        """
        request = {
            "updateSpreadsheetProperties": {"properties": self._jsonsheet['properties'], "fields": "*"}}
        self.client.sheet.batch_update(self.id, request, fields='*')

    def fetch_properties(self, jsonsheet=None, fetch_sheets=True):
        """Update all properties of this spreadsheet with the remote.

        The provided json representation must be the same as the Google Sheets v4 Response. If no sheet is given this
        will simply fetch all data from remote and update the local representation.

        Reference: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets

        :param jsonsheet:       Used to update the spreadsheet.
        :param fetch_sheets:    Fetch sheets from remote.

        """
        if not jsonsheet and len(self.id) > 1:
            self._jsonsheet = self.client.open_as_json(self.id)
        elif not jsonsheet and len(self.id) == 0:
            raise InvalidArgumentValue('jsonsheet')
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
            jsonsheet = self.client.open_as_json(self.id)
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

        jsheet = dict()
        if src_tuple:
            jsheet['properties'] = self.client.sheet.sheets_copy_to(src_tuple[0], src_tuple[1], self.id)
            wks = self.worksheet_cls(self, jsheet)
            wks.title = title
            wks.index = index
        elif src_worksheet:
            if type(src_worksheet) != Worksheet:
                raise InvalidArgumentValue("src_worksheet")
            jsheet['properties'] = self.client.sheet.sheets_copy_to(src_worksheet.spreadsheet.id, src_worksheet.id, self.id)
            wks = self.worksheet_cls(self, jsheet)
            wks.title = title
            wks.index = index
        else:
            request = {"addSheet": {"properties": {'title': title, "gridProperties": {"rowCount": rows, "columnCount": cols}}}}
            if index is not None:
                request["addSheet"]["properties"]["index"] = index
            result = self.client.sheet.batch_update(self.id, request, fields='replies/addSheet')
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
        self.client.sheet.batch_update(self.id, request)
        self._sheet_list.remove(worksheet)

    def replace(self, pattern, replacement=None, **kwargs):
        """Replace values in any cells matched by pattern in all worksheets.

        Keyword arguments not specified will use the default value. If the spreadsheet is -

        Unlinked:
            Uses `self.find(pattern, **kwargs)` to find the cells and then replace the values in each cell.

        Linked:
            The replacement will be done by a `findReplaceRequest <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request#findreplacerequest>`_
            as defined by the Google Sheets API. After the request the local copy is updated.

        :param pattern:             Match cell values.
        :param replacement:         Value used as replacement.
        :arg searchByRegex:         Consider pattern a regex pattern. (default False)
        :arg matchCase:             Match case sensitive. (default False)
        :arg matchEntireCell:       Only match on full match. (default False)
        :arg includeFormulas:       Match fields with formulas too. (default False)
        """
        for wks in self.worksheets():
            wks.replace(pattern, replacement=replacement, **kwargs)

    def find(self, pattern, **kwargs):
        """Searches through all worksheets.

        Search all worksheets with the options given. If an option is not given, the default will be used.
        Will return a list of cells for each worksheet packed into a list. If a worksheet has no cell which
        matches pattern an empty list is added.

        :param pattern:             The value to search.
        :arg searchByRegex:         Consider pattern a regex pattern. (default False)
        :arg matchCase:             Match case sensitive. (default False)
        :arg matchEntireCell:       Only match on full match. (default False)
        :arg includeFormulas:       Match fields with formulas too. (default False)

        :returns: A list of lists of :class:`Cells <Cell>`
        """
        found_cells = []
        for sheet in self.worksheets():
            found_cells.append(sheet.find(pattern, **kwargs))
        return found_cells

    def share(self, email_or_domain, role='reader', type='user', **kwargs):
        """Share this file with a user, group or domain.

        User and groups need an e-mail address and domain needs a domain for a permission.
        Share sheet with a person and send an email message.

        >>> spreadsheet.share('example@gmail.com', role='commenter', type='user', emailMessage='Here is the spreadsheet we talked about!')

        Make sheet public with read only access:

        >>> spreadsheet.share('', role='reader', type='anyone')

        :param email_or_domain: The email address or domain this file should be shared to.
        :param role:            The role of the new permission.
        :param type:            The type of the new permission.
        :param kwargs:          Optional arguments. See DriveAPIWrapper.create_permission documentation for details.
        """
        if type in ['user', 'group']:
            kwargs['emailAddress'] = email_or_domain
        elif type == 'domain':
            kwargs['domain'] = email_or_domain
        self.client.drive.create_permission(self.id, role=role, type=type, **kwargs)

    @property
    def permissions(self):
        """Permissions for this file."""
        return self.client.drive.list_permissions(self.id)

    def remove_permission(self, email_or_domain, permission_id=None):
        """Remove a permission from this sheet.

        All permissions associated with this email or domain are deleted.

        :param email_or_domain:     Email or domain of the permission.
        :param permission_id:       (optional) permission id if a specific permission should be deleted.
        """
        if permission_id is not None:
            self.client.drive.delete_permission(self.id, permission_id=permission_id)
        else:
            for permission in self.permissions:
                if email_or_domain in [permission.get('domain', ''), permission.get('emailAddress', '')]:
                    self.client.drive.delete_permission(self.id, permission_id=permission['id'])

    def export(self, file_format=ExportType.CSV, path='', filename=''):
        """Export all worksheets.

        The filename must have an appropriate file extension. Each sheet will be exported into a separate file.
        The filename is extended (before the extension) with the index number of the worksheet to not overwrite
        each file.

        :param file_format: ExportType.<?>
        :param path:        Path to the directory where the file will be stored. (default: current working directory)
        :param filename:    Filename (default: spreadsheet id)
        """
        self.client.drive.export(self, file_format=file_format, filename=filename, path=path)

    def delete(self):
        """Deletes this spreadsheet.

        Leaves the local copy intact. The deleted spreadsheet is permanently removed from your drive
        and not moved to the trash.
        """
        self.client.drive.delete(self.id)

    def get_developer_metadata(self, key=None, search_sheets=False):
        """
        Fetch developer metadata associated with this spreadsheet

        :param key:            The key of the metadata to fetch. If unspecified, all metadata will be returned
        :param search_sheets:  Set to True to also include worksheets in the metadata search
        """
        spreadsheet_id = None if search_sheets else self.id
        data_filter = DeveloperMetadataLookupDataFilter(spreadsheet_id, meta_key=key)
        results = self.client.sheet.developer_metadata_search(self.id, data_filter.to_json())
        metadata = []
        if results:
            for result in results["matchedDeveloperMetadata"]:
                meta_id = result["developerMetadata"]["metadataId"]
                key = result["developerMetadata"]["metadataKey"]
                value = result["developerMetadata"]["metadataValue"]
                sheet_id = result["developerMetadata"]["location"].get("sheetId", None)
                metadata.append(DeveloperMetadata(meta_id, key, value, self.client, self.id, sheet_id))
        return metadata

    def create_developer_metadata(self, key, value=None):
        """
        Create a new developer metadata associated with this spreadsheet

        Will return None when in batch mode, otherwise will return a DeveloperMetadata object

        :param key:    the key of the metadata to be created
        :param value:  the value of the metadata to be created (optional)
        """
        return DeveloperMetadata.new(key, value, self.client, self.id)

    def custom_request(self, request, fields, **kwargs):
        """
        Send a custom batch update request to this spreadsheet.

        These requests have to be properly constructed. All possible requests are documented in the reference.

        Reference: api docs <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/request>`__

        :param request: One or several requests as dictionaries.
        :param fields:  Fields which should be included in the response.
        :param kwargs:  Any other params according to refrence.

        :return:   json response <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/response> __
        """
        return self.client.sheet.batch_update(self.id, request, fields=fields, **kwargs)

    def to_json(self):
        """Return this spreadsheet as json resource."""
        return self.client.open_as_json(self.id)

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
