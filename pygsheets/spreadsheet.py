# -*- coding: utf-8 -*-.

"""
pygsheets.spreadsheet
~~~~~~~~~~~~~~~~~~~~~

This module contains spreadsheets model

"""

import warnings

from .worksheet import Worksheet
from .datarange import DataRange
from .exceptions import (WorksheetNotFound, RequestError,
                         InvalidArgumentValue, InvalidUser)
from .custom_types import *


class Spreadsheet(object):

    """ A class for a spreadsheet object."""

    worksheet_cls = Worksheet

    def __init__(self, client, jsonsheet=None, id=None):
        """ spreadsheet init.

        :param client: the client object which links to this ssheet
        :param jsonsheet: the json sheet which has properties of this ssheet
        :param id: id of the spreadsheet
        """
        if type(jsonsheet) != dict and jsonsheet is not None:
            raise InvalidArgumentValue("jsonsheet")
        self.client = client
        self._sheet_list = []
        self._jsonsheet = jsonsheet
        self._id = id
        self._title = ''
        self._named_ranges = []
        self.update_properties(jsonsheet)
        self._permissions = dict()
        self.batch_mode = False

    @property
    def id(self):
        """ id of the spreadsheet """
        return self._id

    @property
    def title(self):
        """ title of the spreadsheet """
        return self._title

    @property
    def sheet1(self):
        """Shortcut property for getting the first worksheet."""
        return self.worksheet()

    @property
    def named_ranges(self):
        """All named ranges in thi spreadsheet"""
        return [DataRange(namedjson=x, name=x['name'], worksheet=self.worksheet('id', x['range'].get('sheetId', 0)))
                for x in self._named_ranges]

    @property
    def defaultformat(self):
        """ deafault cell format"""
        return self._defaultFormat

    @property
    def updated(self):
        """Last time the spreadsheet was modified, in RFC 3339 format"""
        request = self.client.driveService.files().get(fileId=self.id, fields='modifiedTime',
                                                       supportsTeamDrives=self.client.enableTeamDriveSupport)
        response = self.client._execute_request(self.id, request, False)
        return response['modifiedTime']

    def update_properties(self, jsonsheet=None, fetch_sheets=True):
        """ Update all sheet properies.

        :param jsonsheet: json object to update values form \
                if not specified, will fetch it and update (see google api, for json format)
        :param fetch_sheets: if the sheets should be fetched

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
        """update sheets list"""
        self._sheet_list = []
        if not jsonsheet:
            jsonsheet = self.client.open_by_key(self.id, returnas='json')
        for sheet in jsonsheet.get('sheets'):
            self._sheet_list.append(self.worksheet_cls(self, sheet))

    def worksheets(self, sheet_property=None, value=None, force_fetch=False):
        """
        Get all worksheets filtered by a property.

        :param sheet_property: proptery to filter - 'title', 'index', 'id'
        :param value: value of property to match
        :param force_fetch: update the sheets, from cloud

        :returns: list of all :class:`worksheets <Worksheet>`
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
        """Returns a worksheet with specified property.

        :param property: A property of a worksheet. If there're multiple worksheets \
                        with the same title, first one will be returned.
        :param value: value of given property

        :type property: 'title','index','id'

        :returns: instance of :class:`Worksheet`

        Example. Getting worksheet named 'Annual bonuses'

        >>> sht = client.open('Sample one')
        >>> worksheet = sht.worksheet('title','Annual bonuses')

        """
        return self.worksheets(property, value)[0]

    def worksheet_by_title(self, title):
        """
        returns worksheet by title

        :param title: title of the sheet

        :returns: Spresheet instance
        """
        return self.worksheet('title', title)

    def add_worksheet(self, title, rows=100, cols=26, src_tuple=None, src_worksheet=None):
        """Adds a new worksheet to a spreadsheet.

        :param title: A title of a new worksheet.
        :param rows: Number of rows.
        :param cols: Number of columns.
        :param src_tuple: a tuple (spreadsheet id, worksheet id) specifying a worksheet to copy
        :param src_worksheet: source worksheet object to copy values from

        :returns: a newly created :class:`worksheets <Worksheet>`.
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
            result = self.client.sh_batch_update(self.id, request, 'replies/addSheet', False)
            jsheet['properties'] = result['replies'][0]['addSheet']['properties']
            wks = self.worksheet_cls(self, jsheet)
        self._sheet_list.append(wks)
        return wks

    def del_worksheet(self, worksheet):
        """Deletes a worksheet from a spreadsheet.

        :param worksheet: The :class:`worksheets <Worksheet>` to be deleted.

        """
        if worksheet not in self.worksheets():
            raise WorksheetNotFound
        request = {"deleteSheet": {'sheetId': worksheet.id}}
        self.client.sh_batch_update(self.id, request, '', False)
        self._sheet_list.remove(worksheet)

    def find(self, string, replace=None, regex=True, match_case=False, include_formulas=False,
             srange=None, sheet=True):
        """
        Find and replace cells in spreadsheet

        :param string: string to search for
        :param replace: string to replace with
        :param regex: is the search string regex
        :param match_case: match case in search
        :param include_formulas: include seach in formula
        :param srange: range to search in A1 format
        :param sheet: if True - search all sheets, else search specified sheet

        """
        if not replace:
            found_list = []
            for wks in self.worksheets():
                found_list.extend(wks.find(string))
            return found_list
        body = {
            "find": string,
            "replacement": replace,
            "matchCase": match_case,
            "matchEntireCell": False,
            "searchByRegex": regex,
            "includeFormulas": include_formulas,
        }
        if srange:
            body['range'] = srange
        elif type(sheet) == bool:
            body['allSheets'] = True
        elif type(sheet) == int:
            body['sheetId'] = sheet
        body = {'findReplace': body}
        response = self.client.sh_batch_update(self.id, request=body, batch=self.batch_mode)
        return response['replies'][0]['findReplace']
    
    # @TODO impliment expiration time
    def share(self, addr, role='reader', expirationTime=None, is_group=False):
        """
        create/update permission for user/group/domain

        :param addr: this is the email for user/group and domain address for domains
        :param role: permission to be applied ('owner','writer','commenter','reader')
        :param expirationTime: (Not Implimented) time until this permission should last (datetime)
        :param is_group: boolean , Is this a use/group used only when email provided
        """
        return self.client.add_permission(self.id, addr, role=role, is_group=False)

    def list_permissions(self):
        """
        list all the permissions of the spreadsheet

        :returns: list of permissions as json object

        """
        permissions = self.client.list_permissions(self.id)
        self._permissions = permissions['permissions']
        return self._permissions

    def remove_permissions(self, addr):
        """
        Removes all permissions of the user provided

        :param addr: email/domain of the user

        """
        try:
            result = self.client.remove_permissions(self.id, addr, self._permissions)
        except InvalidUser:
            result = self.client.remove_permissions(self.id, addr)
        return result

    def batch_start(self):
        """
        Start batch mode, where all updates to sheet values will be batched

        """
        self.batch_mode = True
        warnings.warn('Batching is only for Update operations')

    def batch_stop(self, discard=False):
        """
        Stop batch Mode and Update the changes

        :param discard: discard all changes done in batch mode
        """
        self.batch_mode = False
        if not discard:
            self.client.send_batch(self.id)

    # @TODO
    def link(self, syncToCloud=False):
        """ Link the spreadsheet with cloud, so all local changes \
            will be updated instantly, so does all data fetches

            :param  syncToCloud: true ->  update the cloud with local changes
                                  false -> update the local copy with cloud
        """
        # just link all child sheets
        warnings.warn("method not implimented")

    # @TODO
    def unlink(self):
        """ Unlink the spread sheet with cloud, so all local changes
            will be made on local copy fetched
        """
        # just unlink all sheets
        warnings.warn("method not implimented")

    def export(self, fformat=ExportType.CSV, filename=None):
        """Export all the worksheet of the worksheet in specified format.

        :param fformat: A format of the output as Enum ExportType
        :param filename: name of file exported with extension
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
        send a custom batch update request for this spreadsheet

        :param request: the json batch update request or a list of requests
        :param fields: fields to include in the response
        :return: json Response
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
