# -*- coding: utf-8 -*-.
import re
import warnings
import os
import logging


from pygsheets.drive import DriveAPIWrapper
from pygsheets.sheet import SheetAPIWrapper
from pygsheets.spreadsheet import Spreadsheet
from pygsheets.exceptions import SpreadsheetNotFound, NoValidUrlKeyFound
from pygsheets.custom_types import ValueRenderOption, DateTimeRenderOption

from google_auth_httplib2 import AuthorizedHttp

GOOGLE_SHEET_CELL_UPDATES_LIMIT = 50000

_url_key_re_v1 = re.compile(r'key=([^&#]+)')
_url_key_re_v2 = re.compile(r"/spreadsheets/d/([a-zA-Z0-9-_]+)")
_email_patttern = re.compile(r"\"?([-a-zA-Z0-9.`?{}]+@[-a-zA-Z0-9.]+\.\w+)\"?")
# _domain_pattern = re.compile("(?!-)[A-Z\d-]{1,63}(?<!-)$", re.IGNORECASE)

_deprecated_keyword_mapping = {
    'parent_id': 'folder',
}


class Client(object):
    """Create or access Google spreadsheets.

    Exposes members to create new spreadsheets or open existing ones. Use `authorize` to instantiate an instance of this
    class.

    >>> import pygsheets
    >>> c = pygsheets.authorize()

    The sheet API service object is stored in the sheet property and the drive API service object in the drive property.

    >>> c.sheet.get('<SPREADSHEET ID>')
    >>> c.drive.delete('<FILE ID>')

    :param credentials:             The credentials object returned by google-auth or google-auth-oauthlib.
    :param retries:                 (Optional) Number of times to retry a connection before raising a TimeOut error.
                                    Default: 3
    :param http:                    The underlying HTTP object to use to make requests. If not specified, a
                                    :class:`httplib2.Http` instance will be constructed.
    :param check:                   Check for quota error and apply rate limiting.
    :param seconds_per_quota:       Default value is 100 seconds

    """

    spreadsheet_cls = Spreadsheet

    def __init__(self, credentials, retries=3, http=None, check=True, seconds_per_quota=100):
        self.oauth = credentials
        self.logger = logging.getLogger(__name__)

        http = AuthorizedHttp(credentials, http=http)
        data_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")

        self.sheet = SheetAPIWrapper(http, data_path, retries=retries, check=check, seconds_per_quota=seconds_per_quota)
        self.drive = DriveAPIWrapper(http, data_path)

    @property
    def teamDriveId(self):
        """ Enable team drive support

            Deprecated: use client.drive.enable_team_drive(team_drive_id=?)
        """
        return self.drive.team_drive_id

    @teamDriveId.setter
    def teamDriveId(self, value):
        warnings.warn("Depricated  please use drive.enable_team_drive")
        self.drive.enable_team_drive(value)

    def spreadsheet_ids(self, query=None):
        """Get a list of all spreadsheet ids present in the Google Drive or TeamDrive accessed."""
        return [x['id'] for x in self.drive.spreadsheet_metadata(query)]

    def spreadsheet_titles(self, query=None):
        """Get a list of all spreadsheet titles present in the Google Drive or TeamDrive accessed."""
        return [x['name'] for x in self.drive.spreadsheet_metadata(query)]

    def create(self, title, template=None, folder=None, **kwargs):
        """Create a new spreadsheet.

        The title will always be set to the given value (even overwriting the templates title). The template
        can either be a `spreadsheet resource <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#resource-spreadsheet>`_
        or an instance of :class:`~pygsheets.Spreadsheet`. In both cases undefined values will be ignored.

        :param title:       Title of the new spreadsheet.
        :param template:    A template to create the new spreadsheet from.
        :param folder:      The Id of the folder this sheet will be stored in.
        :param kwargs:      Standard parameters (see reference for details).
        :return: :class:`~pygsheets.Spreadsheet`
        """
        result = self.sheet.create(title, template=template, **kwargs)
        if folder:
            self.drive.move_file(result['spreadsheetId'],
                                 old_folder=self.drive.spreadsheet_metadata(query="name = '" + title + "'")[0]['parents'][0],
                                 new_folder=folder)
        return self.spreadsheet_cls(self, jsonsheet=result)

    def open(self, title):
        """Open a spreadsheet by title.

        In a case where there are several sheets with the same title, the first one found is returned.

        >>> import pygsheets
        >>> c = pygsheets.authorize()
        >>> c.open('TestSheet')

        :param title:                           A title of a spreadsheet.

        :returns:                               :class:`~pygsheets.Spreadsheet`
        :raises pygsheets.SpreadsheetNotFound:  No spreadsheet with the given title was found.
        """
        try:
            spreadsheet = list(filter(lambda x: x['name'] == title, self.drive.spreadsheet_metadata()))[0]
            return self.open_by_key(spreadsheet['id'])
        except (KeyError, IndexError):
            raise SpreadsheetNotFound('Could not find a spreadsheet with title %s.' % title)

    def open_by_key(self, key):
        """Open a spreadsheet by key.

        >>> import pygsheets
        >>> c = pygsheets.authorize()
        >>> c.open_by_key('0BmgG6nO_6dprdS1MN3d3MkdPa142WFRrdnRRUWl1UFE')

        :param key:                             The key of a spreadsheet. (can be found in the sheet URL)
        :returns:                               :class:`~pygsheets.Spreadsheet`
        :raises pygsheets.SpreadsheetNotFound:  The given spreadsheet ID was not found.
        """
        response = self.sheet.get(key,
                                  fields='properties,sheets/properties,spreadsheetId,namedRanges',
                                  includeGridData=False)
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

    def open_all(self, query=''):
        """Opens all available spreadsheets.

        Result can be filtered when specifying the query parameter. On the details on how to form the query:

        `Reference <https://developers.google.com/drive/v3/web/search-parameters>`_

        :param query:   (Optional) Can be used to filter the returned metadata.
        :returns:       A list of :class:`~pygsheets.Spreadsheet`.
        """
        return [self.open_by_key(key) for key in self.spreadsheet_ids(query=query)]

    def open_as_json(self, key):
        """Return a json representation of the spreadsheet.

        See `Reference <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#Spreadsheet>`__ for details.
        """
        return self.sheet.get(key, fields='properties,sheets/properties,sheets/protectedRanges,'
                                          'spreadsheetId,namedRanges',
                              includeGridData=False)

    def get_range(self, spreadsheet_id,
                  value_range,
                  major_dimension='ROWS',
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
                                            render option is `ValueRenderOption.FORMATTED_VALUE`.
        :param date_time_render_option:     How dates, times, and durations should be represented in the output.
                                            This is ignored if `valueRenderOption` is `FORMATTED_VALUE`. The default
                                            dateTime render option is [`DateTimeRenderOption.SERIAL_NUMBER`].
        :return:                            An array of arrays with the values fetched. Returns an empty array if no
                                            values were fetched. Values are dynamically typed as int, float or string.
        """
        result = self.sheet.values_get(spreadsheet_id, value_range, major_dimension, value_render_option,
                                       date_time_render_option)
        try:
            return result['values']
        except KeyError:
            return [['']]