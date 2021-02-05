from pygsheets.spreadsheet import Spreadsheet
from pygsheets.worksheet import Worksheet
from pygsheets.custom_types import ExportType
from pygsheets.exceptions import InvalidArgumentValue, CannotRemoveOwnerError, FolderNotFound

from googleapiclient import discovery
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.errors import HttpError

import logging
import json
import os
import re


"""
pygsheets.drive
~~~~~~~~~~~~~~

This module provides wrappers for the Google Drive API v3.

"""

PERMISSION_ROLES = ['organizer', 'owner', 'writer', 'commenter', 'reader']
PERMISSION_TYPES = ['user', 'group', 'domain', 'anyone']

FIELDS_TO_INCLUDE = 'files(id, name, parents), nextPageToken, incompleteSearch'

_EMAIL_PATTERN = re.compile(r"\"?([-a-zA-Z0-9.`?{}]+@[-a-zA-Z0-9.]+\.\w+)\"?")


class DriveAPIWrapper(object):
    """A simple wrapper for the Google Drive API.

    Various utility and convenience functions to support access to Google Drive files. By default the
    requests will access the users personal drive. Use enable_team_drive(team_drive_id) to connect to a
    TeamDrive instead.

    Only functions used by pygsheet are wrapped. All other functionality can be accessed through the service
    attribute.

    See `reference <https://developers.google.com/drive/v3/reference/>`__ for details.

    :param http:            HTTP object to make requests with.
    :param data_path:       Path to the drive discovery file.
    """

    _spreadsheet_mime_type = "application/vnd.google-apps.spreadsheet"
    _folder_mime_type = "application/vnd.google-apps.folder"

    def __init__(self, http, data_path, retries=3, logger=logging.getLogger(__name__)):

        try:
            with open(os.path.join(data_path, "drive_discovery.json")) as jd:
                self.service = discovery.build_from_document(json.load(jd), http=http)
        except:
            self.service = discovery.build('drive', 'v3', http=http)
        self.team_drive_id = None
        self.include_team_drive_items = True
        """Include files from TeamDrive when executing requests."""
        self.logger = logger
        self.retries = retries

    def enable_team_drive(self, team_drive_id):
        """Access TeamDrive instead of the users personal drive."""
        self.team_drive_id = team_drive_id

    def disable_team_drive(self):
        """Do not access TeamDrive (default behaviour)."""
        self.team_drive_id = None

    def get_update_time(self, file_id):
        """Returns the time this file was last modified in RFC 3339 format."""
        return self._execute_request(self.service.files().get(fileId=file_id, fields='modifiedTime'))['modifiedTime']

    def list(self, **kwargs):
        """
        Fetch metadata of spreadsheets. Fetches a list of all files present in the users drive
        or TeamDrive. See Google Drive API Reference for details.

        Reference: `Files list request <https://developers.google.com/drive/v3/reference/files/list>`__

        :param kwargs:      Standard parameters (see documentation for details).
        :return:            List of metadata.
        """
        result = list()
        response = self._execute_request(self.service.files().list(**kwargs))
        result.extend(response['files'])
        while 'nextPageToken' in response:
            kwargs['pageToken'] = response['nextPageToken']
            response = self._execute_request(self.service.files().list(**kwargs))
            result.extend(response['files'])

        if 'incompleteSearch' in response and response['incompleteSearch']:
            self.logger.warning('Not all files in the corpora %s were searched. As a result '
                                'the response might be incomplete.', kwargs['corpora'])
        return result

    def create_folder(self, name, folder=None):
        """Create a new folder

        :param name:    The name to give the new folder
        :param parent:  The id of the folder this one will be stored in
        :return:        The new folder id
        """
        body = {
          'name': name,
          'mimeType': self._folder_mime_type
        }
        if folder:
            body['parents'] = [folder]
        return self._execute_request(self.service.files().create(body=body))["id"]

    def get_folder_id(self, name):
        """Fetch the first folder id with a given name

        :param name: The name of the folder to find
        """
        try:
            return list(filter(lambda x: x['name'] == name, self.folder_metadata()))[0]["id"]
        except (KeyError, IndexError):
            raise FolderNotFound('Could not find a folder with name %s.' % name)

    def folder_metadata(self, query='', only_team_drive=False):
        """Fetch folder names, ids & and parent folder ids.

        The query string can be used to filter the returned metadata.

        Reference: `search parameters docs. <https://developers.google.com/drive/v3/web/search-parameters>`__

        :param query:   Can be used to filter the returned metadata.
        """
        return self._metadata_for_mime_type(self._folder_mime_type, query, only_team_drive)

    def spreadsheet_metadata(self, query='', only_team_drive=False):
        """Fetch spreadsheet titles, ids & and parent folder ids.

        The query string can be used to filter the returned metadata.

        Reference: `search parameters docs. <https://developers.google.com/drive/v3/web/search-parameters>`__

        :param query:   Can be used to filter the returned metadata.
        """
        return self._metadata_for_mime_type(self._spreadsheet_mime_type, query, only_team_drive)

    def _metadata_for_mime_type(self, mime_type, query, only_team_drive):
        """
        Implementation for fetching drive object metadata by mime type
        """
        mime_type_query = "mimeType='{}'".format(mime_type)
        if query:
            query = query + ' and ' + str(mime_type_query)
        else:
            query = mime_type_query
        if self.team_drive_id:
            result = self.list(corpora='teamDrive',
                             teamDriveId=self.team_drive_id,
                             supportsTeamDrives=True,
                             includeTeamDriveItems=True,
                             fields=FIELDS_TO_INCLUDE,
                             q=query, pageSize=500, orderBy='recency')
            if not result and not only_team_drive:
                result = self.list(fields=FIELDS_TO_INCLUDE,
                                 supportsTeamDrives=True,
                                 includeTeamDriveItems=self.include_team_drive_items,
                                 q=query, pageSize=500, orderBy='recency')
            return result
        else:
            return self.list(fields=FIELDS_TO_INCLUDE,
                             supportsTeamDrives=True,
                             includeTeamDriveItems=self.include_team_drive_items,
                             q=query, pageSize=500, orderBy='recency')

    def delete(self, file_id, **kwargs):
        """Delete a file by ID.

        Permanently deletes a file owned by the user without moving it to the trash. If the file belongs to a
        Team Drive the user must be an organizer on the parent. If the input id is a folder, all descendants
        owned by the user are also deleted.

        Reference: `delete request <https://developers.google.com/drive/v3/reference/files/delete>`__

        :param file_id:     The Id of the file to be deleted.
        :param kwargs:      Standard parameters (see documentation for details).
        """
        if 'supportsTeamDrives' not in kwargs and self.team_drive_id:
            kwargs['supportsTeamDrives'] = True

        self._execute_request(self.service.files().delete(fileId=file_id, **kwargs))

    def move_file(self, file_id, old_folder, new_folder, body=None, **kwargs):
        """Move a file from one folder to another.

        Requires the current folder to delete it.

        Reference: `update request <https://developers.google.com/drive/v3/reference/files/update>`_

        :param file_id:     ID of the file which should be moved.
        :param old_folder:  Current location.
        :param new_folder:  Destination.
        :param body:        Other fields of the file to change. See reference for details.
        :param kwargs:      Optional arguments. See reference for details.
        """
        return self.update_file(file_id, body, removeParents=old_folder, addParents=new_folder, **kwargs)

    def copy_file(self, file_id, title, folder, body=None, **kwargs):
        """
        Copy a file from one location to another

        Reference: `update request`_

        :param file_id: Id of file to copy.
        :param title:   New title of the file.
        :param folder:  New folder where file should be copied.
        :param body:    Other fields of the file to change. See reference for details.
        :param kwargs:  Optional arguments. See reference for details.
        """
        if 'supportsTeamDrives' not in kwargs and self.team_drive_id:
            kwargs['supportsTeamDrives'] = True

        body = body or {}
        body['name'] = title
        if folder:
            body['parents'] = [folder]
        return self._execute_request(self.service.files().copy(fileId=file_id, body=body, **kwargs))

    def update_file(self, file_id, body=None, **kwargs):
        """Update file body.

        Reference: `update request <https://developers.google.com/drive/v3/reference/files/update>`_

        :param file_id:  ID of the file which should be updated.
        :param body:     The properties of the file to update. See reference for details.
        :param kwargs:   Optional arguments. See reference for details.
        """
        if 'supportsTeamDrives' not in kwargs and self.team_drive_id:
            kwargs['supportsTeamDrives'] = True

        if body is not None:
            kwargs["body"] = body

        return self._execute_request(self.service.files().update(fileId=file_id, **kwargs))

    def _export_request(self, file_id, mime_type, **kwargs):
        """The export request."""
        return self.service.files().export(fileId=file_id, mimeType=mime_type, **kwargs)

    def export(self, sheet, file_format, path='', filename=''):
        """Download a spreadsheet and store it.

         Exports a Google Doc to the requested MIME type and returns the exported content.

        .. warning::
            This can at most export files with 10 MB in size!

        Uses one or several export request to download the files. When exporting to CSV or TSV each worksheet is
        exported into a separate file. The API cannot put them into the same file. In this case the worksheet index
        is appended to the file-name.

        Reference: `request <https://developers.google.com/drive/v3/reference/files/export>`__

        :param sheet:           The spreadsheet or worksheet to be exported.
        :param file_format:     File format (:class:`ExportType`)
        :param path:            Path to where the file should be stored. (default: current working directory)
        :param filename:        Name of the file. (default: Spreadsheet Id)
        """
        request = None
        tmp = None
        mime_type, file_extension = getattr(file_format, 'value', file_format).split(':')

        if isinstance(sheet, Spreadsheet):
            if (file_format == ExportType.CSV or file_format == ExportType.TSV) and len(sheet.worksheets()) > 1:
                for worksheet in sheet:
                    self.export(worksheet, file_format, path=path, filename=filename + str(worksheet.index))
                return
            else:
                request = self._export_request(sheet.id, mime_type)
        elif isinstance(sheet, Worksheet):
            if sheet.index != 0:
                tmp = sheet.index
                try:
                    sheet.index = 0
                except HttpError:
                    raise Exception("Can only export first sheet in readonly mode")
            request = self._export_request(sheet.spreadsheet.id, mime_type)

        import io
        file_name = str(sheet.id or tmp) + file_extension if filename is None else filename + file_extension
        file_path = os.path.join(path, file_name)
        fh = io.FileIO(file_path, 'wb')
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            # logging.info('Download progress: %d%%.', int(status.progress() * 100)) TODO fix this
        logging.info('Download finished. File saved in %s.', file_path)

        if tmp is not None:
            sheet.index = tmp + 1
            if isinstance(sheet, Worksheet):
                sheet.refresh(False)

    def create_permission(self, file_id, role, type, **kwargs):
        """Creates a permission for a file or a TeamDrive.

        See `reference <https://developers.google.com/drive/v3/reference/permissions/create>`__ for more details.

        :param file_id:                 The ID of the file or Team Drive.
        :param role:                    The role granted by this permission.
        :param type:                    The type of the grantee.
        :keyword emailAddress:          The email address of the user or group to which this permission refers.
        :keyword domain:                The domain to which this permission refers.
        :parameter allowFileDiscovery:  Whether the permission allows the file to be discovered through search. This is
                                        only applicable for permissions of type domain or anyone.
        :keyword expirationTime:        The time at which this permission will expire (RFC 3339 date-time). Expiration
                                        times have the following restrictions:

                                            * They can only be set on user and group permissions
                                            * The time must be in the future
                                            * The time cannot be more than a year in the future

        :keyword emailMessage:          A plain text custom message to include in the notification email.
        :keyword sendNotificationEmail: Whether to send a notification email when sharing to users or groups.
                                        This defaults to true for users and groups, and is not allowed for other
                                        requests. It must not be disabled for ownership transfers.
        :keyword supportsTeamDrives:    Whether the requesting application supports Team Drives. (Default: False)
        :keyword transferOwnership:     Whether to transfer ownership to the specified user and downgrade
                                        the current owner to a writer. This parameter is required as an acknowledgement
                                        of the side effect. (Default: False)
        :keyword useDomainAdminAccess:  Whether the request should be treated as if it was issued by a
                                        domain administrator; if set to true, then the requester will be granted
                                        access if they are an administrator of the domain to which the item belongs.
                                        (Default: False)
        :return: `Permission Resource <https://developers.google.com/drive/v3/reference/permissions#resource>`_
        """
        if 'supportsTeamDrives' not in kwargs and self.team_drive_id:
            kwargs['supportsTeamDrives'] = True

        if 'emailAddress' in kwargs and 'domain' in kwargs:
            raise InvalidArgumentValue('A permission can only use emailAddress or domain. Do not specify both.')
        if role not in PERMISSION_ROLES:
            raise InvalidArgumentValue('A permission role can only be one of ' + str(PERMISSION_ROLES) + '.')
        if type not in PERMISSION_TYPES:
            raise InvalidArgumentValue('A permission role can only be one of ' + str(PERMISSION_TYPES) + '.')

        body = {
            'kind': 'drive#permission',
            'type': type,
            'role': role
        }

        if 'emailAddress' in kwargs:
            body['emailAddress'] = kwargs['emailAddress']
            del kwargs['emailAddress']
        elif 'domain' in kwargs:
            body['domain'] = kwargs['domain']
            del kwargs['domain']

        if 'allowFileDiscovery' in kwargs:
            body['allowFileDiscovery'] = kwargs['allowFileDiscovery']
            del kwargs['allowFileDiscovery']

        if 'expirationTime' in kwargs:
            body['expirationTime'] = kwargs['expirationTime']
            del kwargs['expirationTime']

        return self._execute_request(self.service.permissions().create(fileId=file_id, body=body, **kwargs))

    def list_permissions(self, file_id, **kwargs):
        """List all permissions for the specified file.

        See `reference <https://developers.google.com/drive/v3/reference/permissions/list>`__  for more details.

        :param file_id:                     The file to get the permissions for.
        :keyword pageSize:                  Number of permissions returned per request. (Default: all)
        :keyword supportsTeamDrives:        Whether the application supports TeamDrives. (Default: False)
        :keyword useDomainAdminAccess:      Request permissions as domain admin. (Default: False)
        :return: List of `Permission Resources <https://developers.google.com/drive/v3/reference/permissions#resource>`_
        """
        if 'supportsTeamDrives' not in kwargs and self.team_drive_id:
            kwargs['supportsTeamDrives'] = True

        # Ensure that all fields are returned. Default is only id, type & role.
        if 'fields' not in kwargs:
            kwargs['fields'] = '*'

        permissions = list()
        response = self._execute_request(self.service.permissions().list(fileId=file_id, **kwargs))
        permissions.extend(response['permissions'])
        while 'nextPageToken' in response:
            response = self._execute_request(self.service.permissions().list(fileId=file_id,
                                             pageToken=response['nextPageToken'], **kwargs))
            permissions.extend(response['permissions'])
        return permissions

    def delete_permission(self, file_id, permission_id, **kwargs):
        """Deletes a permission.

         See `reference <https://developers.google.com/drive/v3/reference/permissions/delete>`__  for more details.

        :param file_id:                 The ID of the file or Team Drive.
        :param permission_id:           The ID of the permission.
        :keyword supportsTeamDrives:    Whether the requesting application supports Team Drives. (Default: false)
        :keyword useDomainAdminAccess:  Whether the request should be treated as if it was issued by a
                                        domain administrator; if set to true, then the requester will be
                                        granted access if they are an administrator of the domain to which
                                        the item belongs. (Default: false)
        """
        if 'supportsTeamDrives' not in kwargs and self.team_drive_id:
            kwargs['supportsTeamDrives'] = True

        try:
            self._execute_request(self.service.permissions().delete(fileId=file_id, permissionId=permission_id, **kwargs))
        except HttpError as error:
            self.logger.exception(str(error))
            if re.search(r'The owner of a file cannot be removed\.', str(error)):
                raise CannotRemoveOwnerError('The owner of a file cannot be removed!')
            else:
                raise

    def _execute_request(self, request):
        """Executes a request.

        :param request: The request to be executed.
        :return:        Returns the response of the request.
        """
        return request.execute(num_retries=self.retries)
