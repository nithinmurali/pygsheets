from pygsheets.spreadsheet import Spreadsheet
from pygsheets.worksheet import Worksheet
from pygsheets.custom_types import ExportType
from pygsheets.exceptions import InvalidArgumentValue, CannotRemoveOwnerError

from googleapiclient import discovery
from googleapiclient.http import MediaIoBaseDownload
from googleapiclient.errors import HttpError

import logging
import json
import os
import re


PERMISSION_ROLES = ['organizer', 'owner', 'writer', 'commenter', 'reader']
PERMISSION_TYPES = ['user', 'group', 'domain', 'anyone']

_EMAIL_PATTERN = re.compile(r"\"?([-a-zA-Z0-9.`?{}]+@[-a-zA-Z0-9.]+\.\w+)\"?")


class DriveAPIWrapper(object):

    def __init__(self, http, data_path, logger=logging.getLogger(__name__)):
        """A simple wrapper for the Google Drive API.

        Various utility and convenience functions to support access to Google Drive files. By default the
        requests will access the users personal drive. Use enable_team_drive(team_drive_id) to connect to a
        TeamDrive instead.

        Only functions used by pygsheet are wrapped. All other functionality can be accessed through the service
        attribute.

        Reference: https://developers.google.com/drive/v3/reference/

        :param http:            HTTP object to make requests with.
        :param data_path:       Path to the drive discovery file.
        """
        with open(os.path.join(data_path, "drive_discovery.json")) as jd:
            self.service = discovery.build_from_document(json.load(jd), http=http)
        self.team_drive_id = None
        self.logger = logger

    def enable_team_drive(self, team_drive_id):
        """All requests will request files & data from this TeamDrive."""
        self.team_drive_id = team_drive_id

    def disable_team_drive(self):
        """All requests will request files & data from the users personal drive."""
        self.team_drive_id = None

    def _export_request(self, file_id, mime_type, **kwargs):
        """Export a Google Doc.

        Exports a Google Doc to the requested MIME type and returns the exported content.

        IMPORTANT: This can at most export files with 10 MB in size!

        Reference: https://developers.google.com/drive/v3/reference/files/export

        :param file_id:     The file to be exported.
        :param mime_type:   The export MIME Type.
        :param kwargs:      Standard fields. See documentation for details.
        :return: Export request.
        """
        return self.service.files().export(fileId=file_id, mimeType=mime_type, **kwargs)

    def export(self, sheet, file_format, path='', filename=None):
        """Download a spreadsheet and store it.

        Uses one or several export request to download the files.
        Spreadsheets with multiple worksheets are exported into several files if the file format is CSV or TSV.
        In this case the index of each worksheet is appended to the file name.

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
            tmp = sheet.index
            sheet.index = 0
            request = self._export_request(sheet.spreadsheet.id, mime_type)

        import io
        file_name = sheet.id + file_extension if filename is None else filename + file_extension
        fh = io.FileIO(path + file_name, 'wb')
        downloader = MediaIoBaseDownload(fh, request)
        done = False
        while done is False:
            status, done = downloader.next_chunk()
            logging.info('Download progress: %d%%.', int(status.progress() * 100))
        logging.info('Download finished. File saved in %s.', path + file_name + file_extension)

        if tmp is not None:
            sheet.index = tmp

    def create_permission(self, file_id, role, type, **kwargs):
        """Creates a permission for a file or a TeamDrive.
        Reference: https://developers.google.com/drive/v3/reference/permissions/create
        :param file_id:                 The ID of the file or Team Drive.
        :param role:                    The role granted by this permission.
        :param type:                    The type of the grantee.
        :keyword emailAddress           The email address of the user or group to which this permission refers.
        :keyword domain                 The domain to which this permission refers.
        :keyword allowFileDiscovery     Whether the permission allows the file to be discovered through search.
                                        This is only applicable for permissions of type domain or anyone.
        :keyword expirationTime:        The time at which this permission will expire (RFC 3339 date-time).
                                        Expiration times have the following restrictions:
                                            - They can only be set on user and group permissions
                                            - The time must be in the future
                                            - The time cannot be more than a year in the future
        Request Arguments:
        -------------------
        :keyword emailMessage:          A plain text custom message to include in the notification email.
        :keyword sendNotificationEmail: Whether to send a notification email when sharing to users or groups.
                                        This defaults to true for users and groups, and is not allowed for other
                                        requests. It must not be disabled for ownership transfers.
        :keyword supportsTeamDrives:    Whether the requesting application supports Team Drives. (Default: false)
        :keyword transferOwnership:     Whether to transfer ownership to the specified user and downgrade
                                        the current owner to a writer. This parameter is required as an acknowledgement
                                        of the side effect. (Default: false)
        :keyword useDomainAdminAccess:  Whether the request should be treated as if it was issued by a
                                        domain administrator; if set to true, then the requester will be granted
                                        access if they are an administrator of the domain to which the item belongs.
                                        (Default: false)
        :return: Permission Resource: https://developers.google.com/drive/v3/reference/permissions#resource
        """
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
            if _EMAIL_PATTERN.match(kwargs['emailAddress']):
                body['emailAddress'] = kwargs['emailAddress']
                del kwargs['emailAddress']
            else:
                raise InvalidArgumentValue("The provided e-mail address doesn't have a valid format: " +
                                           kwargs['emailAddress'] + '.')
        elif 'domain' in kwargs:
            body['domain'] = kwargs['domain']
            del kwargs['domain']

        if 'allowFileDiscovery' in kwargs:
            body['allowFileDiscovery'] = kwargs['allowFileDiscovery']
            del kwargs['allowFileDiscovery']

        if 'expirationTime' in kwargs:
            body['expirationTime'] = kwargs['expirationTime']
            del kwargs['expirationTime']

        return self.service.permissions().create(fileId=file_id, body=body, **kwargs).execute()

    def list_permissions(self, file_id, **kwargs):
        """List all permissions for the specified file.

        Reference: https://developers.google.com/drive/v3/reference/permissions/list

        :param file_id:                     The file to get the permissions for.
        :keyword pageSize:                  Number of permissions returned per request. (default: all)
        :keyword supportsTeamDrives:        Whether the application supports TeamDrives. (default: False)
        :keyword useDomainAdminAccess:      Request permissions as domain admin. (default: False)
        :return:
        """
        permissions = list()
        response = self.service.permissions().list(fileId=file_id, **kwargs).execute()
        permissions.extend(response['permissions'])
        while 'nextPageToken' in response:
            response = self.service.permissions().list(fileId=file_id,
                                                       pageToken=response['nextPageToken'], **kwargs).execute()
            permissions.extend(response['permissions'])
        return permissions

    def delete_permission(self, file_id, permission_id, **kwargs):
        """Deletes a permission.
        Reference: https://developers.google.com/drive/v3/reference/permissions/delete
        :param file_id:                 The ID of the file or Team Drive.
        :param permission_id:           The ID of the permission.
        :keyword supportsTeamDrives:    Whether the requesting application supports Team Drives. (Default: false)
        :keyword useDomainAdminAccess:  Whether the request should be treated as if it was issued by a
                                        domain administrator; if set to true, then the requester will be
                                        granted access if they are an administrator of the domain to which
                                        the item belongs. (Default: false)
        """
        try:
            self.service.permissions().delete(fileId=file_id, permissionId=permission_id, **kwargs).execute()
        except HttpError as error:
            self.logger.exception(str(error))
            if re.search(r'The owner of a file cannot be removed\.', str(error)):
                raise CannotRemoveOwnerError('The owner of a file cannot be removed!')
            else:
                raise


