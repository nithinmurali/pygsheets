from pygsheets import Spreadsheet, Worksheet, ExportType

from googleapiclient import discovery
from googleapiclient.http import MediaIoBaseDownload

import logging
import json
import os


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
        :param kwargs:    Standard fields. See documentation for details.
        :return: Export request.
        """
        return self.service.files().export(fileId=file_id, mimeType=mime_type, **kwargs)

    def export(self, sheet, file_format, path='', filename=None):
        """Download a spreadsheet and store it.

        Uses one or several export request to download the files.
        Spreadsheets with multiple worksheets are exported into several files if the file format is CSV or TSV.
        For each worksheet the index is appended to the file name.

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


