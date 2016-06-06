# -*- coding: utf-8 -*-

"""
pysheets.exceptions
~~~~~~~~~~~~~~~~~~

Exceptions used in gspread.

"""

class PySheetsException(Exception):
    """A base class for gspread's exceptions."""

class AuthenticationError(PySheetsException):
    """An error during authentication process."""

class SpreadsheetNotFound(PySheetsException):
    """Trying to open non-existent or inaccessible spreadsheet."""

class WorksheetNotFound(PySheetsException):
    """Trying to open non-existent or inaccessible worksheet."""

class CellNotFound(PySheetsException):
    """Cell lookup exception."""

class NoValidUrlKeyFound(PySheetsException):
    """No valid key found in URL."""

class UnsupportedFeedTypeError(PySheetsException):
    pass

class UrlParameterMissing(PySheetsException):
    pass

class IncorrectCellLabel(PySheetsException):
    """The cell label is incorrect."""

class UpdateCellError(PySheetsException):
    """Error while setting cell's value."""

class RequestError(PySheetsException):
    """Error while sending API request."""

class InvalidArgumentValue(PySheetsException):
    '''Invalid value foer argument'''

class HTTPError(RequestError):
    """DEPRECATED. Error while sending API request."""
    def __init__(self, code, msg):
        super(HTTPError, self).__init__(msg)
        self.code = code
