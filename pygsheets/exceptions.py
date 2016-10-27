# -*- coding: utf-8 -*-

"""
pygsheets.exceptions
~~~~~~~~~~~~~~~~~~

Exceptions used in pygsheets.

"""


class PyGsheetsException(Exception):
    """A base class for pygsheets's exceptions."""


class AuthenticationError(PyGsheetsException):
    """An error during authentication process."""


class SpreadsheetNotFound(PyGsheetsException):
    """Trying to open non-existent or inaccessible spreadsheet."""


class WorksheetNotFound(PyGsheetsException):
    """Trying to open non-existent or inaccessible worksheet."""


class CellNotFound(PyGsheetsException):
    """Cell lookup exception."""


class NoValidUrlKeyFound(PyGsheetsException):
    """No valid key found in URL."""


class UnsupportedFeedTypeError(PyGsheetsException):
    pass


class UrlParameterMissing(PyGsheetsException):
    pass


class IncorrectCellLabel(PyGsheetsException):
    """The cell label is incorrect."""


class UpdateCellError(PyGsheetsException):
    """Error while setting cell's value."""


class RequestError(PyGsheetsException):
    """Error while sending API request."""


class InvalidArgumentValue(PyGsheetsException):
    """Invalid value foer argument"""


class HTTPError(RequestError):
    """DEPRECATED. Error while sending API request."""
    def __init__(self, code, msg):
        super(HTTPError, self).__init__(msg)
        self.code = code
