# -*- coding: utf-8 -*-

"""
pysheets
~~~~~~~

Google Spreadsheets client library.

"""

__version__ = '0.1'
__author__ = 'Nithin M'


try:
    from urllib import urlencode
except ImportError:
    from urllib.parse import urlencode


from .client import Client, login, authorize
from .models import Spreadsheet, Worksheet, Cell
from .exceptions import (GSpreadException, AuthenticationError,
                         SpreadsheetNotFound, NoValidUrlKeyFound,
                         IncorrectCellLabel, WorksheetNotFound,
                         UpdateCellError, RequestError, CellNotFound)
