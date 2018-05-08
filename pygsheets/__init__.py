# -*- coding: utf-8 -*-

"""
pygsheets
~~~~~~~~~

Google Spreadsheets client library.

"""

__version__ = '1.1.4'
__author__ = 'Nithin Murali'


from .client import Client, authorize
from .spreadsheet import Spreadsheet
from .worksheet import Worksheet
from .cell import Cell
from .datarange import DataRange
from .utils import format_addr
from .custom_types import (FormatType, WorkSheetProperty,
                           ValueRenderOption, ExportType)
from .exceptions import (PyGsheetsException, AuthenticationError,
                         SpreadsheetNotFound, NoValidUrlKeyFound,
                         IncorrectCellLabel, WorksheetNotFound,
                         RequestError, CellNotFound, InvalidUser,
                         InvalidArgumentValue)


# Set default logging handler to avoid "No handler found" warnings.
import logging
try:  # Python 2.7+
    from logging import NullHandler
except ImportError:
    class NullHandler(logging.Handler):
        def emit(self, record):
            pass

logging.getLogger(__name__).addHandler(NullHandler())
