# -*- coding: utf-8 -*-

"""
pygsheets
~~~~~~~~~

Google Spreadsheets client library.

"""

__version__ = '2.0.5'
__author__ = 'Nithin Murali'

from pygsheets.authorization import authorize
from pygsheets.spreadsheet import Spreadsheet
from pygsheets.worksheet import Worksheet
from pygsheets.cell import Cell
from pygsheets.datarange import DataRange
from pygsheets.address import GridRange, Address
from pygsheets.chart import Chart
from pygsheets.utils import format_addr
from pygsheets.custom_types import (FormatType, WorkSheetProperty, DateTimeRenderOption,
                                    ValueRenderOption, ExportType, ChartType, HorizontalAlignment,
                                    VerticalAlignment)
from pygsheets.exceptions import (PyGsheetsException, AuthenticationError,
                                  SpreadsheetNotFound, NoValidUrlKeyFound,
                                  IncorrectCellLabel, WorksheetNotFound,
                                  RequestError, CellNotFound, InvalidUser)


# Set default logging handler to avoid "No handler found" warnings.
import logging

logging.getLogger(__name__).addHandler(logging.NullHandler())
