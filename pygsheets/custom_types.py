# -*- coding: utf-8 -*-.

"""
pygsheets.custom_types
~~~~~~~~~~~~~~~~~~~~~~

This module contains common Enums used in pygsheets

"""

from enum import Enum


# @TODO use this
class WorkSheetProperty(Enum):
    """available properties of worksheets"""
    TITLE = 'title'
    ID = 'id'
    INDEX = 'index'


class ValueRenderOption(Enum):
    """Determines how values should be rendered in the output.

    `ValueRenderOption Docs <https://developers.google.com/sheets/api/reference/rest/v4/ValueRenderOption>`_

    FORMATTED_VALUE: Values will be calculated & formatted in the reply according to the cell's formatting.
    Formatting is based on the spreadsheet's locale, not the requesting user's locale.
    For example, if A1 is 1.23 and A2 is =A1 and formatted as currency,
    then A2 would return "$1.23".

    UNFORMATTED_VALUE : Values will be calculated, but not formatted in the reply.
    For example, if A1 is 1.23 and A2 is =A1 and formatted as currency,
    then A2 would return the number 1.23.

    FORMULA : Values will not be calculated. The reply will include the formulas.
    For example, if A1 is 1.23 and A2 is =A1 and formatted as currency, then A2 would return "=A1".
    """
    FORMATTED_VALUE = 'FORMATTED_VALUE'
    UNFORMATTED_VALUE = 'UNFORMATTED_VALUE'
    FORMULA = 'FORMULA'


class DateTimeRenderOption(Enum):
    """Determines how dates should be rendered in the output.

    `DateTimeRenderOption Doc <https://developers.google.com/sheets/api/reference/rest/v4/DateTimeRenderOption>`_

    SERIAL_NUMBER: Instructs date, time, datetime, and duration fields to be output as doubles in "serial number"
    format, as popularized by Lotus 1-2-3. The whole number portion of the value (left of the
    decimal) counts the days since December 30th 1899. The fractional portion (right of the decimal)
    counts the time as a fraction of the day. For example, January 1st 1900 at noon would be 2.5, 2
    because it's 2 days after December 30st 1899, and .5 because noon is half a day. February 1st
    1900 at 3pm would be 33.625. This correctly treats the year 1900 as not a leap year.

    FORMATTED_STRING: Instructs date, time, datetime, and duration fields to be output as strings in their given
    number format (which is dependent on the spreadsheet locale).
    """
    SERIAL_NUMBER = 'SERIAL_NUMBER'
    FORMATTED_STRING = 'FORMATTED_STRING'


class FormatType(Enum):
    """Enum for cell formats."""
    CUSTOM = None
    TEXT = 'TEXT'
    NUMBER = 'NUMBER'
    PERCENT = 'PERCENT'
    CURRENCY = 'CURRENCY'
    DATE = 'DATE'
    TIME = 'TIME'
    DATE_TIME = 'DATE_TIME'
    SCIENTIFIC = 'SCIENTIFIC'


class ExportType(Enum):
    """Enum for possible export types"""
    XLS = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet:.xls"
    ODT = "application/x-vnd.oasis.opendocument.spreadsheet:.odt"
    PDF = "application/pdf:.pdf"
    CSV = "text/csv:.csv"
    TSV = 'text/tab-separated-values:.tsv'
    HTML = 'application/zip:.zip'


class HorizontalAlignment(Enum):
    """Horizontal alignment of the cell.

    `HorizontalAlignment doc <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#horizontalalign>`_

    """
    LEFT = 'LEFT'
    RIGHT = 'RIGHT'
    CENTER = 'CENTER'
    NONE = None


class VerticalAlignment(Enum):
    """Vertical alignment of the cell.

    `VerticalAlignment doc <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#verticalalign>`_

    """
    TOP = 'TOP'
    MIDDLE = 'MIDDLE'
    BOTTOM = 'BOTTOM'
    NONE = None


class ChartType(Enum):
    """Enum for basic chart types

    Reference: `insert request <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#BasicChartType>`_
    """
    BAR = "BAR"
    LINE = "LINE"
    AREA = "AREA"
    COLUMN = "COLUMN"
    SCATTER = "SCATTER"
    COMBO = "COMBO"
    STEPPED_AREA = "STEPPED_AREA"
