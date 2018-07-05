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
    """Enum for how cell values are rendreed"""
    FORMATTED = 'FORMATTED_VALUE'
    UNFORMATTED = 'UNFORMATTED_VALUE'
    FORMULA = 'FORMULA'


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

    Reference: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#horizontalalign
    """
    LEFT = 'LEFT'
    RIGHT = 'RIGHT'
    CENTER = 'CENTER'
    NONE = None


class VerticalAlignment(Enum):
    """Vertical alignment of the cell.

    Reference: https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets#verticalalign
    """
    TOP = 'TOP'
    MIDDLE = 'MIDDLE'
    BOTTOM = 'BOTTOM'
    NONE = None
