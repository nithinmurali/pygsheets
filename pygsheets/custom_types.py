from enum import Enum


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
    """Enum for format of cell value"""
    CUSTOM = None
    TEXT = 'TEXT'
    NUMBER = 'NUMBER'
    PERCENT = 'PERCENT'
    CURRENCY = 'CURRENCY'
    DATE = 'DATE'
    TIME = 'TIME'
    DATE_TIME = 'DATE_TIME'
    SCIENTIFIC = 'SCIENTIFIC'

