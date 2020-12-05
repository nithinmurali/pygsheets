# -*- coding: utf-8 -*-.

"""
pygsheets.datafilter
~~~~~~~~~~~~~~~~~~~~~

This module represents a datafilter that describes what data should be selected
or returned from a request.

"""

import logging
import warnings


class DataFilter(object):
    """Base class for DataFilter types"""
    def serialize(self):
        raise NotImplemented()


class A1RangeDataFilter(DataFilter):
    """Class for filtering data based on an A1 range

    :param a1_range:  The cell range to filter on
    """
    def __init__(self, a1_range):
        self.a1_range = a1_range

    def serialize(self):
        return {"a1Range": self.a1_range}


class GridRangeDataFilter(DataFilter):
    """Class for filtering data based on start and end indices

    This method uses zero-based indices. Start indices must be less
    than their corrisponding end index.

    :param sheet_id:            Worksheet id to filter on
    :param start_row_index:     Start row of the range to filter on
    :param end_row_index:       End row of the range to filter on
    :param start_column_index:  Start column of the range to filter on
    :param end_column_index:    End end of the range to filter on
    """
    def __init__(self, sheet_id, start_row_index, end_row_index, start_column_index, end_column_index):
        #TODO: make sure start and end indexes are > 0 and start < end
        self.sheet_id = sheet_id
        self.start_row_index = start_row_index
        self.end_row_index = end_row_index
        self.start_column_index = start_column_index
        self.end_column_index = end_column_index

    def serialize(self):
        return {
            "gridRange": {
                "sheetId": self.sheet_id,
                "startRowIndex": self.start_row_index,
                "endRowIndex": self.end_row_index,
                "startColumnIndex": self.start_column_index,
                "endColumnIndex": self.end_column_index
            }
        }


class DeveloperMetadataLookupDataFilter(DataFilter):
    """Class for filtering developer metadata queries

    This class only supports filtering for metadata on a whole spreadsheet or
    worksheet.

    :param spreadsheet_id:  Spreadsheet id to filter on (leave at None to search all metadata)
    :param sheet_id:        Worksheet id to filter on (leave at None for whole-spreadsheet metadata)
    :param meta_id:         Developer metadata id to filter on (optional)
    :param meta_key:        Developer metadata key to filter on (optional)
    :param meta_value:      Developer metadata value to filter on (optional)
    """
    def __init__(self, spreadsheet_id=None, sheet_id=None, meta_id=None, meta_key=None, meta_value=None):
        self.spreadsheet_id = spreadsheet_id
        self.sheet_id = sheet_id
        self.meta_filters = {
            "metadataId": meta_id,
            "metadataKey": meta_key,
            "metadataValue": meta_value,
            "metadataLocation": self.location
        }

    def serialize(self):
        lookup = dict((k,v) for k,v in self.meta_filters.items() if v is not None)
        return {"developerMetadataLookup": lookup}

    @property
    def location(self):
        if self.spreadsheet_id is not None:
            if self.sheet_id is None:
                return {"spreadsheet": True}
            elif self.sheet_id is not None:
                return {"sheetId": self.sheet_id}
        return None
