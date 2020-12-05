# -*- coding: utf-8 -*-.

from pygsheets.datafilter import DeveloperMetadataLookupDataFilter

class DeveloperMetadata(object):
    @classmethod
    def new(cls, key, value, client, spreadsheet_id, sheet_id=None):
        """Create a new developer metadata entry

        Will return None when in batch mode, otherwise will return a DeveloperMetadata object

        :param key:             They key of the new developer metadata entry to create
        :param value:           They value of the new developer metadata entry to create
        :param client:          The client which is responsible to connect the sheet with the remote.
        :param spreadsheet_id:  The id of the spreadsheet where metadata will be created.
        :param sheet_id:        The id of the worksheet where the metadata will be created (optional)
        """
        filter = DeveloperMetadataLookupDataFilter(spreadsheet_id, sheet_id)
        meta_id = client.sheet.developer_metadata_create(spreadsheet_id, key, value, filter.location)
        if meta_id is None:
            # we're in batch mode
            return
        return cls(meta_id, key, value, client, spreadsheet_id, sheet_id)


    def __init__(self, meta_id, key, value, client, spreadsheet_id, sheet_id=None):
        """Create a new developer metadata entry

        Will return None when in batch mode, otherwise will return a DeveloperMetadata object
        :param meta_id:         The id of the developer metadata entry this represents
        :param key:             They key of the new developer metadata entry this represents
        :param value:           They value of the new developer metadata entry this represents
        :param client:          The client which is responsible to connect the sheet with the remote.
        :param spreadsheet_id:  The id of the spreadsheet where metadata is stored
        :param sheet_id:        The id of the worksheet where the metadata is stored (optional)
        """
        self._id = meta_id
        self.key = key
        self.value = value
        self.client = client
        self.spreadsheet_id = spreadsheet_id
        self.sheet_id = sheet_id
        self._filter = DeveloperMetadataLookupDataFilter(self.spreadsheet_id,
                                                         self.sheet_id, self.id)

    def __repr__(self):
        return "<DeveloperMetadata id={} key={} value={}>".format(repr(self.id),
                                                                  repr(self.key),
                                                                  repr(self.value))

    @property
    def id(self):
        return self._id

    def fetch(self):
        """Refresh this developer metadata entry from the spreadsheet"""
        response = self.client.sheet.developer_metadata_get(self.spreadsheet_id, self.id)
        self.key = response["metadataKey"]
        self.value = response["metadataValue"]

    def update(self):
        """Push the current local values to the spreadsheet"""
        self.client.sheet.developer_metadata_update(self.spreadsheet_id, self.key,
                                                    self.value, self._filter.location,
                                                    self._filter.serialize())

    def delete(self):
        """Delete this developer metadata entry"""
        self.client.sheet.developer_metadata_delete(self.spreadsheet_id,
                                                    self._filter.serialize())
