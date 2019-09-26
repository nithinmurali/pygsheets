from pygsheets import utils
from pygsheets import exceptions


class GridRange(object):
    """ A range on a sheet. All indexes are zero-based.
    Indexes are half open, e.g the start index is inclusive and the end index is exclusive -- [startIndex, endIndex).
    Missing indexes indicate the range is unbounded on that side. """

    def __init__(self, label=None, worksheet_id=None, start=None, end=None, spreadsheet=None):
        self._label = label
        self._worksheet_name = None
        self._worksheet_id = worksheet_id
        self._start = start
        self._end = end
        self.spreadsheet = spreadsheet

        if label:
            self._update_values()
        else:
            self._update_label()

    @property
    def start(self):
        """ address of top left cell (index) """
        return self._start

    @start.setter
    def start(self, value):
        value = utils.format_addr(value, 'label')
        self._start = value
        self._update_label()

    @property
    def end(self):
        """ address of bottom right cell (index) """
        return self._end

    @end.setter
    def end(self, value):
        value = utils.format_addr(value, 'label')
        self._end = value
        self._update_label()

    @property
    def label(self):
        """ Label in grid range format """
        return self._label

    @label.setter
    def label(self, value):
        self._label = value
        self._update_values()

    @property
    def worksheet_id(self):
        return self._worksheet_id

    @worksheet_id.setter
    def worksheet_id(self, value):
        self._worksheet_id = value
        self._update_label()

    def set_worksheet(self, value):
        """ set the worksheet of this grid range. """
        self.spreadsheet = value.spreadsheet
        self._worksheet_id = value.id
        self._update_label()

    def _update_label(self):
        """update label from values """
        label = self._worksheet_name
        if self._start and self._end:
            label += "!" + self._start + ":" + self._end
        self._label = label

    def _update_values(self):
        """ update values from label """
        label = self._label
        self._worksheet_name = label.split('!')
        if len(label.split('!')) > 1:
            rem = label.split('!')[1]
            if ":" in rem:
                self._start = rem.split(":")[0]
                self._end = rem.split(":")[1]

    def to_json(self):
        self._update_values()
        # TODO fix
        return {
            "sheetId": self._worksheet_id,
            "startRowIndex": self._start[0]-1,
            "endRowIndex": self._end[0],
            "startColumnIndex": self._start[1]-1,
            "endColumnIndex": self._end[1],
        }
