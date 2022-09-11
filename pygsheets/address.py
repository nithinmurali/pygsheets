from pygsheets.exceptions import InvalidArgumentValue, IncorrectCellLabel
import re


class Address(object):
    """
    Represents the address of a cell.
    This can also be unbound in an axes. So 'A' is also a valid address but this
    requires explict setting of param `allow_non_single`.
    First index correspond to the rows, second index corresponds to columns.
    Integer Indexes start from 1.

    >>> a = Address('A2')
    >>> a.index
    (2, 1)
    >>> a.label
    'A2'
    >>> a[0]
    2
    >>> a[1]
    1
    >>> a = Address((1, 4))
    >>> a.index
    (1, 4)
    >>> a.label
    D1
    >>> b = a + (3,0)
    >>> b
    <Address D4>
    >>> b == (4, 4)
    True
    >>> column_a = Address((None, 1), True)
    >>> column_a
    <Address A>
    >>> row_2 = Address('2', True)
    >>> row_2
    <Address 2>
    """

    _MAGIC_NUMBER = 64

    def __init__(self, value, allow_non_single=False):
        self._is_single = True
        self.allow_non_single = allow_non_single

        if isinstance(value, str):
            self._value = self._label_to_coordinates(value)
        elif isinstance(value, tuple) or isinstance(value, list):
            assert len(value) == 2, 'tuple should be of length 2'
            assert type(value[0]) is int or value[0] is None, 'address row should be int'
            assert type(value[1]) is int or value[1] is None, 'address col should be int'
            self._value = tuple(value)
            self._validate()
        elif not value and self.allow_non_single:
            self._value = (None, None)
            self._validate()
        elif isinstance(value, Address):
            self._value = self._label_to_coordinates(value.label)
        else:
            raise IncorrectCellLabel('Only labels in A1 notation, coordinates as a tuple or '
                                     'pygsheets.Address objects are accepted.')

    @property
    def label(self):
        """ Label of the current address in A1 format."""
        return self._value_as_label()

    @property
    def row(self):
        """Row of the address"""
        return self._value[0]

    @property
    def col(self):
        """Column of the address"""
        return self._value[1]

    @row.setter
    def row(self, value):
        self._value = value, self._value[1]

    @col.setter
    def col(self, value):
        self._value = self._value[0], value

    @property
    def index(self):
        """Current Address in tuple format. Both axes starts at 1."""
        return tuple(self._value)

    def _validate(self):
        if not self.allow_non_single and (self._value[0] is None or self._value[0] is None):
            raise InvalidArgumentValue("Address cannot be unbounded if allow_non_single is not set.")

        if self._value[0]:
            row = int(self._value[0])
            if row < 1:
                raise InvalidArgumentValue('Address coordinates may not be below zero: ' + repr(self._value))

        if self._value[1]:
            col = int(self._value[1])
            if col < 1:
                raise InvalidArgumentValue('Address coordinates may not be below zero: ' + repr(self._value))

    def _value_as_label(self):
        """Transforms tuple coordinates into a label of the form A1."""
        self._validate()

        row_label, column_label = '', ''
        if self._value[0]:
            row_label = str(self._value[0])

        if self._value[1]:
            col = int(self._value[1])
            div = col
            column_label = ''
            while div:
                (div, mod) = divmod(div, 26)
                if mod == 0:
                    mod = 26
                    div -= 1
                column_label = chr(mod + self._MAGIC_NUMBER) + column_label

        return '{}{}'.format(column_label, row_label)

    def _label_to_coordinates(self, label):
        """Transforms a label in A1 notation into numeric coordinates and returns them as tuple."""
        m = re.match(r'([A-Za-z]*)(\d*)', label)
        if m:
            column_label = m.group(1).upper()
            row, col = m.group(2), 0
            if column_label:
                for i, c in enumerate(reversed(column_label)):
                    col += (ord(c) - self._MAGIC_NUMBER) * (26 ** i)
                col = int(col)
            else:
                col = None
            row = int(row) if row else None
        if not m or (not self.allow_non_single and not (row and col)):
            raise IncorrectCellLabel('Not a valid cell label format: {}.'.format(label))
        return row, col

    def __repr__(self):
        return '<%s %s>' % (self.__class__.__name__, str(self.label))

    def __iter__(self):
        return iter(self._value)

    def __getitem__(self, item):
        return self._value[item]

    def __setitem__(self, key, value):
        current_value = list(self._value)
        current_value[key] = value
        self._value = tuple(current_value)

    def __add__(self, other):
        if type(other) is tuple or isinstance(other, Address):
            return Address((self._value[0] + other[0], self._value[1] + other[1]))
        else:
            raise NotImplementedError

    def __sub__(self, other):
        if type(other) is tuple or isinstance(other, Address):
            return Address((self._value[0] - other[0], self._value[1] - other[1]))
        else:
            raise NotImplementedError

    def __eq__(self, other):
        if isinstance(other, Address):
            return self.label == other.label
        elif type(other) is str:
            return self.label == other
        elif type(other) is tuple or type(other) is list:
            return self._value == tuple(other)
        else:
            return super(Address, self).__eq__(other)

    def __ne__(self, other):
        return not self.__eq__(other)

    def __bool__(self):
        return not (self._value[0] is None and self._value[1] is None)

    __nonzero__ = __bool__


class GridRange(object):
    """
    Represents a rectangular (can be unbounded) range of adresses on a sheet.
    All indexes are one-based and are closed, ie the start index and the end index is inclusive
    Missing indexes indicate the range is unbounded on that side.

    A:B, A1:B3, 1:2 are all valid index, but A:1, 2:D are not

    grange.start = (1, None) will make the range unbounded on column
    grange.indexes = ((None, None), (None, None)) will make the range completely unbounded, ie. whole sheet

    Example:

    >>> grange = GridRange(worksheet=wks, start='A1', end='D4')
    >>> grange
    <GridRange Sheet1!A1:D4>
    >>> grange.start = 'A' # will remove bounding in rows
    <GridRange Sheet1!A:D>
    >>> grange.start = 'A1' # cannot add bounding at just start
    <GridRange Sheet1!A:D>
    >>> grange.indexes = ('A1', 'D4') # cannot add bounding at just start
    <GridRange Sheet1!A1:D4>
    >>> grange.end = (3, 5) # tuples will also work
    <GridRange Sheet1!A1:C5>
    >>> grange.end = (None, 5) # make unbounded on rows
    <GridRange Sheet1!1:5>
    >>> grange.end = (None, None) # make it unbounded on one index
    <GridRange Sheet1!1:1>
    >>> grange.start = None # make it unbounded on both indexes
    <GridRange Sheet1>
    >>> grange.start = 'A1' # make it unbounded on single index,now AZ100 is bottom right cell of worksheet
    <GridRange Sheet1:A1:AZ100>
    >>> 'A1' in grange
    True
    >>> (100,100) in grange
    False
    >>> for address in grange:
    >>>     print(address)
    Address((1,1))
    Address((1,2))
    ...

    Reference: `GridRange API docs <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/other#GridRange>`__

    """

    def __init__(self, label=None, worksheet=None, start=None, end=None, worksheet_title=None,
                 worksheet_id=None, propertiesjson=None):
        """
        :param label: label in A1 format
        :param worksheet: worksheet object this grange belongs to
        :param start: start address of the range
        :param end: end address of the range
        :param worksheet_title: worksheet title if worksheet not given
        :param worksheet_id: worksheet id if worksheet not given
        :param propertiesjson: json with all range properties
        """
        self._worksheet_title = worksheet_title
        self._worksheet_id = worksheet_id
        self._worksheet = worksheet
        self._start = Address(start, True)
        self._end = Address(end, True)
        # if fill_bounds and self._start and not self._end:
        #     if not worksheet:
        #         raise InvalidArgumentValue('worksheet need to fill bounds.')
        #     self._end = Address((worksheet.rows, worksheet.cols), True)
        # if fill_bounds and self._end and not self._start:
        #     self._start = Address('A1', True)
        if propertiesjson:
            self.set_json(propertiesjson)
        elif label:
            self._calculate_addresses(label)
        else:
            self._apply_index_constraints()
            self._calculate_label()

    @property
    def start(self):
        """ address of top left cell (index). """
        if not self._start and self._end:
            top_index = [1, 1]
            if not self._end[0]:
                top_index[0] = None
            if not self._end[1]:
                top_index[1] = None
            return Address(tuple(top_index), True)
        return self._start

    @start.setter
    def start(self, value):
        # prev = self._start, self._end
        self._start = Address(value, allow_non_single=True)
        self._apply_index_constraints()
        self._calculate_label()

    @property
    def end(self):
        """ address of bottom right cell (index) """
        if not self._end and self._start:
            if not self._worksheet:
                raise InvalidArgumentValue('worksheet is required for unbounded ranges')
            bottom_index = [self._worksheet.rows, self._worksheet.cols]
            if not self._start[0]:
                bottom_index[0] = None
            if not self._start[1]:
                bottom_index[1] = None
            return Address(tuple(bottom_index), True)
        return self._end

    @end.setter
    def end(self, value):
        # prev = self._start, self._end
        self._end = Address(value, allow_non_single=True)
        self._apply_index_constraints()
        self._calculate_label()

    @property
    def indexes(self):
        """ Indexes of this range as a tuple """
        return self.start, self.end

    @indexes.setter
    def indexes(self, value):
        if type(value) is not tuple:
            raise InvalidArgumentValue("Please provide a tuple")
        self._start, self._end = Address(value[0], True), Address(value[1], True)
        self._apply_index_constraints()
        self._calculate_label()

    @property
    def label(self):
        """ Label in A1 notation format """
        return self._calculate_label()

    @label.setter
    def label(self, value):
        if type(value) is not str:
            raise InvalidArgumentValue('non string value for label')
        self._calculate_addresses(value)

    @property
    def worksheet_id(self):
        """ Id of woksheet this range belongs to """
        if self._worksheet:
            return self._worksheet.id
        return self._worksheet_id

    @worksheet_id.setter
    def worksheet_id(self, value):
        if self._worksheet:
            if self._worksheet.id == value:
                return
            else:
                raise InvalidArgumentValue("This range already has a worksheet with different id set.")
        self._worksheet_id = value

    @property
    def worksheet_title(self):
        """ Title of woksheet this range belongs to """
        if self._worksheet:
            return self._worksheet.title
        return self._worksheet_title

    @worksheet_title.setter
    def worksheet_title(self, value):
        if not value:
            return
        if self._worksheet:
            if self._worksheet.title == value:
                return
            else:
                raise InvalidArgumentValue("This range already has a worksheet with different title set.")
        self._worksheet_title = value
        self._calculate_label()

    @staticmethod
    def create(data, wks=None):
        """
         create a Gridrange from various type of data
        :param data: can be string in A format,tuple or list, dict in GridRange format, GridRange object
        :param wks: worksheet to link to (optional)
        :return: GridRange object
        """
        if isinstance(data, GridRange):
            grange = data
        elif isinstance(data, str):
            grange = GridRange(label=data, worksheet=wks)
        elif isinstance(data, tuple) or isinstance(data, list):
            if len(data) < 2: raise InvalidArgumentValue("start and end required")
            grange = GridRange(start=data[0], end=data[1], worksheet=wks)
        elif isinstance(data, dict):
            grange = GridRange(propertiesjson=data, worksheet=wks)
        else:
            raise InvalidArgumentValue(data)
        if wks:
            grange.set_worksheet(wks)
        return grange

    def set_worksheet(self, value):
        """ set the worksheet of this grid range. """
        self._worksheet = value
        self._worksheet_id = value.id
        self._worksheet_title = value.title
        self._calculate_label()

    def _apply_index_constraints(self):
        if not self._start or not self._end:
            return

        # # If range was single celled, and one is set to none, make both unbound
        # if prev and prev[0] == prev[1]:
        #     if not self._start:
        #         self._end = self._start
        #         return
        #     if not self._end:
        #         self._start = self._end
        #         return
        # # if range is not single celled, and one index is unbounded make it single celled
        # if not self._end:
        #     self._end = self._start
        # if not self._start:
        #     self._start = self._end

        # Check if unbound on different axes
        if ((self._start[0] and not self._start[1]) and (not self._end[0] and self._end[1])) or \
           (not self._start[0] and self._start[1]) and (self._end[0] and not self._end[1]):
            self._start, self._end = Address(None, True), Address(None, True)
            raise InvalidArgumentValue('Invalid indexes set. Indexes should be unbounded at same axes.')

        # If one axes is unbounded on an index, make other index also unbounded on same axes
        if self._start[0] is None or self._end[0] is None:
            self._start[0], self._end[0] = None, None
        elif self._start[1] is None or self._end[1] is None:
            self._start[1], self._end[1] = None, None

        # verify
        # if (self._start[0] and not self._end[0]) or (not self._start[0] and self._end[0]) or \
        #    (self._start[1] and not self._end[1]) or (not self._start[1] and self._end[1]):
        #     self._start, self._end = Address(None, True), Address(None, True)
        #     raise InvalidArgumentValue('Invalid start and end set for this range')

        if self._start and self._end:
            if self._start[0]:
                assert self._start[0] <= self._end[0]
            if self._start[1]:
                assert self._start[1] <= self._end[1]

        self._calculate_label()

    def _calculate_label(self):
        """update label from values """
        label = "'" + self.worksheet_title + "'" if self.worksheet_title else ''
        if self.start and self.end:
            label += "!" + self.start.label + ":" + self.end.label
        return label

    def _calculate_addresses(self, label):
        """ update values from label """
        self._start, self._end = Address(None, True), Address(None, True)

        if len(label.split('!')) > 1:
            self.worksheet_title = label.split('!')[0]
            rem = label.split('!')[1]
            if ":" in rem:
                self._start = Address(rem.split(":")[0], allow_non_single=True)
                self._end = Address(rem.split(":")[1], allow_non_single=True)
            else:
                self._start = Address(rem, allow_non_single=True)
        elif self._worksheet:
            if ":" in label:
                self._start = Address(label.split(":")[0], allow_non_single=True)
                self._end = Address(label.split(":")[1], allow_non_single=True)
            else:
                self._start = Address(label, allow_non_single=True)
        else:
            pass

        self._apply_index_constraints()

    def to_json(self):
        """ Get json representation of this grid range. """
        if self.worksheet_id is None:
            raise Exception("worksheet id not set for this range.")
        return_dict = {"sheetId": self.worksheet_id}
        if self.start[0]:
            return_dict["startRowIndex"] = self.start[0] - 1
        if self.start[1]:
            return_dict["startColumnIndex"] = self.start[1] - 1
        if self.end[0]:
            return_dict["endRowIndex"] = self.end[0]
        if self.end[1]:
            return_dict["endColumnIndex"] = self.end[1]
        return return_dict

    def set_json(self, namedjson):
        """
        Apply a Gridrange json to this named range.

        :param namedjson: json object of the GridRange format

        Reference: `GridRange docs <https://developers.google.com/sheets/api/reference/rest/v4/spreadsheets/other#GridRange>`__
        """
        if 'sheetId' in namedjson:
            self.worksheet_id = namedjson['sheetId']
        start_row_idx = namedjson.get('startRowIndex', None)
        end_row_idx = namedjson.get('endRowIndex', None)
        start_col_idx = namedjson.get('startColumnIndex', None)
        end_col_idx = namedjson.get('endColumnIndex', None)

        start_row_idx = start_row_idx + 1 if start_row_idx is not None else start_row_idx
        start_col_idx = start_col_idx + 1 if start_col_idx is not None else start_col_idx

        self._start = Address((start_row_idx, start_col_idx), True)
        self._end = Address((end_row_idx, end_col_idx), True)

        self._calculate_label()

    def get_bounded_indexes(self):
        """ get bounded indexes of this range based on worksheet size, if the indexes are unbounded """
        start_r, start_c = tuple(iter(self.start)) if self.start else (None, None)
        end_r, end_c = tuple(iter(self.end)) if self.end else (None, None)
        start_r = start_r if start_r else 1
        start_c = start_c if start_c else 1
        if not self._worksheet and not (end_r or end_c):
            raise InvalidArgumentValue('Worksheet not set for calculating size.')
        end_r = end_r if end_r else self._worksheet.rows
        end_c = end_c if end_c else self._worksheet.cols
        return Address((start_r, start_c)), Address((end_r, end_c))

    @property
    def height(self):
        """ Height of this gridrange """
        start, end = self.get_bounded_indexes()
        return end[0] - start[0] + 1

    @property
    def width(self):
        """ Width of this gridrange """
        start, end = self.get_bounded_indexes()
        return end[1] - start[1] + 1

    def contains(self, address):
        return self.start[0] <= address.row <= self.end[0] and self.start[1] <= address.col <= self.end[1]

    def __repr__(self):
        return '<%s %s>' % (self.__class__.__name__, str(self.label))

    def __eq__(self, other):
        if isinstance(other, GridRange):
            return self.label == other.label
        elif type(other) is str:
            return self.label == other
        else:
            return super(GridRange, self).__eq__(other)

    def __ne__(self, other):
        return not self.__eq__(other)

    def __contains__(self, item):
        try:
            item = Address(item)
        except IncorrectCellLabel:
            raise InvalidArgumentValue("Gridrange can only contain an address")
        return self.contains(item)

    def __iter__(self):
        for r in range(self.start[0], self.end[0]+1):
            for c in range(self.start[1], self.end[1]+1):
                yield Address((r, c))
