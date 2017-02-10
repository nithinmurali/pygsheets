# -*- coding: utf-8 -*-

"""
pygsheets.utils
~~~~~~~~~~~~~~~

This module contains utility functions.

"""

from .exceptions import (IncorrectCellLabel, InvalidArgumentValue)
import re


def finditem(func, seq):
    """Finds and returns first item in iterable for which func(item) is True.
    """
    return next((item for item in seq if func(item)))


def numericise(value, empty_value=''):
    """Returns a value that depends on the input string:
        - Float if input can be converted to Float
        - Integer if input can be converted to integer
        - Zero if the input string is empty and empty2zero flag is set
        - The same input string, empty or not, otherwise.

    Executable examples:

    >>> numericise("faa")
    'faa'
    >>> numericise("3")
    3
    >>> numericise("3.1")
    3.1
    >>> numericise("", empty2zero=True)
    0
    >>> numericise("", empty2zero=False)
    ''
    >>> numericise("")
    ''
    >>> numericise(None)
    >>>
    """
    if value == '':
        return empty_value
    if value is not None:
        try:
            value = int(value)
        except ValueError:
            try:
                value = float(value)
            except ValueError:
                pass
    return value


def numericise_all(input, empty_value=''):
    """Returns a list of numericised values from strings"""
    return [numericise(s, empty_value) for s in input]


def format_addr(addr, output='flip'):
        """
        function to convert address format of cells from one to another

        :param addr: address as tuple or label
        :param output: -'label' will output label
                      - 'tuple' will output tuple
                      - 'flip' will convert to other type
        :returns: tuple or label
        """
        _MAGIC_NUMBER = 64
        if type(addr) == tuple:
            if output == 'label' or output == 'flip':
                # return self.get_addr_int(*addr)
                row = int(addr[0])
                col = int(addr[1])
                if row < 1 or col < 1:
                    raise IncorrectCellLabel('(%s, %s)' % (row, col))
                div = col
                column_label = ''
                while div:
                    (div, mod) = divmod(div, 26)
                    if mod == 0:
                        mod = 26
                        div -= 1
                    column_label = chr(mod + _MAGIC_NUMBER) + column_label
                label = '%s%s' % (column_label, row)
                return label

            elif output == 'tuple':
                return addr

        elif type(addr) == str:
            if output == 'tuple' or output == 'flip':
                _cell_addr_re = re.compile(r'([A-Za-z]+)(\d+)')
                m = _cell_addr_re.match(addr)
                if m:
                    column_label = m.group(1).upper()
                    row, col = int(m.group(2)), 0
                    for i, c in enumerate(reversed(column_label)):
                        col += (ord(c) - _MAGIC_NUMBER) * (26 ** i)
                else:
                    raise IncorrectCellLabel(addr)
                return int(row), int(col)
            elif output == 'label':
                return addr
        else:
            raise InvalidArgumentValue

