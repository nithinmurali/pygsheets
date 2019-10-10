# -*- coding: utf-8 -*-

"""
pygsheets.utils
~~~~~~~~~~~~~~~

This module contains utility functions.

"""

from pygsheets.exceptions import (IncorrectCellLabel, InvalidArgumentValue)
from functools import wraps
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


def is_number(n):
    try:
        float(n)
    except ValueError:
        return False
    return True


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
                if addr[0] is None:
                    row_label = ''
                else:
                    row = int(addr[0])
                    if row < 1:
                        raise IncorrectCellLabel(repr(addr))
                    row_label = str(row)

                if addr[1] is None:
                    column_label = ''
                else:
                    col = int(addr[1])
                    if col < 1:
                        raise IncorrectCellLabel(repr(addr))
                    div = col
                    column_label = ''
                    while div:
                        (div, mod) = divmod(div, 26)
                        if mod == 0:
                            mod = 26
                            div -= 1
                        column_label = chr(mod + _MAGIC_NUMBER) + column_label
                label = '%s%s' % (column_label, row_label)
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
            raise InvalidArgumentValue("addr of type " + str(type(addr)))


def fullmatch(regex, string, flags=0):
    """Emulate python-3.4 re.fullmatch()."""
    return re.match("(?:" + regex + r")\Z", string, flags=flags)


def format_color(data, to='dict'):
    """
    change color format
    :param data: coloe data as dict or tuple
    :param to: 'dict' or 'tuple'
    """
    if not (type(data) is dict or type(data) is tuple):
        InvalidArgumentValue('data should be tuple or dict')

    if type(data) is tuple and to == 'dict':
        return {"red": data[0], "green": data[1], "blue": data[2], "alpha": data[3]}
    elif type(data) is dict and to == 'tuple':
            return data.get('red', 1), data.get('green', 1), data.get('blue', 1), data.get('alpha', 1)
    else:
        return data


def batchable(func):
    """ Function generator to make a model member function batachable.
        Model needs to have an attribute named _func_calls """
    @wraps(func)
    def wrapper(*args, **kwargs):
        obj = args[0]
        if obj.linked:
            return func(*args, **kwargs)
        else:
            obj._func_calls.append((func, (args, kwargs)))
            return False
    return wrapper


def allow_gridrange(method):
    """
    Decorator function casts gridrange in argument to start and end
    in range method calls.
    """
    @wraps(method)
    def wrapper(self, *args, **kwargs):
        if 'grange' in kwargs:
            grid_range = kwargs.pop('grange')
            kwargs['start'], kwargs['end'] = [x.index for x in grid_range.get_bounded_indexes()]
        return method(self, *args, **kwargs)
    return wrapper
