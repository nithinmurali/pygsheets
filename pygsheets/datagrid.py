# -*- coding: utf-8 -*-.

"""
pygsheets.datagrid
~~~~~~~~~~~~~~~~~~

This module contains DataGrid class for strong a range of data in spreadsheet.

"""


class Datagrid:

    def __init__(self):
        self.data = [[]]
        self.id = None
        self.name = ''
        self.worksheet = None
        self.startAddr = None
        self.endAddr = None
