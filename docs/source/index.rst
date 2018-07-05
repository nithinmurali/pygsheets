.. pygsheets documentation master file, created by
   sphinx-quickstart on Thu Dec 16 14:44:32 2011.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

pygsheets
=========

A simple, intuitive library to access google spreadsheets through the `Google Sheets API v4`_.

.. _Google Sheets API v4: https://developers.google.com/sheets/api/

Features
--------

- Google Sheets API v4 support.
- Limited Google Drive API v3 support.
- Open and create spreadsheets by **title**.
- Add or remove permissions from you spreadsheets.
- Simple calls to get a row, column or defined range of values.
- Change the formatting properties of a cell.
- Supports named ranges & protected ranges.
- Queue up requests in batch mode and then process them in one go.

Small Example
-------------
First example : Share a numpy array with a friend.
::

   import pygsheets

   client = pygsheets.authorize()

   # Open the spreadsheet and the first sheet.
   sh = client.open('spreadsheet-title')
   wks = sh.sheet1

   # Update a single cell.
   wks.update_cell('A1', "Numbers on Stuff")

   # Update the worksheet with the numpy array values. Beginning at cell 'A2'.
   wks.update_cells('A2', my_numpy_array.to_list())

   # Share the sheet with your friend. (read access only)
   sh.share('friend@gmail.com')
   # sharing with write access
   sh.share('friend@gmail.com', role='writer')

Example two: Store some data and change cell formatting:
::
   # open a worksheet as in the first example.

   header = wks.cell('A1')
   header.value = 'Names'
   header.text_format['bold'] = True # make the header bold
   header.update()

   # The same can be achieved in one line
   wks.cell('B1').set_text_format('bold', True).value = 'heights'

   # set the names
   wks.update_cells('A2:A5',[['name1'],['name2'],['name3'],['name4']])

   # set the heights
   heights = wks.range('B2:B5')  # get the range
   heights.name = "heights"  # name the range
   heights.update_values([[50],[60],[67],[66]]) # update the vales
   wks.update_cell('B6','=average(heights)') # set get the avg value



Installation
------------
Install recent::

   pip install https://github.com/nithinmurali/pygsheets/archive/master.zip

Install stable::

   pip install pygsheets (stable)

Overview
--------
The :ref:`client` is used to access or create spreadsheets. It has the two attributes `drive` and `service` which
are the base googleapiclient service objects. Any part of the API which is not implemented can be accessed through these
attributes.

A Google Spreadsheet is represented by the :ref:`spreadsheet` class. Each spreadsheet contains one or more :ref:`worksheet`.
The data inside of a worksheet can be accessed as plain values or inside of a :ref:`cell` object. The cell has properties
and attributes to change formatting, formulas and more. To work with several cells at once a :ref:`datarange` can be
used.

Authors and License
-------------------

The ``pygsheets`` package is written by Nithin Murali and is inspired by gspread.
The package has a MIT license.

Feel free to improve this package and send a pull request to GitHub_.

.. _GitHub: https://github.com/nithinmurali/pygsheets/issues


Contents:

.. toctree::
   :maxdepth: 3

   authorizing
   reference



Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`

