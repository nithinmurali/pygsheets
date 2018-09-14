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
First example - Share a numpy array with a friend::

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

Second example - Store some data and change cell formatting::

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

   pip install pygsheets

Overview
--------
The :ref:`client` is used to create and access spreadsheets. The property `drive` exposes some Google Drive API
functionality and the `sheet` property exposes used Google Sheets API functionality.

A Google Spreadsheet is represented by the :ref:`spreadsheet` class. Each spreadsheet contains one or more :ref:`worksheet`.
The data inside of a worksheet can be accessed as plain values or inside of a :ref:`cell` object. The cell has properties
and attributes to change formatting, formulas and more. To work with several cells at once a :ref:`datarange` can be
used.

Changelog
---------
This version is not backwards compatible with 1.x
There is major rework in the library with this release.
Some functions are renamed to have better consistency in naming and clear meaning.

- update_cell() renamed to update_value()
- update_cells() renamed to update_values()
- update_cells_prop() renamed to update_cells()
- teamDriveId, enableTeamDriveSupport changed to client.drive.enable_team_drive, include_team_drive_items
- parameter changes for all get_* functions : include_empty, include_all changed to include_tailing_empty, include_tailing_empty_rows
- in created_protected_range(), gridrange param changed to start, end
- removed batch mode
- find() splited into find() and replace()
- removed (show/hide)_(row/column), use (show/hide)_dimensions instead
- removed link/unlink from spreadsheet

**New Features added**
- chart Support added
- sort feature added
- better support for protected ranges
- multi header/index support in dataframes
- removes the dependency on oauth2client and uses google-auth and google-auth-oauth.

Other bug fixes and performance improvements

Authors and License
-------------------

The ``pygsheets`` package is written by Nithin M and is inspired by gspread.

Licensed under the MIT-License.

Feel free to improve this package and send a pull request to GitHub_.

.. _GitHub: https://github.com/nithinmurali/pygsheets/issues


Contents:

.. toctree::
   :maxdepth: 3

   authorization
   reference
   tips



Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`

