.. pygsheets documentation master file, created by
   sphinx-quickstart on Thu Dec 16 14:44:32 2011.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

pygsheets
=========

A simple, intuitive library to access google spreadsheets through the Google Sheets API v4.

Features
--------

- Google Sheets API v4 support.
- Limited Google Drive API v3 support.
- Open and create spreadsheets by __title__.
- Add or remove permissions from you spreadsheets.
- Simple calls to get a row, column or defined range of values.
- Change the formatting properties of a cell.
- Supports named ranges & protected ranges.
- Queue up requests in batch mode and then process them in one go.

Small Example
-------------
Sample scenario : you want to share a numpy array with your remote friend
::

   import pygsheets

   gc = pygsheets.authorize()

   # Open spreadsheet and then workseet
   sh = gc.open('my new ssheet')
   wks = sh.sheet1

   # Update a cell with value (just to let him know values is updated ;) )
   wks.update_cell('A1', "Hey yank this numpy array")

   # update the sheet with array
   wks.update_cells('A2', my_nparray.to_list())

   # share the sheet with your friend
   sh.share("myFriend@gmail.com")


Sample scenario: you want to store students name and their heights
::

  ## import and open the sheet as given above

  header = wks.cell('A1')
  header.value = 'Names'
  header.text_format['bold'] = True # make the header bold
  header.update()

  # or achive the same in oneliner
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
::

   pip install https://github.com/nithinmurali/pygsheets/archive/master.zip (recent)
   pip install pygsheets (stable)


Overview
--------
The entry into this package is through pygsheets.authorize() which will return a ``Client``.
With the client a spreadsheet can be opened or created.

A Google Spreadsheet is represented by the ``spreadsheet`` class. Each spreadsheet contains one or more ``worksheets``.
The data inside of a worksheet can be accessed as plain values or inside of a ``cell`` object. The cell has properties
and attributes to change formatting, formulas and more.

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

