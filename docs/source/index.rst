.. pygsheets documentation master file, created by
   sphinx-quickstart on Thu Dec 16 14:44:32 2011.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

pygsheets
=========

A simple, intutive library for google spreadsheets based on api v4 which gets most of your work done.

Features
--------

- Google spreadsheet api v4 support
- Open, create, delete and share spreadsheets using _title_
- Control permissions of spreadsheets.
- Extract range, entire row or column values.
- Do all the updates and push the changes in a batch


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


Installation
------------
::

   pip install https://github.com/nithinmurali/pygsheets/archive/master.zip (recent)
   pip install pygsheets (stable)


Overview
--------

There are mainly 3 models - ``spreadsheet``, ``worksheet``, ``cell``, they are defined in their respectivr files.
The communication with google api is implimented in ``client.py``. The client.py also impliments the autorization functions.

Authors and License
-------------------

The ``pygsheets`` package is written by Nithin Murali and is inspried by gspread.  It's MIT licensed and freely available.

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

