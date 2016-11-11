.. pygsheets documentation master file, created by
   sphinx-quickstart on Thu Dec 16 14:44:32 2011.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

pygsheets
=========

A simple, intutive library for google spreadsheets based on api v4 which gets most of your work done.

Features
--------

- Simple to use
- Google spreadsheet api __v4__ support
- Open, create, delete and share spreadsheets using _title_ or _key_
- Control permissions of spreadsheets.
- Extract range, entire row or column values.
- Do all the updates and push the changes in a batch


Simple Exmaple
--------------
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
   wks.update_cells('A2:Z100', my_nparray.to_list())

   # share the sheet with your friend
   sh.share("myFriend@gmail.com")


Installation
------------
::

   pip install https://github.com/nithinmurali/pygsheets/archive/master.zip


Authors and License
-------------------

The ``pygsheets`` package is written by Nithin Murali and is based on gspread.  It's MIT
licensed and freely available.

Feel free to improve this package and send a pull request to GitHub_.


.. _GitHub: https://github.com/burnash/pygsheets/issues


Contents:

.. toctree::
   :maxdepth: 2

   authorizing
   reference


Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`

