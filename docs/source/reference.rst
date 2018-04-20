.. pygsheets documentation master file, created by
   sphinx-quickstart on Thu Dec 16 14:44:32 2011.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

pygsheets Reference
===================

`pygsheets <https://github.com/nithinmurali/pygsheets>` is a simple `Google Sheets API v4` Wrapper.

.. _Google Spreadsheets v4 : http://www.google.com/drive/apps.html

.. module:: pygsheets


Top Level Interface
-------------------

.. autofunction:: authorize

.. autoclass:: Client
   :members: spreadsheet_ids, spreadsheet_titles, create, open, open_by_key, open_by_url, open_all, open_as_json

Models
------

The models represent common spreadsheet objects: :class:`spreadsheet <Spreadsheet>`,
:class:`worksheet <Worksheet>` and :class:`cell <Cell>`.


.. toctree::

   spreadsheet
   worksheet
   datarange
   cell

API
---------

The Drive API is wrapped by :class:`DriveAPIWrapper <DriveAPIWrapper>`. Only implements
functionality used by this package.

.. toctree::

   driveapiwrapper


Exceptions
----------

.. autoexception:: AuthenticationError
.. autoexception:: SpreadsheetNotFound
.. autoexception:: WorksheetNotFound
.. autoexception:: NoValidUrlKeyFound
.. autoexception:: IncorrectCellLabel
.. autoexception:: RequestError
.. autoexception:: InvalidUser
.. autoexception:: InvalidArgumentValue


.. _github issue: https://github.com/burnash/pygsheets/issues

Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`

