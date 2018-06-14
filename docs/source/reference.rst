.. pygsheets documentation master file, created by
   sphinx-quickstart on Thu Dec 16 14:44:32 2011.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

pygsheets Reference
===================

.. _pygsheets: https://github.com/nithinmurali/pygsheets

.. _Google Sheets API v4: https://developers.google.com/sheets/api/reference/rest/

.. _Google Drive API v3: https://developers.google.com/drive/v3/reference/

`pygsheets`_ is a simple `Google Sheets API v4`_ Wrapper. Some functionality uses the `Google Drive API v3`_ as well.

.. module:: pygsheets

Top Level Interface
-------------------

.. autofunction:: authorize

.. _client:

Client
------

.. autoclass:: Client
   :members: spreadsheet_ids, spreadsheet_titles, create, open, open_by_key, open_by_url, open_all, open_as_json

Models
------

The models represent common spreadsheet objects: :class:`spreadsheet <Spreadsheet>`,
:class:`worksheet <Worksheet>`,  :class:`cell <Cell>` and :class:`datarange <DataRange>`.


.. toctree::

   spreadsheet
   worksheet
   datarange
   cell

API
---

The Drive API is wrapped by :class:`DriveAPIWrapper <DriveAPIWrapper>`. Only implements
functionality used by this package.

.. toctree::

   drive_api
   sheet_api


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

