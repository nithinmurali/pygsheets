.. pygsheets documentation master file, created by
   sphinx-quickstart on Thu Dec 16 14:44:32 2011.
   You can adapt this file completely to your liking, but it should at least
   contain the root `toctree` directive.

pygsheets Reference
===================

`pygsheets <https://github.com/nithinmurali/pygsheets>`_ is a simple `Google Spreadsheets v4 `_ API wrapper.

.. _Google Spreadsheets v4 : http://www.google.com/drive/apps.html

.. module:: pygsheets


Top Level Interface
-------------------

.. autofunction:: authorize

.. autoclass:: Client
   :members: create, delete, open, open_by_key, open_by_url, open_all, list_ssheets

Models
------

The models represent common spreadsheet objects: :class:`spreadsheet <Spreadsheet>`,
:class:`worksheet <Worksheet>` and :class:`cell <Cell>`.


.. toctree::

   spreadsheet
   worksheet
   datarange
   cell


Exceptions
----------

.. autoexception:: AuthenticationError
.. autoexception:: SpreadsheetNotFound
.. autoexception:: WorksheetNotFound
.. autoexception:: NoValidUrlKeyFound
.. autoexception:: IncorrectCellLabel
.. autoexception:: RequestTimeout
.. autoexception:: InvalidUser
.. autoexception:: InvalidArgumentValue


.. _github issue: https://github.com/burnash/pygsheets/issues

Indices and tables
==================

* :ref:`genindex`
* :ref:`modindex`
* :ref:`search`

