
Some Tips
=========

Note that in this article, wks means worksheet, ss means spreadsheet, gc means google client

easier way to acess sheet values::

    for row in wks:
        print row[0]

Acess sheets by id::


    wks1 = ss[0]


Conversion of sheet data

usually all the values are converted to string while using `get_*` functions. But if you want then to retain
their type, the change the `value_render` option to ValueRenderOption.UNFORMATTED_VALUE.

