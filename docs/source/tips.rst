
Some Tips
=========

Note that in this article, wks means worksheet, ss means spreadsheet, gc means google client

easier way to acess sheet values::

    for row in wks:
        print row[0]

Acess sheets by id::


    wks1 = ss[0]


Create a protected range::

    wks.create_protected_range(wks.get_gridrange('A1', 'C2'))

