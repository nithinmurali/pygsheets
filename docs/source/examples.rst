
Examples
========

**Batching of api calls**

.. code:: python

    wks.unlink()
    for i in range(10):
        wks.update_value((1, i), i) # wont call api
    wks.link() # will do all the updates

**Protect an whole sheet**

.. code:: python

    r = Datarange(worksheet=wks)
    >>> r # this is a datarange unbounded on both indexes
    <Datarange Sheet1>
    >>> r.protected = True # this will make the whole sheet protected


**Formatting columns**
column A as percentage format, column B as currency format. Then formatting
row 1 as white , row 2 as grey colour. By @cwdjankoski

.. code:: python

    model_cell = pygsheets.Cell("A1")

    model_cell.set_number_format(
        format_type = pygsheets.FormatType.PERCENT,
        pattern = "0%"
    )
    # first apply the percentage formatting
    pygsheets.DataRange(
        left_corner_cell , right_corner_cell , worksheet = wks
     ).apply_format(model_cell)

    # now apply the row-colouring interchangeably
    gray_cell = pygsheets.Cell("A1")
    gray_cell.color = (0.9529412, 0.9529412, 0.9529412, 0)

    white_cell = pygsheets.Cell("A2")
    white_cell.color = (1, 1, 1, 0)

    cells = [gray_cell, white_cell]

    for r in range(start_row, end_row + 1):
        print(f"Doing row {r} ...", flush = True, end = "\r")
        wks.get_row(r, returnas = "range").apply_format(cells[ r % 2 ], fields = "userEnteredFormat.backgroundColor")

**Note**:
If you have any intresting examples that you think will be helpful to others, please raise a PR or issue.