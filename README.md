# pygsheets - Google Spreadsheets Python API v4
[![Build Status](https://travis-ci.org/nithinmurali/pygsheets.svg?branch=staging)](https://travis-ci.org/nithinmurali/pygsheets)  [![PyPI version](https://badge.fury.io/py/pygsheets.svg)](https://badge.fury.io/py/pygsheets)    [![Documentation Status](https://readthedocs.org/projects/pygsheets/badge/?version=latest)](http://pygsheets.readthedocs.io/en/latest/?badge=latest) [![Gitpod ready-to-code](https://img.shields.io/badge/Gitpod-ready--to--code-blue?logo=gitpod)](https://gitpod.io/#https://github.com/nithinmurali/pygsheets)

A simple, intuitive library for google sheets which gets your work done.
 
Features:

* Open, create, delete and share spreadsheets using _title_ or _key_
* Intuitive models - spreadsheet, worksheet, cell, datarange
* Control permissions of spreadsheets.
* Set cell format, text format, color, write notes
* Named and Protected Ranges Support
* Work with range of cells easily with DataRange and Gridrange
* Data validation support. checkboxes, drop-downs etc.
* Conditional formatting support
* get multiple ranges with get_values_batch and update wit update_values_batch

## Updates
* version [2.0.5](https://github.com/nithinmurali/pygsheets/releases/tag/2.0.5) released

## Installation

#### From PyPi (Stable)

```sh
pip install pygsheets

```

If you are installing from pypi please see the docs [here](https://pygsheets.readthedocs.io/en/stable/).


#### From GitHub (Recommended)

```sh
pip install https://github.com/nithinmurali/pygsheets/archive/staging.zip

```

If you are installing from github please see the docs [here](https://pygsheets.readthedocs.io/en/staging/).

## Basic Usage

Basic features are shown here, for complete set of features see the full documentation [here](http://pygsheets.readthedocs.io/en/staging/).

1. Obtain OAuth2 credentials from Google Developers Console for __google spreadsheet api__ and __drive api__ and save the file as `client_secret.json` in same directory as project. [read more here.](https://pygsheets.readthedocs.io/en/latest/authorization.html)

2. Start using pygsheets: 
   
Sample scenario : you want to share a numpy array with your remote friend 

```python
import pygsheets
import numpy as np

gc = pygsheets.authorize()

# Open spreadsheet and then worksheet
sh = gc.open('my new sheet')
wks = sh.sheet1

# Update a cell with value (just to let him know values is updated ;) )
wks.update_value('A1', "Hey yank this numpy array")
my_nparray = np.random.randint(10, size=(3, 4))

# update the sheet with array
wks.update_values('A2', my_nparray.tolist())

# share the sheet with your friend
sh.share("myFriend@gmail.com")

```

Sample Scenario: you want to fill height values of students
```python

## import pygsheets and open the sheet as given above

header = wks.cell('A1')
header.value = 'Names'
header.text_format['bold'] = True # make the header bold
header.update()

# or achive the same in oneliner
wks.cell('B1').set_text_format('bold', True).value = 'heights'

# set the names
wks.update_values('A2:A5',[['name1'],['name2'],['name3'],['name4']])

# set the heights
heights = wks.range('B2:B5', returnas='range')  # get the range as DataRange object
heights.name = "heights"  # name the range
heights.update_values([[50],[60],[67],[66]]) # update the vales
wks.update_value('B6','=average(heights)') # set the avg value of heights using named range

```

## More Examples

### Opening a Spreadsheet

```python
# You can open a spreadsheet by its title as it appears in Google Docs 
sh = gc.open("pygsheetTest")

# If you want to be specific, use a key
sht1 = gc.open_by_key('1mwA-NmvjDqd3A65c8hsxOpqdfdggPR0fgfg5nXRKScZAuM')

# create a spreadsheet in a folder (by id)
sht2 = gc.create("new sheet", folder_name="my worksheets")

# open enable TeamDrive support
gc.drive.enable_team_drive("Dqd3A65c8hsxOpqdfdggPR0fgfg")

```

### Operations on Spreadsheet [doc](http://pygsheets.readthedocs.io/en/latest/spreadsheet.html)

<details> <summary>show code</summary>

```python

import pygsheets
c = pygsheets.authorize()
sh = c.open('spreadsheet')

# create a new sheet with 50 rows and 60 colums
wks = sh.add_worksheet("new sheet",rows=50,cols=60)

# create a new sheet with 50 rows and 60 colums at the begin of worksheets
wks = sh.add_worksheet("new sheet",rows=50,cols=60,index=0)

# or copy from another worksheet
wks = sh.add_worksheet("new sheet", src_worksheet='<other worksheet instance>')

# delete this wroksheet
sh.del_worksheet(wks)

# unshare the sheet
sh.remove_permissions("myNotSoFriend@gmail.com")

```

</details>

### Selecting a Worksheet

<details> <summary>show code</summary>

```python
import pygsheets
c = pygsheets.authorize()
sh = c.open('spreadsheet')

# Select worksheet by id, index, title.
wks = sh.worksheet_by_title("my test sheet")

# By any property
wks = sh.worksheet('index', 0)

# Get a list of all worksheets
wks_list = sh.worksheets()

# Or just
wks = sh[0]
```

</details>

### Operations on Worksheet [doc](http://pygsheets.readthedocs.io/en/latest/worksheet.html)

<details> <summary>show code</summary>

```python
# Get values as 2d array('matrix') which can easily be converted to an numpy aray or as 'cell' list
values_mat = wks.get_values(start=(1,1), end=(20,20), returnas='matrix')

# Get values of - rows A1 to B10, column C, 1st row, 10th row
wks.get_values_batch(['A1:B10', 'C', '1', (10, None)])

# Get all values of sheet as 2d list of cells
cell_matrix = wks.get_all_values(returnas='matrix')

# update a range of values with a cell list or matrix
wks.update_values(crange='A1:E10', values=values_mat)

# update multiple ranges with bath update
wks.update_values_batch(['A1:A2', 'B1:B2'], [[[1],[2]], [[3],[4]]])

# Insert 2 rows after 20th row and fill with values
wks.insert_rows(row=20, number=2, values=values_list)

# resize by changing rows and colums
wks.rows=30

# use the worksheet as a csv
for row in wks:
    print(row)

# get values by indexes
 A1_value = wks[0][0]

# clear all values
wks.clear()

# Search for a table in the worksheet and append a row to it
wks.append_table(values=[1,2,3,4])

# export a worksheet as csv
wks.export(pygsheets.ExportType.CSV)

# Find/Replace cells with string value
cell_list = worksheet.find("query string")

# Find/Replace cells with regexp
filter_re = re.compile(r'(small|big) house')
cell_list = worksheet.find(filter_re, searchByRegex=True)
cell_list = worksheet.replace(filter_re, 'some house', searchByRegex=True)

# Move a worksheet in the same spreadsheet (update index)
wks.index = 2 # index start at 1 , not 0

# Update title
wks.title = "NewTitle"

# Update hidden state
wks.hidden = False

# working with named ranges
wks.create_named_range('A1', 'A10', 'prices')
wks.get_named_range('prices')
wks.get_named_ranges()  # will return a list of DataRange objects
wks.delete_named_range('prices')

# Plot a chart/graph
wks.add_chart(('A1', 'A6'), [('B1', 'B6')], 'Health Trend')

# create drop-downs
wks.set_data_validation(start='C4', end='E7', condition_type='NUMBER_BETWEEN', condition_values=[2,10], strict=True, inputMessage="inut between 2 and 10")


```

</details>

#### Pandas integration
If you work with pandas, you can directly use the dataframes
```python
#set the values of a pandas dataframe to sheet
wks.set_dataframe(df,(1,1))

#you can also get the values of sheet as dataframe
df = wks.get_as_df()

```


### Cell Object [doc](http://pygsheets.readthedocs.io/en/latest/cell.html)

Each cell has a __value__ and cordinates (__row__, __col__, __label__) properties.

Getting cell objects

<details open> <summary>show code</summary>

```python
c1 = Cell('A1',"hello")  # create a unlinked cell
c1 = worksheet.cell('A1')  # creates a linked cell whose changes syncs instantanously
cl.value  # Getting cell value
c1.value_unformatted #Getting cell unformatted value
c1.formula # Getting cell formula if any
c1.note # any notes on the cell
c1.address # address object with cell position

cell_list = worksheet.range('A1:C7')  # get a range of cells 
cell_list = worksheet.col(5, returnas='cell')  # return all cells in 5th column(E)

```

</details>

Most of the functions has `returnas` param, if whose value is `cell` it will return a list of cell objects. Also you can use *label* or *(row,col)* tuple interchangbly as a cell adress

### Cell Operations

Each cell is directly linked with its cell in spreadsheet, hence changing the value of cell object will update the corresponding cell in spreadsheet unless you explictly unlink it
Also not that bu default only the value of cell is fetched, so if you are directly accessing any cell properties call `cell.fetch()` beforehand. 

Different ways of updating Cells

<details> <summary>show code</summary>

```python
# using linked cells
c1 = worksheet.cell('B1') # created from worksheet, so linked cell
c1.col = 5  # Now c1 correponds to E1
c1.value = "hoho"  # will change the value of E1

# Or onliner
worksheet.update_value('B1', 'hehe')

# get a range of cells
cell_list = worksheet.range('A1:C7')
cell_list = worksheet.get_values(start='A1', end='C7', returnas='cells')
cell_list = worksheet.get_row(2, returnas='cells')


# add formula
c1.formula = 'A1+C2'
c1.formula # '=A1+C2'

# get neighbouring cells
c2 = c1.neighbour('topright') # you can also specify relative position as tuple eg (1,1)

# set cell format
c1.set_number_format(pygsheets.FormatType.NUMBER, '00.0000')

# write notes on cell
c1.note = "yo mom"

# set cell color
c1.color = (1.0, 1.0, 1.0, 1.0) # Red, Green, Blue, Alpha

# set text format
c1.text_format['fontSize'] = 14
c1.set_text_format('bold', True)

# sync the changes
 c1.update()

# you can unlink a cell and set all required properties and then link it
# So yu could create a model cell and update multiple sheets
c.unlink()
c.note = "offine note"
c.link(wks1, True)
c.link(wks2, True)

```

</details>

### DataRange Object [doc](http://pygsheets.readthedocs.io/en/latest/datarange.html)

The DataRange is used to represent a range of cells in a worksheet. They can be named or protected.
Almost all `get_` functions has a `returnas` param, set it to `range` to get a range object.

<details open> <summary>show code</summary>

```python
# Getting a Range object
rng = wks.get_values('A1', 'C5', returnas='range')
rng.start_addr = 'A' # make the range unbounded on rows <Datarange Sheet1!A:B>
drange.end_addr = None # make the range unbounded on both axes <Datarange Sheet1>

# Named ranges
rng.name = 'pricesRange'  # will make this range a named range
rng = wks.get_named_ranges('commodityCount') # directly get a named range
rng.name = ''  # will delete this named range

#Protected ranges
rng.protected = True
rng.editors = ('users', 'someemail@gmail.com')

# Setting Format
 # first create a model cell with required properties
model_cell = Cell('A1')
model_cell.color = (1.0,0,1.0,1.0) # rose color cell
model_cell.format = (pygsheets.FormatType.PERCENT, '')

 # Setting format to multiple cells in one go
rng.apply_format(model_cell)  # will make all cell in this range rose color and percent format
# Or if you just want to apply format, you can skip fetching data while creating datarange
Datarange('A1','A10', worksheet=wks).apply_format(model_cell)

# get cells in range
cell = rng[0][1]

```
</details>

### Batching calls

If you are calling a lot of spreadsheet modification functions (non value update). you can merge them into a single call.
By doing so all the requests will be merged into a single call.

```python
gc.set_batch_mode(True)
wks.merge_cells("A1", "A2")
wks.merge_cells("B1", "B2")
Datarange("D1", "D5", wks).apply_format(cell)
gc.run_batch() # All the above requests are executed here
gc.set_batch_mode(False)

```
Batching also happens when you unlink worksheet. But in that case the requests are not merged.


## How to Contribute

This library is still in development phase.
 
* Follow the [Contributing to Open Source](https://guides.github.com/activities/contributing-to-open-source/) Guide.
* Branch off of the `staging` branch, and submit Pull Requests back to
  that branch.  Note that the `master` branch is used for version
  bumps and hotfixes only.
* For quick testing the changes you have made to source, run the file `tests/manual_testing.py`. It will give you an IPython shell with lastest code loaded.

### Report Issues/Features

* Please report bugs and suggest features via the [GitHub Issues](https://github.com/nithinmurali/pygsheets/issues).
* Before opening an issue, search the tracker for possible duplicates.
* If you have any usage questions, ask a question on stackoverflow with `pygsheets` Tag

## Run Tests
* install `pip install -r requirements-dev.txt`
* run `make test`

Now that you have scrolled all the way down, **finding this library useful?**  <a href="https://www.buymeacoffee.com/pygsheets" target="_blank"><img src="https://www.buymeacoffee.com/assets/img/custom_images/orange_img.png" alt="Buy Me A Coffee" style="height: auto !important;width: auto !important;" ></a>
