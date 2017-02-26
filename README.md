# pygsheets - Google Spreadsheets Python API v4
[![Build Status](https://travis-ci.org/nithinmurali/pygsheets.svg?branch=master)](https://travis-ci.org/nithinmurali/pygsheets)  [![PyPI version](https://badge.fury.io/py/pygsheets.svg)](https://badge.fury.io/py/pygsheets)

A simple, intutive library for google sheets which gets most of your work done.
 
Features:

* Google spreadsheet api __v4__ support
* Open, create, delete and share spreadsheets using _title_ or _key_
* Control permissions of spreadsheets.
* Set cell format, text format, color, write notes
* __NamedRanges__ Support
* Work with range of cells easily with DataRange
* Do all the updates and push the changes in a batch

## Requirements

Python 2.6+ or 3+

## Installation

#### From PyPi

```sh
pip install pygsheets

```

#### From GitHub (Recommended)

```sh
pip install https://github.com/nithinmurali/pygsheets/archive/master.zip

```


## Basic Usage

Basic features are shown here, for complete set of features see the full documentation [here](http://pygsheets.readthedocs.io/en/latest/).

1. Obtain OAuth2 credentials from Google Developers Console for __google spreadsheet api__ and __drive api__ and save the file as `client_secret.json` in same directory as project. [read more here.](https://pygsheets.readthedocs.io/en/latest/authorizing.html)

2. Start using pygsheets: 
   
Sample scenario : you want to share a numpy array with your remote friend 

```python
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
wks.update_cells('A2:A5',[['name1'],['name2'],['name3'],['name4']])

# set the heights
heights = wks.range('B2:B5')  # get the range
heights.name = "heights"  # name the range
heights.update_values([[50],[60],[67],[66]]) # update the vales
wks.update_cell('B6','=average(heights)') # set get the avg value

```

## More Examples

### Opening a Spreadsheet

```python
# You can open a spreadsheet by its title as it appears in Google Docs 
sh = gc.open("pygsheetTest")

# If you want to be specific, use a key
sht1 = gc.open_by_key('1mwA-NmvjDqd3A65c8hsxOpqdfdggPR0fgfg5nXRKScZAuM')

# Or,paste the entire url
sht2 = gc.open_by_url('https://docs.google.com/spreadsheets/d/1mwA...AuM/edit')

```

### Operations on Spreadsheet

```python

# create a new sheet with 50 rows and 60 colums
wks = sh.add_worksheet("new sheet",rows=50,cols=60)

# or copy from another worksheet
wks = sh.add_worksheet("new sheet", src_worksheet=another_wks)

# delete this wroksheet
del_worksheet(wks)

# unshare the sheet
sh.remove_permissions("myNotSoFriend@gmail.com")

```

### Selecting a Worksheet

```python
# Select worksheet by id, index, title.
wks = sh.worksheet_by_title("my test sheet")

# By any property
wks = sh.worksheet('index', 0)

# Get a list of all worksheets
wks_list = sh.worksheets()

# Or just
wks = sh[0]
```

### Operations on Worksheet

```python
# Get values as 2d array('matrix') which can easily be converted to an numpy aray or as 'cell' list
values_mat = wks.values(start=(1,1), end=(20,20), returnas='matrix')

# Get all values of sheet as 2d list of cells
cell_matrix = wks.all_values('cell')

# update a range of values with a cell list or matrix
wks.update_cells(range='A1:E10', values=values_mat)

# Insert 2 rows after 20th row and fill with values
wks.insert_rows(row=20, number=2, values=values_list)

# resize by changing rows and colums
wks.rows=30

# use the worksheet as a csv
for row in wks:
    print(row)

# get values by indexes
 A1_value = wks[0][0]

# Search for a table in the worksheet and append a row to it
wks.append_row(values=[1,2,3,4])

# export a worksheet as csv
wks.export(pygsheets.ExportType.CSV)

# Find/Replace cells with string value
cell_list = worksheet.find("query string")

# Find/Replace cells with regexp
filter_re = re.compile(r'(small|big) house')
cell_list = worksheet.find(filter_re)

# Move a worksheet in the same spreadsheet (update index)
wks.index = 2 # index start at 1 , not 0

# Update title
wks.title = "NewTitle"
# working with named ranges
wks.create_named_range('A1', 'A10', 'prices')
wks.get_named_ranges()  # will return a list of DataRange objects
wks.delete_named_range('prices')

```

#### Pandas integration
If you work with pandas, you can directly use the dataframes
```python
#set the values of a pandas dataframe to sheet
wks.set_dataframe(df,(1,1))

#you can also get the values of sheet as dataframe
df = wks.get_as_df()

```


### Cell Object

Each cell has a __value__ and cordinates (__row__, __col__, __label__) properties.

Getting cell objects

```python
c1 = Cell('A1',"hello")  # create a unlinked cell
c1 = worksheet.cell('A1')  # creates a linked cell whose changes syncs instantanously
cl.value  # Getting cell value
c1.value_unformatted #Getting cell unformatted value
c1.formula # Getting cell formula if any
c1.note # any notes on the cell

cell_list = worksheet.range('A1:C7')  # get a range of cells 
cell_list = col(5, returnas='cell')  # return all cells in 5th column(E)

```

Most of the functions has `returnas` param, if whose value is `cell` it will return a list of cell objects. Also you can use *label* or *(row,col)* tuple interchangbly as a cell adress

### Cell Operations

Each cell is directly linked with its cell in spreadsheet, hence changing the value of cell object will update the corresponding cell in spreadsheet unless you explictly unlink it

Different ways of updating Cells
```python
# using linked cells
c1 = worksheet.cell('B1') # created from worksheet, so linked cell
c1.col = 5  # Now c1 correponds to E1
c1.value = "hoho"  # will change the value of E1

# Or onliner
worksheet.update_cell('B1', 'hehe')

# Or Update a range
cell_list = worksheet.range('A1:C7')
for cell in cell_list:
    cell.value = 'O_0'

# add formula
c1.formula = 'A1+C2'

# get neighbouring cells
c2 = c1.neighbour('topright') # you can also specify relative position as tuple eg (1,1)

# set cell format
c1.format = pygsheets.FormatType.NUMBER, '00.0000' # format is optional

# write notes on cell
c1.note = "yo mom"

# set cell color
c1.color = (1,1,1,1) # Red Green Blue Alpha

# set text format
c1.text_format['fontSize'] = 14
c1.text_format['bold'] = True

# sync the changes
 c1.update()

# you can unlink a cell and set all required properties and then link it
# So yu could create a model cell and update multiple sheets
c.unlink()
c.note = "offine note"
c.link(wks1, True)
c.link(wks2, True)

```

### DataRange Object

The DataRange is used to represent a range of cells in a worksheet. They can be named or protected.
Almost all `get_` functions has a `returnas` param, set it to `range` to get a range object.
```python
# Getting a Range object
rng = wks.get_values('A1', 'C5', returnas='range')
rng.unlink()  # linked ranges will sync the changes as they are changed

# Named ranges
rng.name = 'pricesRange'  # will make this range a named range
rng = wks.get_named_ranges('commodityCount') # directly get a named range
rng.name = ''  # will delete this named range

# Setting Format
 # first create a model cell with required properties
model_cell = Cell('A1')
model_cell.color = (1.0,0,1.0,1.0) # rose color cell
model_cell.format = pygsheets.FormatType.PERCENT

 # now set its format to all cells in the range
rng.applay_format(model_cell)  # will make all cell in this range rose color and percent format

# get cells in range
cell = rng[0][1]

```


## How to Contribute

This library is still in development phase. So there is a lot of work to be done. Checkout the [TO DO's](TODO.md).
 
* Follow the [Contributing to Open Source](https://guides.github.com/activities/contributing-to-open-source/) Guide.
* Please Create Pull Requests to the `staging` branch

### Report Issues/Features

* Please report bugs and suggest features via the [GitHub Issues](https://github.com/nithinmurali/pygsheets/issues).
* I have listed some possible features in the [TO DO's](TODO.md). If you would like to see any of that implimented or would like to work on any, lemme know (Just create an Issue).
* Before opening an issue, search the tracker for possible duplicates.
* If you have any usage questions, ask a question on stackoverflow with `pygsheets` Tag

## Disclaimer
The gspread library is used as an outline for developing pygsheets, much of the skeleton code is copied from there.
