# PyGsheets - Google Spreadsheets Python API v4
[![Downloads](https://img.shields.io/pypi/dm/pygsheets.svg)](https://pypi.python.org/pypi/pygsheets)

Do everything you can do from gui from python terminal
 
Features:

* Simple to use
* Google spreadsheet api __v4__ support
* Open, create, delete and share spreadsheets using _title_ or _key
* Control permissions of spreadsheets.
* Extract range, entire row or column values.
* Work offline and update the requests in batch

## Requirements

Python 2.6+

## Installation

#### From GitHub

```sh
pip install --upgrade google-api-python-client
pip install https://github.com/nithinmurali/pygsheets/archive/master.zip
```

#### From PyPI (TBD)


## Basic Usage

1. [Obtain OAuth2 credentials from Google Developers Console](https://console.developers.google.com/start/api?id=sheets.googleapis.com) for __google spreadsheet api__ and __drive api__ and save the file as client_secret.json in same directory as project

2. Start using pygsheets:

```python
import pygsheets

gc = pygsheets.authorize()

# Open a worksheet from spreadsheet with one shot
wks = gc.open('my new ssheet').sheet1

wks.update_cell('B2', "let me take another look.")

# Fetch a cell range
cell_list = wks.range('A1:B7')
```

## More Examples

### Opening a Spreadsheet

```python
# You can open a spreadsheet by its title as it appears in Google Docs 
sh = gc.open("My poor gym results") # <-- Look ma, no keys!

# If you want to be specific, use a key (which can be extracted from
# the spreadsheet's url)
sht1 = gc.open_by_key('0BmgG6nO_6dprdS1MN3d3MkdPa142WFRrdnRRUWl1UFE')

# Or, if you feel really lazy to extract that key, paste the entire url
sht2 = gc.open_by_url('https://docs.google.com/spreadsheet/ccc?key=0Bm...FE&hl')

# Now share the spreadsheet
sh.share("myFriend@gmail.com","writer")

```

### More Operations on Spreadsheet

```python

# create a new sheet with 50 rows and 60 colums
sh.add_worksheet("new sheet",50,60)

# unshare the sheet
sh.remove_permissions("myNotSoFriend@gmail.com")

```

### Selecting a Worksheet

```python
# Select worksheet by id, index, title. Worksheet indexes start from zero
wks = sh.worksheet_by_title("my test sheet")

# By any property
wks = sh.worksheet('index',0)

# Get a list of all worksheets
wks_list = sh.worksheets()
```

### Manupulating Worksheet

```python
# Get values as 2d array('matrix') which can easily be converted to an numpy aray or as 'cell' list
values_mat = wks.values((1,1),(20,20),'matrix')

cell_matrix = wks.all_values('cell')

# update a range of values with a cell list or matrix
wks.update_cells(range='A1:E10',values=values_mat)

#insert 2 rows after 20th row and fill with values
wks.insert_rows(20, 2, values)

#resize by changing rows and colums
wks.row_count=30

```

### Cell Object

Each cell has a __value__ and coordinates (__row__, __col__, __label__) properties.

Getting cell objects

```python
c1 = Cell('A1',"hello") # create a unlinked cell
c1 = worksheet.acell('A1') # creates a linked cell whose changes syncs instantanously
cell_list = worksheet.range('A1:C7') # get a range of cells 
cell_list = col(5,returnas='cell') #return all cells in 5th column(E)
```

Also most functions has `returnas` if whose value is 'cell' it will return a list of cell objects

### Getting a Cell Value

```python
# With label
val = worksheet.cell('B1').value

# With coords
val = worksheet.cell((1, 2)).value
```

### Updating Cells

Each cell is directly linked with its cell in spreadsheet. hence to changing the value of cell object will update the corresponding cell in spreadsheet unless you explictly unlink it

Different ways of updating Spreadsheet
```python

c1 = worksheet.cell('B1')
c1.value = 'hehe'

# using linked cells
c1.col = 5  # Now c1 correponds to E1
c1.value = "hoho"  # will change the value of E1

# Or onliner
worksheet.update_cell('B1', 'hehe')

# Or Update a range
cell_list = worksheet.range('A1:C7')

for cell in cell_list:
    cell.value = 'O_0'

```

## [Contributors](https://github.com/nithinmurali/pygsheets/graphs/contributors)

## How to Contribute

This library is Still in development phase. So there is a lot of work to be done. The `models.py` is the file which defines the models used in this library. There are mainly 3 models - `spreadsheet`, `worksheet`, `cell`. Fuctions which are yet to be implimented are left out empty with an _@TODO_ comment. you can start by implimenting them. The communication with google api using google-python-client is implimented in `client.py` and the exceptions in exceptions.py

### Report Issues

Please report bugs and suggest features via the [GitHub Issues](https://github.com/nithinmurali/pygsheets/issues).

Before opening an issue, search the tracker for possible duplicates. If you find a duplicate, please add a comment saying that you encountered the problem as well.

### Contribute code

* Check the [GitHub Issues](https://github.com/nithinmurali/pygsheets/issues) for open issues that need attention.
* Follow the [Contributing to Open Source](https://guides.github.com/activities/contributing-to-open-source/) Guide.


## Disclaimer
The gspread library is used as an outline for developing pygsheets, much of the skelton code is copied from there.
