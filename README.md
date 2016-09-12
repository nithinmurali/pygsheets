# Google Spreadsheets Python API v4
[![Downloads](https://img.shields.io/pypi/dm/pygsheets.svg)](https://pypi.python.org/pypi/pygsheets)

Manage your spreadsheets with _pygsheets_ in Python.

Features:

* Simple to use
* Open spreadsheets using _title_ or _key_
* Extract range, entire row or column values.
* Google spreadsheet api __v4__ support

## Requirements

Python 2.6+

## Installation

#### From GitHub

```sh
pip install --upgrade google-api-python-client
git clone https://github.com/nithinmurali/pygsheets.git
cd pygsheets
python setup.py install
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

wks.update_acell('B2', "it's down there somewhere, let me take another look.")

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
```

### Selecting a Worksheet

```python
# Select worksheet by index. Worksheet indexes start from zero
worksheet = sh.get_worksheet('index',0)

# By title
worksheet = sh.worksheet('title',"January")

# Most common case: Sheet1
worksheet = sh.sheet1

# Get a list of all worksheets
worksheet_list = sh.worksheets()
```

### Getting a Cell Value

```python
# With label
val = worksheet.acell('B1').value

# With coords
val = worksheet.cell(1, 2).value
```

### Getting All Values From a Row or a Column

```python
# Get all values from the first row
values_list = worksheet.row_values(1)

# Get all values from the first column
values_list = worksheet.col_values(1)
```

### Cell Object

Each cell has a __value__ and coordinates (__row__, __col__, __label__) properties.

Getting cell objects

```python
c1 = Cell('A1',"hello") # create a unlinked cell
c1 = worksheet.acell('A1') # creates a linked cell
cell_list = worksheet.range('A1:C7')
cell_list = col_values(5,returnas='cell') #return all cells in 5th column(E)
```

Also most functions has `returnas` if whose value is 'cell' it will return a list of cell objects

### Updating Cells

Each cell is directly linked with its cell in spreadsheet. hence to changing the value of cell object will update the corresponding cell in spreadsheet

Different ways of updating Spreadsheet
```python

c1 = worksheet.acell('B1')
c1.value = 'hehe'

# using linked cells
c1.col=5 #Now c1 correponds to E1
c1.value = "hoho" # will change the value of E1

# Or onliner
worksheet.update_acell('B1', 'hehe')

# Or Update a range
cell_list = worksheet.range('A1:C7')

for cell in cell_list:
    cell.value = 'O_0'

```

## [Contributors](https://github.com/nithinmurali/pygsheets/graphs/contributors)

## How to Contribute

This library is Still in development phase. I have only implimented the basic features that i required. So there is a lot of work to be done. The models.py is the file which defines the models used in this library. There are mainly 3 models - spreadsheet, worksheet, cell. Fuctions which are yet to be implimented are left out empty with an @TODO comment. you can start by implimenting them. The communication with google api using google-python-client is implimented in client.py and the exceptions in exceptions.py

### Report Issues

Please report bugs and suggest features via the [GitHub Issues](https://github.com/nithinmurali/pygsheets/issues).

Before opening an issue, search the tracker for possible duplicates. If you find a duplicate, please add a comment saying that you encountered the problem as well.

### Contribute code

* Check the [GitHub Issues](https://github.com/nithinmurali/pygsheets/issues) for open issues that need attention.
* Follow the [Contributing to Open Source](https://guides.github.com/activities/contributing-to-open-source/) Guide.


## Disclaimer
The gspread library is used as an outline for developing pygsheets, much of the skelton code is copied from there.
