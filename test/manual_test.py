"""
This file is for manual testing of pygsheets
"""
import sys
import IPython
from os import path
sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))

import pygsheets
gc = pygsheets.authorize(application_name='testapp1')

# wks = gc.open_by_key('18WX-VFi_yaZ6LkXWLH856sgAsH5CQHgzxjA5T2PGxIY')
wks =gc.open('pygsheetTest')
print wks

s1 = wks.worksheets()
print s1

# s1.update_acell('A1',"yoyo")
# print s1.col_values(2,"cell")
# s1.update_cell(5,5,"again yoypo")
# s1.insert_cols(3,1)

IPython.embed()
