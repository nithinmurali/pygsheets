"""
This file is for manual testing of pygsheets
"""
import sys
import IPython
from os import path
sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))

import pygsheets

from oauth2client.service_account import ServiceAccountCredentials

SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.metadata.readonly']
CREDS_FILENAME = path.join(path.dirname(__file__), 'data/creds.json')

credentials = ServiceAccountCredentials.from_json_keyfile_name('data/service_creds.json', SCOPES)

# gc = pygsheets.authorize(service_file='./data/service_creds.json')

gc = pygsheets.authorize(outh_file=CREDS_FILENAME)

# wks = gc.open_by_key('18WX-VFi_yaZ6LkXWLH856sgAsH5CQHgzxjA5T2PGxIY')
ss =gc.open('pygsheetTest')
print ss

# try:
#     ss.del_worksheet(ss.worksheet_by_title('testtt'))
# except:
#     pass
# ss.add_worksheet('testtt',50,50)

wks = ss.sheet1
print wks

# s1.update_acell('A1',"yoyo")
# print s1.col_values(2,"cell")
# s1.update_cell(5,5,"again yoypo")
# s1.insert_cols(3,1)
pass
IPython.embed()
