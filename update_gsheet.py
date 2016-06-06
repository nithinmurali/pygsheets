#!/usr/bin/python
# @author nithin
# @email nithin.m@olacabs.com

import pygsheets
import sys, argparse, csv, os
from collections import defaultdict
import json
import webbrowser
from time import sleep
import requests.exceptions as re

import httplib2
from oauth2client import client
from oauth2client.file import Storage
from oauth2client.service_account import ServiceAccountCredentials

# command arguments
parser = argparse.ArgumentParser(description='This python script will update a google spread sheet form an mapped csv file ')
#parser.add_argument('file', help='csv file to import', action='store')
parser.add_argument('gs_ws', help='google sheet id to update', action='store')
parser.add_argument('-si','--gs_ss', help='google worksheet index to update', action='store', type=int)
parser.add_argument('-sn','--sheet_name', help='google worksheet name to update', action='store')
parser.add_argument('-ch','--col_header', help='new colum header', action='store')
parser.add_argument('-rh','--replace', help='replace the header if exists', action='store_true')
args = parser.parse_args()
#csv_file = args.file;
gs_link = args.gs_ws;

#print(args.gs_ss)

__location__ = os.path.realpath(os.path.join(os.getcwd(), os.path.dirname(__file__)))
weekly_data = defaultdict(list) # each value in each column is appended to a list
with open(os.path.join(__location__,'data_final.csv')) as f:
    reader = csv.DictReader(f) # read rows into a dictionary format
    for row in reader: # read a row as {column1: value1, column2: value2,...}
        for (k,v) in row.items(): # go over each column name and value 
            weekly_data[k].append(v) # append the value into the appropriate list

gc = pygsheets.authorize()
wks = ''
if args.sheet_name:
    wks = gc.open_by_key(gs_link).worksheet('title',args.sheet_name)
else:
    wks = gc.open_by_key(gs_link).get_worksheet('id',args.gs_ss)

print "updating the sheet, ", wks.title

max_rows = 125;
max_cols = 4; #cols with matrices
week_start_index = 5
colsDicts = []

def createDict(ilist):
    #print(ilist)
    nkey = "none"
    idict={"none":[1]}
    for x in xrange(0,len(ilist)):
        if ilist[x] != "":
            if len(idict[nkey]) > 1 and ilist[x-1] == "":
                #idict[nkey]=idict[nkey][:-1]
                pass
            nkey = ilist[x].strip().lower()
            if not idict.has_key(nkey):
                idict[nkey] = [x+1];
            else:
                idict[nkey].append(x+1);
        else:
            idict[nkey].append(x+1)
    return idict

def intersect(a, b):
    return list( set(a) & set(b))

def getCol(header):
    headers = header.lower().split('_')
    try:
        i=0
        final_list = colsDicts[0][headers[0]]
        for i in xrange(1,len(headers)):
            final_list = intersect(colsDicts[i][headers[i]],final_list)
    except KeyError:
        print "WARNING : CANT FIND KEY ", headers[i]
        return []
    if len(final_list) == 0:
        print "no header found ",
        return final_list
    else:
        return final_list

def updateCol(header,value):
    index = getCol(header)
    
    if len(index) == 1:
        update_values[sorted(index)[0]-1] = value
    elif len(index) > 1:
        print "more than 1 matching column"
        update_values[sorted(index)[0]-1] = value
    else:
        print "no matching colum found"

# create the map dict
for x in range(1,max_cols+1):
    co = wks.col_values(x)[:max_rows]
    colsDicts.append(createDict(co))

#search for approriate column
ncell = ''
first_row = wks.row_values(1,'cell')
for i in xrange(week_start_index,len(first_row)):
    if ((args.replace and ( args.col_header == first_row[i].value ))):
        ncell = first_row[i]
        #print "replaceing ", ncell
        break

# populate the data
# if the key has '%' and not '%ile' then it will *100 else
update_values = ['']*max_rows
update_values[0] = args.col_header

for key,value in weekly_data.iteritems():
    if key.find('%') != -1 and not key.find('%ile') != -1:
        if(value[0]=='NA'):value[0] = 0  #it is not used value (as of now)
        rvalue = str(float(value[0])*100.0) + '%'
    else:
        if(value[0]=='NA'):value[0] = 0  #it is not used value (as of now)
        rvalue = value[0]
    #print "updating " + str(key) + " " + str(rvalue)
    #updateCol(key,rvalue)     
    for i in xrange(1,5):
        try:
            updateCol(key,rvalue)
            break
        except Exception as e:
            print e.__doc__ , ' ... Retrying'
            sleep(5)
            continue

#remove tailing empty values
while update_values[-1] == '':
    update_values.pop()
print update_values

#update the sheet
if ncell != "" :
    wks.update_col(ncell.col,update_values)
else:
    wks.insert_cols(week_start_index-1,1,update_values)
    
