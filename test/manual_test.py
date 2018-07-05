"""
This file is for manual testing of pygsheets
"""
if __name__ == '__main__':

    import sys
    import IPython
    from os import path
    sys.path.append(path.dirname(path.dirname(path.abspath(__file__))))

    import pygsheets
    import logging

    from oauth2client.service_account import ServiceAccountCredentials

    # SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive.metadata.readonly']
    # CREDS_FILENAME = path.join(path.dirname(__file__), 'data/creds.json')

    # credentials = ServiceAccountCredentials.from_json_keyfile_name('data/service_creds.json', SCOPES)

    # gc = pygsheets.authorize(service_file='./data/service_creds.json')
    gc = pygsheets.authorize(client_secret='auth_test_data/client_secret.json')
    sheet = gc.create('sheet')
    sheet.share('pygsheettest@gmail.com')

    # wks = gc.open_by_key('18WX-VFi_yaZ6LkXWLH856sgAsH5CQHgzxjA5T2PGxIY')
    # ss = gc.open('pygsheetTest')
    # print (ss)

    # wks = ss.sheet1
    # print (wks)

    # import  pandas as pd
    # import numpy as np
    #
    # tuples = list(zip(*[['bar', 'bar', 'baz', 'baz', 'foo', 'foo', 'qux', 'qux'],
    # ['one', 'two', 'one', 'two', 'one', 'two', 'one', 'two']]))
    # index = pd.MultiIndex.from_tuples(tuples, names=['first', 'second'])
    # df = pd.DataFrame(np.random.randn(8, 2), index=index, columns=['A', 'B'])

    # glogger = logging.getLogger('pygsheets')
    #glogger.setLevel(logging.DEBUG)

    IPython.embed()
