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

    # gc = pygsheets.authorize(service_file='auth_test_data/pygsheettest_service_account.json')
    gc = pygsheets.authorize(client_secret='auth_test_data/client_secret.json',
                             credentials_directory='auth_test_data')
    # sheet = gc.open('sheet')
    # sheet.share('pygsheettest@gmail.com')

    ss = gc.open('manualTestSheet')
    print (ss)

    wks = ss.sheet1
    print (wks)

    # import  pandas as pd
    # import numpy as np
    #
    # arrays = [np.array(['bar', 'bar', 'baz', 'baz', 'foo', 'foo', 'qux', 'qux']),
    #           np.array(['one', 'two', 'one', 'two', 'one', 'two', 'one', 'two'])]
    # tuples = list(zip(*arrays))
    # index = pd.MultiIndex.from_tuples(tuples, names=['first', 'second'])
    # df = pd.DataFrame(np.random.randn(8, 8), index=index, columns=index)

    # glogger = logging.getLogger('pygsheets')
    # glogger.setLevel(logging.DEBUG)

    # import pandas.util.testing as tm;
    # tm.N = 3
    # def unpivot(frame):
    #     N, K = frame.shape
    #     data = {'value': frame.values.ravel('F'),
    #             'variable': np.asarray(frame.columns).repeat(N),
    #             'date': np.tile(np.asarray(frame.index), K)}
    #     return pd.DataFrame(data, columns=['date', 'variable', 'value'])
    #
    # df = unpivot(tm.makeTimeDataFrame())
    # df['value2'] = df['value'] * 3
    # df = df.pivot(index='date', columns='variable')

    IPython.embed()
