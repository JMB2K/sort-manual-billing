import pandas as pd
from collections import OrderedDict as OD
import os
import requests
import re

file = os.path.join(os.path.expanduser('~'), 'Desktop', 'manualEntryLog.xlsx')

df = pd.read_excel(file)

def getNum(name):
    url = "http://www.corganinet.com/_applications/whois_V1/view_list_01.cfm"
    payload = {'MySearch_01': name}
    r=requests.post(url, data=payload)
    try:
        return re.findall('EMID=([0-9]+)', r.text)[0]
    except Exception:
        return 'Name not found in locator'

for i in range(1, len(df)):
    df.loc[i, 'ORDERED BY'] = '{} - {}'.format(df.loc[i, 'ORDERED BY'], getNum(df.loc[i, 'ORDERED BY']))

def filter_empty_entries(xDict):
    """
    This is going through the pandas_df
     and creating a dict out of the data,
     skipping anything with no value, certain
     depts needed a dict of their own and they
     are created here and stored in the
     final dict as a value to the dept
     """
    final_dict = OD()
    s, sc, lc, o = dict(), dict(), dict(), dict()
    # dicts for specs, lg color, sm color, and oce
    check_set = xDict.notnull()
    """
    returns True if there is data in the value
    of the dict, weeding out what isn't needed
    """
    for k, v in xDict.items():
        if k != 'PROJECT NUMBER' and check_set[k]:
        # project num already used as key
            if k.split()[1] in ['COPIES', 'ORIGS', 'COSTSHARE', 'TIME']:
                v = int(v)
            if k.startswith('OCE'):
                o[k.split()[1]] = v
            elif k.startswith('SMCLR'):
                sc[k.split()[1]] = v
            elif k.startswith('SPEC'):
                s[k.split()[1]] = v
            elif k.startswith('LRGCLR'):
                lc[k.split()[1]] = v
            else:
                final_dict[k] = v
    for i in [(s, 'SPECS'),(o, 'OCE'),(sc, 'SMALL COLOR'),(lc, 'LARGE COLOR')]:
        if len(i[0]) != 0:
            final_dict[i[1]] = i[0]
    return final_dict



entryList = (x[1] for x in df.iterrows())  # creating a generator of pandas series
entries = OD()
flag = False
for entry in entryList:
    if flag:
        if entry['PROJECT NUMBER'] not in entries:
        # skipping the first line - is null and ruins everything after
            entries[entry['PROJECT NUMBER']] = [filter_empty_entries(entry)]
            # either create a new entry in the dict
        else:
            entries[entry['PROJECT NUMBER']].append(filter_empty_entries(entry))
            # or append to the list already in the value
    else:
        flag = True

with open(os.path.join(os.path.expanduser('~'), 'Desktop', 'manualEntries.txt'), 'w') as f:
    """
    this is all just formatting the way it
    prints to file so that its easy to read
    """
    for job in entries:
        f.write('{}:\n'.format(job))
        for more in entries[job]:
            for k, v in more.items():
                if k in ['SPECS', 'LARGE COLOR', 'OCE', 'SMALL COLOR']:
                    f.write('\t{}:\n'.format(k))
                    for key, value in more[k].items():
                        f.write('\t\t{}: {}\n'.format(key, value))
                    f.write('\n')
                else:
                    f.write('\t{}: {}'.format(k, v))
                    f.write('\n')
            f.write('\n')
    f.close()




