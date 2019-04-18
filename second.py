# 041819 - Vlookup for Salesforce 18 Character ID
import os
import glob
import time
import shutil
import webbrowser
import pandas as pd
import datetime as dt

dl_dir = '/Users/smk/Downloads/'
root_dir = '/Users/smk/Desktop/vlookup/'
dest_dir = '/Users/smk/Google Drive/vlook/'
goog_url = 'https://drive.google.com/drive/folders/'

def drop_y(df):
    to_drop = [x for x in df if x.endswith('_y')]
    df.drop(to_drop, axis=1, inplace=True)


files = glob.glob(dl_dir + '*.xlsx')
latest_file = max(files, key=os.path.getctime)
print(latest_file)

ct = dt.datetime.fromtimestamp(os.path.getmtime(latest_file))
create_time = int(ct.strftime("%Y%m%d%H%M"))
print(create_time)

now = int(time.strftime("%Y%m%d%H%M"))
print(now)

how_old = now - create_time
print(how_old)

if how_old <= 0:
    print(latest_file)
    
    df = pd.read_excel(latest_file, sheet_name='1')
    df2 = pd.read_excel(latest_file, sheet_name='2')

    vlook = df.merge(df2, on='Contact ID (18 character)', how='left')

    drop_y(vlook)

    date_created = time.strftime("%m.%d.%Y_%H:%M:%S")

    vlook.to_csv('python_vlook_' + date_created + '.csv', sep=',', index=False)

    for find_csv in sorted(glob.glob(root_dir + '*.csv')):
        print(find_csv)

    shutil.copy(find_csv, dest_dir)

    try:
        os.remove(find_csv)
        print("SUCCESS!")
        time.sleep(4)
        webbrowser.open_new_tab(goog_url + '1V4waIWHHAtF2kIXQ-J8a6zbv1flfp8DI')
    except OSError as e:
       print ("Error: %s - %s." % (e.filename, e.strerror)) 

else:
    print("🤷‍ No recent .xlsx files, build again! 🤷‍")



## Imporvements
## check to see if there are two workbooks if root_dir.. if true then assign to df2.
## else use current methods of 1 workbook w/ two


### Misc

# for workbook in sorted(glob.glob(root_dir + '*.xlsx')):
#     print(workbook)
