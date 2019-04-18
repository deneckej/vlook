import pandas as pd
from pandas import ExcelWriter
import shutil  
import glob
import os

root_dir = '/Users/smk/Desktop/vlookup/'
destination_folder = '/Users/smk/Google Drive/'
file_path = '/Users/smk/Desktop/vlookup/PythonExport.csv'

def drop_y(df):
    to_drop = [x for x in df if x.endswith('_y')]
    df.drop(to_drop, axis=1, inplace=True)


for workbook1 in sorted(glob.glob(root_dir + '*.xlsx')):
    print(workbook1)

df = pd.read_excel(workbook1, sheet_name='sheet1')
df2 = pd.read_excel(workbook1, sheet_name='sheet2')

results = df.merge(df2, on='Contact ID (18 character)', how='left')

drop_y(results)

results.to_csv('PythonExport.csv', sep=',', index=False)
shutil.copy('PythonExport.csv', destination_folder)

try:
    os.remove(file_path)
except OSError as e:
    print ("Error: %s - %s." % (e.filename, e.strerror))


# check to see if there are two workbooks if root_dir.. if true then assign to df2.
# else use current methods of 1 workbook w/ two
