# 041819 - Vlookup for Salesforce 18 Character ID
import glob as gb
import pandas as pd
import tkinter as tk
import datetime as dt
import os, time, shutil
import webbrowser as wb
from tkinter.font import Font
from tkinter import filedialog
from tkinter.filedialog import askopenfilename

user_path = "/Users/smk"
dl_dir   = user_path + '/Downloads/'
root_dir = user_path + '/Desktop/vlookup/'
dest_dir = user_path + '/Google Drive/vlook/'
goog_url = 'https://drive.google.com/drive/folders/'


def getFolderPath(self):
    file_selected = filedialog.askopenfilename(initialdir = dl_dir, title = "Select File", filetypes = (("Excel Files","*.xlsx"),("All Files","*.*")))
    filePath.set(file_selected)


def drop_y(df):
    to_drop = [x for x in df if x.endswith('_y')]
    df.drop(to_drop, axis=1, inplace=True)


height = 400
width = 500

root = tk.Tk()
canvas = tk.Canvas(root, height=height, width=width)
canvas.pack()

frame = tk.Frame(root, bg='#88c1ff')
frame.place(relwidth=1, relheight=1)

text = tk.Text(root)
headerFont = Font(size=12)
text.configure(font=headerFont)

filePath = tk.StringVar()

files = gb.glob(dl_dir + '*.xlsx')
latest_file = max(files, key=os.path.getctime)

ct = dt.datetime.fromtimestamp(os.path.getmtime(latest_file))
create_time = int(ct.strftime("%Y%m%d%H%M"))
now = int(time.strftime("%Y%m%d%H%M"))
how_old = now - create_time
print(how_old)

if how_old <= 1:
    files = gb.glob(dl_dir + '*.xlsx')
    latest_file = max(files, key=os.path.getctime)
else:
    latest_file = "No recent *.xlsx file in your downloads folder"

drop_entry_options = ['left','right','outer','inner']
drop_entry = tk.StringVar()

drop_subheader = tk.Label(frame, bg='#88c1ff', font=headerFont, text="Select your join:")
drop_subheader.place(relx=-0.21, rely=0.17, relwidth=0.8, relheight=0.1)

drop_down = tk.OptionMenu(frame, drop_entry, *drop_entry_options)
drop_down.place(relx=0.1, rely=0.25, relwidth=0.8, relheight=0.1)

file1_subheader = tk.Label(frame, bg='#88c1ff', font=headerFont, text="Latest spreadsheet in Downloads:")
file1_subheader.place(relx=-0.11, rely=0.42, relwidth=0.8, relheight=0.1)

file1 = tk.Label(frame, bg='#ffffff', text=latest_file)
file1.place(relx=0.1, rely=0.50, relwidth=0.8, relheight=0.1)

file2_subheader = tk.Label(frame, bg='#88c1ff', font=headerFont, text="Select a spreadsheet:")
file2_subheader.place(relx=-0.18, rely=0.62, relwidth=0.8, relheight=0.1)

file2 = tk.Entry(frame, bg='#ffffff', textvariable=filePath)
file2.place(relx=0.1, rely=0.70, relwidth=0.8, relheight=0.1)

button = tk.Button(frame, text="Select a file", bg='#333333', command=lambda: getFolderPath(file2.get()))
button.place(relx=0.19, rely=0.84, relwidth=0.3, relheight=0.088)

submit = tk.Button(frame, text="Submit", bg='#333333', command = root.destroy)
submit.place(relx=0.51, rely=0.84, relwidth=0.3, relheight=0.088)

root.mainloop()


if how_old <= 1:

    df = pd.read_excel(filePath.get())
    df2 = pd.read_excel(latest_file)

    key_headers = df.columns.values.tolist()

    vlook = df.merge(df2, on=key_headers[0], how=drop_entry)
    drop_y(vlook)
    date_created = time.strftime("%m.%d.%Y_%H:%M:%S")
    vlook.to_csv('python_vlook_' + date_created + '.csv', sep=',', index=False)

    for find_csv in sorted(gb.glob(root_dir + '*.csv')):
        print(find_csv)

    shutil.copy(find_csv, dest_dir)

    try:
        os.remove(find_csv)
        print("🔥🔥🔥🔥🔥🔥 SUCCESS! 🔥🔥🔥🔥🔥🔥")
        time.sleep(1)
        wb.open(goog_url + '1V4waIWHHAtF2kIXQ-J8a6zbv1flfp8DI', new=2, autoraise=False)
    except OSError as e:
       print ("Error: %s - %s." % (e.filename, e.strerror)) 

else:
    print("🤷‍ No recent .xlsx files, build again! 🤷‍")




## Imporvements

# 0) add a refresh button to update the most recent file
# 1) add settings for type of "Join"
# 2) add setting to ask if if you have one file w/ two sheets or two 
#    workbooks.
# 3) DONE: add a way to get the name of the first column in the first workbook and use that
#    as the uID to match on.
# 4) Investigate the use of a class for the GUI

### Misc

# for workbook in sorted(glob.glob(root_dir + '*.xlsx')):
#     print(workbook)
