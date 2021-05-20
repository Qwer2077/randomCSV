# import io
import json
import os
import tkinter as tk
import tkinter.filedialog

import pandas as pd

try:
    os.makedirs("tmp/")
except FileExistsError:
    # print("dir exist")
    pass


def startupCheck(PATH):
    if os.path.isfile(PATH) and os.access(PATH, os.R_OK):
        # checks if file exists
        # print("File exists and is readable")
        pass
    else:
        # print("Either file is missing or is not readable, creating file...")
        with open(PATH, 'w+') as fp:
            config = {
                "defaultPath": "/"
            }

            json.dump(config, fp, sort_keys=True, indent=4)


startupCheck("tmp/config.json")

json_file = json.load(open('tmp/config.json'))

# print(json_file["defaultPath"])

root = tk.Tk()
filename = ""
label1 = None
root.title("Convert xls to csv")

root.geometry("500x100+350+250")


def openfile():
    global filename, label1

    if label1 is not None:
        label1.destroy()

    filename = tk.filedialog.askopenfilename(initialdir=json_file["defaultPath"], title="Select a xlsx file",
                                             filetypes=(("xls files", "*.xls"), ("all files", "*.*")))
    label1 = tk.Label(root, text=filename)
    label1.pack()


def confirm(f1):
    df = pd.read_excel(f1)
    col = list(df.columns)

    col[0], col[1] = col[1], col[0]
    col[1], col[2] = col[2], col[1]
    col[2], col[3] = col[3], col[2]
    col[3], col[15] = col[15], col[3]
    df = df[col]

    df = df.loc[::-1]

    df.to_csv(f1[:-3] + "csv", index=False)


def selectDefault():
    directory = tk.filedialog.askdirectory()
    json_file["defaultPath"] = directory

    with open("tmp/config.json", "w") as f:
        json.dump(json_file, f, indent=4)


bt1 = tk.Button(root, text="Open File", command=openfile).pack()

bt2 = tk.Button(root, text="Confirm", command=lambda: confirm(filename)).pack()

bt3 = tk.Button(root, text="Select default directory", command=selectDefault).pack()

root.mainloop()
