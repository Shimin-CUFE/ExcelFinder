import tkinter as tk
from tkinter.filedialog import askdirectory
from guiver_backend import FindMethod
from threading import Thread  # TODO:

root = tk.Tk()
root.title('ExcelFinder')
root.attributes("-topmost", 1)
root.geometry('500x600')
path = tk.StringVar()
value = tk.StringVar()
msg_list = []
msg = tk.StringVar()





def get_path():
    p = askdirectory()
    path.set(p)

def click():
    print(tx)
    fm = FindMethod()
    d = {
        'v': value.get(),
        'p': path.get(),
        'text': tx
    }
    fm.search_path(d)


def show_msg(m):
    msg_list.append(m)
    msg.set(msg.get() + '\n' + m)


tk.Label(root, text='Directory: ').grid(row=0, column=0)
entry1 = tk.Entry(root, textvariable=path).grid(row=0, column=1)
tk.Button(root, text='Choose', command=get_path).grid(row=0, column=2)
tk.Label(root, text='Value to Search: ').grid(row=1, column=0)
entry2 = tk.Entry(root, textvariable=value).grid(row=1, column=1)
tk.Button(root, text='Search!', command=click).grid(row=1, column=2)
tx = tk.Text(root).grid(row=2, column=0, columnspan=3)

root.mainloop()
