import re
import tkinter as tk
import tkinter.filedialog as fd
from tkinter import ttk, W
from openpyxl.workbook import Workbook
import pandas as pd
from tkinter.ttk import *

unique_codes = []
checked_file = False

def check_for_error(line):
    code = re.search("Проба с номером (.*) не найдена", line)
    return code

def open_file():
    global listbox
    global file_label
    global found_numbers

    try:
        listbox.destroy()
        file_label.destroy()
        found_numbers.destroy()

        file_label = tk.Label(root)
        found_numbers = tk.Label(root)

        #file_label.grid_forget()
        #found_numbers.grid_forget()

    except:
        pass

    file_input_name = get_file()
    if file_input_name != None:
        error_list = []
        pb = Progressbar(root, orient='horizontal', mode='indeterminate', length=280)
        pb.grid(sticky=W, padx=10, pady=10)
        pb.start()
        with open(file_input_name, 'r', encoding='utf-8') as data:
            for line in data:
                code = check_for_error(line)
                if code:
                    error_list.append(code.group(1))
                else:
                    pass

        checked_file = True
        print('finished checking')
        global unique_codes
        unique_codes = pd.unique(error_list)
        print('till listbox')
        pb.stop()
        pb.destroy()

        if checked_file:
            #file_label = tk.Label(root)
            #found_numbers = tk.Label(root)
            file_label.config(text='Файл: ' + file_input_name)
            file_label.grid(sticky=W, padx=10, pady=10)
            found_numbers.config(text='Номера не найденных проб: ')
            found_numbers.grid(sticky=W, padx=10, pady=10)

            listbox = tk.Listbox(root, height=10, width=40)
            listbox.grid(sticky=W, padx=10, pady=10)

            print('displaying objects')

            display_items(unique_codes, listbox)

def get_file():
    file = fd.askopenfile(parent=root, mode='r', title='Выберите файл')
    return file.name

def file_save():

    f = fd.asksaveasfile(mode='w', defaultextension=".xlsx")
    if f is None: # asksaveasfile return `None` if dialog closed with "cancel".
        return

    out = pd.DataFrame(unique_codes, columns=['Номер пробы'])
    print(out.shape)
    out.to_excel(f.name, index=False)

def display_items(code_list, listbox):
    global export_btn

    try:
        export_btn.destroy()

        export_btn = Button(root)

        #export_btn.grid_forget()
    except:
        pass

    if code_list.size > 0:
        r = 1
        for i in code_list:
            listbox.insert(r, i)
            r += 1

        #export_btn = Button(root)
        export_btn.config(text="Экспорт в Экзель", width=25, command=file_save)
        export_btn.grid(sticky=W, padx=10, pady=10)

    else:
        listbox.insert(1, 'Нет данных')

root = tk.Tk()
#root.withdraw()

#p1 = tk.PhotoImage(file='icon.ico')

#root.iconphoto(False, p1)
root.iconbitmap("icon.ico")

root.title("Log Finder")

root.geometry("300x400")

browse_label = Label(root, text="Нажмите ниже чтобы выбрать файл для анализа").grid(sticky=W, padx=10, pady=10)

browse_btn = Button(root, text='Выбрать файл', width=25, command=open_file).grid(sticky=W, padx=10, pady=10)
#browse_btn.pack(side='bottom')

file_label = tk.Label(root)

found_numbers = tk.Label(root)

export_btn = Button(root)

#file_label = tk.Label(text="Файл: " + fn)

#for c in sorted(root.children):
#    root.children[c].pack()

root.mainloop()
