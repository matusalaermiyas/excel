from tkinter import filedialog, Tk, Frame, ttk
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
from openpyxl.cell import Cell

import os

app = Tk()
frame = Frame(app)
frame.place(relx=0.5, rely=0.5, anchor="c")

x = ttk.Style()
x.configure("BW.TLabel", foreground="white",
            background="darkgrey", font=('calibri', 15, 'bold'), padding=10)


def get_column_input() -> str:
    return column_input.get()


def get_row_input() -> str:
    return row_input.get()


def handle_open():

    path = filedialog.askopenfilename()
    file_name = os.path.basename(path)

    pattern = PatternFill(start_color='FF0000',
                          end_color='FF0000', fill_type="solid")

    workbook: Workbook = load_workbook(path)
    workspace = workbook.active

    for row in workspace.iter_rows():
        make_it_red = False

        cell: Cell

        for cell in row:
            if cell.column_letter == get_column_input().upper() or cell.row < int(get_row_input()):
                break

            if cell.value == "I" or cell.value == "NG" or cell.value == None:
                make_it_red = True

        if make_it_red:
            for cell in row:
                cell.fill = pattern

    workbook.save(filename=file_name)


label = ttk.Label(frame, text="Max Column")
label.pack(side="left")

column_input = ttk.Entry(frame)
column_input.pack(side="right")
column_input.insert(0, "Max Column")

row_input = ttk.Entry(frame)
row_input.pack(side="right")
row_input.insert(0, "Starting Row")

open_button = ttk.Button(frame, text='Open Excel',
                         command=handle_open, style='BW.TLabel')
open_button.pack(side="bottom")


app.geometry("500x500")
app.title("Excel")
app.mainloop()
