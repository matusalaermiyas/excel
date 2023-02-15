from tkinter import filedialog, Tk, Frame, ttk
from openpyxl import load_workbook, Workbook
from openpyxl.styles import PatternFill
import os

app = Tk()
frame = Frame(app)
frame.place(relx=0.5, rely=0.5, anchor="c")

x = ttk.Style()
x.configure("BW.TLabel", foreground="white",
            background="darkgrey", font=('calibri', 15, 'bold'), padding=10)


def handle_open():
    path = filedialog.askopenfilename()
    file_name = os.path.basename(path)

    pattern = PatternFill(start_color='FF0000',
                          end_color='FF0000', fill_type="solid")

    workbook: Workbook = load_workbook(path)
    workspace = workbook.active

    for row in workspace.iter_rows():
        make_it_red = False

        for cell in row:
            if cell.value == "I" or cell.value == "NG":
                make_it_red = True

        if make_it_red:
            for cell in row:
                cell.fill = pattern

    workbook.save(filename=file_name)


open_button = ttk.Button(frame, text='Open Excel',
                         command=handle_open, style='BW.TLabel')
open_button.pack()


app.geometry("500x500")
app.title("Excel")
app.mainloop()
