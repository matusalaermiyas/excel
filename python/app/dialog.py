from tkinter import messagebox


def error(title="Error", body="Make Sure To Select A File"):
    messagebox.showerror(title, body)


def success(title="Success", body="Task Finished Successfully"):
    messagebox.showinfo(title, body)
