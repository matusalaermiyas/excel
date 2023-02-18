
from tkinter import ttk,  Tk

from app.mark_invalid_data import MakeRowRedWithInvalidData
from app.get_invalid_ids import CreateExcelWithInvalidIds
from app.remove_slip import RemoveARowWithInvalidId

from app import dialog


def start():

    try:
        app = Tk()

        x = ttk.Style()
        x.configure("BW.TLabel", foreground="white",
                    background="darkgrey", font=('calibri', 15, 'bold'), padding=10)

        MakeRowRedWithInvalidData(app)  # Make a row red
        CreateExcelWithInvalidIds(app)  # Get ids from excel
        RemoveARowWithInvalidId(app)  # Remove invalid slip

        app.geometry("800x800")
        app.title("Excel")
        app.mainloop()

    except Exception as ex:
        print(ex)
        dialog.error(body="Error occured try again!")


if __name__ == "__main__":
    start()
