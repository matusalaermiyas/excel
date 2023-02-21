
from app import dialog
from tkinter import filedialog, Tk, ttk
from pathlib import Path
from pdf2docx import Converter

from threading import Thread


class PdfToWord:
    def __init__(self, app: Tk) -> None:
        open_button = ttk.Button(app, text='Convert All Files In Directory To Word',
                                 command=self.handle_convert, style='BW.TLabel')

        open_button.place(x=150, y=550, width=400, height=50)

    def handle_convert(self):
        print("Kicking here..")
        thread = Thread(target=self.thread_convert)
        thread.start()

        print("Convert finished ....")

    def thread_convert(self):

        directory = filedialog.askdirectory()

        pdfs = Path(directory).glob("*.pdf")

        for pdf in pdfs:
            converted_pdf = Converter(pdf)
            file_name = f"{Path(pdf).stem}.docx"
            converted_pdf.convert(file_name)
            converted_pdf.close()

        dialog.success()
