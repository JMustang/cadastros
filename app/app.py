import customtkinter as ctk
from tkinter import *
from tkinter import messagebox
import openpyxl, xlrd
import pathlib
from openpyxl import Workbook

ctk.set_appearance_mode("System")
ctk.set_default_color_theme("blue")


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.layout_config()

    def layout_config(self):
        self.title("Cadastros")
        self.geometry("700x500")


if __name__ == "__main__":
    app = App()
    app.mainloop()
