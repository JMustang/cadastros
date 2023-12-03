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
        self.appearence()
        self.sistema()

    def layout_config(self):
        self.title("Cadastros")
        self.geometry("700x500")

    def appearence(self):
        self.lb_apm = ctk.CTkLabel(
            self, text="Tema", bg_color="transparent", text_color=["#000", "#fff"]
        ).place(x=50, y=430)
        self.opt_apm = ctk.CTkOptionMenu(
            self, values=["Light", "Dark", "System"], command=self.change_apm
        ).place(x=50, y=460)

    def sistema(self):
        frame = ctk.CTkFrame(
            self,
            width=700,
            height=50,
            corner_radius=0,
            bg_color="darkmagenta",
            fg_color="darkmagenta",
        ).place(x=0, y=10)
        title = ctk.CTkLabel(
            frame, text="Cadastros", font=("Arial bold", 24), text_color="#fff"
        )
        span = ctk.CTkLabel(
            self,
            text="Por favor, preencha o formulário!",
            font=("Arial bold", 16),
            text_color=["#000", "#fff"],
        ).place(x=50, y=70)

        # Entrys
        en_name = ctk.CTkEntry(
            self,
            width=350,
            font=("Century Gohtic bold", 16),
            fg_color="transparent",
            placeholder_text="Nome",
        )
        en_contact = ctk.CTkEntry(
            self,
            width=200,
            font=("Century Gohtic bold", 16),
            fg_color="transparent",
            placeholder_text="Contato",
        )
        en_age = ctk.CTkEntry(
            self,
            width=150,
            font=("Century Gohtic bold", 16),
            fg_color="transparent",
            placeholder_text="Idade",
        )
        en_address = ctk.CTkEntry(
            self,
            width=200,
            font=("Century Gohtic bold", 16),
            fg_color="transparent",
            placeholder_text="Endereço",
        )

        # Combobox
        opt_gender = ctk.CTkComboBox(
            self, values=["Masculino", "Feminino"], font=("Century Gohtic bold", 14)
        )
        opt_gender.set("Masculino")

        # TextBox
        txt_opt = ctk.CTkTextbox(
            self,
            width=500,
            height=150,
            font=("Arial", 18),
            border_color="#aaa",
            border_width=2,
            fg_color="transparent",
        )

        # Labels
        lb_name = ctk.CTkLabel(
            self,
            text="Por favor, preencha o formulário!",
            font=("Arial bold", 16),
            text_color=["#000", "#fff"],
        ).place(x=50, y=70)
        lb_contact = ctk.CTkLabel(
            self,
            text="Por favor, preencha o formulário!",
            font=("Arial bold", 16),
            text_color=["#000", "#fff"],
        )
        lb_age = ctk.CTkLabel(
            self,
            text="Por favor, preencha o formulário!",
            font=("Arial bold", 16),
            text_color=["#000", "#fff"],
        )
        lb_gender = ctk.CTkLabel(
            self,
            text="Por favor, preencha o formulário!",
            font=("Arial bold", 16),
            text_color=["#000", "#fff"],
        )
        lb_opt = ctk.CTkLabel(
            self,
            text="Por favor, preencha o formulário!",
            font=("Arial bold", 16),
            text_color=["#000", "#fff"],
        )
        lb_name = ctk.CTkLabel(
            self,
            text="Por favor, preencha o formulário!",
            font=("Arial bold", 16),
            text_color=["#000", "#fff"],
        )

    def change_apm(self, theme):
        ctk.set_appearance_mode(theme)
        self.appearence()


if __name__ == "__main__":
    app = App()
    app.mainloop()
