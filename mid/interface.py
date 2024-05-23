import customtkinter # type: ignore
from PIL import Image, ImageTk
import os

import tkinter
from tkinter import filedialog

# tkinter.Tk().withdraw() # prevents an empty tkinter window from appearing


customtkinter.set_appearance_mode("dark")   # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"

class App(customtkinter.CTk):
    width = 900
    height = 600

    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.title("MAXIMELT - screw generator")
        self.geometry(f"{self.width}x{self.height}")
        self.resizable(False, False)

        self.iconpath = ImageTk.PhotoImage(file="C:\\repo_git\\SCREW\\img\\MAXIMELT.ico")
        self.wm_iconbitmap()
        self.iconphoto(False, self.iconpath)

        # create main frame
        self.main_frame = customtkinter.CTkFrame(self, corner_radius=0)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_label = customtkinter.CTkLabel(self.main_frame, text="CustomTkinter\nMain Page",
                                                 font=customtkinter.CTkFont(size=20, weight="bold"))
        self.main_label.grid(row=0, column=0, padx=30, pady=(30, 15))
        self.back_button = customtkinter.CTkButton(self.main_frame, text="Back", command=self.back_event, width=200)
        self.back_button.grid(row=1, column=0, padx=30, pady=(15, 15))

    def login_event(self):
        # print("Login pressed - username:", self.username_entry.get(), "password:", self.password_entry.get())

        self.login_frame.grid_forget()  # remove login frame
        self.main_frame.grid(row=0, column=0, sticky="nsew", padx=100)  # show main frame

    def back_event(self):
        self.main_frame.grid_forget()  # remove main frame
        self.login_frame.grid(row=0, column=0, sticky="ns")  # show login frame


if __name__ == "__main__":
    app = App()
    folder_path = filedialog.askdirectory()
    print(folder_path)
    app.mainloop()