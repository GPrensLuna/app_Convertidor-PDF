import os
from tkinter import messagebox
from config import get_global_carpeta_pdf

def abrir_carpeta_pdf():
    carpeta_pdf = get_global_carpeta_pdf()
    print(f"Intentando abrir la carpeta: {carpeta_pdf}")
    if carpeta_pdf:
        os.startfile(carpeta_pdf)
    else:
        messagebox.showinfo(
            "Información", "No se ha generado ninguna carpeta PDF todavía."
        )
