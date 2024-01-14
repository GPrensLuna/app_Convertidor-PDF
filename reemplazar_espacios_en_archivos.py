import os
import tkinter as tk
from tkinter import filedialog, messagebox


def reemplazar_espacios_en_archivos():
    root = tk.Tk()
    root.withdraw()  # Ocultamos la ventana raíz
    directorio = filedialog.askdirectory()
    if not directorio:  # Si el usuario cancela la selección
        return

    archivos_modificados = 0
    for nombre_archivo in os.listdir(directorio):
        if "%20" in nombre_archivo:
            nuevo_nombre = nombre_archivo.replace("%20", " ")
            os.rename(
                os.path.join(directorio, nombre_archivo),
                os.path.join(directorio, nuevo_nombre),
            )
            archivos_modificados += 1
            print(f"Cambiado: {nombre_archivo} a {nuevo_nombre}")

    messagebox.showinfo("Terminado", f"Se modificaron {archivos_modificados} archivos.")
