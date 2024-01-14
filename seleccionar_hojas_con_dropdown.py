import tkinter as tk
from tkinter import messagebox, ttk
import win32com.client


def seleccionar_hojas_con_dropdown(archivo_excel, callback, app):
    def confirmar_seleccion():
        inicio = lista_hojas_inicio.current()
        fin = lista_hojas_fin.current()
        if inicio <= fin:
            rango_seleccionado = hojas_disponibles[inicio : fin + 1]
            rango.set(rango_seleccionado)
            ventana_seleccion.destroy()
            callback(archivo_excel, rango_seleccionado, app)
        else:
            messagebox.showerror(
                "Error", "La hoja de inicio debe ser anterior a la hoja final."
            )

    ventana_seleccion = tk.Toplevel(app)
    ventana_seleccion.title("Seleccionar Rango de Hojas")
    rango = tk.StringVar()
    excel = win32com.client.DispatchEx("Excel.Application")
    workbook = excel.Workbooks.Open(archivo_excel)
    hojas_disponibles = [sheet.Name for sheet in workbook.Sheets]
    workbook.Close()
    excel.Quit()

    tk.Label(ventana_seleccion, text="Hoja de inicio:").pack()
    lista_hojas_inicio = ttk.Combobox(ventana_seleccion, values=hojas_disponibles)
    lista_hojas_inicio.pack()
    lista_hojas_inicio.current(0)  # Seleccionar el primer elemento por defecto

    tk.Label(ventana_seleccion, text="Hoja final:").pack()
    lista_hojas_fin = ttk.Combobox(ventana_seleccion, values=hojas_disponibles)
    lista_hojas_fin.pack()
    lista_hojas_fin.current(0)  # Seleccionar el primer elemento por defecto

    boton_confirmar = tk.Button(
        ventana_seleccion, text="Confirmar", command=confirmar_seleccion
    )
    boton_confirmar.pack()
    ventana_seleccion.mainloop()


# Ejemplo de cómo se podría usar la función
# app = tk.Tk() # Suponiendo que 'app' es la instancia de Tkinter principal
# archivo_excel = "ruta/a/tu/archivo.xlsx" # Ruta al archivo Excel
# callback = tu_funcion_callback # La función a llamar con el rango seleccionado
# seleccionar_hojas_con_dropdown(archivo_excel, callback, app)
