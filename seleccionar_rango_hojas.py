from tkinter import filedialog
from seleccionar_hojas_con_dropdown import seleccionar_hojas_con_dropdown
from convertir_excel_a_pdf import convertir_excel_a_pdf


def seleccionar_rango_hojas(app):
    archivo_excel = filedialog.askopenfilename(filetypes=[("Archivos Excel", "*.xlsx")])
    if archivo_excel:
        seleccionar_hojas_con_dropdown(archivo_excel, convertir_excel_a_pdf, app)
