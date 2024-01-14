import os
import win32com.client
from pywintypes import com_error
import logging
from tkinter import filedialog, messagebox
from config import set_global_carpeta_pdf

def convertir_excel_a_pdf(
    app, barra_progreso, archivo_excel=None, hojas_a_convertir=None
):
    global global_carpeta_pdf
    excel = None
    workbook = None
    try:
        if archivo_excel is None:
            archivo_excel = filedialog.askopenfilename(
                filetypes=[("Archivos Excel", "*.xlsx")]
            )
            if not archivo_excel:
                logging.error("No se seleccionó un archivo.")
                messagebox.showerror("Error", "No se seleccionó un archivo válido.")
                return

        archivo_ruta, archivo_nombre = os.path.split(archivo_excel)
        nombre_sin_extension = os.path.splitext(archivo_nombre)[0]
        carpeta_pdf = os.path.join(archivo_ruta, nombre_sin_extension).replace(" ", "_")

        if not os.path.exists(carpeta_pdf):
            os.makedirs(carpeta_pdf)
        set_global_carpeta_pdf(carpeta_pdf)

        excel = win32com.client.DispatchEx("Excel.Application")
        excel.DisplayAlerts = False
        workbook = excel.Workbooks.Open(archivo_excel)

        if hojas_a_convertir is None:
            hojas_a_convertir = [sheet.Name for sheet in workbook.Sheets]

        total_hojas = len(hojas_a_convertir)
        barra_progreso.establecer_maximo(total_hojas)

        for index, sheet_name in enumerate(hojas_a_convertir, start=1):
            try:
                pdf_file_name = (f"{sheet_name}.pdf").replace(" ", "%20")
                pdf_path = os.path.join(carpeta_pdf, pdf_file_name)

                if os.path.exists(pdf_path):
                    os.remove(pdf_path)

                sheet = workbook.Sheets[sheet_name]
                sheet.ExportAsFixedFormat(0, pdf_path)
                barra_progreso.actualizar(index)
                app.update_idletasks()
            except com_error as e:
                logging.error(f"Error al exportar la hoja '{sheet_name}' a PDF: {e}")
                messagebox.showerror(
                    "Error de Exportación",
                    f"No se pudo exportar la hoja '{sheet_name}' a PDF.",
                )

        messagebox.showinfo(
            "Conversión Completa", "Todos los archivos han sido convertidos a PDF."
        )

    except com_error as e:
        manejo_errores_com(e)
    except Exception as e:
        logging.exception("Error desconocido")
        messagebox.showerror("Error", f"Se produjo un error: {str(e)}")
    finally:
        if workbook is not None:
            workbook.Close()
        if excel is not None:
            excel.Quit()



def manejo_errores_com(e):
    if "Excel cannot access" in str(e):
        messagebox.showerror(
            "Error",
            "El archivo Excel está abierto. Ciérrelo y guarde los cambios antes de continuar.",
        )
    else:
        logging.exception("Error al manejar el archivo Excel")
        messagebox.showerror("Error", f"Se produjo un error: {str(e)}")