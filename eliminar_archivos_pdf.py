import os


def eliminar_archivos_pdf(carpeta_pdf, workbook):
    for sheet in workbook.Sheets:
        pdf_nombre = f"{sheet.Name}.pdf"
        pdf_ruta = os.path.join(carpeta_pdf, pdf_nombre)
        if os.path.exists(pdf_ruta):
            os.remove(pdf_ruta)
