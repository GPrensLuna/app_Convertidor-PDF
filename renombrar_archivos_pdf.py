import os


def renombrar_archivos_pdf(carpeta_pdf):
    for archivo in os.listdir(carpeta_pdf):
        if archivo.endswith(".pdf"):
            nuevo_nombre = archivo
            os.rename(
                os.path.join(carpeta_pdf, archivo),
                os.path.join(carpeta_pdf, nuevo_nombre),
            )
