import tkinter as tk
from tkinter import ttk


class BarraDeCarga:
    def __init__(self, parent, max_value=100):
        self.parent = parent
        self.max_value = max_value

        self.barra = ttk.Progressbar(parent, maximum=max_value, mode="determinate")
        self.etiqueta = tk.Label(parent, text="Progreso: 0%")

    def mostrar(self):
        self.barra.pack()
        self.etiqueta.pack()

    def ocultar(self):
        self.barra.pack_forget()
        self.etiqueta.pack_forget()

    def actualizar(self, valor_actual):
        porcentaje = (valor_actual / self.max_value) * 100
        self.barra["value"] = valor_actual
        self.etiqueta.config(text=f"Progreso: {porcentaje:.1f}%")

    def establecer_maximo(self, max_value):
        self.max_value = max_value
        self.barra["maximum"] = max_value

    def reset(self):
        self.barra["value"] = 0  # Corrección aquí
        self.etiqueta.config(text="Progreso: 0%")
