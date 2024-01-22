import os
import pandas as pd
import tkinter as tk
import openpyxl
from tkinter import ttk

from conversor import conversor

def convertir():
    cantidad = entry_cantidad.get()
    
    # Llamar a la función conversor y obtener el resultado
    resultado = conversor(ui.get(), uf.get(), cantidad, tabla_seleccionada)

    # Actualizar la variable asociada a la etiqueta
    resultado_var.set(resultado)

# Crear una ventana
window = tk.Tk()
window.title("Conversor de Unidades")

# Ruta del archivo Excel con varias hojas
path = './tablas/tablas.xlsx'

# Leer el archivo Excel con varias hojas
wb = openpyxl.load_workbook(path, read_only=True)
hojas_excel = wb.sheetnames
wb.close()

def cargar_hojas(*args):
    # Obtener la hoja seleccionada del archivo Excel
    hoja_seleccionada = sheets.get()

    # Leer la hoja seleccionada
    global tabla_seleccionada
    tabla_seleccionada = pd.read_excel(path, sheet_name=hoja_seleccionada, header=None)

    # Actualizar los desplegables de unidades con las columnas de la tabla seleccionada
    ui['values'] = tabla_seleccionada.iloc[1:, 0].tolist()
    uf['values'] = tabla_seleccionada.iloc[0, 1:].tolist()

# Crear etiquetas y controles en la interfaz gráfica
label_hojas = tk.Label(window, text="Unidades de:")
label_hojas.grid(row=0, column=0, padx=10, pady=10)

# Crear desplegable para seleccionar la hoja del archivo Excel
sheets = ttk.Combobox(window, values=hojas_excel)
sheets.grid(row=0, column=1, padx=10, pady=10)

label_ui = tk.Label(window, text="Unidad de Entrada:")
label_ui.grid(row=1, column=0, padx=10, pady=10)

ui = ttk.Combobox(window, values=[])
ui.grid(row=1, column=1, padx=10, pady=10)

label_uf = tk.Label(window, text="Unidad de Salida:")
label_uf.grid(row=2, column=0, padx=10, pady=10)

uf = ttk.Combobox(window, values=[])
uf.grid(row=2, column=1, padx=10, pady=10)

label_cantidad = tk.Label(window, text="Cantidad:")
label_cantidad.grid(row=3, column=0, padx=10, pady=10)

entry_cantidad = tk.Entry(window)
entry_cantidad.grid(row=3, column=1, padx=10, pady=10)

button_convertir = tk.Button(window, text="Convertir", command=convertir)
button_convertir.grid(row=4, column=0, columnspan=2, pady=10)

# Utilizar StringVar para la variable asociada a la etiqueta
resultado_var = tk.StringVar()
label_resultado = tk.Label(window, textvariable=resultado_var, wraplength=300)
label_resultado.grid(row=5, column=0, columnspan=2, pady=10)

# Vincular la función cargar_hojas al evento de selección del desplegable de hojas
sheets.bind("<<ComboboxSelected>>", cargar_hojas)

# Iniciar el bucle de eventos
window.mainloop()