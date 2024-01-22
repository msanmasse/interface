import os
import pandas as pd
import tkinter as tk
import openpyxl
from tkinter import ttk

from conversor import conversor

def convertir(event=None):
    cantidad = entry_cantidad.get()
    # Usar la hoja seleccionada actualmente para la conversión
    hoja_seleccionada = selected_sheet.get()
    tabla_conversor = pd.read_excel(path, sheet_name=hoja_seleccionada, header=None)
    resultado = conversor(ui.get(), uf.get(), cantidad, tabla_conversor)
    resultado_var.set(resultado)

def mostrar_tabla(dataframe):
    # Limpiar el Treeview
    for i in tree.get_children():
        tree.delete(i)

    # Establecer encabezados
    tree["column"] = list(dataframe.columns)
    tree["show"] = "headings"

    # Formato de columnas
    for columna in tree["column"]:
        tree.heading(columna, text=columna)

    # Añadir datos
    for fila in dataframe.itertuples():
        tree.insert("", "end", values=fila[1:])

def actualizar_comboboxes(hoja_seleccionada):
    # Leer la hoja seleccionada
    tabla_conversor = pd.read_excel(path, sheet_name=hoja_seleccionada, header=None)

    # Actualizar los desplegables de unidades
    ui['values'] = tabla_conversor.iloc[1:, 0].tolist()
    uf['values'] = tabla_conversor.iloc[0, 1:].tolist()

def seleccionar_hoja(event):
    hoja_seleccionada = selected_sheet.get()
    actualizar_comboboxes(hoja_seleccionada)

# Crear una ventana
window = tk.Tk()
window.title("Conversor de Unidades")

# Establecer estilos generales
estilo = ttk.Style()
estilo.configure("TLabel", font=("Arial", 10))
estilo.configure("TButton", font=("Arial", 10), background="lightblue")
estilo.configure("TEntry", font=("Arial", 10))
estilo.configure("TCombobox", font=("Arial", 10))

# Organización con Frames
frame_superior = tk.Frame(window)
frame_superior.grid(row=0, column=0, padx=10, pady=10)

frame_central = tk.Frame(window)
frame_central.grid(row=1, column=0, padx=10, pady=10)

frame_inferior = tk.Frame(window)
frame_inferior.grid(row=2, column=0, padx=10, pady=10)

# Ruta del archivo Excel con varias hojas
path = './tablas/tablas.xlsx'

# Leer el archivo Excel con varias hojas
wb = openpyxl.load_workbook(path, read_only=True)
hojas_excel = wb.sheetnames
wb.close()

# Cargar la última hoja del archivo Excel para mostrar en la tabla
tabla_ultima_hoja = pd.read_excel(path, sheet_name=hojas_excel[-1], header=None)

# Diccionario para almacenar los datos de cada hoja en caché
datos_hojas = {}

# Leer el archivo Excel con varias hojas y almacenar los datos en el diccionario
for hoja in hojas_excel[:-1]:  # Excluyendo la última hoja que se supone es diferente
    datos_hojas[hoja] = pd.read_excel(path, sheet_name=hoja, header=None)

def actualizar_comboboxes(hoja_seleccionada):
    # Utilizar los datos en caché en lugar de leer el archivo
    tabla_conversor = datos_hojas[hoja_seleccionada]

    # Actualizar los desplegables de unidades
    ui['values'] = tabla_conversor.iloc[1:, 0].tolist()
    uf['values'] = tabla_conversor.iloc[0, 1:].tolist()

# Inicializar el diccionario de datos en caché y leer el archivo
datos_hojas = {}
for hoja in hojas_excel[:-1]:
    datos_hojas[hoja] = pd.read_excel(path, sheet_name=hoja, header=None)

# Crear etiquetas y controles en la interfaz gráfica
label_hojas = ttk.Label(frame_superior, text="Unidades de:")
label_hojas.pack(side="left", padx=5, pady=5)

selected_sheet = tk.StringVar()

sheets = ttk.Combobox(window, textvariable=selected_sheet, values=hojas_excel[:-1], state="readonly")
sheets.grid(row=0, column=1, padx=10, pady=10)
sheets.bind("<<ComboboxSelected>>", seleccionar_hoja)

label_ui = tk.Label(window, text="Unidad de Entrada:")
label_ui.grid(row=1, column=0, padx=10, pady=10)

ui = ttk.Combobox(window, values=[], state="readonly")
ui.grid(row=1, column=1, padx=10, pady=10)

label_uf = tk.Label(window, text="Unidad de Salida:")
label_uf.grid(row=2, column=0, padx=10, pady=10)

uf = ttk.Combobox(window, values=[], state="readonly")
uf.grid(row=2, column=1, padx=10, pady=10)

label_cantidad = tk.Label(window, text="Cantidad:")
label_cantidad.grid(row=3, column=0, padx=10, pady=10)

def validar_entrada(cadena):
    color_normal = entry_cantidad.cget('background')  # Obtener el color de fondo normal del Entry
    if cadena == "":
        entry_cantidad.config(highlightbackground=color_normal, highlightcolor=color_normal)  # Restablecer color normal
        error_label.config(text='')  # Ocultar mensaje de error
        return True
    try:
        float(cadena)  # Intenta convertir a float
        entry_cantidad.config(highlightbackground=color_normal, highlightcolor=color_normal)  # Restablecer color normal
        error_label.config(text='')  # Ocultar mensaje de error
        return True
    except ValueError:
        entry_cantidad.config(highlightbackground='red', highlightcolor='red', highlightthickness=2)  # Cambiar color del borde a rojo para entrada inválida
        error_label.config(text='Por favor, ingrese solo números.')  # Mostrar mensaje de error
        return False

# Crear un validador de entrada
validador = window.register(validar_entrada)

# Crear el campo de entrada y el validador
entry_cantidad = tk.Entry(window, validate="key", validatecommand=(validador, '%P'), highlightthickness=2)
entry_cantidad.grid(row=3, column=1, padx=10, pady=10)
entry_cantidad.bind('<Return>', convertir)  # Vinculación del evento Enter a la función convertir

button_convertir = tk.Button(window, text="Convertir", command=convertir)
button_convertir.grid(row=4, column=0, columnspan=2, pady=10)

# Etiqueta de error justo debajo del botón "Convertir"
error_label = tk.Label(window, text='', fg='red')  # Ajusta el valor de 'width' según sea necesario
error_label.grid(row=5, column=0, columnspan=2)  # Asegúrate de que el error_label esté centrado debajo del botón

resultado_var = tk.StringVar()
label_resultado = tk.Label(window, textvariable=resultado_var, wraplength=300)
label_resultado.grid(row=5, column=0, columnspan=2, pady=10)

# Crear un Treeview para mostrar la tabla de la última hoja
tree = ttk.Treeview(window)
tree.grid(row=6, column=0, columnspan=2, padx=10, pady=10, sticky='nsew')

# Mostrar la tabla de la última hoja
mostrar_tabla(tabla_ultima_hoja)

# Iniciar el bucle de eventos
window.mainloop()