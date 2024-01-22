import pandas as pd

def conversor(ui, uf, cantidad, tabla):
    valor_conversion = 0.0

    # Verificar si la unidad de salida se encuentra en la primera fila
    if uf not in tabla.iloc[0].values:
        print(f'La unidad de salida "{uf}" no se encuentra en la tabla.')
        return

    # Verificar si la unidad de entrada se encuentra en la primera columna
    if ui not in tabla[0].values:
        print(f'La unidad de entrada "{ui}" no se encuentra en la tabla.')
        return

    # Obtener la posición de la columna de salida
    posicion_salida = tabla.iloc[0].tolist().index(uf)

    # Buscar la fila correspondiente a la unidad de entrada
    fila_entrada = tabla[tabla[0] == ui]

    # Verificar si la unidad de entrada existe en la tabla
    if fila_entrada.empty:
        print(f'La unidad de entrada "{ui}" no se encuentra en la tabla.')
        return

    # Obtener el valor de conversión
    valor_conversion = fila_entrada.iloc[0, posicion_salida]

    # Calcular y mostrar el resultado de la conversión
    resultado = float(cantidad) * float(valor_conversion)

    # Verificar si el resultado es menor que 1
    if resultado < 1:
        resultado_formato = "{:.2e}".format(resultado)
    else:
        resultado_formato = round(resultado, 2)

    return(f'{cantidad} {ui} equivale a {resultado_formato} {uf}.')