import os
import csv
import pandas as pd
import re
from datetime import datetime
from openpyxl import Workbook

# Definir el nombre del archivo CSV para almacenar los datos
archivo_csv = "notas_servicio.csv"

# Verificar si el archivo CSV existe y cargar los datos si está presente
if os.path.exists(archivo_csv):
    df = pd.read_csv(archivo_csv, parse_dates=["Fecha"])
else:
    df = pd.DataFrame(columns=["Folio", "Fecha", "Cliente", "RFC", "Correo", "Monto", "Detalle"])

# Función para generar un folio único
def generar_folio():
    folio = len(df) + 1
    return folio

# Función para validar el formato de RFC
def validar_rfc(rfc):
    # Expresión regular para validar RFC (simplificada)
    rfc_pattern = r'^[A-Z&Ñ]{3,4}[\d]{6}[A-V1-9][A-Z1-9][0-9A]'
    return re.match(rfc_pattern, rfc)

# Función para validar el formato de correo electrónico
def validar_correo(correo):
    # Expresión regular para validar correo electrónico
    correo_pattern = r'^[\w\.-]+@[\w\.-]+$'
    return re.match(correo_pattern, correo)

# Función para calcular el monto total de una nota
def calcular_monto(detalle):
    total = sum(item["Costo"] for item in detalle)
    return round(total, 2)

# Función para registrar una nueva nota
def registrar_nota():
    folio = generar_folio()
    while True:
        fecha_str = input("Ingrese la fecha (YYYY-MM-DD): ")
        try:
            fecha = datetime.strptime(fecha_str, "%Y-%m-%d")
        except ValueError:
            print("Fecha ingresada no válida. Formato válido: YYYY-MM-DD")
            continue

        if fecha > datetime.now():
            print("La fecha no puede ser posterior a la fecha actual.")
        else:
            break

    cliente = input("Ingrese el nombre del cliente: ")
    rfc = input("Ingrese el RFC del cliente: ")
    correo = input("Ingrese el correo electrónico del cliente: ")

    # Validar el formato de RFC y correo
    if not validar_rfc(rfc):
        print("El RFC ingresado no tiene un formato válido.")
        return

    if not validar_correo(correo):
        print("El correo electrónico ingresado no tiene un formato válido.")
        return

    detalle = []
    while True:
        nombre_servicio = input("Ingrese el nombre del servicio (o 'fin' para finalizar): ")
        if nombre_servicio.lower() == "fin":
            break
        costo_servicio = float(input("Ingrese el costo del servicio: "))
        if costo_servicio <= 0:
            print("El costo del servicio debe ser mayor que cero.")
            continue
        detalle.append({"Nombre": nombre_servicio, "Costo": costo_servicio})

    monto = calcular_monto(detalle)
    df.loc[len(df)] = [folio, fecha, cliente, rfc, correo, monto, detalle]
    df.to_csv(archivo_csv, index=False)
    print("Nota registrada con éxito.")

