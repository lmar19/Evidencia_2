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

# Función para consultar por período
def consultar_por_periodo():
    fecha_inicial = input("Ingrese la fecha inicial (YYYY-MM-DD, o presione Enter para utilizar la fecha predeterminada '2000-01-01'): ")
    fecha_final = input("Ingrese la fecha final (YYYY-MM-DD, o presione Enter para utilizar la fecha actual): ")

    # Establecer las fechas predeterminadas si el usuario las omite
    if fecha_inicial == "":
        fecha_inicial = datetime(2000, 1, 1)
        print("Fecha inicial predeterminada: 2000-01-01")
    else:
        fecha_inicial = datetime.strptime(fecha_inicial, "%Y-%m-%d")

    if fecha_final == "":
        fecha_final = datetime.now()
        print(f"Fecha final predeterminada: {fecha_final.strftime('%Y-%m-%d')}")
    else:
        fecha_final = datetime.strptime(fecha_final, "%Y-%m-%d")

    # Validar las fechas
    if fecha_final < fecha_inicial:
        print("La fecha final debe ser igual o posterior a la fecha inicial.")
        return

    notas_periodo = df[(df["Fecha"] >= fecha_inicial) & (df["Fecha"] <= fecha_final)]

    if notas_periodo.empty:
        print("No hay notas emitidas para el período especificado.")
    else:
        promedio_monto = notas_periodo["Monto"].mean()
        print("Notas del período:")
        print(notas_periodo[["Folio", "Fecha", "Cliente"]])
        print(f"Monto promedio de notas del período: ${round(promedio_monto, 2)}")

# Función para consultar por folio
def consultar_por_folio():
    folio_consulta = input("Ingrese el folio de la nota a consultar: ")
    folio_consulta = int(folio_consulta)

    nota = df[df["Folio"] == folio_consulta]

    if nota.empty:
        print("La nota no existe o está cancelada.")
    else:
        print("Datos de la nota:")
        print(nota[["Folio", "Fecha", "Cliente", "RFC", "Correo", "Monto"]])
        print("Detalle de la nota:")
        print(nota["Detalle"].iloc[0])

# Función para consultar por cliente
def consultar_por_cliente():
    rfc_clientes = df["RFC"].unique()
    rfc_clientes.sort()

    for i, rfc in enumerate(rfc_clientes, start=1):
        print(f"{i}. RFC: {rfc}")

    seleccion = int(input("Seleccione el número del cliente a consultar: "))

    if seleccion < 1 or seleccion > len(rfc_clientes):
        print("Selección no válida.")
        return

    rfc_seleccionado = rfc_clientes[seleccion - 1]

    notas_cliente = df[df["RFC"] == rfc_seleccionado]

    if notas_cliente.empty:
        print("No hay notas emitidas para este cliente.")
    else:
        promedio_monto = notas_cliente["Monto"].mean()
        print("Notas del cliente:")
        print(notas_cliente[["Folio", "Fecha", "Cliente"]])
        print(f"Monto promedio de notas del cliente: ${round(promedio_monto, 2)}")

        exportar_excel = input("¿Desea exportar esta información a un archivo de Excel? (S/N): ")

        if exportar_excel.lower() == "s":
            exportar_a_excel(notas_cliente, rfc_seleccionado)

# Función para exportar notas de un cliente a un archivo Excel
def exportar_a_excel(notas_cliente, rfc):
    wb = Workbook()
    ws = wb.active
    ws.title = "Notas de Cliente"

    # Encabezados
    ws.append(["Folio", "Fecha", "Cliente", "Monto"])
    for _, row in notas_cliente.iterrows():
        ws.append([row["Folio"], row["Fecha"].strftime("%Y-%m-%d"), row["Cliente"], row["Monto"]])

    # Guardar el archivo Excel
    fecha_emision = datetime.now().strftime("%Y-%m-%d")
    archivo_excel = f"{rfc}_{fecha_emision}.xlsx"
    wb.save(archivo_excel)
    print(f"Archivo Excel guardado como '{archivo_excel}'.")
    # Función para cancelar una nota
def cancelar_nota():
    folio_cancelar = input("Ingrese el folio de la nota a cancelar: ")
    folio_cancelar = int(folio_cancelar)

    nota = df[df["Folio"] == folio_cancelar]

    if nota.empty:
        print("La nota no existe o ya está cancelada.")
    else:
        print("Datos de la nota a cancelar:")
        print(nota[["Folio", "Fecha", "Cliente", "RFC", "Correo", "Monto"]])
        print("Detalle de la nota:")
        print(nota["Detalle"].iloc[0])

        confirmar = input("¿Está seguro de que desea cancelar esta nota? (S/N): ")

        if confirmar.lower() == "s":
            # Marcar la nota como cancelada
            df.loc[df["Folio"] == folio_cancelar, "Fecha"] = None
            df.to_csv(archivo_csv, index=False)
            print("Nota cancelada con éxito.")
        else:
            print("Operación de cancelación cancelada.")
# Función para recuperar una nota
def recuperar_nota():
    notas_canceladas = df[df["Fecha"].isna()]

    if notas_canceladas.empty:
        print("No hay notas canceladas para recuperar.")
    else:
        print("Notas canceladas:")
        print(notas_canceladas[["Folio", "Cliente"]])

        folio_recuperar = input("Ingrese el folio de la nota a recuperar o '0' para cancelar: ")
        folio_recuperar = int(folio_recuperar)

        if folio_recuperar == 0:
            print("Operación de recuperación cancelada.")
        else:
            nota = notas_canceladas[notas_canceladas["Folio"] == folio_recuperar]

            if nota.empty:
                print("La nota no existe o no está cancelada.")
            else:
                print("Datos de la nota a recuperar:")
                print(nota[["Folio", "Fecha", "Cliente", "RFC", "Correo", "Monto"]])
                print("Detalle de la nota:")
                print(nota["Detalle"].iloc[0])

                confirmar = input("¿Está seguro de que desea recuperar esta nota? (S/N): ")

                if confirmar.lower() == "s":
                    # Restaurar la fecha de la nota (recuperarla)
                    df.loc[df["Folio"] == folio_recuperar, "Fecha"] = datetime.now()
                    df.to_csv(archivo_csv, index=False)
                    print("Nota recuperada con éxito.")
                else:
                    print("Operación de recuperación cancelada.")

# Menú principal
while True:
    print("\n===== Menú Principal =====")
    print("1. Registrar una nota")
    print("2. Consultas y reportes")
    print("3. Cancelar una nota")
    print("4. Recuperar una nota")
    print("5. Salir")

    opcion_menu = input("Seleccione una opción (1/2/3/4/5): ")

    if opcion_menu == "1":
        registrar_nota()
    elif opcion_menu == "2":
        while True:
            print("\n===== Consultas y Reportes =====")
            print("1. Consulta por período")
            print("2. Consulta por folio")
            print("3. Consulta por cliente")
            print("4. Volver al menú principal")

            opcion_consulta = input("Seleccione una opción (1/2/3/4): ")

            if opcion_consulta == "1":
                consultar_por_periodo()
            elif opcion_consulta == "2":
                consultar_por_folio()
            elif opcion_consulta == "3":
                consultar_por_cliente()
            elif opcion_consulta == "4":
                break
            else:
                print("Opción no válida.")
    elif opcion_menu == "3":
        cancelar_nota()
    elif opcion_menu == "4":
        recuperar_nota()
    elif opcion_menu == "5":
        confirmar_salida = input("¿Está seguro de que desea salir? (S/N): ")
        if confirmar_salida.lower() == "s":
            break
    else:
        print("Opción no válida.")
