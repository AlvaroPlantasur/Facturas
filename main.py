import os
import psycopg2
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from datetime import datetime
import sys

def main():
    db_params = {
        'dbname': os.environ.get('DB_NAME'),
        'user': os.environ.get('DB_USER'),
        'password': os.environ.get('DB_PASSWORD'),
        'host': os.environ.get('DB_HOST'),
        'port': os.environ.get('DB_PORT'),
        'sslmode': 'require'
    }
    file_path = os.environ.get('EXCEL_FILE_PATH')

    fecha_inicio_str = '2025-01-01'
    fecha_fin = datetime.now().date().strftime('%Y-%m-%d')

    print(f"Rango de fechas para la consulta: Desde {fecha_inicio_str} hasta {fecha_fin}")

    # Tu consulta SQL completa aquí (omitida por brevedad)
    query = """ ... """  # Usa la misma consulta SQL que ya tienes

    # Ejecutar la consulta
    try:
        with psycopg2.connect(**db_params) as conn:
            with conn.cursor() as cur:
                cur.execute(query)
                resultados = cur.fetchall()
                headers = [desc[0] for desc in cur.description]
    except Exception as e:
        print(f"Error al conectar o ejecutar la consulta: {e}")
        sys.exit(1)

    if not resultados:
        print("No se obtuvieron resultados.")
        return
    else:
        print(f"Se obtuvieron {len(resultados)} filas.")

    # Cargar el archivo Excel
    try:
        book = load_workbook(file_path)
        sheet = book.active
    except FileNotFoundError:
        print(f"No se encontró el archivo '{file_path}'.")
        return

    # Limpiar completamente la hoja
    sheet.delete_rows(1, sheet.max_row)

    # Escribir los encabezados
    sheet.append(headers)

    # Escribir todos los datos
    for row in resultados:
        sheet.append(row)

    print(f"Se sobrescribieron {len(resultados)} filas.")

    # Actualizar la tabla si existe
    if "Lineas2025" in sheet.tables:
        tabla = sheet.tables["Lineas2025"]
        last_col_letter = get_column_letter(len(headers))
        max_row = sheet.max_row
        tabla.ref = f"A1:{last_col_letter}{max_row}"
        print(f"Tabla 'Lineas2025' actualizada a rango: A1:{last_col_letter}{max_row}")
    else:
        print("No se encontró la tabla 'Lineas2025'.")

    # Guardar cambios
    try:
        book.save(file_path)
        print(f"Archivo guardado con los datos actualizados en '{file_path}'.")
    except Exception as e:
        print(f"Error al guardar el archivo '{file_path}': {e}")

if __name__ == '__main__':
    main()
