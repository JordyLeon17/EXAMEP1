import os
import pyodbc
import csv

# Dirección completa de la carpeta donde están los archivos CSV
ruta_archivos_csv = r"C:\Users\jordy\Desktop\ExamenP1"

# Establecer conexión con SQL Server
conexion = pyodbc.connect(
    'Driver={SQL Server};'
    'Server=ASUS-ROG-SCAR-I;'
    'Database=Ventas;'
    'Trusted_Connection=yes;'
)

# Verificar si la conexión fue exitosa
if conexion:
    print("Conexión exitosa a la base de datos")

    # Crear cursor para ejecutar consultas
    cursor = conexion.cursor()

    # Iterar sobre los archivos CSV y cargar los datos a la base de datos
    for nombre_archivo in os.listdir(ruta_archivos_csv):
        if nombre_archivo.endswith('.csv'):
            numero_local = int(nombre_archivo.replace('Local', '').replace('.csv', ''))
            ruta_completa_archivo = os.path.join(ruta_archivos_csv, nombre_archivo)
            with open(ruta_completa_archivo, newline='') as archivo_csv:
                lector_csv = csv.reader(archivo_csv)
                # Omitir la primera fila (encabezados)
                next(lector_csv)
                # Insertar los datos en la tabla
                for fila in lector_csv:
                    # Verificar la longitud de la fila
                    if len(fila) != 8:
                        print(f"Error: Longitud de fila incorrecta en el archivo {nombre_archivo}: {fila}")
                        continue
                    # Agregar el número de local (IdLocal) al principio de la fila
                    fila_con_local = [numero_local] + fila[:-1]  # Excluimos el último elemento (Producto)
                    # Eliminar comillas dobles de Producto si existen
                    fila_con_local[4] = fila_con_local[4].replace('"', '')  # Índice 4 es la columna Producto
                    # Ejecutar la consulta SQL
                    cursor.execute("INSERT INTO Ventas_Consolidadas (IdLocal, IdTransaccion, Fecha, IdCategoria, IdProducto, Cantidad, PrecioUnitario, TotalVenta) VALUES (?, ?, ?, ?, ?, ?, ?, ?)", fila_con_local)

    # Hacer commit para guardar los cambios en la base de datos
    conexion.commit()

    # Cerrar cursor y conexión
    cursor.close()
    conexion.close()

    print("Los datos se han cargado correctamente en la base de datos")
else:
    print("No se pudo establecer conexión a la base de datos")
