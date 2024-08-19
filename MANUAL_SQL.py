import pyodbc
import os
import re
import xml.etree.ElementTree as ET
import datetime

connection_string = (

    "Driver={ODBC Driver 17 for SQL Server};"

    "Server=192.168.1.53;"

    "Database=Scharfstein;"

    "UID=sa;"

    "PWD=Vegam123;"

    "TrustServerCertificate=yes;"

)

conn = pyodbc.connect(connection_string)

cursor = conn.cursor()

ruta_caf = "C:\\CAF\\"

nombre_archivo = "FoliosSII761268763957590012024871652.xml"

estado = 2

script_dir = os.getcwd()  # DESARROLLO, distinta en converter.py

# Carpeta DESCARGAS
# xml_test = os.path.join(script_dir, 'caf.xml')

xml_test = os.path.join('C:\\CAF\\', nombre_archivo)

try:

    with open(xml_test, 'r', encoding='utf-8') as file:

        archivo_xml = file.read()

        # Eliminar declaración de codificación si existe
        archivo_xml = re.sub(r'<\?xml version="1\.0" encoding="[^"]*"\?>', '<?xml version="1.0"?>', archivo_xml)

        # Parsear el contenido XML
        root = ET.fromstring(archivo_xml)

        # Encontrar los valores entre las etiquetas <D> y <H>
        valor_desde = int(root.find('.//D').text)

        valor_hasta = int(root.find('.//H').text)

        valor_tipo_documento = int(root.find('.//TD').text)

        valor_rut = root.find('.//RE').text

        valor_razon_social = root.find('.//RS').text

        diferencia_folios = (valor_hasta - valor_desde) + 1

        sql = """
        INSERT INTO CAF (nombre_ruta, nombre_archivo, cod_estado, xml_archivo)
        VALUES (?, ?, ?, ?);
        """

        cursor.execute(sql, ruta_caf, nombre_archivo, estado, archivo_xml)

        conn.commit()

        print("Archivo XML insertado correctamente en CAF.")

        sql = """
        SELECT TOP 1 id FROM CAF ORDER BY 1 DESC;
        """
        cursor.execute(sql)

        caf_id = cursor.fetchone()[0]

        print(caf_id)

        conn.commit()

        print("CAF ID obtenido.")

        sql = """
        INSERT INTO datos_caf (caf_id, rut_emisor, razon_social, tipo_documento, desde, hasta, fecha, cod_estado)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """
        cursor.execute(sql, caf_id, valor_rut, valor_razon_social, valor_tipo_documento, valor_desde, valor_hasta, datetime.date.today(), 2 )

        conn.commit()

        print("Dato insertado correctamente en datos_caf.")

        sql = """
        INSERT INTO folios_asignados (caf_id, desde, hasta, saldo, cod_estado)
        VALUES (?, ?, ?, ?, ?)
        """
        cursor.execute(sql, caf_id, valor_desde, valor_hasta, diferencia_folios, 2 )

        conn.commit()

        print("Dato insertado correctamente en folios_asignados.")

except pyodbc.Error as ex:

    print(f"Error insertando el archivo XML: {ex}")

finally:

    cursor.close()

    conn.close()