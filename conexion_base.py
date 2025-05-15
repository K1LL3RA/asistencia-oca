# conexion.py
# Establece conexión con SQL Server usando pyodbc
import pyodbc

def obtener_conexion():
    try:
        connection = pyodbc.connect(
            'DRIVER={ODBC Driver 17 for SQL Server};'
            'SERVER=98.84.221.29,1434;'
            'DATABASE=SafeSmart_Checklist;'
            'UID=oterrazas;'
            'PWD=oca@2025'
        )
        return connection
    except pyodbc.Error as e:
        print(f"Error de conexión: {e}")
        return None
