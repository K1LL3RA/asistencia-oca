# conexion.py
# Establece conexión con SQL Server usando pyodbc
import pytds

def obtener_conexion():
    try:
        conn = pytds.connect(
            server='98.84.221.29,1434',
            database='SafeSmart_Checklist',
            user='jarbildo',
            password='oca@2025',
            port=1433  # Asegúrate de que tu SQL Server acepte conexiones remotas en este puerto
        )
        return conn
    except Exception as e:
        print("Error al conectar con pytds:", e)
        return None
