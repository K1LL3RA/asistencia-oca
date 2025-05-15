# config.py
# Archivo de configuración centralizada para las variables del sistema
# Datos de conexión a la base de datos SQL Server
import os
server = '98.84.221.29'
database = 'SafeSmart_Checklist'
username = 'jarbildo'
password = 'oca$2025$'

# Ruta fija a la plantilla de Excel base (puedes hacerla seleccionable desde la GUI si deseas)
plantilla_path = os.path.join(os.path.dirname(__file__), "base.xlsx")#CAMBIARLO SEGUN SEA NESARIO