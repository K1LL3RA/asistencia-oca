import pandas as pd

def ejecutar_consulta(connection, checklist_id):
    """
    Ejecuta el procedimiento almacenado GetCheckListData
    con el ID ingresado y devuelve los datos en un DataFrame.
    """
    try:
        query = f"EXEC dbo.GetCheckListData @CheckListId = {checklist_id}"
        cursor = connection.cursor()
        cursor.execute(query)
        columns = [col[0] for col in cursor.description]
        rows = cursor.fetchall()
        cursor.close()
        return pd.DataFrame.from_records(rows, columns=columns)
    except Exception as e:
        print(f"Error en consulta: {e}")
        return None
