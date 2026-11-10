import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string
from pathlib import Path

def procesar_psa(ruta_plantilla, ruta_psa):
    """
    Módulo Automatizado PSA
    Cruza los datos del archivo PSA.xlsx hacia la plantilla maestra.
    Reglas:
    - Col A (índice 0): Número de identificación del usuario.
    - Col B (índice 1): Fecha -> se inyecta en la columna BV.
    - Col C (índice 2): Resultado -> se inyecta en la columna DF.
    """
    if not Path(ruta_psa).exists():
        raise Exception(f"No se encontró el archivo PSA en: {ruta_psa}")

    # Leemos el archivo
    df_psa = pd.read_excel(ruta_psa)
    
    # Asumimos que la cédula está en la Columna A (índice 0)
    col_doc = df_psa.columns[0]
    
    # Eliminamos duplicados dejando el primer registro que encuentre por paciente
    df_unico = df_psa.drop_duplicates(subset=[col_doc], keep='first')

    # Cargamos la plantilla en memoria
    wb = openpyxl.load_workbook(ruta_plantilla)
    ws = wb.active

    # Creamos un diccionario para búsqueda súper rápida O(1)
    dict_psa = {}
    for _, row in df_unico.iterrows():
        doc_val = row.iloc[0]
        if pd.notnull(doc_val):
            # Limpieza segura del documento (evitar decimales tipo 123.0)
            try:
                doc = str(int(float(doc_val)))
            except:
                doc = str(doc_val).strip()
            
            # Guardamos los valores de la Columna B y C (validando que existan)
            dict_psa[doc] = {
                "val_bv": row.iloc[1] if len(row) > 1 else None, 
                "val_df": row.iloc[2] if len(row) > 2 else None  
            }

    # Iteramos sobre las filas de la plantilla para inyectar la información
    for r in range(2, ws.max_row + 1):
        cell_val = ws.cell(row=r, column=5).value # Columna E (Documento en la plantilla)
        
        if cell_val is not None:
            try:
                doc_plantilla = str(int(float(cell_val)))
            except:
                doc_plantilla = str(cell_val).strip()
            
            # Si el documento de la plantilla está en nuestro archivo PSA
            if doc_plantilla in dict_psa:
                datos = dict_psa[doc_plantilla]
                
                # ----------------------------------------------------
                # Asignamos Columna B a BV (Fecha)
                # ----------------------------------------------------
                if pd.notnull(datos["val_bv"]):
                    # Blindaje: Extraer solo la fecha sin la hora
                    if isinstance(datos["val_bv"], pd.Timestamp):
                        ws.cell(row=r, column=column_index_from_string("BV"), value=datos["val_bv"].date())
                    else:
                        # Forzamos la conversión a fecha por si viene como texto
                        parsed_fec = pd.to_datetime(str(datos["val_bv"]), errors='coerce')
                        if pd.notnull(parsed_fec):
                            ws.cell(row=r, column=column_index_from_string("BV"), value=parsed_fec.date())
                        else:
                            ws.cell(row=r, column=column_index_from_string("BV"), value=datos["val_bv"])
                
                # ----------------------------------------------------
                # Asignamos Columna C a DF (Resultado)
                # ----------------------------------------------------
                if pd.notnull(datos["val_df"]):
                    # Por si de casualidad el resultado también tiene formato de fecha
                    if isinstance(datos["val_df"], pd.Timestamp):
                        ws.cell(row=r, column=column_index_from_string("DF"), value=datos["val_df"].date())
                    else:
                        ws.cell(row=r, column=column_index_from_string("DF"), value=datos["val_df"])

    # Guardamos los cambios en el archivo físico
    wb.save(ruta_plantilla)
    wb.close()