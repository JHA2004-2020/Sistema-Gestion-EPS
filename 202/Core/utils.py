import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string
from pathlib import Path  # <-- Esto es vital para que no dé error de lectura

def procesar_informe_hc(ruta_plantilla, ruta_hc):
    """
    Paso 5: Módulo Automatizado INFORME HC
    Cruza el Peso, Talla y sus Fechas usando la nueva estructura del archivo de Peso y Talla.
    """
    if not Path(ruta_hc).exists():
        raise Exception(f"No se encontró el archivo INFORME HC en: {ruta_hc}")

    # Leemos el archivo nuevo de Peso y Talla
    df_hc = pd.read_excel(ruta_hc)
    
    # Identificamos las columnas basándonos en la nueva estructura (índices 0, 1, 2, 3)
    col_doc = df_hc.columns[0]      # Codigo_KC
    col_fecha_hc = df_hc.columns[1] # Fecha HC
    col_peso = df_hc.columns[2]     # Peso
    col_talla = df_hc.columns[3]    # qm1_talla_cm

    # Agrupar por documento (tomando el registro más reciente usando la Fecha HC)
    df_hc = df_hc.sort_values(by=col_fecha_hc, ascending=False)
    df_hc_unico = df_hc.drop_duplicates(subset=[col_doc], keep='first')

    wb = openpyxl.load_workbook(ruta_plantilla)
    ws = wb.active

    dict_hc = {}
    for _, row in df_hc_unico.iterrows():
        doc_val = row[col_doc]
        if pd.notnull(doc_val):
            # Limpieza segura del documento (evitar decimales)
            try:
                doc = str(int(float(doc_val)))
            except:
                doc = str(doc_val).strip()
                
            dict_hc[doc] = {
                "peso": row[col_peso],
                "talla": row[col_talla],
                "fecha": row[col_fecha_hc]
            }

    for r in range(2, ws.max_row + 1):
        cell_val = ws.cell(row=r, column=5).value
        if cell_val is not None:
            try:
                doc_plantilla = str(int(float(cell_val)))
            except:
                doc_plantilla = str(cell_val).strip()
            
            if doc_plantilla in dict_hc:
                datos = dict_hc[doc_plantilla]
                
                fec_str = ""
                if pd.notnull(datos["fecha"]):
                    # Usamos .date() para enviar la fecha pura a Excel y evitar horas adicionales
                    parsed_date = pd.to_datetime(str(datos["fecha"]), errors='coerce')
                    if pd.notnull(parsed_date):
                        fec_str = parsed_date.date()

                if pd.notnull(datos["peso"]):
                    ws.cell(row=r, column=column_index_from_string("AE"), value=datos["peso"])
                    if fec_str:
                        ws.cell(row=r, column=column_index_from_string("AD"), value=fec_str)
                
                if pd.notnull(datos["talla"]):
                    ws.cell(row=r, column=column_index_from_string("AG"), value=datos["talla"])
                    if fec_str:
                        ws.cell(row=r, column=column_index_from_string("AF"), value=fec_str)

    wb.save(ruta_plantilla)
    wb.close()


def dividir_nombre_compuesto(nombre_completo):
    """
    Divide un nombre completo colombiano inteligentemente respetando conectores.
    """
    if pd.isnull(nombre_completo):
        return "", "", "", ""
    
    palabras = str(nombre_completo).strip().upper().split()
    conectores = {"DE", "DEL", "LA", "LAS", "LOS", "SAN", "SANTA", "MAC", "MC", "VAN", "VON", "Y"}
    
    partes_agrupadas = []
    i = 0
    while i < len(palabras):
        parte = palabras[i]
        while i < len(palabras) - 1 and palabras[i] in conectores:
            i += 1
            parte += " " + palabras[i]
        partes_agrupadas.append(parte)
        i += 1

    p_ape, s_ape, p_nom, s_nom = "", "", "", ""
    
    if len(partes_agrupadas) == 1:
        p_nom = partes_agrupadas[0]
    elif len(partes_agrupadas) == 2:
        p_ape = partes_agrupadas[0]
        p_nom = partes_agrupadas[1]
    elif len(partes_agrupadas) == 3:
        p_ape = partes_agrupadas[0]
        s_ape = partes_agrupadas[1]
        p_nom = partes_agrupadas[2]
    elif len(partes_agrupadas) >= 4:
        p_ape = partes_agrupadas[0]
        s_ape = partes_agrupadas[1]
        p_nom = partes_agrupadas[2]
        s_nom = " ".join(partes_agrupadas[3:])
        
    return p_ape, s_ape, p_nom, s_nom


def procesar_ajustes_demograficos(ruta_plantilla, ruta_produccion):
    """
    Paso 6: Módulo Automatizado AJUSTES ADICIONALES (Nombres y Demografía)
    Extrae Nombres, Sexo, Fecha de Nacimiento y Tipo de Identificación.
    """
    if not Path(ruta_produccion).exists():
        raise Exception(f"No se encontró el INFORME PRODUCCION en: {ruta_produccion}")

    df_prod = pd.read_excel(ruta_produccion)
    col_doc = df_prod.columns[8]
    
    df_prod_unico = df_prod.drop_duplicates(subset=[col_doc], keep='first')

    wb = openpyxl.load_workbook(ruta_plantilla)
    ws = wb.active

    dict_demo = {}
    for _, row in df_prod_unico.iterrows():
        doc_val = row[col_doc]
        if pd.notnull(doc_val):
            try:
                doc = str(int(float(doc_val)))
            except:
                doc = str(doc_val).strip()
                
            dict_demo[doc] = {
                "tipo_id": row.iloc[66],  # Columna BO (Índice 66)
                "nombre": row.iloc[10],
                "sexo": row.iloc[11],
                "fecha_nace": row.iloc[12]
            }

    for r in range(2, ws.max_row + 1):
        cell_val = ws.cell(row=r, column=5).value
        if cell_val is not None:
            try:
                doc_plantilla = str(int(float(cell_val)))
            except:
                doc_plantilla = str(cell_val).strip()
            
            if doc_plantilla in dict_demo:
                datos = dict_demo[doc_plantilla]
                
                # --- TIPO DE IDENTIFICACIÓN ---
                if pd.notnull(datos["tipo_id"]):
                    ws.cell(row=r, column=column_index_from_string("D"), value=str(datos["tipo_id"]).strip().upper())
                
                # --- DIVISIÓN DE NOMBRES ---
                p_ape, s_ape, p_nom, s_nom = dividir_nombre_compuesto(datos["nombre"])
                
                ws.cell(row=r, column=column_index_from_string("F"), value=p_ape)
                ws.cell(row=r, column=column_index_from_string("G"), value=s_ape)
                ws.cell(row=r, column=column_index_from_string("H"), value=p_nom)
                ws.cell(row=r, column=column_index_from_string("I"), value=s_nom)

                # --- SEXO ---
                if pd.notnull(datos["sexo"]):
                    ws.cell(row=r, column=column_index_from_string("K"), value=str(datos["sexo"]).strip())

                # --- FECHA DE NACIMIENTO ---
                if pd.notnull(datos["fecha_nace"]):
                    parsed_fec_nace = pd.to_datetime(str(datos["fecha_nace"]), errors='coerce')
                    if pd.notnull(parsed_fec_nace):
                        ws.cell(row=r, column=column_index_from_string("J"), value=parsed_fec_nace.date())

    wb.save(ruta_plantilla)
    wb.close()