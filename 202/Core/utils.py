import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string
from pathlib import Path  # <-- Esto es vital para que no dé error de lectura

def procesar_informe_hc(ruta_plantilla, ruta_hc):
    """
    Paso 5: Módulo Automatizado INFORME HC
    Cruza el Peso, Talla y sus Fechas.
    """
    if not Path(ruta_hc).exists():
        raise Exception(f"No se encontró el archivo INFORME HC en: {ruta_hc}")

    df_hc = pd.read_excel(ruta_hc)
    
    # Identificar columna de documento en HC
    col_doc = df_hc.columns[1]

    # Agrupar por documento (tomando el registro más reciente usando la Fecha HC)
    df_hc = df_hc.sort_values(by=df_hc.columns[22], ascending=False)
    df_hc_unico = df_hc.drop_duplicates(subset=[col_doc], keep='first')

    wb = openpyxl.load_workbook(ruta_plantilla)
    ws = wb.active

    idx_peso = 23      # X -> Peso
    idx_talla = 66     # BO -> qm1_talla_cm
    idx_fecha_hc = 22  # W -> Fecha HC

    dict_hc = {}
    for _, row in df_hc_unico.iterrows():
        doc = str(row[col_doc]).split('.')[0].strip()
        dict_hc[doc] = {
            "peso": row.iloc[idx_peso],
            "talla": row.iloc[idx_talla],
            "fecha": row.iloc[idx_fecha_hc]
        }

    for r in range(2, ws.max_row + 1):
        doc_plantilla = str(ws.cell(row=r, column=5).value).split('.')[0].strip()
        
        if doc_plantilla in dict_hc:
            datos = dict_hc[doc_plantilla]
            
            fec_str = ""
            if pd.notnull(datos["fecha"]):
                fec_str = pd.to_datetime(str(datos["fecha"]), errors='coerce').strftime('%Y/%m/%d')
                if fec_str == "NaT": fec_str = ""

            if pd.notnull(datos["peso"]):
                ws.cell(row=r, column=column_index_from_string("AE"), value=datos["peso"])
                ws.cell(row=r, column=column_index_from_string("AD"), value=fec_str)
            
            if pd.notnull(datos["talla"]):
                ws.cell(row=r, column=column_index_from_string("AG"), value=datos["talla"])
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
        doc = str(row[col_doc]).split('.')[0].strip()
        dict_demo[doc] = {
            "nombre": row.iloc[10],
            "sexo": row.iloc[11],
            "fecha_nace": row.iloc[12]
        }

    for r in range(2, ws.max_row + 1):
        doc_plantilla = str(ws.cell(row=r, column=5).value).split('.')[0].strip()
        
        if doc_plantilla in dict_demo:
            datos = dict_demo[doc_plantilla]
            
            p_ape, s_ape, p_nom, s_nom = dividir_nombre_compuesto(datos["nombre"])
            
            ws.cell(row=r, column=column_index_from_string("F"), value=p_ape)
            ws.cell(row=r, column=column_index_from_string("G"), value=s_ape)
            ws.cell(row=r, column=column_index_from_string("H"), value=p_nom)
            ws.cell(row=r, column=column_index_from_string("I"), value=s_nom)

            if pd.notnull(datos["sexo"]):
                ws.cell(row=r, column=column_index_from_string("K"), value=str(datos["sexo"]).strip())

            if pd.notnull(datos["fecha_nace"]):
                fec_nace_str = pd.to_datetime(str(datos["fecha_nace"]), errors='coerce').strftime('%Y/%m/%d')
                if fec_nace_str != "NaT":
                    ws.cell(row=r, column=column_index_from_string("J"), value=fec_nace_str)

    wb.save(ruta_plantilla)
    wb.close()