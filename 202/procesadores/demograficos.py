import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string
from pathlib import Path
from utils.excel_helpers import limpiar_documento, formatear_fecha_date
from utils.texto_helpers import dividir_nombre_compuesto


def procesar_ajustes_demograficos(ruta_plantilla, ruta_produccion):
    """
    Módulo Automatizado AJUSTES DEMOGRÁFICOS
    Extrae desde INFORME PRODUCCION.xlsx:
      - Tipo de Identificación (Columna BO / índice 66) -> Col D
      - Nombres y Apellidos divididos (índice 10) -> Col F, G, H, I
      - Sexo (índice 11) -> Col K
      - Fecha de Nacimiento (índice 12) -> Col J
    """
    if not Path(ruta_produccion).exists():
        raise FileNotFoundError(f"No se encontró el INFORME PRODUCCION en: {ruta_produccion}")

    df_prod = pd.read_excel(ruta_produccion)
    col_doc = df_prod.columns[8]
    df_prod_unico = df_prod.drop_duplicates(subset=[col_doc], keep="first")

    wb = openpyxl.load_workbook(ruta_plantilla)
    ws = wb.active

    dict_demo = {}
    for _, row in df_prod_unico.iterrows():
        doc = limpiar_documento(row[col_doc])
        if doc:
            dict_demo[doc] = {
                "tipo_id": row.iloc[66],     # Columna BO (Índice 66)
                "nombre": row.iloc[10],      # Nombre completo
                "sexo": row.iloc[11],        # Sexo
                "fecha_nace": row.iloc[12],  # Fecha de nacimiento
            }

    col_d = column_index_from_string("D")
    col_f = column_index_from_string("F")
    col_g = column_index_from_string("G")
    col_h = column_index_from_string("H")
    col_i = column_index_from_string("I")
    col_j = column_index_from_string("J")
    col_k = column_index_from_string("K")

    for r in range(2, ws.max_row + 1):
        doc_plantilla = limpiar_documento(ws.cell(row=r, column=5).value)
        if doc_plantilla and doc_plantilla in dict_demo:
            datos = dict_demo[doc_plantilla]

            if pd.notnull(datos["tipo_id"]):
                ws.cell(row=r, column=col_d, value=str(datos["tipo_id"]).strip().upper())

            p_ape, s_ape, p_nom, s_nom = dividir_nombre_compuesto(datos["nombre"])
            ws.cell(row=r, column=col_f, value=p_ape)
            ws.cell(row=r, column=col_g, value=s_ape)
            ws.cell(row=r, column=col_h, value=p_nom)
            ws.cell(row=r, column=col_i, value=s_nom)

            if pd.notnull(datos["sexo"]):
                ws.cell(row=r, column=col_k, value=str(datos["sexo"]).strip())

            fec_nace = formatear_fecha_date(datos["fecha_nace"])
            if fec_nace:
                ws.cell(row=r, column=col_j, value=fec_nace)

    wb.save(ruta_plantilla)
    wb.close()
