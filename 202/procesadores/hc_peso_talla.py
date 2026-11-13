import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string
from pathlib import Path
from utils.excel_helpers import limpiar_documento, formatear_fecha_date


def procesar_informe_hc(ruta_plantilla, ruta_hc):
    """
    Módulo Automatizado INFORME HC (Pesos y Tallas)
    Cruza el Peso (AE), Fecha Peso (AD), Talla (AG) y Fecha Talla (AF)
    tomando el registro más reciente por paciente según Fecha HC.
    """
    if not Path(ruta_hc).exists():
        raise FileNotFoundError(f"No se encontró el archivo INFORME HC en: {ruta_hc}")

    df_hc = pd.read_excel(ruta_hc)

    col_doc = df_hc.columns[0]       # Codigo_KC
    col_fecha_hc = df_hc.columns[1]  # Fecha HC
    col_peso = df_hc.columns[2]      # Peso
    col_talla = df_hc.columns[3]     # qm1_talla_cm

    df_hc = df_hc.sort_values(by=col_fecha_hc, ascending=False)
    df_hc_unico = df_hc.drop_duplicates(subset=[col_doc], keep="first")

    wb = openpyxl.load_workbook(ruta_plantilla)
    ws = wb.active

    dict_hc = {}
    for _, row in df_hc_unico.iterrows():
        doc = limpiar_documento(row[col_doc])
        if doc:
            dict_hc[doc] = {
                "peso": row[col_peso],
                "talla": row[col_talla],
                "fecha": row[col_fecha_hc],
            }

    col_ad = column_index_from_string("AD")
    col_ae = column_index_from_string("AE")
    col_af = column_index_from_string("AF")
    col_ag = column_index_from_string("AG")

    for r in range(2, ws.max_row + 1):
        doc_plantilla = limpiar_documento(ws.cell(row=r, column=5).value)
        if doc_plantilla and doc_plantilla in dict_hc:
            datos = dict_hc[doc_plantilla]
            fec_date = formatear_fecha_date(datos["fecha"])

            if pd.notnull(datos["peso"]):
                ws.cell(row=r, column=col_ae, value=datos["peso"])
                if fec_date:
                    ws.cell(row=r, column=col_ad, value=fec_date)

            if pd.notnull(datos["talla"]):
                ws.cell(row=r, column=col_ag, value=datos["talla"])
                if fec_date:
                    ws.cell(row=r, column=col_af, value=fec_date)

    wb.save(ruta_plantilla)
    wb.close()
