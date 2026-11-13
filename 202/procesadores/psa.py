import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string
from pathlib import Path
from utils.excel_helpers import limpiar_documento, formatear_fecha_date


def procesar_psa(ruta_plantilla, ruta_psa):
    """
    Módulo Automatizado PSA
    Cruza los datos del archivo PSA.xlsx hacia la plantilla maestra para pacientes existentes.
    Reglas:
    - Col A (índice 0): Número de identificación del usuario.
    - Col B (índice 1): Fecha -> se inyecta en la columna BV.
    - Col C (índice 2): Resultado -> se inyecta en la columna DF.
    """
    if not Path(ruta_psa).exists():
        raise FileNotFoundError(f"No se encontró el archivo PSA en: {ruta_psa}")

    df_psa = pd.read_excel(ruta_psa)
    col_doc = df_psa.columns[0]
    df_unico = df_psa.drop_duplicates(subset=[col_doc], keep="first")

    wb = openpyxl.load_workbook(ruta_plantilla)
    ws = wb.active

    dict_psa = {}
    for _, row in df_unico.iterrows():
        doc = limpiar_documento(row.iloc[0])
        if doc:
            dict_psa[doc] = {
                "val_bv": row.iloc[1] if len(row) > 1 else None,
                "val_df": row.iloc[2] if len(row) > 2 else None,
            }

    col_bv = column_index_from_string("BV")
    col_df = column_index_from_string("DF")

    for r in range(2, ws.max_row + 1):
        doc_plantilla = limpiar_documento(ws.cell(row=r, column=5).value)
        if doc_plantilla and doc_plantilla in dict_psa:
            datos = dict_psa[doc_plantilla]

            if pd.notnull(datos["val_bv"]):
                fec_date = formatear_fecha_date(datos["val_bv"])
                ws.cell(
                    row=r,
                    column=col_bv,
                    value=fec_date if fec_date is not None else datos["val_bv"],
                )

            if pd.notnull(datos["val_df"]):
                if isinstance(datos["val_df"], pd.Timestamp):
                    ws.cell(row=r, column=col_df, value=datos["val_df"].date())
                else:
                    ws.cell(row=r, column=col_df, value=datos["val_df"])

    wb.save(ruta_plantilla)
    wb.close()
