import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string
from pathlib import Path
from utils.excel_helpers import (
    limpiar_documento,
    formatear_fecha_str,
    indexar_documentos_plantilla,
    obtener_o_crear_fila,
)


def procesar_cop_generico(nombre_eps, ruta_plantilla_in, ruta_plantilla_out, ruta_cop=None):
    """
    Procesa el archivo COP.xlsx según la EPS indicada ('NUEVA EPS' o 'SAVIA').
    Mapeo:
      Col C (2) 'Número de identificación del usuario' -> Col E (Plantilla)
      Col L (11) 'Fecha de atención en salud bucal...' -> Col BY (Plantilla)
      Col S (18) 'CALCULO AUTOMATICO VALOR A TRANSFERIR' -> Col CY (Plantilla)
    """
    if ruta_cop is None:
        base_dir = Path(__file__).resolve().parent.parent
        ruta_cop = base_dir / "Insumos" / "COP.xlsx"
    else:
        ruta_cop = Path(ruta_cop)

    if not ruta_cop.exists():
        raise FileNotFoundError("No se encontró el archivo COP.xlsx en Insumos.")

    try:
        df_cop = pd.read_excel(ruta_cop, sheet_name=nombre_eps, skiprows=12, dtype=str)
    except Exception as e:
        raise ValueError(f"Error al leer la hoja '{nombre_eps}' en COP.xlsx: {e}")

    df_cop = df_cop.dropna(subset=[df_cop.columns[2]])

    wb = openpyxl.load_workbook(ruta_plantilla_in)
    ws = wb.active

    dict_cop = {}
    for _, row in df_cop.iterrows():
        doc = limpiar_documento(row.iloc[2])
        if doc:
            calc_val = row.iloc[18]
            calc_final = (
                str(calc_val).strip()
                if pd.notnull(calc_val) and str(calc_val).strip().lower() != "nan"
                else None
            )

            dict_cop[doc] = {
                "fecha_bucal": row.iloc[11],
                "calculo": calc_final,
            }

    indice_docs, fila_destino_nueva = indexar_documentos_plantilla(ws)
    col_by = column_index_from_string("BY")
    col_cy = column_index_from_string("CY")

    for doc, datos in dict_cop.items():
        fila_destino, fila_destino_nueva = obtener_o_crear_fila(
            ws, indice_docs, fila_destino_nueva, doc
        )

        fec_str = formatear_fecha_str(datos["fecha_bucal"])
        if fec_str:
            ws.cell(row=fila_destino, column=col_by, value=fec_str)

        if datos["calculo"] is not None:
            ws.cell(row=fila_destino, column=col_cy, value=f"{datos['calculo']}")

    wb.save(ruta_plantilla_out)
    wb.close()
