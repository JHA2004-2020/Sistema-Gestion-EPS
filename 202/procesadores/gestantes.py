import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string
from pathlib import Path
from utils.excel_helpers import (
    limpiar_documento,
    formatear_fecha_str,
    formatear_fecha_date,
    indexar_documentos_plantilla,
    obtener_o_crear_fila,
)


def procesar_gestantes_nueva_eps(ruta_plantilla, ruta_salida, ruta_gestantes=None):
    """
    Procesa GESTANTES.xlsx filtrando por NUEVA EPS.
    Mapeo:
      Codigo_KC -> Col E (5) (Busca si existe o crea fila nueva para no duplicar)
      FPP -> Col AH (34)
      Fecha HC -> Col BG (59)
      Valores fijos = 1 en columnas O (15), BH (60), BI (61), BJ (62)
    """
    if ruta_gestantes is None:
        base_dir = Path(__file__).resolve().parent.parent
        archivo_gestantes = base_dir / "Insumos" / "GESTANTES.xlsx"
    else:
        archivo_gestantes = Path(ruta_gestantes)

    if not archivo_gestantes.exists():
        raise FileNotFoundError("No se encontró el archivo GESTANTES.xlsx en Insumos.")

    df_gestantes = pd.read_excel(archivo_gestantes, sheet_name=0)

    if "NOMBRE ENTIDAD" in df_gestantes.columns:
        df_gestantes["NOMBRE ENTIDAD"] = df_gestantes["NOMBRE ENTIDAD"].astype(str).str.upper()
        df_filtrado = df_gestantes[
            df_gestantes["NOMBRE ENTIDAD"].str.contains("NUEVA EPS", na=False)
        ].copy()
    else:
        df_filtrado = df_gestantes.copy()

    if "Codigo_KC" in df_filtrado.columns:
        df_filtrado = df_filtrado.dropna(subset=["Codigo_KC"])

    wb = openpyxl.load_workbook(ruta_plantilla)
    ws = wb.active

    indice_docs, fila_destino_nueva = indexar_documentos_plantilla(ws)

    for _, row in df_filtrado.reset_index(drop=True).iterrows():
        doc = limpiar_documento(row.get("Codigo_KC", ""))
        if not doc:
            continue

        fila_actual, fila_destino_nueva = obtener_o_crear_fila(
            ws, indice_docs, fila_destino_nueva, doc
        )

        fpp = formatear_fecha_str(row.get("FPP", "")) or ""
        fecha_hc = formatear_fecha_str(row.get("Fecha HC", "")) or ""

        if fpp:
            ws.cell(row=fila_actual, column=34, value=fpp)       # AH: FPP
        if fecha_hc:
            ws.cell(row=fila_actual, column=59, value=fecha_hc)  # BG: Fecha_HC

        ws.cell(row=fila_actual, column=15, value=1)  # O
        ws.cell(row=fila_actual, column=60, value=1)  # BH
        ws.cell(row=fila_actual, column=61, value=1)  # BI
        ws.cell(row=fila_actual, column=62, value=1)  # BJ

    wb.save(ruta_salida)
    wb.close()
    return True


def procesar_gestantes_savia(ruta_plantilla_in, ruta_plantilla_out, ruta_gestantes=None):
    """
    Procesa GESTANTES SAVIA.xlsx aplicando reglas de prioridad en cascada para
    Sífilis (CC, CD), VIH (CE, CF), Hepatitis B (CA, CB), FPP (AH), Riesgo (AJ),
    Fecha Control (BE) y marcadores fijos en O, BH, BI, BJ = 1.
    """
    if ruta_gestantes is None:
        base_dir = Path(__file__).resolve().parent.parent
        ruta_gestantes = base_dir / "Insumos" / "GESTANTES SAVIA.xlsx"
    else:
        ruta_gestantes = Path(ruta_gestantes)

    if not ruta_gestantes.exists():
        raise FileNotFoundError("No se encontró el archivo GESTANTES SAVIA.xlsx en Insumos.")

    df_gest = pd.read_excel(ruta_gestantes)
    df_gest = df_gest.dropna(subset=[df_gest.columns[0]])

    wb = openpyxl.load_workbook(ruta_plantilla_in)
    ws = wb.active

    indice_docs, fila_destino_nueva = indexar_documentos_plantilla(ws)

    col_o = column_index_from_string("O")
    col_bh = column_index_from_string("BH")
    col_bi = column_index_from_string("BI")
    col_bj = column_index_from_string("BJ")
    col_ah = column_index_from_string("AH")
    col_aj = column_index_from_string("AJ")
    col_ca = column_index_from_string("CA")
    col_cb = column_index_from_string("CB")
    col_be = column_index_from_string("BE")
    col_cc = column_index_from_string("CC")
    col_cd = column_index_from_string("CD")
    col_ce = column_index_from_string("CE")
    col_cf = column_index_from_string("CF")

    for _, row in df_gest.iterrows():
        doc = limpiar_documento(row.iloc[0])
        if not doc:
            continue

        fila_destino, fila_destino_nueva = obtener_o_crear_fila(
            ws, indice_docs, fila_destino_nueva, doc
        )

        # Marcadores fijos de gestante
        ws.cell(row=fila_destino, column=col_o, value=1)
        ws.cell(row=fila_destino, column=col_bh, value=1)
        ws.cell(row=fila_destino, column=col_bi, value=1)
        ws.cell(row=fila_destino, column=col_bj, value=1)

        # Col C (2) -> AH (FPP)
        fec_fpp = formatear_fecha_str(row.iloc[2])
        if fec_fpp:
            ws.cell(row=fila_destino, column=col_ah, value=fec_fpp)

        # Col D (3) -> AJ (Riesgo)
        if pd.notnull(row.iloc[3]):
            ws.cell(row=fila_destino, column=col_aj, value=row.iloc[3])

        # Col Q (16) -> CA (Hepatitis B Fecha)
        fec_hb = formatear_fecha_str(row.iloc[16])
        if fec_hb:
            ws.cell(row=fila_destino, column=col_ca, value=fec_hb)

        # Col R (17) -> CB (Hepatitis B Resultado)
        if pd.notnull(row.iloc[17]):
            ws.cell(row=fila_destino, column=col_cb, value=row.iloc[17])

        # Col B (1) -> BE (Fecha control prenatal)
        fec_be = formatear_fecha_date(row.iloc[1])
        if fec_be:
            ws.cell(row=fila_destino, column=col_be, value=fec_be)

        # Cascada 1: Sífilis (Resultados J->H->F en CD y Fechas I->G->E en CC)
        pares_sifilis = [
            (row.iloc[9], row.iloc[8]),  # Trimestre III (J, I)
            (row.iloc[7], row.iloc[6]),  # Trimestre II (H, G)
            (row.iloc[5], row.iloc[4]),  # Trimestre I (F, E)
        ]
        for res_val, fec_val in pares_sifilis:
            if pd.notnull(res_val) and str(res_val).strip() not in ("", "21"):
                ws.cell(row=fila_destino, column=col_cd, value=res_val)
                fec_sif = formatear_fecha_str(fec_val)
                if fec_sif:
                    ws.cell(row=fila_destino, column=col_cc, value=fec_sif)
                break

        # Cascada 2: VIH (Resultados P->N->L en CF y Fechas O->M->K en CE)
        pares_vih = [
            (row.iloc[15], row.iloc[14]),  # Trimestre III (P, O)
            (row.iloc[13], row.iloc[12]),  # Trimestre II (N, M)
            (row.iloc[11], row.iloc[10]),  # Trimestre I (L, K)
        ]
        for res_val, fec_val in pares_vih:
            if pd.notnull(res_val) and str(res_val).strip() not in ("", "21"):
                ws.cell(row=fila_destino, column=col_cf, value=res_val)
                fec_vih = formatear_fecha_str(fec_val)
                if fec_vih:
                    ws.cell(row=fila_destino, column=col_ce, value=fec_vih)
                break

    wb.save(ruta_plantilla_out)
    wb.close()
    return True
