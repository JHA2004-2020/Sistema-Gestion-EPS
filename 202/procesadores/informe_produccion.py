import streamlit as st
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
from utils.texto_helpers import convertir_edad_a_meses


SECCIONES_PRODUCCION = [
    "BA_seccion",
    "BC_BD_seccion",
    "AR_AU_seccion",
    "BB_seccion",
    "BQ_seccion",
    "DI_DJ_seccion",
    "AY_AX_seccion",
]


def crear_config_produccion_vacia():
    return {sec: [] for sec in SECCIONES_PRODUCCION}


@st.cache_data
def _leer_excel_produccion(ruta_str, mtime):
    return pd.read_excel(ruta_str)


def cargar_informe_produccion():
    """Carga el archivo INFORME PRODUCCION.xlsx invalidando caché si cambia en disco."""
    base_dir = Path(__file__).resolve().parent.parent
    ruta_inf = base_dir / "Insumos" / "INFORME PRODUCCION.xlsx"
    if not ruta_inf.exists():
        ruta_inf = base_dir / "INFORME PRODUCCION.xlsx"
    if not ruta_inf.exists():
        return None
    return _leer_excel_produccion(str(ruta_inf), ruta_inf.stat().st_mtime)


def procesar_informe_produccion(
    config, ruta_plantilla, entidades_seleccionadas=None, entidad_sugerida=""
):
    """
    Cruza los datos del INFORME PRODUCCION.xlsx según las 7 secciones configuradas
    manualmente por el usuario y las entidades seleccionadas.
    """
    df_general = cargar_informe_produccion()
    if df_general is None:
        raise FileNotFoundError("No se encontró el archivo 'INFORME PRODUCCION.xlsx'.")

    col_entidad = df_general.columns[34] if len(df_general.columns) > 34 else "ENTIDAD"
    col_doc = df_general.columns[8] if len(df_general.columns) > 8 else "CODIGO"
    col_fec = df_general.columns[3] if len(df_general.columns) > 3 else "FECHA"
    col_articulo = df_general.columns[14] if len(df_general.columns) > 14 else "NOMBRE_ARTICULO"
    col_edad = df_general.columns[64] if len(df_general.columns) > 64 else "EDAD"
    col_un_edad = df_general.columns[65] if len(df_general.columns) > 65 else "UNIDAD EDAD"

    df_general = df_general.copy()
    df_general[col_entidad] = df_general[col_entidad].apply(
        lambda x: str(x).upper().strip() if pd.notnull(x) else ""
    )

    if entidades_seleccionadas:
        df_filtrado = df_general[df_general[col_entidad].isin(entidades_seleccionadas)].copy()
    else:
        ent_sug = str(entidad_sugerida).upper().strip()
        df_filtrado = df_general[df_general[col_entidad].str.contains(ent_sug, na=False)].copy()

    if df_filtrado.empty:
        return

    df_filtrado["ART_NORM"] = df_filtrado[col_articulo].apply(
        lambda x: " ".join(str(x).upper().split()) if pd.notnull(x) else ""
    )

    todos_filtros = set()
    for sec in SECCIONES_PRODUCCION:
        for item in config.get(sec, []):
            todos_filtros.add(" ".join(str(item["filtro"]).upper().split()))

    if not todos_filtros:
        return

    df_match_total = df_filtrado[df_filtrado["ART_NORM"].isin(todos_filtros)].copy()
    usuarios_grupo = df_match_total.groupby(col_doc)

    wb = openpyxl.load_workbook(ruta_plantilla)
    ws = wb.active

    indice_docs, fila_destino_nueva = indexar_documentos_plantilla(ws)

    for doc_usuario, grupo in usuarios_grupo:
        doc_limpio = limpiar_documento(doc_usuario)
        if not doc_limpio:
            continue

        fila_destino, fila_destino_nueva = obtener_o_crear_fila(
            ws, indice_docs, fila_destino_nueva, doc_limpio
        )

        # 1. Sección BA (y reglas por edad en meses: BS, BT, BZ, BK, AN, AB, AC)
        cumplio_bk = False
        for item in config.get("BA_seccion", []):
            p_filtro_norm = " ".join(str(item["filtro"]).upper().split())
            match = grupo[grupo["ART_NORM"] == p_filtro_norm]
            if not match.empty:
                fecha_ba_str = formatear_fecha_str(match.iloc[0].get(col_fec, ""))
                if fecha_ba_str:
                    ws.cell(
                        row=fila_destino,
                        column=column_index_from_string("BA"),
                        value=fecha_ba_str,
                    )

                    edad_val = match.iloc[0].get(col_edad, 0)
                    un_val = match.iloc[0].get(col_un_edad, "")
                    meses_vida = convertir_edad_a_meses(edad_val, un_val)

                    if 6 <= meses_vida <= 23:
                        ws.cell(row=fila_destino, column=column_index_from_string("BS"), value=1)
                    if 24 <= meses_vida <= 60:
                        ws.cell(row=fila_destino, column=column_index_from_string("BT"), value=1)
                    if 24 <= meses_vida <= 59:
                        ws.cell(row=fila_destino, column=column_index_from_string("BZ"), value=1)

                    if meses_vida > 36:
                        ws.cell(
                            row=fila_destino,
                            column=column_index_from_string("BK"),
                            value=fecha_ba_str,
                        )
                        cumplio_bk = True

                    if meses_vida < 60:
                        ws.cell(row=fila_destino, column=column_index_from_string("AN"), value=0)
                    break

        if cumplio_bk:
            ws.cell(row=fila_destino, column=column_index_from_string("AC"), value=3)
            ws.cell(row=fila_destino, column=column_index_from_string("AB"), value=3)

        # 2. Sección BC y BD
        for item in config.get("BC_BD_seccion", []):
            p_filtro_norm = " ".join(str(item["filtro"]).upper().split())
            match_bc = grupo[grupo["ART_NORM"] == p_filtro_norm]
            if not match_bc.empty:
                val_bc = item.get("valor_bc", "23")
                ws.cell(row=fila_destino, column=column_index_from_string("BC"), value=val_bc)

                fec_bc_str = formatear_fecha_str(match_bc.iloc[0].get(col_fec, ""))
                if fec_bc_str:
                    ws.cell(
                        row=fila_destino,
                        column=column_index_from_string("BD"),
                        value=fec_bc_str,
                    )

        # 3. Sección AR, AS, AT, AU (Valor fijo = 5)
        for item in config.get("AR_AU_seccion", []):
            p_filtro_norm = " ".join(str(item["filtro"]).upper().split())
            match_ar = grupo[grupo["ART_NORM"] == p_filtro_norm]
            if not match_ar.empty:
                for col_letra in ["AR", "AS", "AT", "AU"]:
                    ws.cell(row=fila_destino, column=column_index_from_string(col_letra), value=5)
                break

        # 4. Sección BB (Fecha)
        for item in config.get("BB_seccion", []):
            p_filtro_norm = " ".join(str(item["filtro"]).upper().split())
            match_bb = grupo[grupo["ART_NORM"] == p_filtro_norm]
            if not match_bb.empty:
                fec_bb_str = formatear_fecha_str(match_bb.iloc[0].get(col_fec, ""))
                if fec_bb_str:
                    ws.cell(
                        row=fila_destino,
                        column=column_index_from_string("BB"),
                        value=fec_bb_str,
                    )
                break

        # 5. Sección BQ y asociadas (CJ=Fecha, CI=1, CK=17, CL=1, CM=053610450501)
        for item in config.get("BQ_seccion", []):
            p_filtro_norm = " ".join(str(item["filtro"]).upper().split())
            match_bq = grupo[grupo["ART_NORM"] == p_filtro_norm]
            if not match_bq.empty:
                fecha_limpia = formatear_fecha_date(match_bq.iloc[0].get(col_fec, ""))
                if fecha_limpia:
                    ws.cell(
                        row=fila_destino,
                        column=column_index_from_string("BQ"),
                        value=fecha_limpia,
                    )
                    ws.cell(
                        row=fila_destino,
                        column=column_index_from_string("CJ"),
                        value=fecha_limpia,
                    )

                ws.cell(row=fila_destino, column=column_index_from_string("CI"), value=1)
                ws.cell(row=fila_destino, column=column_index_from_string("CK"), value=17)
                ws.cell(row=fila_destino, column=column_index_from_string("CL"), value=1)
                ws.cell(
                    row=fila_destino,
                    column=column_index_from_string("CM"),
                    value="053610450501",
                )
                break

        # 6. Sección DI y DJ (Fecha o 'X')
        for item in config.get("DI_DJ_seccion", []):
            p_filtro_norm = " ".join(str(item["filtro"]).upper().split())
            match_di = grupo[grupo["ART_NORM"] == p_filtro_norm]
            if not match_di.empty:
                val_di = formatear_fecha_str(match_di.iloc[0].get(col_fec, "")) or "X"
                ws.cell(row=fila_destino, column=column_index_from_string("DI"), value=val_di)
                ws.cell(row=fila_destino, column=column_index_from_string("DJ"), value=val_di)
                break

        # 7. Sección AY (Fecha) y AX (Fecha - 1 día)
        for item in config.get("AY_AX_seccion", []):
            p_filtro_norm = " ".join(str(item["filtro"]).upper().split())
            match_ay = grupo[grupo["ART_NORM"] == p_filtro_norm]
            if not match_ay.empty:
                fec_raw_ay = match_ay.iloc[0].get(col_fec, "")
                parsed_fec_ay = pd.to_datetime(str(fec_raw_ay), errors="coerce")
                if pd.notnull(parsed_fec_ay):
                    ws.cell(
                        row=fila_destino,
                        column=column_index_from_string("AY"),
                        value=parsed_fec_ay.strftime("%Y/%m/%d"),
                    )
                    fec_ax = parsed_fec_ay - pd.Timedelta(days=1)
                    ws.cell(
                        row=fila_destino,
                        column=column_index_from_string("AX"),
                        value=fec_ax.strftime("%Y/%m/%d"),
                    )
                break

    wb.save(ruta_plantilla)
    wb.close()
