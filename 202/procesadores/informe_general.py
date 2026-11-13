import streamlit as st
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


@st.cache_data
def _leer_excel_general(ruta_str, mtime):
    return pd.read_excel(ruta_str, sheet_name=0)


def cargar_informe_general():
    """Carga el archivo INFORME GENERAL.xlsx invalidando caché si el archivo cambia."""
    base_dir = Path(__file__).resolve().parent.parent
    ruta_inf = base_dir / "Insumos" / "INFORME GENERAL.xlsx"
    if not ruta_inf.exists():
        return None
    return _leer_excel_general(str(ruta_inf), ruta_inf.stat().st_mtime)


def procesar_informe_general(df_filtrado, config_dict, ruta_plantilla):
    """
    Cruza el DataFrame ya filtrado por las entidades seleccionadas aplicando
    las reglas de prioridad y mapeo manual de columnas configuradas por el usuario.
    """
    if df_filtrado is None or df_filtrado.empty or not config_dict:
        return

    wb_final = openpyxl.load_workbook(ruta_plantilla)
    ws_final = wb_final.active

    indice_docs, fila_destino_nueva = indexar_documentos_plantilla(ws_final)
    usuarios_informe = df_filtrado.groupby("DOCUMENTO")

    mapeo_por_celda = {}
    for prueba_nombre, config_val in config_dict.items():
        pos_resultado = str(config_val.get("Posición Plantilla", "")).upper().strip()
        pos_fecha = str(config_val.get("Columna Fecha Plantilla", "")).upper().strip()

        if pos_resultado:
            if pos_resultado not in mapeo_por_celda:
                mapeo_por_celda[pos_resultado] = []
            mapeo_por_celda[pos_resultado].append(
                {
                    "prueba": prueba_nombre,
                    "pos_fecha": pos_fecha,
                }
            )

    for doc_usuario, grupo in usuarios_informe:
        doc_limpio = limpiar_documento(doc_usuario)
        if not doc_limpio:
            continue

        fila_destino, fila_destino_nueva = obtener_o_crear_fila(
            ws_final, indice_docs, fila_destino_nueva, doc_limpio
        )

        for pos_res, lista_candidatas in mapeo_por_celda.items():
            valor_asignado = None
            fecha_asignada_str = None
            p_nombre = None

            for cand in lista_candidatas:
                p_nombre = cand["prueba"]
                prueba_match = grupo[grupo["PRUEBA"].astype(str) == p_nombre]

                if not prueba_match.empty:
                    val_res = prueba_match.iloc[0].get("RESULTADO", "")
                    val_res_clean = str(val_res).strip() if pd.notnull(val_res) else ""

                    if val_res_clean != "" and val_res_clean.lower() not in ("nan", "-"):
                        col_destino_res = pos_res.upper().strip()
                        if col_destino_res in ["CB", "CD", "CF"]:
                            val_res_upper = val_res_clean.upper()
                            if val_res_upper in ("NEGATIVO", "NO REACTIVO"):
                                valor_asignado = 5
                            elif val_res_upper in ("POSITIVO", "REACTIVO"):
                                valor_asignado = 4
                            else:
                                valor_asignado = val_res
                        else:
                            valor_asignado = val_res

                        fecha_asignada_str = formatear_fecha_str(
                            prueba_match.iloc[0].get("FECHA", "")
                        )
                        break

            if valor_asignado is not None:
                try:
                    col_idx_res = column_index_from_string(pos_res)
                    ws_final.cell(row=fila_destino, column=col_idx_res, value=valor_asignado)
                except Exception:
                    pass

                for cand in lista_candidatas:
                    if cand["prueba"] == p_nombre and cand["pos_fecha"] and fecha_asignada_str:
                        try:
                            col_idx_fec = column_index_from_string(cand["pos_fecha"])
                            ws_final.cell(
                                row=fila_destino, column=col_idx_fec, value=fecha_asignada_str
                            )
                        except Exception:
                            pass
                        break

    wb_final.save(ruta_plantilla)
    wb_final.close()
