import streamlit as st
import pandas as pd
import json
import shutil
from pathlib import Path
import openpyxl
from openpyxl.utils import column_index_from_string

@st.cache_data
def cargar_informe_produccion():
    base_dir = Path(__file__).resolve().parent.parent
    ruta_inf = base_dir / "Insumos" / "INFORME PRODUCCION.xlsx"
    if not ruta_inf.exists():
        ruta_inf = base_dir / "INFORME PRODUCCION.xlsx"
    if not ruta_inf.exists():
        return None
    return pd.read_excel(ruta_inf)

def convertir_edad_a_meses(edad, unidad):
    """Convierte la edad a meses exactos según la unidad (DIAS, MESES, AÑOS)."""
    if pd.isnull(edad):
        return 0
    try:
        val = float(edad)
    except:
        return 0
    
    unidad_str = str(unidad).strip().upper()
    if "D" in unidad_str or "DÍA" in unidad_str:
        return val / 30.0
    elif "M" in unidad_str:
        return val
    elif "A" in unidad_str or "ANO" in unidad_str:
        return val * 12.0
    return val

def ejecutar_modulo_informe_produccion(entidad_sugerida):
    st.title(f"🏭 Módulo Informe de Producción - {str(entidad_sugerida).upper()}")
    st.markdown("---")

    base_dir = Path(__file__).resolve().parent.parent
    insumos_dir = base_dir / "Insumos"
    if not insumos_dir.exists():
        insumos_dir = base_dir

    df_general = cargar_informe_produccion()

    if df_general is None:
        st.error("No se encontró el archivo 'INFORME PRODUCCION.xlsx'.")
        return

    # Extraer columnas principales
    col_entidad = df_general.columns[34] if len(df_general.columns) > 34 else 'ENTIDAD'
    col_articulo = df_general.columns[14] if len(df_general.columns) > 14 else 'NOMBRE_ARTICULO'

    # Blindaje contra celdas vacías (NaN / float)
    df_general[col_entidad] = df_general[col_entidad].apply(lambda x: str(x).upper().strip() if pd.notnull(x) else "")
    
    # Extraemos valores únicos para los desplegables (blindado)
    entidades_unicas = sorted(df_general[col_entidad][df_general[col_entidad] != ""].unique())
    articulos_unicos = sorted(df_general[col_articulo].dropna().astype(str).str.strip().unique())

    # ==========================================
    # SECCIÓN 1: FILTRO DESPLEGABLE DE ENTIDAD
    # ==========================================
    st.subheader("🏢 1. Filtro de Entidades y Derivados")
    st.markdown("Selecciona exactamente cuáles derivados de la entidad deseas cruzar en este informe:")
    
    entidad_sug_str = str(entidad_sugerida).upper().strip()
    entidades_defecto = [e for e in entidades_unicas if entidad_sug_str in e]
    
    if 'filtro_entidades_prod' not in st.session_state:
        st.session_state.filtro_entidades_prod = entidades_defecto

    entidades_seleccionadas = st.multiselect(
        "Despliega y selecciona las entidades:",
        options=entidades_unicas,
        default=st.session_state.filtro_entidades_prod,
        key="ms_entidades_prod"
    )
    
    st.session_state.filtro_entidades_prod = entidades_seleccionadas

    if entidades_seleccionadas:
        df_filtrado = df_general[df_general[col_entidad].isin(entidades_seleccionadas)].copy()
        st.success(f"📊 Registros totales encontrados para las entidades seleccionadas: {len(df_filtrado)}")
    else:
        df_filtrado = pd.DataFrame()
        st.warning("⚠️ Debes seleccionar al menos una entidad para procesar.")
    
    st.markdown("---")

    # ==========================================
    # CONFIGURACIÓN DE SECCIONES
    # ==========================================
    if 'config_prod_nueva_eps' not in st.session_state:
        st.session_state.config_prod_nueva_eps = {
            "BA_seccion": [], "BC_BD_seccion": [], "AR_AU_seccion": [],
            "BB_seccion": [], "BQ_seccion": [], "DI_DJ_seccion": [], "AY_AX_seccion": []
        }

    config = st.session_state.config_prod_nueva_eps
    for sec in ["BC_BD_seccion", "AR_AU_seccion", "BB_seccion", "BQ_seccion", "DI_DJ_seccion", "AY_AX_seccion"]:
        if sec not in config: config[sec] = []

    st.sidebar.markdown("### 💾 Gestión de Configuración")
    json_str = json.dumps(config, indent=4, ensure_ascii=False)
    st.sidebar.download_button(
        label="💾 Guardar Configuración Actual",
        data=json_str,
        file_name=f"filtros_informe_produccion_{entidad_sugerida}.json",
        mime="application/json",
        use_container_width=True
    )

    st.sidebar.markdown("<br>", unsafe_allow_html=True)
    uploaded_config = st.sidebar.file_uploader("📂 Cargar configuración guardada (.json)", type=["json"], key="up_json_prod")
    if uploaded_config is not None:
        try:
            content = uploaded_config.getvalue()
            loaded_data = json.loads(content.decode("utf-8"))
            if loaded_data != st.session_state.config_prod_nueva_eps:
                st.session_state.config_prod_nueva_eps = loaded_data
                st.sidebar.success("¡Configuración cargada con éxito!")
                st.rerun()
        except Exception as e:
            st.sidebar.error(f"Error al leer el archivo JSON: {e}")

    # ==========================================
    # SECCIÓN 1: COLUMNA BA
    # ==========================================
    st.subheader("2️⃣ Sección Columna BA (Fecha de consulta de valoración integral)")
    c1, c2 = st.columns([1, 1.5])
    with c1:
        disp_ba = [a for a in articulos_unicos if a not in [x['filtro'] for x in config['BA_seccion']]]
        sel_ba = st.selectbox("🔍 Agregar filtro para BA:", ["-- Selecciona --"] + disp_ba, key="sel_ba")
        if sel_ba != "-- Selecciona --":
            config['BA_seccion'].append({"filtro": sel_ba, "pos_resultado": "BA"})
            st.rerun()
    with c2:
        if config['BA_seccion']:
            df_ba = pd.DataFrame(config['BA_seccion'])
            st.data_editor(df_ba, hide_index=True, use_container_width=True, key="ed_ba")
            item_ba = st.selectbox("Selecciona ítem para eliminar (BA):", options=[x['filtro'] for x in config['BA_seccion']], key="item_ba_sel")
            if st.button("🗑️ Eliminar Filtro BA", type="secondary", use_container_width=True):
                config['BA_seccion'] = [x for x in config['BA_seccion'] if x['filtro'] != item_ba]
                st.rerun()
        else:
            st.info("No hay filtros agregados.")

    st.markdown("---")

    # ==========================================
    # SECCIÓN 2: COLUMNAS BC y BD
    # ==========================================
    st.subheader("3️⃣ Sección Columnas BC y BD")
    d1, d2 = st.columns([1, 1.5])
    with d1:
        disp_bc = [a for a in articulos_unicos if a not in [x['filtro'] for x in config['BC_BD_seccion']]]
        sel_bc = st.selectbox("🔍 Agregar filtro para BC/BD:", ["-- Selecciona --"] + disp_bc, key="sel_bc")
        if sel_bc != "-- Selecciona --":
            config['BC_BD_seccion'].append({"filtro": sel_bc, "valor_bc": "23", "pos_resultado_bc": "BC", "pos_resultado_bd": "BD"})
            st.rerun()
    with d2:
        if config['BC_BD_seccion']:
            df_bc_bd = pd.DataFrame(config['BC_BD_seccion'])
            edited_df_bc_bd = st.data_editor(df_bc_bd, hide_index=True, use_container_width=True, key="ed_bc_bd")
            config['BC_BD_seccion'] = edited_df_bc_bd.to_dict('records')
            item_bc_sel = st.selectbox("Selecciona ítem para eliminar (BC/BD):", options=[x['filtro'] for x in config['BC_BD_seccion']], key="item_bc_sel")
            if st.button("🗑️ Eliminar Filtro BC/BD", type="secondary", use_container_width=True):
                config['BC_BD_seccion'] = [x for x in config['BC_BD_seccion'] if x['filtro'] != item_bc_sel]
                st.rerun()
        else:
            st.info("No hay filtros agregados.")

    st.markdown("---")

    # ==========================================
    # SECCIÓN 3: COLUMNAS AR, AS, AT, AU
    # ==========================================
    st.subheader("4️⃣ Sección Columnas AR, AS, AT, AU (Valor fijo = 5)")
    e1, e2 = st.columns([1, 1.5])
    with e1:
        disp_ar = [a for a in articulos_unicos if a not in [x['filtro'] for x in config['AR_AU_seccion']]]
        sel_ar = st.selectbox("🔍 Agregar filtro para AR-AU:", ["-- Selecciona --"] + disp_ar, key="sel_ar")
        if sel_ar != "-- Selecciona --":
            config['AR_AU_seccion'].append({"filtro": sel_ar})
            st.rerun()
    with e2:
        if config['AR_AU_seccion']:
            df_ar_au = pd.DataFrame(config['AR_AU_seccion'])
            st.data_editor(df_ar_au, hide_index=True, use_container_width=True, key="ed_ar_au")
            item_ar_sel = st.selectbox("Selecciona ítem para eliminar (AR-AU):", options=[x['filtro'] for x in config['AR_AU_seccion']], key="item_ar_sel")
            if st.button("🗑️ Eliminar Filtro AR-AU", type="secondary", use_container_width=True):
                config['AR_AU_seccion'] = [x for x in config['AR_AU_seccion'] if x['filtro'] != item_ar_sel]
                st.rerun()
        else:
            st.info("No hay filtros agregados.")

    st.markdown("---")

    # ==========================================
    # SECCIÓN 4: COLUMNA BB
    # ==========================================
    st.subheader("5️⃣ Sección Columna BB (Fecha)")
    f1, f2 = st.columns([1, 1.5])
    with f1:
        disp_bb = [a for a in articulos_unicos if a not in [x['filtro'] for x in config['BB_seccion']]]
        sel_bb = st.selectbox("🔍 Agregar filtro para BB:", ["-- Selecciona --"] + disp_bb, key="sel_bb")
        if sel_bb != "-- Selecciona --":
            config['BB_seccion'].append({"filtro": sel_bb})
            st.rerun()
    with f2:
        if config['BB_seccion']:
            df_bb = pd.DataFrame(config['BB_seccion'])
            st.data_editor(df_bb, hide_index=True, use_container_width=True, key="ed_bb")
            item_bb_sel = st.selectbox("Selecciona ítem para eliminar (BB):", options=[x['filtro'] for x in config['BB_seccion']], key="item_bb_sel")
            if st.button("🗑️ Eliminar Filtro BB", type="secondary", use_container_width=True):
                config['BB_seccion'] = [x for x in config['BB_seccion'] if x['filtro'] != item_bb_sel]
                st.rerun()
        else:
            st.info("No hay filtros agregados.")

    st.markdown("---")

    # ==========================================
    # SECCIÓN 5: COLUMNA BQ y ASOCIADAS (CI, CK, CL, CM)
    # ==========================================
    st.subheader("6️⃣ Sección Columna BQ (Fecha) y asociadas (CI=1, CK=17, CL=1, CM=053610450501)")
    g1, g2 = st.columns([1, 1.5])
    with g1:
        disp_bq = [a for a in articulos_unicos if a not in [x['filtro'] for x in config['BQ_seccion']]]
        sel_bq = st.selectbox("🔍 Agregar filtro para BQ:", ["-- Selecciona --"] + disp_bq, key="sel_bq")
        if sel_bq != "-- Selecciona --":
            config['BQ_seccion'].append({"filtro": sel_bq})
            st.rerun()
    with g2:
        if config['BQ_seccion']:
            df_bq = pd.DataFrame(config['BQ_seccion'])
            st.data_editor(df_bq, hide_index=True, use_container_width=True, key="ed_bq")
            item_bq_sel = st.selectbox("Selecciona ítem para eliminar (BQ):", options=[x['filtro'] for x in config['BQ_seccion']], key="item_bq_sel")
            if st.button("🗑️ Eliminar Filtro BQ", type="secondary", use_container_width=True):
                config['BQ_seccion'] = [x for x in config['BQ_seccion'] if x['filtro'] != item_bq_sel]
                st.rerun()
        else:
            st.info("No hay filtros agregados.")

    st.markdown("---")

    # ==========================================
    # SECCIÓN 6: COLUMNAS DI y DJ
    # ==========================================
    st.subheader("7️⃣ Sección Columnas DI y DJ (Fecha)")
    h1, h2 = st.columns([1, 1.5])
    with h1:
        disp_di = [a for a in articulos_unicos if a not in [x['filtro'] for x in config['DI_DJ_seccion']]]
        sel_di = st.selectbox("🔍 Agregar filtro para DI/DJ:", ["-- Selecciona --"] + disp_di, key="sel_di")
        if sel_di != "-- Selecciona --":
            config['DI_DJ_seccion'].append({"filtro": sel_di})
            st.rerun()
    with h2:
        if config['DI_DJ_seccion']:
            df_di_dj = pd.DataFrame(config['DI_DJ_seccion'])
            st.data_editor(df_di_dj, hide_index=True, use_container_width=True, key="ed_di_dj")
            item_di_sel = st.selectbox("Selecciona ítem para eliminar (DI/DJ):", options=[x['filtro'] for x in config['DI_DJ_seccion']], key="item_di_sel")
            if st.button("🗑️ Eliminar Filtro DI/DJ", type="secondary", use_container_width=True):
                config['DI_DJ_seccion'] = [x for x in config['DI_DJ_seccion'] if x['filtro'] != item_di_sel]
                st.rerun()
        else:
            st.info("No hay filtros agregados.")

    st.markdown("---")

    # ==========================================
    # SECCIÓN 7: COLUMNAS AY y AX
    # ==========================================
    st.subheader("8️⃣ Sección Columna AY (Fecha) y AX (Fecha - 1 día)")
    i1, i2 = st.columns([1, 1.5])
    with i1:
        disp_ay = [a for a in articulos_unicos if a not in [x['filtro'] for x in config['AY_AX_seccion']]]
        sel_ay = st.selectbox("🔍 Agregar filtro para AY/AX:", ["-- Selecciona --"] + disp_ay, key="sel_ay")
        if sel_ay != "-- Selecciona --":
            config['AY_AX_seccion'].append({"filtro": sel_ay})
            st.rerun()
    with i2:
        if config['AY_AX_seccion']:
            df_ay_ax = pd.DataFrame(config['AY_AX_seccion'])
            st.data_editor(df_ay_ax, hide_index=True, use_container_width=True, key="ed_ay_ax")
            item_ay_sel = st.selectbox("Selecciona ítem para eliminar (AY/AX):", options=[x['filtro'] for x in config['AY_AX_seccion']], key="item_ay_sel")
            if st.button("🗑️ Eliminar Filtro AY/AX", type="secondary", use_container_width=True):
                config['AY_AX_seccion'] = [x for x in config['AY_AX_seccion'] if x['filtro'] != item_ay_sel]
                st.rerun()
        else:
            st.info("No hay filtros agregados.")

    st.markdown("---")
    st.info(f"💡 **Nota:** La ejecución se realiza desde el botón final en el módulo de 'Informe General'.")


# ==============================================================================
# FUNCIÓN PURA DE EJECUCIÓN (Llamada desde Informe General)
# ==============================================================================
def procesar_informe_produccion(config, ruta_plantilla, entidad_sugerida=""):
    """
    Toma la configuración guardada y cruza los datos según las entidades seleccionadas.
    Incluye blindaje estricto contra valores flotantes o celdas vacías.
    """
    df_general = cargar_informe_produccion()
    if df_general is None:
        raise Exception("No se encontró el archivo 'INFORME PRODUCCION.xlsx'.")

    col_entidad = df_general.columns[34] if len(df_general.columns) > 34 else 'ENTIDAD'
    col_doc = df_general.columns[8] if len(df_general.columns) > 8 else 'CODIGO'
    col_fec = df_general.columns[3] if len(df_general.columns) > 3 else 'FECHA'
    col_articulo = df_general.columns[14] if len(df_general.columns) > 14 else 'NOMBRE_ARTICULO'
    col_edad = df_general.columns[64] if len(df_general.columns) > 64 else 'EDAD'
    col_un_edad = df_general.columns[65] if len(df_general.columns) > 65 else 'UNIDAD EDAD'

    # Blindaje al limpiar Entidades
    df_general[col_entidad] = df_general[col_entidad].apply(lambda x: str(x).upper().strip() if pd.notnull(x) else "")

    if 'filtro_entidades_prod' in st.session_state and st.session_state.filtro_entidades_prod:
        entidades_a_filtrar = st.session_state.filtro_entidades_prod
        df_filtrado = df_general[df_general[col_entidad].isin(entidades_a_filtrar)].copy()
    else:
        ent_sug = str(entidad_sugerida).upper().strip()
        df_filtrado = df_general[df_general[col_entidad].str.contains(ent_sug, na=False)].copy()

    if df_filtrado.empty:
        return

    # Blindaje total contra celdas vacías (NaN) en la columna ARTICULO al convertir a mayúsculas
    df_filtrado['ART_NORM'] = df_filtrado[col_articulo].apply(lambda x: " ".join(str(x).upper().split()) if pd.notnull(x) else "")

    filtros_ba = [" ".join(str(item['filtro']).upper().split()) for item in config.get('BA_seccion', [])]
    filtros_bc_bd = [" ".join(str(item['filtro']).upper().split()) for item in config.get('BC_BD_seccion', [])]
    filtros_ar_au = [" ".join(str(item['filtro']).upper().split()) for item in config.get('AR_AU_seccion', [])]
    filtros_bb = [" ".join(str(item['filtro']).upper().split()) for item in config.get('BB_seccion', [])]
    filtros_bq = [" ".join(str(item['filtro']).upper().split()) for item in config.get('BQ_seccion', [])]
    filtros_di_dj = [" ".join(str(item['filtro']).upper().split()) for item in config.get('DI_DJ_seccion', [])]
    filtros_ay_ax = [" ".join(str(item['filtro']).upper().split()) for item in config.get('AY_AX_seccion', [])]
    
    todos_filtros = list(set(filtros_ba + filtros_bc_bd + filtros_ar_au + filtros_bb + filtros_bq + filtros_di_dj + filtros_ay_ax))

    if not todos_filtros:
        return 

    df_match_total = df_filtrado[df_filtrado['ART_NORM'].isin(todos_filtros)].copy()
    usuarios_grupo = df_match_total.groupby(col_doc)

    wb = openpyxl.load_workbook(ruta_plantilla)
    ws = wb.active

    usuarios_existentes = {}
    for r in range(2, ws.max_row + 1):
        val = ws.cell(row=r, column=5).value
        if val is not None:
            val_str = str(val).split('.')[0].strip()
            if val_str:
                usuarios_existentes[val_str] = r

    fila_destino_nueva = ws.max_row + 1 if ws.cell(row=2, column=5).value is not None else 2

    for doc_usuario, grupo in usuarios_grupo:
        doc_limpio = str(doc_usuario).split('.')[0].strip()
        if not doc_limpio or doc_limpio == 'nan':
            continue

        if doc_limpio in usuarios_existentes:
            fila_destino = usuarios_existentes[doc_limpio]
        else:
            fila_destino = fila_destino_nueva
            ws.cell(row=fila_destino, column=5, value=doc_limpio)
            usuarios_existentes[doc_limpio] = fila_destino
            fila_destino_nueva += 1

        fecha_ba_str = None
        cumplio_bk = False
        for item in config.get('BA_seccion', []):
            p_filtro_norm = " ".join(str(item['filtro']).upper().split())
            match = grupo[grupo['ART_NORM'] == p_filtro_norm]
            if not match.empty:
                fec_raw = match.iloc[0].get(col_fec, '')
                parsed_fec = pd.to_datetime(str(fec_raw), errors='coerce')
                if pd.notnull(parsed_fec):
                    fecha_ba_str = parsed_fec.strftime('%Y/%m/%d')
                    ws.cell(row=fila_destino, column=column_index_from_string("BA"), value=fecha_ba_str)
                    
                    edad_val = match.iloc[0].get(col_edad, 0)
                    un_val = match.iloc[0].get(col_un_edad, '')
                    meses_vida = convertir_edad_a_meses(edad_val, un_val)

                    if 6 <= meses_vida <= 23:
                        ws.cell(row=fila_destino, column=column_index_from_string("BS"), value=1)
                    if 24 <= meses_vida <= 60:
                        ws.cell(row=fila_destino, column=column_index_from_string("BT"), value=1)
                    if 24 <= meses_vida <= 59:
                        ws.cell(row=fila_destino, column=column_index_from_string("BZ"), value=1)

                    if meses_vida > 36:
                        ws.cell(row=fila_destino, column=column_index_from_string("BK"), value=fecha_ba_str)
                        cumplio_bk = True

                    if meses_vida < 60:
                        ws.cell(row=fila_destino, column=column_index_from_string("AN"), value=0)
                    break

        if cumplio_bk:
            ws.cell(row=fila_destino, column=column_index_from_string("AC"), value=3)
            ws.cell(row=fila_destino, column=column_index_from_string("AB"), value=3)

        for item in config.get('BC_BD_seccion', []):
            p_filtro_norm = " ".join(str(item['filtro']).upper().split())
            match_bc = grupo[grupo['ART_NORM'] == p_filtro_norm]
            if not match_bc.empty:
                val_bc = item.get('valor_bc', '23')
                ws.cell(row=fila_destino, column=column_index_from_string("BC"), value=val_bc)

                fec_raw_bc = match_bc.iloc[0].get(col_fec, '')
                parsed_fec_bc = pd.to_datetime(str(fec_raw_bc), errors='coerce')
                if pd.notnull(parsed_fec_bc):
                    ws.cell(row=fila_destino, column=column_index_from_string("BD"), value=parsed_fec_bc.strftime('%Y/%m/%d'))

        for item in config.get('AR_AU_seccion', []):
            p_filtro_norm = " ".join(str(item['filtro']).upper().split())
            match_ar = grupo[grupo['ART_NORM'] == p_filtro_norm]
            if not match_ar.empty:
                for col_letra in ["AR", "AS", "AT", "AU"]:
                    ws.cell(row=fila_destino, column=column_index_from_string(col_letra), value=5)
                break

        for item in config.get('BB_seccion', []):
            p_filtro_norm = " ".join(str(item['filtro']).upper().split())
            match_bb = grupo[grupo['ART_NORM'] == p_filtro_norm]
            if not match_bb.empty:
                fec_raw_bb = match_bb.iloc[0].get(col_fec, '')
                parsed_fec_bb = pd.to_datetime(str(fec_raw_bb), errors='coerce')
                if pd.notnull(parsed_fec_bb):
                    ws.cell(row=fila_destino, column=column_index_from_string("BB"), value=parsed_fec_bb.strftime('%Y/%m/%d'))
                break

        # =======================================================
        # SECCIÓN BQ Y ASOCIADAS (Incluyendo CJ con fecha Columna D)
        # =======================================================
        for item in config.get('BQ_seccion', []):
            p_filtro_norm = " ".join(str(item['filtro']).upper().split())
            match_bq = grupo[grupo['ART_NORM'] == p_filtro_norm]
            if not match_bq.empty:
                # col_fec corresponde al índice 3 (Columna D del INFORME PRODUCCION)
                fec_raw_bq = match_bq.iloc[0].get(col_fec, '')
                parsed_fec_bq = pd.to_datetime(str(fec_raw_bq), errors='coerce')
                
                if pd.notnull(parsed_fec_bq):
                    fecha_limpia = parsed_fec_bq.date()
                    # Inyectamos la fecha en BQ
                    ws.cell(row=fila_destino, column=column_index_from_string("BQ"), value=fecha_limpia)
                    # Inyectamos la misma fecha de la Columna D en CJ
                    ws.cell(row=fila_destino, column=column_index_from_string("CJ"), value=fecha_limpia)
                
                # Inyectamos los valores fijos asociados a esta misma regla
                ws.cell(row=fila_destino, column=column_index_from_string("CI"), value=1)
                ws.cell(row=fila_destino, column=column_index_from_string("CK"), value=17)
                ws.cell(row=fila_destino, column=column_index_from_string("CL"), value=1)
                ws.cell(row=fila_destino, column=column_index_from_string("CM"), value="053610450501")
                
                break

        for item in config.get('DI_DJ_seccion', []):
            p_filtro_norm = " ".join(str(item['filtro']).upper().split())
            match_di = grupo[grupo['ART_NORM'] == p_filtro_norm]
            if not match_di.empty:
                fec_raw_di = match_di.iloc[0].get(col_fec, '')
                parsed_fec_di = pd.to_datetime(str(fec_raw_di), errors='coerce')
                val_di = parsed_fec_di.strftime('%Y/%m/%d') if pd.notnull(parsed_fec_di) else "X"
                ws.cell(row=fila_destino, column=column_index_from_string("DI"), value=val_di)
                ws.cell(row=fila_destino, column=column_index_from_string("DJ"), value=val_di)
                break

        for item in config.get('AY_AX_seccion', []):
            p_filtro_norm = " ".join(str(item['filtro']).upper().split())
            match_ay = grupo[grupo['ART_NORM'] == p_filtro_norm]
            if not match_ay.empty:
                fec_raw_ay = match_ay.iloc[0].get(col_fec, '')
                parsed_fec_ay = pd.to_datetime(str(fec_raw_ay), errors='coerce')
                if pd.notnull(parsed_fec_ay):
                    ws.cell(row=fila_destino, column=column_index_from_string("AY"), value=parsed_fec_ay.strftime('%Y/%m/%d'))
                    fec_ax = parsed_fec_ay - pd.Timedelta(days=1)
                    ws.cell(row=fila_destino, column=column_index_from_string("AX"), value=fec_ax.strftime('%Y/%m/%d'))
                break

    wb.save(ruta_plantilla)
    wb.close()