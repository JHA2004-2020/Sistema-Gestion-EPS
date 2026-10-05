import streamlit as st
import pandas as pd
import json
import shutil
from pathlib import Path
import openpyxl
from openpyxl.utils import column_index_from_string

@st.cache_data
def cargar_informe():
    base_dir = Path(__file__).resolve().parent.parent
    ruta_inf = base_dir / "Insumos" / "INFORME GENERAL.xlsx"
    if not ruta_inf.exists():
        return None
    return pd.read_excel(ruta_inf, sheet_name=0)

def ejecutar_modulo_informe_general(entidad):
    st.title(f"📊 Módulo Informe General - {entidad.upper()}")
    st.markdown("---")

    base_dir = Path(__file__).resolve().parent.parent
    insumos_dir = base_dir / "Insumos"

    df_general = cargar_informe()

    if df_general is None:
        st.error("No se encontró el archivo 'INFORME GENERAL.xlsx' en la carpeta Insumos.")
        return

    # Validamos y preparamos las columnas
    # Validamos y preparamos las columnas
    if 'EPS' in df_general.columns:
        # Aplicamos strip() a toda la columna para borrar espacios invisibles
        df_general['EPS'] = df_general['EPS'].astype(str).str.upper().str.strip()
        entidades_unicas = sorted(df_general['EPS'].dropna().unique())
    else:
        entidades_unicas = []
        df_general['EPS'] = "N/A"

    # ==========================================
    # NUEVA SECCIÓN: FILTRO DESPLEGABLE DE ENTIDAD
    # ==========================================
    st.subheader("🏢 1. Filtro de Entidades y Derivados")
    st.markdown("Selecciona los derivados de la entidad para el Informe General:")
    
    # Pre-seleccionar los derivados que contengan la entidad actual
    entidades_defecto = [e for e in entidades_unicas if entidad.upper() in e]
    
    if 'filtro_entidades_general' not in st.session_state:
        st.session_state.filtro_entidades_general = entidades_defecto

    entidades_seleccionadas = st.multiselect(
        "Despliega y selecciona las entidades:",
        options=entidades_unicas,
        default=st.session_state.filtro_entidades_general,
        key="ms_entidades_general"
    )
    st.session_state.filtro_entidades_general = entidades_seleccionadas

    # Filtrar el DataFrame según la selección manual
    if entidades_seleccionadas:
        df_filtrado = df_general[df_general['EPS'].isin(entidades_seleccionadas)].copy()
        st.success(f"📊 Registros totales encontrados para las entidades seleccionadas: {len(df_filtrado)}")
    else:
        df_filtrado = df_general.copy()
        st.warning("⚠️ No has seleccionado entidades. El sistema usará la base completa (no recomendado).")

    # Extraer las pruebas únicas basándonos en la base completa para evitar pérdida de opciones
    if 'PRUEBA' in df_general.columns:
        pruebas_unicas = sorted(df_general['PRUEBA'].dropna().astype(str).unique())
    else:
        pruebas_unicas = []

    if 'config_dict_nueva_eps' not in st.session_state:
        st.session_state.config_dict_nueva_eps = {}

    # ==========================================
    # BARRA LATERAL PARA GUARDAR Y CARGAR CONFIGURACIÓN
    # ==========================================
    st.sidebar.header("💾 Gestión de Configuración")
    
    if st.sidebar.button("💾 Guardar Configuración Actual"):
        config_data = st.session_state.config_dict_nueva_eps
        json_str = json.dumps(config_data, indent=4, ensure_ascii=False)
        st.sidebar.download_button(
            label="📥 Descargar archivo de configuración (.json)",
            data=json_str,
            file_name=f"filtros_informe_general_{entidad}.json",
            mime="application/json"
        )

    uploaded_config = st.sidebar.file_uploader("📂 Cargar configuración guardada (.json)", type=["json"], key="up_json_nueva_eps")
    if uploaded_config is not None:
        try:
            content = uploaded_config.getvalue()
            loaded_data = json.loads(content.decode("utf-8"))
            if loaded_data != st.session_state.config_dict_nueva_eps:
                st.session_state.config_dict_nueva_eps = loaded_data
                st.sidebar.success("¡Configuración cargada con éxito!")
                st.rerun()
        except Exception as e:
            st.sidebar.error(f"Error al leer el archivo JSON: {e}")

    st.markdown("---")
    st.subheader("📋 Configuración y Dinamización de Pruebas")

    col1, col2 = st.columns([1, 1.5])

    with col1:
        st.markdown("### 2. Agregar Prueba")
        st.markdown("Usa el buscador para filtrar y agregar pruebas:")
        
        pruebas_disponibles = [p for p in pruebas_unicas if p not in st.session_state.config_dict_nueva_eps]
        
        prueba_a_agregar = st.selectbox(
            "🔍 Buscar y seleccionar prueba:",
            options=["-- Selecciona una prueba --"] + pruebas_disponibles
        )

        if prueba_a_agregar != "-- Selecciona una prueba --":
            st.session_state.config_dict_nueva_eps[prueba_a_agregar] = {
                "Posición Plantilla": "AA",
                "Columna Fecha Plantilla": "AA"
            }
            st.rerun()

    with col2:
        st.markdown("### 3. Tabla de Configuración, Prioridades y Destinos")
        st.markdown("Define columnas y ajusta la **prioridad (arriba = mayor prioridad)**:")

        if len(st.session_state.config_dict_nueva_eps) > 0:
            keys_list = list(st.session_state.config_dict_nueva_eps.keys())
            
            lista_tabla = []
            for prueba_key in keys_list:
                val = st.session_state.config_dict_nueva_eps[prueba_key]
                lista_tabla.append({
                    "Filtro": prueba_key,
                    "Posición Plantilla (Resultado)": val["Posición Plantilla"],
                    "Posición Plantilla (Fecha opcional)": val.get("Columna Fecha Plantilla", "")
                })

            df_config = pd.DataFrame(lista_tabla)
            
            edited_df = st.data_editor(
                df_config,
                hide_index=True,
                use_container_width=True,
                key="editor_config_tabla_nueva_eps"
            )

            for idx, row in edited_df.iterrows():
                p_name = row['Filtro']
                if p_name in st.session_state.config_dict_nueva_eps:
                    st.session_state.config_dict_nueva_eps[p_name] = {
                        "Posición Plantilla": row['Posición Plantilla (Resultado)'],
                        "Columna Fecha Plantilla": row['Posición Plantilla (Fecha opcional)']
                    }

            st.markdown("---")
            st.markdown("### ↕️ Ajustar Prioridades (Subir / Bajar) o Eliminar")
            
            item_seleccionado = st.selectbox("Selecciona una prueba para mover o borrar:", options=keys_list, key="select_gestionar_nueva_eps")
            
            c_subir, c_bajar, c_borrar = st.columns(3)
            with c_subir:
                if st.button("⬆️ Subir Prioridad", use_container_width=True):
                    idx = keys_list.index(item_seleccionado)
                    if idx > 0:
                        keys_list[idx], keys_list[idx-1] = keys_list[idx-1], keys_list[idx]
                        st.session_state.config_dict_nueva_eps = {k: st.session_state.config_dict_nueva_eps[k] for k in keys_list}
                        st.rerun()
            with c_bajar:
                if st.button("⬇️ Bajar Prioridad", use_container_width=True):
                    idx = keys_list.index(item_seleccionado)
                    if idx < len(keys_list) - 1:
                        keys_list[idx], keys_list[idx+1] = keys_list[idx+1], keys_list[idx]
                        st.session_state.config_dict_nueva_eps = {k: st.session_state.config_dict_nueva_eps[k] for k in keys_list}
                        st.rerun()
            with c_borrar:
                if st.button("🗑️ Eliminar", type="secondary", use_container_width=True):
                    if item_seleccionado in st.session_state.config_dict_nueva_eps:
                        del st.session_state.config_dict_nueva_eps[item_seleccionado]
                        st.rerun()
        else:
            st.info("No hay pruebas seleccionadas. Usa el buscador de la izquierda.")

    st.markdown("---")

    # ==========================================
    # EJECUCIÓN DE LA CONSOLIDACIÓN FINAL (ORQUESTADOR)
    # ==========================================
    st.subheader("🚀 Ejecución de la Consolidación General")
    st.markdown("Se procesará aplicando COP, Gestantes, Informe General, Informe de Producción, HC y Ajustes Demográficos en orden secuencial.")

    if st.button("✨ Generar Plantilla Procesada Final", type="primary", use_container_width=True):
        if not entidades_seleccionadas:
            st.error("Por favor, selecciona al menos una entidad en el Paso 1 antes de procesar.")
            return

        plantilla_original = insumos_dir / "Plantilla.xlsx"
        plantilla_resultado = insumos_dir / f"Plantilla_Procesada_{entidad}.xlsx"

        if not plantilla_original.exists():
            st.error("No se encontró el archivo base 'Plantilla.xlsx'.")
        else:
            # 1. Copia limpia de la plantilla desde cero
            shutil.copy(plantilla_original, plantilla_resultado)
            
            wb = openpyxl.load_workbook(plantilla_resultado)
            ws = wb.active
            for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
                for cell in row:
                    cell.value = None
            wb.save(plantilla_resultado)
            wb.close()

            # 2. Ejecutar COP
            try:
                from Core.cop import procesar_cop_generico
                if entidad.upper() == "NUEVA EPS" or entidad.upper() == "NUEVA":
                    procesar_cop_generico("NUEVA EPS", plantilla_resultado, plantilla_resultado)
                    st.write("✅ Módulo **COP** procesado correctamente.")
                else:
                    procesar_cop_generico("SAVIA", plantilla_resultado, plantilla_resultado)
                    st.write("✅ Módulo **COP** procesado correctamente.")
            except Exception as e:
                st.warning(f"Aviso en COP: {e}")

            # 4. Procesar INFORME GENERAL
            try:
                wb_final = openpyxl.load_workbook(plantilla_resultado)
                ws_final = wb_final.active

                usuarios_informe = df_filtrado.groupby('DOCUMENTO')

                mapeo_por_celda = {}
                for prueba_nombre, config_val in st.session_state.config_dict_nueva_eps.items():
                    pos_resultado = config_val['Posición Plantilla'].upper().strip()
                    pos_fecha = config_val.get('Columna Fecha Plantilla', '').upper().strip()
                    
                    if pos_resultado:
                        if pos_resultado not in mapeo_por_celda:
                            mapeo_por_celda[pos_resultado] = []
                        mapeo_por_celda[pos_resultado].append({
                            "prueba": prueba_nombre,
                            "pos_fecha": pos_fecha
                        })

                for doc_usuario, grupo in usuarios_informe:
                    fila_destino = None
                    for r in range(2, ws_final.max_row + 1):
                        if str(ws_final.cell(row=r, column=5).value) == str(doc_usuario):
                            fila_destino = r
                            break
                    
                    if fila_destino is None:
                        fila_destino = ws_final.max_row + 1 if ws_final.cell(row=2, column=5).value is not None else 2
                        ws_final.cell(row=fila_destino, column=5, value=doc_usuario)

                    for pos_res, lista_candidatas in mapeo_por_celda.items():
                        valor_asignado = None
                        fecha_asignada_str = None
                        p_nombre = None
                        
                        for cand in lista_candidatas:
                            p_nombre = cand["prueba"]
                            
                            prueba_match = grupo[grupo['PRUEBA'].astype(str) == p_nombre]
                            if not prueba_match.empty:
                                val_res = prueba_match.iloc[0].get('RESULTADO', '')
                                
                                if pd.notnull(val_res) and str(val_res).strip() != "" and str(val_res).lower() != "nan":
                                    valor_asignado = val_res
                                    
                                    fec_raw = prueba_match.iloc[0].get('FECHA', '')
                                    if pd.notnull(fec_raw):
                                        parsed_date = pd.to_datetime(str(fec_raw), errors='coerce')
                                        fecha_asignada_str = parsed_date.strftime('%Y/%m/%d') if pd.notnull(parsed_date) else ""
                                    
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
                                        ws_final.cell(row=fila_destino, column=col_idx_fec, value=fecha_asignada_str)
                                    except Exception:
                                        pass
                                    break

                wb_final.save(plantilla_resultado)
                wb_final.close()
                st.write("✅ Módulo **INFORME GENERAL** procesado correctamente.")
            except Exception as e:
                st.error(f"Error al dinamizar Informe General: {e}")

            # 3. Ejecutar GESTANTES
            try:
                if entidad.upper() == "NUEVA EPS" or entidad.upper() == "NUEVA":
                    from ModulosNuevaEPS.gestantes import procesar_gestantes_nueva_eps
                    procesar_gestantes_nueva_eps(plantilla_resultado, plantilla_resultado)
                    st.write("✅ Módulo **GESTANTES NUEVA** procesado correctamente.")
                else:
                    from ModulosSAVIA.gestantes import procesar_gestantes_savia
                    procesar_gestantes_savia(plantilla_resultado, plantilla_resultado)
                    st.write("✅ Módulo **GESTANTES SAVIA** procesado correctamente.")
            except Exception as e:
                st.warning(f"Aviso en GESTANTES: {e}")
            
            # ==========================================
            # 5. Ejecutar INFORME PRODUCCIÓN
            # ==========================================
            try:
                if 'config_prod_nueva_eps' in st.session_state and st.session_state.config_prod_nueva_eps:
                    from Core.informe_produccion import procesar_informe_produccion
                    procesar_informe_produccion(st.session_state.config_prod_nueva_eps, plantilla_resultado, entidad)
                    st.write("✅ Módulo **INFORME DE PRODUCCIÓN** cruzado correctamente.")
                else:
                    st.warning("⚠️ No se cruzó Informe de Producción (no hay filtros configurados en sesión).")
            except Exception as e:
                st.error(f"Error al dinamizar Informe de Producción: {e}")

            # ==========================================
            # 6. Ejecutar INFORME HC (Pesos y Tallas)
            # ==========================================
            try:
                from Core.utils import procesar_informe_hc
                ruta_hc = insumos_dir / "INFORME HC.xlsx"
                procesar_informe_hc(plantilla_resultado, ruta_hc)
                st.write("✅ Módulo **INFORME HC (Pesos y Tallas)** cruzado correctamente.")
            except Exception as e:
                st.error(f"Error al procesar Informe HC: {e}")

            # ==========================================
            # 7. Ejecutar AJUSTES DEMOGRÁFICOS (Nombres, Sexo, Fecha Nace)
            # ==========================================
            try:
                from Core.utils import procesar_ajustes_demograficos
                ruta_produccion = insumos_dir / "INFORME PRODUCCION.xlsx"
                procesar_ajustes_demograficos(plantilla_resultado, ruta_produccion)
                st.write("✅ **AJUSTES DEMOGRÁFICOS** (Nombres divididos y datos de usuario) procesados correctamente.")
            except Exception as e:
                st.error(f"Error al procesar Ajustes Demográficos: {e}")

            st.success("🎉 ¡Todos los procesos terminaron con éxito!")

            # Descargar la plantilla final
            with open(plantilla_resultado, "rb") as file:
                st.download_button(
                    label="📥 Descargar Plantilla Final y Consolidada",
                    data=file,
                    file_name=f"Plantilla_Procesada_Final_{entidad}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )