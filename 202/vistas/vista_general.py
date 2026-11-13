import json
from pathlib import Path
import pandas as pd
import streamlit as st
from procesadores.informe_general import cargar_informe_general
from procesadores.informe_produccion import crear_config_produccion_vacia
from procesadores.orquestador import ejecutar_consolidacion_completa


def obtener_llaves_sesion_general(entidad):
    ent_key = str(entidad).strip().upper().replace(" ", "_")
    return (
        f"filtro_entidades_general_{ent_key}",
        f"config_dict_general_{ent_key}",
    )


def ejecutar_modulo_informe_general(entidad):
    entidad_upper = str(entidad).strip().upper()
    st.title(f"📊 Módulo Informe General - {entidad_upper}")
    st.markdown("---")

    base_dir = Path(__file__).resolve().parent.parent
    insumos_dir = base_dir / "Insumos"

    df_general = cargar_informe_general()
    if df_general is None:
        st.error("No se encontró el archivo 'INFORME GENERAL.xlsx' en la carpeta Insumos.")
        return

    df_general = df_general.copy()
    if "EPS" in df_general.columns:
        df_general["EPS"] = df_general["EPS"].astype(str).str.upper().str.strip()
        entidades_unicas = sorted(df_general["EPS"].dropna().unique())
    else:
        entidades_unicas = []
        df_general["EPS"] = "N/A"

    key_entidades, key_config = obtener_llaves_sesion_general(entidad_upper)

    # ==========================================
    # 1. FILTRO DESPLEGABLE DE ENTIDAD
    # ==========================================
    st.subheader("🏢 1. Filtro de Entidades y Derivados")
    st.markdown("Selecciona los derivados de la entidad para el Informe General:")

    entidades_defecto = [e for e in entidades_unicas if entidad_upper in e]
    if key_entidades not in st.session_state:
        st.session_state[key_entidades] = entidades_defecto

    # Filtramos valores por defecto que sigan existiendo en el archivo actual
    default_validos = [e for e in st.session_state[key_entidades] if e in entidades_unicas]

    entidades_seleccionadas = st.multiselect(
        "Despliega y selecciona las entidades:",
        options=entidades_unicas,
        default=default_validos,
        key=f"ms_entidades_general_{entidad_upper}",
    )
    st.session_state[key_entidades] = entidades_seleccionadas

    if entidades_seleccionadas:
        df_filtrado = df_general[df_general["EPS"].isin(entidades_seleccionadas)].copy()
        st.success(
            f"📊 Registros totales encontrados para las entidades seleccionadas: {len(df_filtrado)}"
        )
    else:
        df_filtrado = df_general.copy()
        st.warning(
            "⚠️ No has seleccionado entidades. El sistema usará la base completa (no recomendado)."
        )

    if "PRUEBA" in df_general.columns:
        pruebas_unicas = sorted(df_general["PRUEBA"].dropna().astype(str).unique())
    else:
        pruebas_unicas = []

    if key_config not in st.session_state:
        st.session_state[key_config] = {}

    config_actual = st.session_state[key_config]

    # ==========================================
    # BARRA LATERAL: GUARDAR Y CARGAR CONFIGURACIÓN JSON
    # ==========================================
    st.sidebar.header("💾 Gestión de Configuración")

    json_str = json.dumps(config_actual, indent=4, ensure_ascii=False)
    st.sidebar.download_button(
        label="💾 Guardar Configuración Actual (.json)",
        data=json_str,
        file_name=f"filtros_informe_general_{entidad_upper}.json",
        mime="application/json",
        use_container_width=True,
    )

    uploaded_config = st.sidebar.file_uploader(
        "📂 Cargar configuración guardada (.json)",
        type=["json"],
        key=f"up_json_general_{entidad_upper}",
    )
    if uploaded_config is not None:
        try:
            content = uploaded_config.getvalue()
            loaded_data = json.loads(content.decode("utf-8"))
            if loaded_data != st.session_state[key_config]:
                st.session_state[key_config] = loaded_data
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

        pruebas_disponibles = [p for p in pruebas_unicas if p not in config_actual]
        prueba_a_agregar = st.selectbox(
            "🔍 Buscar y seleccionar prueba:",
            options=["-- Selecciona una prueba --"] + pruebas_disponibles,
            key=f"sel_agregar_prueba_{entidad_upper}",
        )

        if prueba_a_agregar != "-- Selecciona una prueba --":
            config_actual[prueba_a_agregar] = {
                "Posición Plantilla": "AA",
                "Columna Fecha Plantilla": "AA",
            }
            st.rerun()

    with col2:
        st.markdown("### 3. Tabla de Configuración, Prioridades y Destinos")
        st.markdown("Define columnas y ajusta la **prioridad (arriba = mayor prioridad)**:")

        if len(config_actual) > 0:
            keys_list = list(config_actual.keys())
            lista_tabla = [
                {
                    "Filtro": k,
                    "Posición Plantilla (Resultado)": config_actual[k]["Posición Plantilla"],
                    "Posición Plantilla (Fecha opcional)": config_actual[k].get(
                        "Columna Fecha Plantilla", ""
                    ),
                }
                for k in keys_list
            ]

            df_config = pd.DataFrame(lista_tabla)
            edited_df = st.data_editor(
                df_config,
                hide_index=True,
                use_container_width=True,
                key=f"editor_config_tabla_{entidad_upper}",
            )

            for _, row in edited_df.iterrows():
                p_name = row["Filtro"]
                if p_name in config_actual:
                    config_actual[p_name] = {
                        "Posición Plantilla": row["Posición Plantilla (Resultado)"],
                        "Columna Fecha Plantilla": row["Posición Plantilla (Fecha opcional)"],
                    }

            st.markdown("---")
            st.markdown("### ↕️ Ajustar Prioridades (Subir / Bajar) o Eliminar")

            item_seleccionado = st.selectbox(
                "Selecciona una prueba para mover o borrar:",
                options=keys_list,
                key=f"select_gestionar_{entidad_upper}",
            )

            c_subir, c_bajar, c_borrar = st.columns(3)
            with c_subir:
                if st.button("⬆️ Subir Prioridad", use_container_width=True):
                    idx = keys_list.index(item_seleccionado)
                    if idx > 0:
                        keys_list[idx], keys_list[idx - 1] = keys_list[idx - 1], keys_list[idx]
                        st.session_state[key_config] = {k: config_actual[k] for k in keys_list}
                        st.rerun()
            with c_bajar:
                if st.button("⬇️ Bajar Prioridad", use_container_width=True):
                    idx = keys_list.index(item_seleccionado)
                    if idx < len(keys_list) - 1:
                        keys_list[idx], keys_list[idx + 1] = keys_list[idx + 1], keys_list[idx]
                        st.session_state[key_config] = {k: config_actual[k] for k in keys_list}
                        st.rerun()
            with c_borrar:
                if st.button("🗑️ Eliminar", type="secondary", use_container_width=True):
                    if item_seleccionado in config_actual:
                        del config_actual[item_seleccionado]
                        st.rerun()
        else:
            st.info("No hay pruebas seleccionadas. Usa el buscador de la izquierda.")

    st.markdown("---")

    # ==========================================
    # EJECUCIÓN DE LA CONSOLIDACIÓN FINAL
    # ==========================================
    st.subheader("🚀 Ejecución de la Consolidación General")
    st.markdown(
        "Se procesará aplicando **COP, PSA, Informe General, Gestantes, Informe de Producción, "
        "HC (Pesos y Tallas) y Ajustes Demográficos** en orden secuencial."
    )

    if st.button(
        "✨ Generar Plantilla Procesada Final",
        type="primary",
        use_container_width=True,
        key=f"btn_generar_desde_general_{entidad_upper}",
    ):
        if not entidades_seleccionadas:
            st.error("Por favor, selecciona al menos una entidad en el Paso 1 antes de procesar.")
            return

        key_ent_prod = f"filtro_entidades_prod_{entidad_upper}"
        key_cfg_prod = f"config_prod_{entidad_upper}"
        config_prod = st.session_state.get(key_cfg_prod, crear_config_produccion_vacia())
        entidades_prod = st.session_state.get(key_ent_prod, [])

        ruta_final = ejecutar_consolidacion_completa(
            entidad=entidad_upper,
            insumos_dir=insumos_dir,
            df_general_filtrado=df_filtrado,
            config_general=config_actual,
            config_produccion=config_prod,
            entidades_produccion=entidades_prod,
        )

        if ruta_final and Path(ruta_final).exists():
            with open(ruta_final, "rb") as file:
                st.download_button(
                    label="📥 Descargar Plantilla Final y Consolidada",
                    data=file,
                    file_name=f"Plantilla_Procesada_Final_{entidad_upper}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
