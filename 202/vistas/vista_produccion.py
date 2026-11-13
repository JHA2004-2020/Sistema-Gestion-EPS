import json
import pandas as pd
import streamlit as st
from procesadores.informe_produccion import (
    cargar_informe_produccion,
    crear_config_produccion_vacia,
    SECCIONES_PRODUCCION,
)


def obtener_llaves_sesion_produccion(entidad):
    ent_key = str(entidad).strip().upper().replace(" ", "_")
    return (
        f"filtro_entidades_prod_{ent_key}",
        f"config_prod_{ent_key}",
    )


def _render_seccion_simple(
    titulo,
    clave_seccion,
    etiqueta_corto,
    config,
    articulos_unicos,
    entidad_key,
    item_extra=None,
):
    """Renderiza un bloque estándar de selección y eliminación de artículos para una sección."""
    st.subheader(titulo)
    c1, c2 = st.columns([1, 1.5])

    existentes = [x["filtro"] for x in config[clave_seccion]]
    with c1:
        disponibles = [a for a in articulos_unicos if a not in existentes]
        seleccion = st.selectbox(
            f"🔍 Agregar filtro para {etiqueta_corto}:",
            ["-- Selecciona --"] + disponibles,
            key=f"sel_{clave_seccion}_{entidad_key}",
        )
        if seleccion != "-- Selecciona --":
            nuevo_registro = {"filtro": seleccion}
            if item_extra:
                nuevo_registro.update(item_extra)
            config[clave_seccion].append(nuevo_registro)
            st.rerun()

    with c2:
        if config[clave_seccion]:
            df_sec = pd.DataFrame(config[clave_seccion])
            st.data_editor(
                df_sec,
                hide_index=True,
                use_container_width=True,
                key=f"ed_{clave_seccion}_{entidad_key}",
            )
            item_eliminar = st.selectbox(
                f"Selecciona ítem para eliminar ({etiqueta_corto}):",
                options=[x["filtro"] for x in config[clave_seccion]],
                key=f"item_del_{clave_seccion}_{entidad_key}",
            )
            if st.button(
                f"🗑️ Eliminar Filtro {etiqueta_corto}",
                type="secondary",
                use_container_width=True,
                key=f"btn_del_{clave_seccion}_{entidad_key}",
            ):
                config[clave_seccion] = [
                    x for x in config[clave_seccion] if x["filtro"] != item_eliminar
                ]
                st.rerun()
        else:
            st.info("No hay filtros agregados.")

    st.markdown("---")


def ejecutar_modulo_informe_produccion(entidad_sugerida):
    entidad_upper = str(entidad_sugerida).strip().upper()
    st.title(f"🏭 Módulo Informe de Producción - {entidad_upper}")
    st.markdown("---")

    df_general = cargar_informe_produccion()
    if df_general is None:
        st.error("No se encontró el archivo 'INFORME PRODUCCION.xlsx' en la carpeta Insumos.")
        return

    df_general = df_general.copy()
    col_entidad = df_general.columns[34] if len(df_general.columns) > 34 else "ENTIDAD"
    col_articulo = df_general.columns[14] if len(df_general.columns) > 14 else "NOMBRE_ARTICULO"

    df_general[col_entidad] = df_general[col_entidad].apply(
        lambda x: str(x).upper().strip() if pd.notnull(x) else ""
    )

    entidades_unicas = sorted(df_general[col_entidad][df_general[col_entidad] != ""].unique())
    articulos_unicos = sorted(
        df_general[col_articulo].dropna().astype(str).str.strip().unique()
    )

    key_entidades, key_config = obtener_llaves_sesion_produccion(entidad_upper)

    # ==========================================
    # SECCIÓN 1: FILTRO DESPLEGABLE DE ENTIDAD
    # ==========================================
    st.subheader("🏢 1. Filtro de Entidades y Derivados")
    st.markdown("Selecciona exactamente cuáles derivados de la entidad deseas cruzar en este informe:")

    entidades_defecto = [e for e in entidades_unicas if entidad_upper in e]
    if key_entidades not in st.session_state:
        st.session_state[key_entidades] = entidades_defecto

    default_validos = [e for e in st.session_state[key_entidades] if e in entidades_unicas]

    entidades_seleccionadas = st.multiselect(
        "Despliega y selecciona las entidades:",
        options=entidades_unicas,
        default=default_validos,
        key=f"ms_entidades_prod_{entidad_upper}",
    )
    st.session_state[key_entidades] = entidades_seleccionadas

    if entidades_seleccionadas:
        df_filtrado = df_general[df_general[col_entidad].isin(entidades_seleccionadas)].copy()
        st.success(
            f"📊 Registros totales encontrados para las entidades seleccionadas: {len(df_filtrado)}"
        )
    else:
        st.warning("⚠️ Debes seleccionar al menos una entidad para procesar.")

    st.markdown("---")

    # ==========================================
    # GESTIÓN DE CONFIGURACIÓN EN MEMORIA Y JSON
    # ==========================================
    if key_config not in st.session_state:
        st.session_state[key_config] = crear_config_produccion_vacia()

    config = st.session_state[key_config]
    for sec in SECCIONES_PRODUCCION:
        if sec not in config:
            config[sec] = []

    st.sidebar.markdown("### 💾 Gestión de Configuración")
    json_str = json.dumps(config, indent=4, ensure_ascii=False)
    st.sidebar.download_button(
        label="💾 Guardar Configuración Actual (.json)",
        data=json_str,
        file_name=f"filtros_informe_produccion_{entidad_upper}.json",
        mime="application/json",
        use_container_width=True,
    )

    uploaded_config = st.sidebar.file_uploader(
        "📂 Cargar configuración guardada (.json)",
        type=["json"],
        key=f"up_json_prod_{entidad_upper}",
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

    # ==========================================
    # SECCIÓN BA
    # ==========================================
    _render_seccion_simple(
        titulo="2️⃣ Sección Columna BA (Fecha de consulta de valoración integral)",
        clave_seccion="BA_seccion",
        etiqueta_corto="BA",
        config=config,
        articulos_unicos=articulos_unicos,
        entidad_key=entidad_upper,
        item_extra={"pos_resultado": "BA"},
    )

    # ==========================================
    # SECCIÓN BC y BD (Editable en valor_bc)
    # ==========================================
    st.subheader("3️⃣ Sección Columnas BC y BD")
    d1, d2 = st.columns([1, 1.5])
    with d1:
        disp_bc = [
            a for a in articulos_unicos if a not in [x["filtro"] for x in config["BC_BD_seccion"]]
        ]
        sel_bc = st.selectbox(
            "🔍 Agregar filtro para BC/BD:",
            ["-- Selecciona --"] + disp_bc,
            key=f"sel_bc_{entidad_upper}",
        )
        if sel_bc != "-- Selecciona --":
            config["BC_BD_seccion"].append(
                {
                    "filtro": sel_bc,
                    "valor_bc": "23",
                    "pos_resultado_bc": "BC",
                    "pos_resultado_bd": "BD",
                }
            )
            st.rerun()
    with d2:
        if config["BC_BD_seccion"]:
            df_bc_bd = pd.DataFrame(config["BC_BD_seccion"])
            edited_df_bc_bd = st.data_editor(
                df_bc_bd,
                hide_index=True,
                use_container_width=True,
                key=f"ed_bc_bd_{entidad_upper}",
            )
            config["BC_BD_seccion"] = edited_df_bc_bd.to_dict("records")
            item_bc_sel = st.selectbox(
                "Selecciona ítem para eliminar (BC/BD):",
                options=[x["filtro"] for x in config["BC_BD_seccion"]],
                key=f"item_bc_sel_{entidad_upper}",
            )
            if st.button(
                "🗑️ Eliminar Filtro BC/BD",
                type="secondary",
                use_container_width=True,
                key=f"btn_del_bc_{entidad_upper}",
            ):
                config["BC_BD_seccion"] = [
                    x for x in config["BC_BD_seccion"] if x["filtro"] != item_bc_sel
                ]
                st.rerun()
        else:
            st.info("No hay filtros agregados.")

    st.markdown("---")

    # ==========================================
    # SECCIONES RESTANTES (AR-AU, BB, BQ, DI/DJ, AY/AX)
    # ==========================================
    _render_seccion_simple(
        titulo="4️⃣ Sección Columnas AR, AS, AT, AU (Valor fijo = 5)",
        clave_seccion="AR_AU_seccion",
        etiqueta_corto="AR-AU",
        config=config,
        articulos_unicos=articulos_unicos,
        entidad_key=entidad_upper,
    )

    _render_seccion_simple(
        titulo="5️⃣ Sección Columna BB (Fecha)",
        clave_seccion="BB_seccion",
        etiqueta_corto="BB",
        config=config,
        articulos_unicos=articulos_unicos,
        entidad_key=entidad_upper,
    )

    _render_seccion_simple(
        titulo="6️⃣ Sección Columna BQ (Fecha) y asociadas (CJ=Fecha, CI=1, CK=17, CL=1, CM=053610450501)",
        clave_seccion="BQ_seccion",
        etiqueta_corto="BQ",
        config=config,
        articulos_unicos=articulos_unicos,
        entidad_key=entidad_upper,
    )

    _render_seccion_simple(
        titulo="7️⃣ Sección Columnas DI y DJ (Fecha)",
        clave_seccion="DI_DJ_seccion",
        etiqueta_corto="DI/DJ",
        config=config,
        articulos_unicos=articulos_unicos,
        entidad_key=entidad_upper,
    )

    _render_seccion_simple(
        titulo="8️⃣ Sección Columna AY (Fecha) y AX (Fecha - 1 día)",
        clave_seccion="AY_AX_seccion",
        etiqueta_corto="AY/AX",
        config=config,
        articulos_unicos=articulos_unicos,
        entidad_key=entidad_upper,
    )

    st.info(
        "💡 **Nota:** Puedes generar el archivo final desde la sección **Consolidación Final** "
        "o directamente al final de **Informe General**."
    )
