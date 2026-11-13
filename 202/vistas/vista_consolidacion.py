from pathlib import Path
import streamlit as st
from procesadores.informe_general import cargar_informe_general
from procesadores.informe_produccion import crear_config_produccion_vacia
from procesadores.orquestador import ejecutar_consolidacion_completa
from vistas.vista_general import obtener_llaves_sesion_general
from vistas.vista_produccion import obtener_llaves_sesion_produccion


def ejecutar_vista_consolidacion(entidad):
    entidad_upper = str(entidad).strip().upper()
    st.title(f"🚀 Consolidación Final Formato 202 - {entidad_upper}")
    st.markdown("---")

    base_dir = Path(__file__).resolve().parent.parent
    insumos_dir = base_dir / "Insumos"

    key_ent_gral, key_cfg_gral = obtener_llaves_sesion_general(entidad_upper)
    key_ent_prod, key_cfg_prod = obtener_llaves_sesion_produccion(entidad_upper)

    config_general = st.session_state.get(key_cfg_gral, {})
    entidades_general = st.session_state.get(key_ent_gral, [])

    config_prod = st.session_state.get(key_cfg_prod, crear_config_produccion_vacia())
    entidades_prod = st.session_state.get(key_ent_prod, [])
    total_filtros_prod = sum(len(v) for v in config_prod.values())

    col1, col2 = st.columns(2)
    with col1:
        st.markdown("#### 📊 Resumen Informe General")
        st.write(f"- **Entidades seleccionadas:** {len(entidades_general)}")
        st.write(f"- **Pruebas configuradas:** {len(config_general)}")
        if not config_general:
            st.caption("⚠️ Puedes configurar o cargar el JSON en el módulo *Informe General*.")

    with col2:
        st.markdown("#### 🏭 Resumen Informe de Producción")
        st.write(f"- **Entidades seleccionadas:** {len(entidades_prod)}")
        st.write(f"- **Filtros en las 7 secciones:** {total_filtros_prod}")
        if total_filtros_prod == 0:
            st.caption(
                "⚠️ Puedes configurar o cargar el JSON en el módulo *Informe de Producción*."
            )

    st.markdown("---")
    st.markdown(
        "Al ejecutar la consolidación se aplicarán en orden: "
        "**1. COP → 2. PSA → 3. Informe General → 4. Gestantes → "
        "5. Informe de Producción → 6. Informe HC (Pesos y Tallas) → 7. Ajustes Demográficos**."
    )

    if st.button(
        "✨ Generar Plantilla Procesada Final",
        type="primary",
        use_container_width=True,
        key=f"btn_consolidar_vista_{entidad_upper}",
    ):
        df_general = cargar_informe_general()
        df_filtrado = None
        if df_general is not None:
            df_general = df_general.copy()
            if "EPS" in df_general.columns:
                df_general["EPS"] = df_general["EPS"].astype(str).str.upper().str.strip()
                if entidades_general:
                    df_filtrado = df_general[df_general["EPS"].isin(entidades_general)].copy()
                else:
                    df_filtrado = df_general[
                        df_general["EPS"].str.contains(entidad_upper, na=False)
                    ].copy()
            else:
                df_filtrado = df_general

        ruta_final = ejecutar_consolidacion_completa(
            entidad=entidad_upper,
            insumos_dir=insumos_dir,
            df_general_filtrado=df_filtrado,
            config_general=config_general,
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
