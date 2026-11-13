import streamlit as st
from vistas import (
    mostrar_centro_carga_global,
    ejecutar_modulo_informe_general,
    ejecutar_modulo_informe_produccion,
    ejecutar_vista_consolidacion,
)

st.set_page_config(
    page_title="Sistema de Gestión y Procesamiento EPS - Formato 202",
    page_icon="🏥",
    layout="wide",
)


def main():
    st.sidebar.title("🏥 Navegación del Sistema")

    opcion_principal = st.sidebar.radio(
        "Menú Principal",
        ["Cargar Archivos", "Procesar EPS"],
    )

    st.sidebar.markdown("---")

    if opcion_principal == "Cargar Archivos":
        mostrar_centro_carga_global()

    elif opcion_principal == "Procesar EPS":
        eps_seleccionada = st.sidebar.selectbox(
            "Seleccione la EPS a procesar:",
            ["Nueva EPS", "Savia Salud"],
        )

        codigo_eps = "NUEVA" if eps_seleccionada == "Nueva EPS" else "SAVIA"

        st.sidebar.markdown("---")
        st.sidebar.subheader(f"Módulos - {eps_seleccionada}")

        modulo = st.sidebar.selectbox(
            "Seleccione el Módulo:",
            [
                "Informe General",
                "Informe de Producción",
                "Consolidación Final",
            ],
        )

        if modulo == "Informe General":
            ejecutar_modulo_informe_general(codigo_eps)
        elif modulo == "Informe de Producción":
            ejecutar_modulo_informe_produccion(codigo_eps)
        elif modulo == "Consolidación Final":
            ejecutar_vista_consolidacion(codigo_eps)


if __name__ == "__main__":
    main()
