import streamlit as st
from pathlib import Path


ARCHIVOS_SISTEMA = [
    ("Plantilla.xlsx", "Plantilla Base"),
    ("COP.xlsx", "Archivo COP"),
    ("INFORME HC.xlsx", "Informe HC (Pesos y Tallas)"),
    ("PSA.xlsx", "Archivo PSA"),
    ("INFORME GENERAL.xlsx", "Informe General"),
    ("INFORME PRODUCCION.xlsx", "Informe de Producción"),
    ("GESTANTES.xlsx", "Gestantes (Nueva EPS)"),
    ("GESTANTES SAVIA.xlsx", "Gestantes (Savia Salud)"),
]


def guardar_insumo(uploaded_file, nombre_estricto, insumos_dir):
    """
    Guarda el archivo Excel subido por el usuario con el 'nombre_estricto'
    dentro de la carpeta Insumos, reemplazando la versión anterior.
    """
    if uploaded_file is not None:
        ruta_destino = insumos_dir / nombre_estricto

        if ruta_destino.exists():
            try:
                ruta_destino.unlink()
            except Exception as e:
                st.warning(f"No se pudo eliminar el archivo anterior {nombre_estricto}: {e}")

        with open(ruta_destino, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.success(f"✅ Archivo recibido. Reemplazado en el sistema como: **{nombre_estricto}**")


def mostrar_centro_carga_global():
    """Muestra un panel único de carga de archivos e indicadores de estado para todo el sistema."""
    st.title("📥 Centro de Carga Global de Insumos")
    st.markdown("---")
    st.info(
        "Sube aquí los archivos correspondientes al mes actual. "
        "Al subir un archivo nuevo, este reemplazará automáticamente al del mes anterior."
    )

    base_dir = Path(__file__).resolve().parent.parent
    insumos_dir = base_dir / "Insumos"
    insumos_dir.mkdir(parents=True, exist_ok=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### Archivos Base y Cruzados")

        f_cop = st.file_uploader("Subir archivo COP", type=["xlsx"], key="up_cop")
        guardar_insumo(f_cop, "COP.xlsx", insumos_dir)

        f_hc = st.file_uploader("Subir INFORME HC (Pesos y Tallas)", type=["xlsx"], key="up_hc")
        guardar_insumo(f_hc, "INFORME HC.xlsx", insumos_dir)

        f_psa = st.file_uploader("Subir archivo PSA", type=["xlsx"], key="up_psa")
        guardar_insumo(f_psa, "PSA.xlsx", insumos_dir)

    with col2:
        st.markdown("#### Informes Generales y EPS")

        f_general = st.file_uploader("Subir INFORME GENERAL", type=["xlsx"], key="up_gral")
        guardar_insumo(f_general, "INFORME GENERAL.xlsx", insumos_dir)

        f_prod = st.file_uploader("Subir INFORME PRODUCCIÓN", type=["xlsx"], key="up_prod")
        guardar_insumo(f_prod, "INFORME PRODUCCION.xlsx", insumos_dir)

    st.markdown("---")
    st.markdown("#### Archivos de Gestantes")
    st.markdown("Sube los archivos específicos de gestantes según la entidad a procesar:")

    col3, col4 = st.columns(2)
    with col3:
        f_gest_nueva = st.file_uploader(
            "Subir Gestantes (NUEVA EPS)", type=["xlsx"], key="up_gest_nueva"
        )
        guardar_insumo(f_gest_nueva, "GESTANTES.xlsx", insumos_dir)
    with col4:
        f_gest_savia = st.file_uploader(
            "Subir Gestantes (SAVIA SALUD)", type=["xlsx"], key="up_gest_savia"
        )
        guardar_insumo(f_gest_savia, "GESTANTES SAVIA.xlsx", insumos_dir)

    st.markdown("---")
    with st.expander("📋 Ver estado actual de archivos en carpeta Insumos", expanded=False):
        cols_estado = st.columns(4)
        for idx, (archivo_nombre, descripcion) in enumerate(ARCHIVOS_SISTEMA):
            existe = (insumos_dir / archivo_nombre).exists()
            icono = "🟢" if existe else "⚪"
            estado_txt = "Disponible" if existe else "No cargado"
            with cols_estado[idx % 4]:
                st.markdown(f"{icono} **{descripcion}**  \n`{archivo_nombre}` ({estado_txt})")
