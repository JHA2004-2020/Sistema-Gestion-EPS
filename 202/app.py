import streamlit as st
import os
from pathlib import Path

# Importaciones desde Core
from Core.informe_general import ejecutar_modulo_informe_general
from Core.informe_produccion import ejecutar_modulo_informe_produccion

st.set_page_config(
    page_title="Sistema de Gestión y Procesamiento EPS",
    page_icon="🏥",
    layout="wide",
)

def guardar_insumo(uploaded_file, nombre_estricto, insumos_dir):
    """
    Toma cualquier archivo Excel subido por el usuario y lo guarda estrictamente 
    con el 'nombre_estricto' (ej. 'INFORME GENERAL.xlsx') dentro de la carpeta Insumos,
    eliminando previamente la versión del mes anterior.
    """
    if uploaded_file is not None:
        ruta_destino = insumos_dir / nombre_estricto
        
        # Eliminamos el archivo anterior si existe para garantizar un reemplazo limpio
        if ruta_destino.exists():
            try:
                ruta_destino.unlink()
            except Exception as e:
                st.warning(f"No se pudo eliminar el archivo anterior {nombre_estricto}: {e}")

        # Guardamos el nuevo archivo
        with open(ruta_destino, "wb") as f:
            f.write(uploaded_file.getbuffer())
        st.success(f"✅ Archivo recibido. Reemplazado en el sistema como: **{nombre_estricto}**")

def mostrar_centro_carga_global():
    """Muestra un panel único de carga de archivos para todo el sistema."""
    st.title("📥 Centro de Carga Global de Insumos")
    st.markdown("---")
    st.info("Sube aquí los archivos correspondientes al mes actual. Al subir un archivo nuevo, este reemplazará automáticamente al del mes anterior.")

    base_dir = Path(__file__).resolve().parent
    insumos_dir = base_dir / "Insumos"
    insumos_dir.mkdir(parents=True, exist_ok=True)

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### Archivos Base y Cruzados")
        # Se eliminó la carga de la Plantilla, ya que es un archivo base del sistema.
        
        f_cop = st.file_uploader("Subir archivo COP", type=["xlsx"], key="up_cop")
        guardar_insumo(f_cop, "COP.xlsx", insumos_dir)

        f_hc = st.file_uploader("Subir INFORME HC (Pesos y Tallas)", type=["xlsx"], key="up_hc")
        guardar_insumo(f_hc, "INFORME HC.xlsx", insumos_dir)

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
        f_gest_nueva = st.file_uploader("Subir Gestantes (NUEVA EPS)", type=["xlsx"], key="up_gest_nueva")
        guardar_insumo(f_gest_nueva, "GESTANTES.xlsx", insumos_dir)
    with col4:
        f_gest_savia = st.file_uploader("Subir Gestantes (SAVIA SALUD)", type=["xlsx"], key="up_gest_savia")
        guardar_insumo(f_gest_savia, "GESTANTES SAVIA.xlsx", insumos_dir)


def main():
    st.sidebar.title("Navegación del Sistema")

    # 1. Menú Principal Global
    opcion_principal = st.sidebar.radio(
        "Menú Principal",
        ["Cargar Archivos", "Procesar EPS"]
    )

    st.sidebar.markdown("---")

    if opcion_principal == "Cargar Archivos":
        mostrar_centro_carga_global()
        
    elif opcion_principal == "Procesar EPS":
        # 2. Selección de EPS solo si están en la sección de procesar
        eps_seleccionada = st.sidebar.selectbox(
            "Seleccione la EPS a procesar:", ["Nueva EPS", "Savia Salud"]
        )
        
        st.sidebar.markdown("---")

        # 3. Selección de Módulo según la EPS
        if eps_seleccionada == "Nueva EPS":
            st.sidebar.subheader("Módulos - Nueva EPS")
            modulo = st.sidebar.selectbox(
                "Seleccione el Módulo", ["Informe General", "Informe de Producción"]
            )
                
            if modulo == "Informe General":
                ejecutar_modulo_informe_general("NUEVA")  
                
            elif modulo == "Informe de Producción":
                ejecutar_modulo_informe_produccion("NUEVA")  

        elif eps_seleccionada == "Savia Salud":
            st.sidebar.subheader("Módulos - Savia Salud")
            modulo = st.sidebar.selectbox(
                "Seleccione el Módulo", ["Informe General", "Informe de Producción"]
            )
                
            if modulo == "Informe General":
                ejecutar_modulo_informe_general("SAVIA")  
                
            elif modulo == "Informe de Producción":
                ejecutar_modulo_informe_produccion("SAVIA")  

if __name__ == "__main__":
    main()