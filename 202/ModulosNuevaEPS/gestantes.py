import streamlit as st
from pathlib import Path
import pandas as pd
import openpyxl

def procesar_gestantes_nueva_eps(ruta_plantilla, ruta_salida):
    base_dir = Path(__file__).resolve().parent.parent # Sube a la carpeta raíz
    archivo_gestantes = base_dir / "Insumos" / "GESTANTES.xlsx"

    print("--> Leyendo el archivo GESTANTES.xlsx para Nueva EPS...")
    
    try:
        df_gestantes = pd.read_excel(archivo_gestantes, sheet_name=0)
    except Exception as e:
        print(f"Error al leer GESTANTES.xlsx. Detalles: {e}")
        return False

    # 1. Filtrar la entidad a cualquiera que sea de la NUEVA EPS en la columna NOMBRE ENTIDAD
    if 'NOMBRE ENTIDAD' in df_gestantes.columns:
        df_gestantes['NOMBRE ENTIDAD'] = df_gestantes['NOMBRE ENTIDAD'].astype(str).str.upper()
        df_filtrado = df_gestantes[df_gestantes['NOMBRE ENTIDAD'].str.contains("NUEVA EPS", na=False)].copy()
    else:
        df_filtrado = df_gestantes.copy()

    # Limpiar filas totalmente vacías basadas en el código de usuario
    if 'Codigo_KC' in df_filtrado.columns:
        df_filtrado = df_filtrado.dropna(subset=['Codigo_KC'])
    
    print(f"Total de registros válidos para NUEVA EPS en GESTANTES: {len(df_filtrado)}")

    # Cargar la plantilla con openpyxl
    wb_plantilla = openpyxl.load_workbook(ruta_plantilla)
    ws_plantilla = wb_plantilla.active 

    # Encontrar la siguiente fila disponible en la columna E (Columna 5)
    fila_inicio = 2
    while ws_plantilla.cell(row=fila_inicio, column=5).value is not None:
        fila_inicio += 1

    for idx, row in df_filtrado.reset_index(drop=True).iterrows():
        fila_actual = fila_inicio + idx
        
        # Extracción del número de identificación
        num_id = row.get('Codigo_KC', '')        # Codigo_KC (Columna B del Excel)
        
        # FPP (Columna K del Excel) va en la columna AH de la plantilla (Columna 34)
        fpp = row.get('FPP', '')
        if pd.notnull(fpp):
            parsed_fpp = pd.to_datetime(str(fpp), errors='coerce')
            fpp = parsed_fpp.strftime('%Y/%m/%d') if pd.notnull(parsed_fpp) else ""
        else:
            fpp = ""

        # Fecha_HC (Columna F del Excel) va en la columna BG de la plantilla (Columna 59)
        fecha_hc = row.get('Fecha HC', '')
        if pd.notnull(fecha_hc):
            parsed_hc = pd.to_datetime(str(fecha_hc), errors='coerce')
            fecha_hc = parsed_hc.strftime('%Y/%m/%d') if pd.notnull(parsed_hc) else ""
        else:
            fecha_hc = ""

        # ==========================================
        # MAPEO ESTRICTO A LAS COLUMNAS DE LA PLANTILLA:
        # ==========================================
        ws_plantilla.cell(row=fila_actual, column=5, value=num_id)       # E: Número de Identificación (Codigo_KC)
        
        # Fechas específicas en sus columnas de destino:
        ws_plantilla.cell(row=fila_actual, column=34, value=fpp)         # AH: FPP 
        ws_plantilla.cell(row=fila_actual, column=59, value=fecha_hc)    # BG: Fecha_HC 

        # Valores fijos de 1 solicitados:
        ws_plantilla.cell(row=fila_actual, column=15, value=1)           # O: Valor de 1
        ws_plantilla.cell(row=fila_actual, column=60, value=1)           # BH: Valor de 1
        ws_plantilla.cell(row=fila_actual, column=61, value=1)           # BI: Valor de 1
        ws_plantilla.cell(row=fila_actual, column=62, value=1)           # BJ: Valor de 1

    # Guardar cambios
    wb_plantilla.save(ruta_salida)
    print("¡Proceso de GESTANTES finalizado con éxito!")
    return True


def ejecutar_modulo_gestantes():
    st.title("🤰 Módulo Gestantes - Nueva EPS")
    st.write("Procesa el archivo de gestantes filtrando exclusivamente por Nueva EPS y mapeando las columnas correspondientes.")

    if st.button("Procesar Gestantes Nueva EPS", type="primary"):
        with st.spinner("Procesando gestantes de Nueva EPS..."):
            base_dir = Path(__file__).resolve().parent.parent
            plantilla = base_dir / "Insumos" / "Plantilla.xlsx"
            salida = base_dir / "Insumos" / "Plantilla_Procesada_Gestantes_NuevaEPS.xlsx"
            
            exito = procesar_gestantes_nueva_eps(plantilla, salida)
            
            if exito:
                st.success("¡Gestantes de Nueva EPS procesado correctamente!")