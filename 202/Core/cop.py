import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string
from pathlib import Path

def procesar_cop_generico(nombre_eps, ruta_plantilla_in, ruta_plantilla_out):
    """
    Procesa el archivo COP.xlsx según la EPS indicada (NUEVA EPS o SAVIA).
    Mapeo:
      Col C (2) 'Número de identificación del usuario' -> Col E (Plantilla)
      Col L (11) 'Fecha de atención en salud bucal...' -> Col BY (Plantilla)
      Col S (18) 'CALCULO AUTOMATICO VALOR A TRANSFEFRIR' -> Col CY (Plantilla)
    """
    base_dir = Path(__file__).resolve().parent.parent
    ruta_cop = base_dir / "Insumos" / "COP.xlsx"

    if not ruta_cop.exists():
        raise Exception(f"No se encontró el archivo COP.xlsx en Insumos.")

    # SOLUCIÓN: Agregamos dtype=str para que lea todo como texto y no borre los ceros a la izquierda
    try:
        df_cop = pd.read_excel(ruta_cop, sheet_name=nombre_eps, skiprows=12, dtype=str)
    except Exception as e:
        raise Exception(f"Error al leer la hoja '{nombre_eps}' en COP.xlsx: {e}")

    # Limpiamos filas vacías asegurando que la columna 2 (Cédula) tenga datos
    df_cop = df_cop.dropna(subset=[df_cop.columns[2]])

    wb = openpyxl.load_workbook(ruta_plantilla_in)
    ws = wb.active

    # Diccionario con los datos del COP para búsqueda rápida O(1)
    dict_cop = {}
    for _, row in df_cop.iterrows():
        doc_val = row.iloc[2] # Columna C (índice 2)
        
        # Validamos que el documento exista y no sea la palabra "nan"
        if pd.notnull(doc_val) and str(doc_val).strip().lower() != 'nan':
            # Limpieza robusta de cédulas
            try:
                doc = str(int(float(doc_val)))
            except:
                doc = str(doc_val).strip()
            
            # Limpieza del cálculo automático respetando ceros
            calc_val = row.iloc[18]
            calc_final = str(calc_val).strip() if pd.notnull(calc_val) and str(calc_val).strip().lower() != 'nan' else None

            dict_cop[doc] = {
                "fecha_bucal": row.iloc[11], # Columna L
                "calculo": calc_final        # Columna S
            }

    # Búsqueda de fila inicial disponible en la plantilla
    fila_destino_nueva = ws.max_row + 1 if ws.cell(row=2, column=5).value is not None else 2

    # Llenamos la plantilla
    for doc, datos in dict_cop.items():
        # Verificamos si ya existe el usuario
        fila_destino = None
        for r in range(2, ws.max_row + 1):
            cell_val = ws.cell(row=r, column=5).value
            if cell_val is not None:
                try:
                    cell_doc = str(int(float(cell_val)))
                except:
                    cell_doc = str(cell_val).strip()
                    
                if cell_doc == doc:
                    fila_destino = r
                    break
        
        if fila_destino is None:
            fila_destino = fila_destino_nueva
            ws.cell(row=fila_destino, column=5, value=doc)
            fila_destino_nueva += 1

        # Mapeo Columna L -> BY (CON SOLUCIÓN AL ERROR NaT)
        if pd.notnull(datos["fecha_bucal"]) and str(datos["fecha_bucal"]).strip().lower() != 'nan':
            parsed_fec = pd.to_datetime(str(datos["fecha_bucal"]), errors='coerce')
            if pd.notnull(parsed_fec): 
                fec_str = parsed_fec.strftime('%Y/%m/%d')
                ws.cell(row=fila_destino, column=column_index_from_string("BY"), value=fec_str)
        
        # Mapeo Columna S -> CY
        if datos["calculo"] is not None:
            # Forzamos escribirlo como string explícitamente en la celda
            ws.cell(row=fila_destino, column=column_index_from_string("CY"), value=f"{datos['calculo']}")

    wb.save(ruta_plantilla_out)
    wb.close()