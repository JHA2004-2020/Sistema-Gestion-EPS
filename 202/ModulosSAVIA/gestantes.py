import pandas as pd
import openpyxl
from openpyxl.utils import column_index_from_string
from pathlib import Path

def procesar_gestantes_savia(ruta_plantilla_in, ruta_plantilla_out):
    """
    Procesa GESTANTES SAVIA.xlsx aplicando reglas estrictas de prioridad para CC, CD, CE, CF y CA, CB.
    Adicionalmente, marca con el valor '1' la columna O para todas las gestantes.
    """
    base_dir = Path(__file__).resolve().parent.parent
    ruta_gestantes = base_dir / "Insumos" / "GESTANTES SAVIA.xlsx"

    if not ruta_gestantes.exists():
        raise Exception(f"No se encontró el archivo GESTANTES SAVIA.xlsx en Insumos.")

    df_gest = pd.read_excel(ruta_gestantes)
    
    # Limpiamos nulos en el documento (Columna A, índice 0)
    df_gest = df_gest.dropna(subset=[df_gest.columns[0]])

    wb = openpyxl.load_workbook(ruta_plantilla_in)
    ws = wb.active

    # Búsqueda de fila inicial disponible
    fila_destino_nueva = ws.max_row + 1 if ws.cell(row=2, column=5).value is not None else 2

    for _, row in df_gest.iterrows():
        doc = str(row.iloc[0]).split('.')[0].strip() # Col A (Identificación)
        if not doc: continue

        # Buscar si el usuario ya existe en la plantilla
        fila_destino = None
        for r in range(2, ws.max_row + 1):
            if str(ws.cell(row=r, column=5).value).split('.')[0].strip() == doc:
                fila_destino = r
                break
        
        if fila_destino is None:
            fila_destino = fila_destino_nueva
            ws.cell(row=fila_destino, column=5, value=doc)
            fila_destino_nueva += 1

        # -------------------------------------------------------------
        # NUEVA REGLA: Marcar con 1 en la Columna O para todas las gestantes
        # -------------------------------------------------------------
        ws.cell(row=fila_destino, column=column_index_from_string("O"), value=1)
        ws.cell(row=fila_destino, column=column_index_from_string("BH"), value=1)
        ws.cell(row=fila_destino, column=column_index_from_string("BI"), value=1)
        ws.cell(row=fila_destino, column=column_index_from_string("BJ"), value=1)
        # -------------------------------------------------------------
        # REGLAS DIRECTAS
        # -------------------------------------------------------------
        # Col C (2) -> AH (FPP)
        if pd.notnull(row.iloc[2]):
            fec = pd.to_datetime(str(row.iloc[2]), errors='coerce').strftime('%Y/%m/%d')
            if fec != "NaT": ws.cell(row=fila_destino, column=column_index_from_string("AH"), value=fec)
        
        # Col D (3) -> AJ (Riesgo)
        if pd.notnull(row.iloc[3]):
            ws.cell(row=fila_destino, column=column_index_from_string("AJ"), value=row.iloc[3])

        # Col Q (16) -> CA (Hepatitis B Fecha)
        if pd.notnull(row.iloc[16]):
            fec = pd.to_datetime(str(row.iloc[16]), errors='coerce').strftime('%Y/%m/%d')
            if fec != "NaT": ws.cell(row=fila_destino, column=column_index_from_string("CA"), value=fec)

        # Col R (17) -> CB (Hepatitis B Resultado)
        if pd.notnull(row.iloc[17]):
            ws.cell(row=fila_destino, column=column_index_from_string("CB"), value=row.iloc[17])

        if pd.notnull(row.iloc[1]):
            # Convertimos el valor a fecha de forma segura y extraemos solo día, mes y año
            parsed_fec = pd.to_datetime(str(row.iloc[1]), errors='coerce')
            if pd.notnull(parsed_fec): ws.cell(row=fila_destino, column=column_index_from_string("BE"), value=parsed_fec.date())
        # -------------------------------------------------------------
        # REGLAS CON PRIORIDAD EN CASCADA
        # -------------------------------------------------------------
        
        # Cascada 1: Sífilis (Resultados J->H->F en CC) y (Fechas I->G->E en CD)
        res_j, res_h, res_f = row.iloc[9], row.iloc[7], row.iloc[5]
        fec_i, fec_g, fec_e = row.iloc[8], row.iloc[6], row.iloc[4]

        # Prioridad 1: Trimestre III (J, I)
        if pd.notnull(res_j) and str(res_j).strip() != "21" and str(res_j).strip() != "":
            ws.cell(row=fila_destino, column=column_index_from_string("CD"), value=res_j)
            if pd.notnull(fec_i): ws.cell(row=fila_destino, column=column_index_from_string("CC"), value=pd.to_datetime(str(fec_i), errors='coerce').strftime('%Y/%m/%d'))
        
        # Prioridad 2: Trimestre II (H, G)
        elif pd.notnull(res_h) and str(res_h).strip() != "21" and str(res_h).strip() != "":
            ws.cell(row=fila_destino, column=column_index_from_string("CD"), value=res_h)
            if pd.notnull(fec_g): ws.cell(row=fila_destino, column=column_index_from_string("CC"), value=pd.to_datetime(str(fec_g), errors='coerce').strftime('%Y/%m/%d'))
            
        # Prioridad 3: Trimestre I (F, E)
        elif pd.notnull(res_f) and str(res_f).strip() != "21" and str(res_f).strip() != "":
            ws.cell(row=fila_destino, column=column_index_from_string("CD"), value=res_f)
            if pd.notnull(fec_e): ws.cell(row=fila_destino, column=column_index_from_string("CC"), value=pd.to_datetime(str(fec_e), errors='coerce').strftime('%Y/%m/%d'))

        # Cascada 2: VIH (Resultados P->N->L en CF) y (Fechas O->M->K en CE)
        res_p, res_n, res_l = row.iloc[15], row.iloc[13], row.iloc[11]
        fec_o, fec_m, fec_k = row.iloc[14], row.iloc[12], row.iloc[10]

        # Prioridad 1: Trimestre III (P, O)
        if pd.notnull(res_p) and str(res_p).strip() != "21" and str(res_p).strip() != "":
            ws.cell(row=fila_destino, column=column_index_from_string("CF"), value=res_p)
            if pd.notnull(fec_o): ws.cell(row=fila_destino, column=column_index_from_string("CE"), value=pd.to_datetime(str(fec_o), errors='coerce').strftime('%Y/%m/%d'))
        
        # Prioridad 2: Trimestre II (N, M)
        elif pd.notnull(res_n) and str(res_n).strip() != "21" and str(res_n).strip() != "":
            ws.cell(row=fila_destino, column=column_index_from_string("CF"), value=res_n)
            if pd.notnull(fec_m): ws.cell(row=fila_destino, column=column_index_from_string("CE"), value=pd.to_datetime(str(fec_m), errors='coerce').strftime('%Y/%m/%d'))
            
        # Prioridad 3: Trimestre I (L, K)
        elif pd.notnull(res_l) and str(res_l).strip() != "21" and str(res_l).strip() != "":
            ws.cell(row=fila_destino, column=column_index_from_string("CF"), value=res_l)
            if pd.notnull(fec_k): ws.cell(row=fila_destino, column=column_index_from_string("CE"), value=pd.to_datetime(str(fec_k), errors='coerce').strftime('%Y/%m/%d'))

    wb.save(ruta_plantilla_out)
    wb.close()