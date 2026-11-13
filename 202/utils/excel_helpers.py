import shutil
import openpyxl
import pandas as pd


def limpiar_documento(doc_val):
    """
    Normaliza un número de documento evitando decimales (.0), espacios o valores nulos.
    Retorna un string limpio o None si el valor no es válido.
    """
    if doc_val is None or pd.isnull(doc_val):
        return None

    texto = str(doc_val).strip()
    if not texto or texto.lower() == "nan":
        return None

    # Si empieza con '0' y no tiene punto decimal, conservamos los ceros a la izquierda
    if texto.startswith("0") and "." not in texto:
        return texto

    try:
        return str(int(float(texto)))
    except (ValueError, TypeError):
        return texto.split(".")[0].strip()


def formatear_fecha_str(fec_val, formato="%Y/%m/%d"):
    """
    Convierte un valor de fecha a string con el formato indicado (por defecto YYYY/MM/DD).
    Retorna None si la fecha es inválida o vacía (evitando errores NaT).
    """
    if fec_val is None or pd.isnull(fec_val):
        return None

    texto = str(fec_val).strip()
    if not texto or texto.lower() in ("nan", "nat", "-"):
        return None

    parsed = pd.to_datetime(texto, errors="coerce")
    if pd.notnull(parsed):
        return parsed.strftime(formato)
    return None


def formatear_fecha_date(fec_val):
    """
    Convierte un valor a objeto datetime.date puro para inyectar en Excel sin horas.
    Retorna None si no se puede convertir.
    """
    if fec_val is None or pd.isnull(fec_val):
        return None

    if isinstance(fec_val, pd.Timestamp):
        return fec_val.date()

    texto = str(fec_val).strip()
    if not texto or texto.lower() in ("nan", "nat", "-"):
        return None

    parsed = pd.to_datetime(texto, errors="coerce")
    if pd.notnull(parsed):
        return parsed.date()
    return None


def indexar_documentos_plantilla(ws, col_doc=5):
    """
    Construye un diccionario {documento_limpio: fila} para búsqueda O(1)
    y calcula la siguiente fila libre en la hoja activa.
    """
    indice = {}
    for r in range(2, ws.max_row + 1):
        val = ws.cell(row=r, column=col_doc).value
        doc = limpiar_documento(val)
        if doc:
            indice[doc] = r

    siguiente_fila = ws.max_row + 1 if ws.cell(row=2, column=col_doc).value is not None else 2
    return indice, siguiente_fila


def obtener_o_crear_fila(ws, indice, siguiente_fila, doc_limpio, col_doc=5):
    """
    Busca si el documento ya existe en el índice de la plantilla.
    Si no existe, lo inserta en `siguiente_fila` y actualiza el índice.
    Retorna (fila_destino, nueva_siguiente_fila).
    """
    if doc_limpio in indice:
        return indice[doc_limpio], siguiente_fila

    fila_destino = siguiente_fila
    ws.cell(row=fila_destino, column=col_doc, value=doc_limpio)
    indice[doc_limpio] = fila_destino
    return fila_destino, siguiente_fila + 1


def limpiar_plantilla_base(ruta_plantilla_original, ruta_plantilla_destino):
    """
    Copia la plantilla original en la ruta de destino y limpia todas las filas de datos
    desde la fila 2 en adelante conservando el encabezado intacto.
    """
    shutil.copy(ruta_plantilla_original, ruta_plantilla_destino)
    wb = openpyxl.load_workbook(ruta_plantilla_destino)
    ws = wb.active
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, min_col=1, max_col=ws.max_column):
        for cell in row:
            cell.value = None
    wb.save(ruta_plantilla_destino)
    wb.close()
