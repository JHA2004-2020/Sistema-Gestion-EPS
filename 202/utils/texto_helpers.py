import pandas as pd


def dividir_nombre_compuesto(nombre_completo):
    """
    Divide un nombre completo colombiano respetando conectores compuestos
    (DE, DEL, LA, LAS, LOS, SAN, SANTA, MAC, MC, VAN, VON, Y).
    Retorna: (primer_apellido, segundo_apellido, primer_nombre, segundo_nombre)
    """
    if pd.isnull(nombre_completo):
        return "", "", "", ""

    palabras = str(nombre_completo).strip().upper().split()
    conectores = {"DE", "DEL", "LA", "LAS", "LOS", "SAN", "SANTA", "MAC", "MC", "VAN", "VON", "Y"}

    partes_agrupadas = []
    i = 0
    while i < len(palabras):
        parte = palabras[i]
        while i < len(palabras) - 1 and palabras[i] in conectores:
            i += 1
            parte += " " + palabras[i]
        partes_agrupadas.append(parte)
        i += 1

    p_ape, s_ape, p_nom, s_nom = "", "", "", ""

    if len(partes_agrupadas) == 1:
        p_nom = partes_agrupadas[0]
    elif len(partes_agrupadas) == 2:
        p_ape = partes_agrupadas[0]
        p_nom = partes_agrupadas[1]
    elif len(partes_agrupadas) == 3:
        p_ape = partes_agrupadas[0]
        s_ape = partes_agrupadas[1]
        p_nom = partes_agrupadas[2]
    elif len(partes_agrupadas) >= 4:
        p_ape = partes_agrupadas[0]
        s_ape = partes_agrupadas[1]
        p_nom = partes_agrupadas[2]
        s_nom = " ".join(partes_agrupadas[3:])

    return p_ape, s_ape, p_nom, s_nom


def convertir_edad_a_meses(edad, unidad):
    """Convierte la edad a meses exactos según la unidad (DIAS, MESES, AÑOS)."""
    if pd.isnull(edad):
        return 0
    try:
        val = float(edad)
    except (ValueError, TypeError):
        return 0

    unidad_str = str(unidad).strip().upper()
    if "D" in unidad_str or "DÍA" in unidad_str:
        return val / 30.0
    elif "M" in unidad_str:
        return val
    elif "A" in unidad_str or "ANO" in unidad_str:
        return val * 12.0
    return val
