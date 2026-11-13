import streamlit as st
from pathlib import Path
from utils.excel_helpers import limpiar_plantilla_base
from procesadores.cop import procesar_cop_generico
from procesadores.psa import procesar_psa
from procesadores.gestantes import procesar_gestantes_nueva_eps, procesar_gestantes_savia
from procesadores.informe_general import procesar_informe_general
from procesadores.informe_produccion import procesar_informe_produccion
from procesadores.hc_peso_talla import procesar_informe_hc
from procesadores.demograficos import procesar_ajustes_demograficos


def ejecutar_consolidacion_completa(
    entidad,
    insumos_dir,
    df_general_filtrado,
    config_general,
    config_produccion,
    entidades_produccion,
):
    """
    Orquesta en orden secuencial todos los módulos automáticos y manuales
    sobre una copia limpia de Plantilla.xlsx.
    """
    insumos_dir = Path(insumos_dir)
    plantilla_original = insumos_dir / "Plantilla.xlsx"
    plantilla_resultado = insumos_dir / f"Plantilla_Procesada_{entidad}.xlsx"

    if not plantilla_original.exists():
        st.error("No se encontró el archivo base 'Plantilla.xlsx' en la carpeta Insumos.")
        return None

    es_nueva_eps = entidad.upper() in ("NUEVA EPS", "NUEVA")

    # 1. Copia limpia de la plantilla desde cero
    limpiar_plantilla_base(plantilla_original, plantilla_resultado)

    # 2. Módulo Automático: COP
    try:
        hoja_cop = "NUEVA EPS" if es_nueva_eps else "SAVIA"
        procesar_cop_generico(hoja_cop, plantilla_resultado, plantilla_resultado)
        st.write("✅ Módulo **COP** procesado correctamente.")
    except Exception as e:
        st.warning(f"⚠️ Aviso en COP: {e}")

    # 3. Módulo Automático: PSA
    try:
        ruta_psa = insumos_dir / "PSA.xlsx"
        if ruta_psa.exists():
            procesar_psa(plantilla_resultado, ruta_psa)
            st.write("✅ Módulo **PSA** cruzado correctamente (Columnas BV y DF).")
        else:
            st.warning("⚠️ No se encontró el archivo PSA.xlsx. Se omitió este paso.")
    except Exception as e:
        st.error(f"❌ Error al procesar PSA: {e}")

    # 4. Módulo Manual: INFORME GENERAL
    try:
        if config_general:
            procesar_informe_general(df_general_filtrado, config_general, plantilla_resultado)
            st.write("✅ Módulo **INFORME GENERAL** procesado correctamente.")
        else:
            st.warning("⚠️ No se cruzó Informe General (no hay pruebas configuradas para esta EPS).")
    except Exception as e:
        st.error(f"❌ Error al dinamizar Informe General: {e}")

    # 5. Módulo Automático: GESTANTES
    try:
        if es_nueva_eps:
            procesar_gestantes_nueva_eps(plantilla_resultado, plantilla_resultado)
            st.write("✅ Módulo **GESTANTES NUEVA EPS** procesado correctamente.")
        else:
            procesar_gestantes_savia(plantilla_resultado, plantilla_resultado)
            st.write("✅ Módulo **GESTANTES SAVIA SALUD** procesado correctamente.")
    except Exception as e:
        st.warning(f"⚠️ Aviso en GESTANTES: {e}")

    # 6. Módulo Manual: INFORME DE PRODUCCIÓN
    try:
        tiene_filtros_prod = config_produccion and any(
            len(lista) > 0 for lista in config_produccion.values()
        )
        if tiene_filtros_prod:
            procesar_informe_produccion(
                config_produccion,
                plantilla_resultado,
                entidades_seleccionadas=entidades_produccion,
                entidad_sugerida=entidad,
            )
            st.write("✅ Módulo **INFORME DE PRODUCCIÓN** cruzado correctamente.")
        else:
            st.warning("⚠️ No se cruzó Informe de Producción (no hay filtros configurados para esta EPS).")
    except Exception as e:
        st.error(f"❌ Error al dinamizar Informe de Producción: {e}")

    # 7. Módulo Automático: INFORME HC (Pesos y Tallas)
    try:
        ruta_hc = insumos_dir / "INFORME HC.xlsx"
        procesar_informe_hc(plantilla_resultado, ruta_hc)
        st.write("✅ Módulo **INFORME HC (Pesos y Tallas)** cruzado correctamente.")
    except Exception as e:
        st.error(f"❌ Error al procesar Informe HC: {e}")

    # 8. Módulo Automático: AJUSTES DEMOGRÁFICOS
    try:
        ruta_produccion = insumos_dir / "INFORME PRODUCCION.xlsx"
        procesar_ajustes_demograficos(plantilla_resultado, ruta_produccion)
        st.write(
            "✅ **AJUSTES DEMOGRÁFICOS** (Nombres divididos y datos de usuario) procesados correctamente."
        )
    except Exception as e:
        st.error(f"❌ Error al procesar Ajustes Demográficos: {e}")

    st.success("🎉 ¡Todos los procesos terminaron con éxito!")
    return plantilla_resultado
