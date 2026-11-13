from .cop import procesar_cop_generico
from .psa import procesar_psa
from .gestantes import procesar_gestantes_nueva_eps, procesar_gestantes_savia
from .hc_peso_talla import procesar_informe_hc
from .demograficos import procesar_ajustes_demograficos
from .informe_general import cargar_informe_general, procesar_informe_general
from .informe_produccion import (
    cargar_informe_produccion,
    procesar_informe_produccion,
    crear_config_produccion_vacia,
)
from .orquestador import ejecutar_consolidacion_completa
