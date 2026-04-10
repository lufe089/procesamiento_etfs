"""
Sistema de Procesamiento ESG de ETFs

Este módulo procesa archivos ETF individuales, los cruza con masters de ESG
y Market Capitalization, calcula promedios simples y ponderados, y genera
reportes de trazabilidad.


Fecha: 2026-04-08
"""

import pandas as pd
import os
import sys
import logging
from typing import Dict, List, Tuple, Optional, Any


# ============================================================================
# CONFIGURACIÓN Y CONSTANTES
# ============================================================================

def obtener_ruta_base() -> str:
    """
    Obtiene la ruta base donde se encuentra el ejecutable o el script.

    Cuando se ejecuta como ejecutable de PyInstaller, sys.executable apunta al .exe
    Cuando se ejecuta como script de Python, __file__ apunta al .py

    Returns:
        str: Ruta del directorio base donde está el ejecutable o script
    """
    if getattr(sys, 'frozen', False):
        # Ejecutando como ejecutable de PyInstaller
        ruta_base = os.path.dirname(sys.executable)
    else:
        # Ejecutando como script de Python
        ruta_base = os.path.dirname(os.path.abspath(__file__))

    return ruta_base


# Obtener ruta base del ejecutable/script
RUTA_BASE = obtener_ruta_base()

# Carpetas
CARPETA_INPUT = os.path.join(RUTA_BASE, "input")
CARPETA_ETFS_INPUT = os.path.join(CARPETA_INPUT, "ETFS")
CARPETA_OUTPUT = os.path.join(RUTA_BASE, "output")
CARPETA_ETFS_OUTPUT = os.path.join(CARPETA_OUTPUT, "ETFS")

# Archivos de entrada
ARCHIVO_ESG_MASTER = os.path.join(CARPETA_INPUT, "data_maria_esg.xlsx")
ARCHIVO_MARKET_CAP_MASTER = os.path.join(CARPETA_INPUT, "data_maria_market_cap.xlsx")
ARCHIVO_BASE_DONNEES = os.path.join(CARPETA_INPUT, "Base de donnees.xlsx")

# Archivos de salida
ARCHIVO_RESULTADO_FINAL = os.path.join(CARPETA_OUTPUT, "resultado_esg_etf.xlsx")
ARCHIVO_TRAZABILIDAD = os.path.join(CARPETA_OUTPUT, "trazabilidad_procesamiento.xlsx")
ARCHIVO_ETFS_OMITIDOS = os.path.join(CARPETA_OUTPUT, "etfs_sin_metadatos.xlsx")
ARCHIVO_BASE_DONNEES_OUTPUT = os.path.join(CARPETA_OUTPUT, "Base_de_donnees.xlsx")

# Configuración de logging
LOG_FILE = os.path.join(RUTA_BASE, "procesamiento_etf.log")


# ============================================================================
# CONFIGURACIÓN DE LOGGING
# ============================================================================

def configurar_logging() -> logging.Logger:
    """
    Configura el sistema de logging para consola y archivo.

    Returns:
        logging.Logger: Logger configurado
    """
    logging.basicConfig(
        level=logging.INFO,
        format='%(levelname)s: %(message)s',
        handlers=[
            logging.FileHandler(LOG_FILE, mode='w', encoding='utf-8'),
            logging.StreamHandler()
        ]
    )
    return logging.getLogger(__name__)


# ============================================================================
# FUNCIONES DE NORMALIZACIÓN
# ============================================================================

def normalizar_nombre(nombre: Any) -> str:
    """
    Normaliza un nombre: quita espacios y convierte a mayúsculas.

    Args:
        nombre: Nombre a normalizar

    Returns:
        str: Nombre normalizado (vacío si es NaN/None)

    Examples:
        >>> normalizar_nombre("  Enel Americas ord  ")
        'ENEL AMERICAS ORD'
        >>> normalizar_nombre(None)
        ''
    """
    if pd.isna(nombre) or nombre is None:
        return ""

    nombre_str = str(nombre).strip().upper()
    return nombre_str


def normalizar_fecha_a_año(fecha: Any) -> Optional[int]:
    """
    Parsea una fecha y extrae solo el año.

    Args:
        fecha: Valor de fecha en cualquier formato

    Returns:
        Optional[int]: Año extraído (2010-2030) o None si no se puede parsear

    Examples:
        >>> normalizar_fecha_a_año("31-12-2024")
        2024
        >>> normalizar_fecha_a_año("2024-01-15 00:00:00")
        2024
        >>> normalizar_fecha_a_año("invalid")
        None
    """
    if pd.isna(fecha) or fecha is None:
        return None

    try:
        # Intentar parsear la fecha
        fecha_parseada = pd.to_datetime(fecha, errors='coerce')

        if pd.isna(fecha_parseada):
            return None

        año = fecha_parseada.year

        # Validar rango razonable
        if 2010 <= año <= 2030:
            return año

        return None

    except (ValueError, TypeError):
        return None


# ============================================================================
# VALIDACIONES INICIALES
# ============================================================================

def validar_archivos_entrada() -> None:
    """
    Valida que existan los archivos y carpetas de entrada requeridos.

    Raises:
        FileNotFoundError: Si falta algún archivo o carpeta requerida
    """
    logger = logging.getLogger(__name__)

    # Validar carpeta ETFS
    if not os.path.exists(CARPETA_ETFS_INPUT):
        raise FileNotFoundError(f"No existe la carpeta: {CARPETA_ETFS_INPUT}")

    archivos_etf = [f for f in os.listdir(CARPETA_ETFS_INPUT)
                    if f.endswith('.xlsx') and not f.startswith('~$')]

    if len(archivos_etf) == 0:
        raise FileNotFoundError(f"No hay archivos .xlsx en {CARPETA_ETFS_INPUT}")

    logger.info(f"Encontrados {len(archivos_etf)} archivos ETF para procesar")

    # Validar masters
    if not os.path.exists(ARCHIVO_ESG_MASTER):
        raise FileNotFoundError(f"No existe el archivo: {ARCHIVO_ESG_MASTER}")

    if not os.path.exists(ARCHIVO_MARKET_CAP_MASTER):
        raise FileNotFoundError(f"No existe el archivo: {ARCHIVO_MARKET_CAP_MASTER}")

    logger.info("Validación de archivos de entrada: EXITOSA")


def crear_carpetas_salida() -> None:
    """
    Crea las carpetas de salida si no existen.
    """
    os.makedirs(CARPETA_OUTPUT, exist_ok=True)
    os.makedirs(CARPETA_ETFS_OUTPUT, exist_ok=True)


# ============================================================================
# CARGA DE MASTERS
# ============================================================================

def cargar_masters() -> Tuple[pd.DataFrame, pd.DataFrame]:
    """
    Carga y normaliza los archivos master de ESG y Market Cap.

    Returns:
        Tuple[pd.DataFrame, pd.DataFrame]: (df_esg, df_market_cap) con fechas normalizadas a año

    Raises:
        ValueError: Si faltan columnas requeridas
    """
    logger = logging.getLogger(__name__)

    # Cargar ESG master
    logger.info("Cargando master ESG...")
    df_esg = pd.read_excel(ARCHIVO_ESG_MASTER)

    # Validar columnas
    columnas_requeridas_esg = ['Instrument', 'Date', 'ESG Score']
    for col in columnas_requeridas_esg:
        if col not in df_esg.columns:
            raise ValueError(f"Columna '{col}' no encontrada en {ARCHIVO_ESG_MASTER}")

    # Normalizar fechas a año
    df_esg['year'] = df_esg['Date'].apply(normalizar_fecha_a_año)
    df_esg = df_esg.dropna(subset=['year'])
    df_esg['year'] = df_esg['year'].astype(int)

    # Filtrar valores inválidos de ESG Score (<0 o NaN)
    df_esg = df_esg[df_esg['ESG Score'].notna()]
    df_esg = df_esg[df_esg['ESG Score'] >= 0]

    logger.info(f"Master ESG cargado: {len(df_esg)} registros válidos")

    # Cargar Market Cap master
    logger.info("Cargando master Market Cap...")
    df_cap = pd.read_excel(ARCHIVO_MARKET_CAP_MASTER)

    # Validar columnas
    columnas_requeridas_cap = ['Instrument', 'Date', 'Company Market Capitalization']
    for col in columnas_requeridas_cap:
        if col not in df_cap.columns:
            raise ValueError(f"Columna '{col}' no encontrada en {ARCHIVO_MARKET_CAP_MASTER}")

    # Normalizar fechas a año
    df_cap['year'] = df_cap['Date'].apply(normalizar_fecha_a_año)
    df_cap = df_cap.dropna(subset=['year'])
    df_cap['year'] = df_cap['year'].astype(int)

    # Filtrar valores inválidos de Market Cap (<0 o NaN)
    df_cap = df_cap[df_cap['Company Market Capitalization'].notna()]
    df_cap = df_cap[df_cap['Company Market Capitalization'] >= 0]

    logger.info(f"Master Market Cap cargado: {len(df_cap)} registros válidos")

    return df_esg, df_cap


# ============================================================================
# LECTURA Y PARSEO DE ETFs
# ============================================================================

def leer_metadatos_etf(ruta_archivo: str) -> Tuple[Optional[str], Optional[str], Optional[int]]:
    """
    Lee los metadatos de un archivo ETF buscando en las primeras 5 filas y 3 columnas.

    Args:
        ruta_archivo: Ruta al archivo Excel del ETF

    Returns:
        Tuple[Optional[str], Optional[str], Optional[int]]: (ticker, nombre, año) o (None, None, None)
    """
    logger = logging.getLogger(__name__)

    try:
        # Leer primeras 5 filas sin encabezado
        df_meta = pd.read_excel(ruta_archivo, header=None, nrows=5)

        ticker = None
        nombre = None
        año = None

        # Buscar ticker en columna 0, primeras 5 filas
        for idx in range(min(5, len(df_meta))):
            valor = df_meta.iloc[idx, 0]
            if pd.notna(valor):
                valor_str = str(valor).strip()
                # Validar: 2-10 caracteres, al menos 2 letras, no solo números
                if 2 <= len(valor_str) <= 10:
                    letras = sum(c.isalpha() for c in valor_str)
                    if letras >= 2 and not valor_str.isdigit():
                        ticker = valor_str
                        break

        # Buscar nombre en primeras 5 filas y 3 columnas
        palabras_clave = ['iShares', 'MSCI', 'ETF', 'Fund', 'Index']
        for idx in range(min(5, len(df_meta))):
            for col in range(min(3, df_meta.shape[1])):
                valor = df_meta.iloc[idx, col]
                if pd.notna(valor):
                    valor_str = str(valor).strip()
                    # Validar: contiene palabra clave y > 10 caracteres
                    if len(valor_str) > 10:
                        if any(palabra in valor_str for palabra in palabras_clave):
                            nombre = valor_str
                            break
            if nombre:
                break

        # Buscar fecha en primeras 5 filas y 3 columnas
        for idx in range(min(5, len(df_meta))):
            for col in range(min(3, df_meta.shape[1])):
                valor = df_meta.iloc[idx, col]
                if pd.notna(valor):
                    año_temp = normalizar_fecha_a_año(valor)
                    if año_temp:
                        año = año_temp
                        break
            if año:
                break

        # Validar que se encontraron los 3 metadatos
        if not all([ticker, nombre, año]):
            return None, None, None

        return ticker, nombre, año

    except Exception as e:
        logger.warning(f"Error leyendo metadatos de {os.path.basename(ruta_archivo)}: {e}")
        return None, None, None


def leer_tabla_etf(ruta_archivo: str) -> Optional[pd.DataFrame]:
    """
    Lee la tabla de datos de un ETF desde la fila 4 (header=3).
    NO descarta filas sin RIC - las mantiene todas para trazabilidad completa.

    Args:
        ruta_archivo: Ruta al archivo Excel del ETF

    Returns:
        Optional[pd.DataFrame]: DataFrame con columnas RIC, Name, Country, Weight, No. Shares
                               (RIC puede ser NaN si no existe)
    """
    logger = logging.getLogger(__name__)

    try:
        # Leer tabla desde fila 4
        df = pd.read_excel(ruta_archivo, header=3)

        # Validar columnas requeridas
        columnas_requeridas = ['RIC', 'Name', 'Country', 'Weight', 'No. Shares']
        for col in columnas_requeridas:
            if col not in df.columns:
                logger.warning(f"Columna '{col}' no encontrada en {os.path.basename(ruta_archivo)}")
                return None

        # Filtrar solo columnas requeridas
        df = df[columnas_requeridas].copy()

        # NO descartar filas sin RIC - mantener todas para trazabilidad
        # Las filas sin RIC tendrán NaN en RIC y no harán match con masters

        return df

    except Exception as e:
        logger.warning(f"Error leyendo tabla de {os.path.basename(ruta_archivo)}: {e}")
        return None


# ============================================================================
# CRUCE CON MASTERS
# ============================================================================

def cruzar_con_masters(df_etf: pd.DataFrame, año: int,
                      df_esg: pd.DataFrame, df_cap: pd.DataFrame) -> pd.DataFrame:
    """
    Cruza el ETF con los masters de ESG y Market Cap por RIC/Instrument y año.

    Args:
        df_etf: DataFrame del ETF con columnas RIC, Name, Country, Weight, No. Shares
        año: Año del ETF
        df_esg: Master de ESG
        df_cap: Master de Market Cap

    Returns:
        pd.DataFrame: DataFrame con columnas adicionales ESG Score y Company Market Capitalization
    """
    # Filtrar masters por año
    df_esg_año = df_esg[df_esg['year'] == año].copy()
    df_cap_año = df_cap[df_cap['year'] == año].copy()

    # Merge con ESG
    df_merged = pd.merge(
        df_etf,
        df_esg_año[['Instrument', 'ESG Score']],
        left_on='RIC',
        right_on='Instrument',
        how='left'
    )
    df_merged = df_merged.drop(columns=['Instrument'], errors='ignore')

    # Merge con Market Cap
    df_merged = pd.merge(
        df_merged,
        df_cap_año[['Instrument', 'Company Market Capitalization']],
        left_on='RIC',
        right_on='Instrument',
        how='left'
    )
    df_merged = df_merged.drop(columns=['Instrument'], errors='ignore')

    return df_merged


# ============================================================================
# CÁLCULO DE INDICADORES ESG
# ============================================================================

def calcular_esg(df_merged: pd.DataFrame) -> Optional[Dict[str, Any]]:
    """
    Calcula promedios simple y ponderado de ESG.

    También genera estadísticas detalladas de trazabilidad:
    - total_instrumentos: Total de filas en el ETF
    - sin_ric: Instrumentos sin RIC (NaN o vacío)
    - con_ric: Instrumentos con RIC válido
    - match_esg: Con RIC y match en master ESG
    - match_market_cap: Con RIC y match en master Market Cap
    - utilizados: Con RIC y match completo (ESG + Market Cap válidos >= 0)

    Args:
        df_merged: DataFrame con columnas RIC, ESG Score y Company Market Capitalization

    Returns:
        Optional[Dict]: Diccionario con resultados o None si no hay datos válidos para cálculo ESG
    """
    total_instrumentos = len(df_merged)

    # Contar instrumentos sin RIC (NaN o vacío después de strip)
    sin_ric = df_merged['RIC'].isna().sum()
    sin_ric += df_merged[df_merged['RIC'].notna()]['RIC'].apply(lambda x: str(x).strip() == '').sum()

    # Contar instrumentos con RIC
    con_ric = total_instrumentos - sin_ric

    # De los que tienen RIC, contar matches con masters
    df_con_ric = df_merged[
        df_merged['RIC'].notna() &
        (df_merged['RIC'].astype(str).str.strip() != '')
    ].copy()

    match_esg = df_con_ric['ESG Score'].notna().sum()
    match_market_cap = df_con_ric['Company Market Capitalization'].notna().sum()

    # Filtrar valores válidos para cálculo: RIC existe, ambos no NaN y >= 0
    df_valido = df_merged[
        df_merged['RIC'].notna() &
        (df_merged['RIC'].astype(str).str.strip() != '') &
        df_merged['ESG Score'].notna() &
        df_merged['Company Market Capitalization'].notna() &
        (df_merged['ESG Score'] >= 0) &
        (df_merged['Company Market Capitalization'] >= 0)
    ].copy()

    instrumentos_utilizados = len(df_valido)

    # Si no hay datos válidos, retornar None para esg_simple y esg_ponderado
    # pero SÍ retornar las estadísticas de trazabilidad
    if instrumentos_utilizados == 0:
        return {
            'esg_simple': None,
            'esg_ponderado': None,
            'total_instrumentos': total_instrumentos,
            'sin_ric': int(sin_ric),
            'con_ric': int(con_ric),
            'match_esg': int(match_esg),
            'match_market_cap': int(match_market_cap),
            'instrumentos_utilizados': 0
        }

    # Promedio simple
    promedio_simple = df_valido['ESG Score'].mean()

    # Promedio ponderado por Market Cap
    numerador = (df_valido['ESG Score'] * df_valido['Company Market Capitalization']).sum()
    denominador = df_valido['Company Market Capitalization'].sum()
    promedio_ponderado = numerador / denominador if denominador > 0 else 0

    return {
        'esg_simple': round(float(promedio_simple), 4),
        'esg_ponderado': round(float(promedio_ponderado), 4),
        'total_instrumentos': total_instrumentos,
        'sin_ric': int(sin_ric),
        'con_ric': int(con_ric),
        'match_esg': int(match_esg),
        'match_market_cap': int(match_market_cap),
        'instrumentos_utilizados': instrumentos_utilizados
    }


# ============================================================================
# EXPORTACIÓN DE ETF INDIVIDUAL
# ============================================================================

def exportar_etf_individual(df_merged: pd.DataFrame, ticker: str, año: int,
                           contadores: Dict[str, int]) -> str:
    """
    Exporta un ETF individual con 2 hojas: Procesados y Descartados.

    Args:
        df_merged: DataFrame con todas las filas del ETF (incluyendo sin RIC)
        ticker: Ticker del ETF
        año: Año del ETF
        contadores: Diccionario con contadores de procesamiento

    Returns:
        str: Ruta del archivo exportado
    """
    # Separar procesados (tienen ESG o Market Cap) y descartados (sin ninguno)
    df_procesados = df_merged[
        df_merged['ESG Score'].notna() |
        df_merged['Company Market Capitalization'].notna()
    ].copy()

    df_descartados = df_merged[
        df_merged['ESG Score'].isna() &
        df_merged['Company Market Capitalization'].isna()
    ].copy()

    # Preparar hoja de descartados con motivo específico
    if len(df_descartados) > 0:
        df_descartados_export = df_descartados[['RIC', 'Name', 'Country', 'Weight', 'No. Shares']].copy()

        # Asignar motivo específico según el caso
        def determinar_motivo(row):
            ric_vacio = pd.isna(row['RIC']) or str(row['RIC']).strip() == ''
            if ric_vacio:
                return 'Sin RIC'
            else:
                return 'Con RIC pero sin match en masters ESG ni Market Cap'

        df_descartados_export['Motivo de descarte'] = df_descartados.apply(determinar_motivo, axis=1)
    else:
        df_descartados_export = pd.DataFrame(columns=['RIC', 'Name', 'Country', 'Weight', 'No. Shares', 'Motivo de descarte'])

    # Exportar a Excel con 2 hojas
    nombre_archivo = f"{ticker}_{año}.xlsx"
    ruta_archivo = os.path.join(CARPETA_ETFS_OUTPUT, nombre_archivo)

    with pd.ExcelWriter(ruta_archivo, engine='openpyxl') as writer:
        df_procesados.to_excel(writer, sheet_name='Procesados', index=False)
        df_descartados_export.to_excel(writer, sheet_name='Descartados', index=False)

    return ruta_archivo


# ============================================================================
# PROCESAMIENTO DE UN ETF
# ============================================================================

def procesar_etf(ruta_archivo: str, df_esg: pd.DataFrame,
                df_cap: pd.DataFrame) -> Optional[Dict[str, Any]]:
    """
    Procesa un archivo ETF completo.

    Args:
        ruta_archivo: Ruta al archivo Excel del ETF
        df_esg: Master de ESG
        df_cap: Master de Market Cap

    Returns:
        Optional[Dict]: Resultado del procesamiento o None si hay error
    """
    logger = logging.getLogger(__name__)
    archivo_nombre = os.path.basename(ruta_archivo)

    # 1. Leer metadatos
    ticker, nombre, año = leer_metadatos_etf(ruta_archivo)
    if not all([ticker, nombre, año]):
        logger.warning(f"ETF {archivo_nombre} - Metadatos incompletos, se omite")
        return {'archivo': archivo_nombre, 'omitido': True, 'motivo': 'Metadatos incompletos'}

    # 2. Leer tabla
    df_etf = leer_tabla_etf(ruta_archivo)
    if df_etf is None or len(df_etf) == 0:
        logger.warning(f"ETF {archivo_nombre} - No se pudo leer tabla o está vacía")
        return {'archivo': archivo_nombre, 'omitido': True, 'motivo': 'Tabla vacía o error lectura'}

    total_instrumentos = len(df_etf)

    # 3. Cruzar con masters por RIC y año
    df_merged = cruzar_con_masters(df_etf, año, df_esg, df_cap)

    # 4. Calcular ESG (siempre retorna estadísticas, incluso si esg_simple=None)
    resultado_esg = calcular_esg(df_merged)

    if resultado_esg is None:
        logger.warning(f"ETF {archivo_nombre} - Error calculando ESG")
        return {'archivo': archivo_nombre, 'omitido': True, 'motivo': 'Error en cálculo ESG'}

    # 5. Contadores para trazabilidad
    contadores = {
        'total_instrumentos': resultado_esg['total_instrumentos'],
        'sin_ric': resultado_esg['sin_ric'],
        'con_ric': resultado_esg['con_ric'],
        'con_esg': resultado_esg['match_esg'],
        'con_market_cap': resultado_esg['match_market_cap'],
        'utilizados': resultado_esg['instrumentos_utilizados']
    }

    # 6. Exportar ETF individual (siempre, incluso sin datos ESG válidos)
    exportar_etf_individual(df_merged, ticker, año, contadores)

    # 7. Si no hay datos ESG válidos, marcar como omitido para resultado final
    if resultado_esg['esg_simple'] is None:
        logger.warning(f"ETF {archivo_nombre} - Sin instrumentos válidos para análisis ESG (archivo exportado con estadísticas)")
        return {
            'archivo': archivo_nombre,
            'omitido': True,
            'motivo': 'Sin datos ESG válidos',
            **contadores
        }

    # 8. Extraer país (más frecuente - cada ETF es de un solo país)
    country = df_etf['Country'].mode()[0] if not df_etf['Country'].isna().all() else 'N/A'

    # 9. Preparar resultado
    logger.info(f"ETF {archivo_nombre} - Procesado: {contadores['utilizados']}/{contadores['total_instrumentos']} instrumentos utilizados")

    return {
        'archivo': archivo_nombre,
        'omitido': False,
        'etf_ticker': ticker,
        'etf_name': nombre,
        'year': año,
        'country': country,
        'esg_score_simple_avg': resultado_esg['esg_simple'],
        'esg_score_weighted_avg': resultado_esg['esg_ponderado'],
        'status': 'OK',
        'instruments_count': resultado_esg['instrumentos_utilizados'],
        **contadores
    }


# ============================================================================
# GENERACIÓN DE REPORTES
# ============================================================================

def generar_reporte_trazabilidad(resultados: List[Dict]) -> None:
    """
    Genera reporte de trazabilidad del procesamiento.

    Args:
        resultados: Lista de resultados de procesamiento
    """
    logger = logging.getLogger(__name__)

    # Filtrar solo ETFs procesados (no omitidos)
    resultados_validos = [r for r in resultados if not r.get('omitido', False)]

    if not resultados_validos:
        logger.warning("No hay resultados válidos para reporte de trazabilidad")
        return

    # Crear DataFrame
    datos_reporte = []
    for r in resultados_validos:
        porcentaje = (r['utilizados'] / r['total_instrumentos'] * 100) if r['total_instrumentos'] > 0 else 0
        datos_reporte.append({
            'etf_name': r['etf_name'],
            'YEAR': r['year'],
            'total_instrumentos': r['total_instrumentos'],
            'sin_ric': r['sin_ric'],
            'con_ric': r['con_ric'],
            'match_esg': r['con_esg'],
            'match_market_cap': r['con_market_cap'],
            'utilizados_analisis': r['utilizados'],
            'porcentaje_utilizado': round(porcentaje, 2)
        })

    df_trazabilidad = pd.DataFrame(datos_reporte)
    df_trazabilidad.to_excel(ARCHIVO_TRAZABILIDAD, index=False)

    logger.info(f"Generado reporte de trazabilidad: {ARCHIVO_TRAZABILIDAD}")


def generar_reporte_etfs_omitidos(resultados: List[Dict]) -> None:
    """
    Genera reporte de ETFs omitidos del procesamiento.

    Args:
        resultados: Lista de resultados de procesamiento
    """
    logger = logging.getLogger(__name__)

    # Filtrar solo ETFs omitidos
    omitidos = [r for r in resultados if r.get('omitido', False)]

    if not omitidos:
        logger.info("No hay ETFs omitidos")
        return

    # Crear DataFrame
    df_omitidos = pd.DataFrame(omitidos)
    df_omitidos = df_omitidos[['archivo', 'motivo']]
    df_omitidos.to_excel(ARCHIVO_ETFS_OMITIDOS, index=False)

    logger.info(f"Generado reporte de ETFs omitidos: {ARCHIVO_ETFS_OMITIDOS} ({len(omitidos)} archivos)")


# ============================================================================
# EXPORTACIÓN DE RESULTADO FINAL
# ============================================================================

def exportar_resultado_final(resultados: List[Dict]) -> pd.DataFrame:
    """
    Exporta el resultado consolidado final.

    Args:
        resultados: Lista de resultados de procesamiento

    Returns:
        pd.DataFrame: DataFrame con resultados para actualizar Base de données
    """
    logger = logging.getLogger(__name__)

    # Filtrar solo ETFs procesados
    resultados_validos = [r for r in resultados if not r.get('omitido', False)]

    if not resultados_validos:
        logger.warning("No hay resultados válidos para exportar")
        return pd.DataFrame()

    # Crear DataFrame
    columnas = [
        'etf_ticker', 'etf_name', 'year', 'country',
        'esg_score_simple_avg', 'esg_score_weighted_avg',
        'status', 'instruments_count'
    ]

    datos = [{k: r[k] for k in columnas} for r in resultados_validos]
    df_resultado = pd.DataFrame(datos)

    # Exportar
    df_resultado.to_excel(ARCHIVO_RESULTADO_FINAL, index=False)

    logger.info(f"Generado resultado final: {ARCHIVO_RESULTADO_FINAL} ({len(df_resultado)} ETFs)")

    return df_resultado


# ============================================================================
# ACTUALIZACIÓN DE BASE DE DONNÉES
# ============================================================================

def actualizar_base_donnees(df_resultado: pd.DataFrame) -> None:
    """
    Actualiza la Base de données con los resultados ESG.

    IMPORTANTE: En la Base de données:
    - La columna 'Country' contiene el TICKER del ETF SIN sufijos (ej: 'ECH', 'EDEN')
    - La columna 'Name' contiene el NOMBRE del ETF (ej: 'iShares MSCI Chile ETF')
    - La columna 'year' contiene el AÑO

    En df_resultado:
    - 'etf_ticker' puede tener sufijos con punto (ej: 'ECH.K', 'EDEN.O')
      → Se extrae solo la parte antes del punto para hacer match con 'Country'
    - 'etf_name' corresponde a 'Name' de la base
    - 'year' corresponde a 'year' de la base

    Args:
        df_resultado: DataFrame con resultados finales
    """
    logger = logging.getLogger(__name__)

    if df_resultado.empty:
        logger.warning("No hay resultados para actualizar Base de données")
        return

    # Verificar que existe el archivo
    if not os.path.exists(ARCHIVO_BASE_DONNEES):
        logger.warning(f"No existe el archivo {ARCHIVO_BASE_DONNEES}, se omite actualización")
        return

    try:
        # Leer Base de données (intentar primero con sheet_name, luego sin)
        logger.info("Leyendo Base de données...")
        try:
            df_base = pd.read_excel(ARCHIVO_BASE_DONNEES, sheet_name='Base de donnes')
        except:
            # Si falla, intentar leer la primera hoja
            df_base = pd.read_excel(ARCHIVO_BASE_DONNEES)

        logger.info(f"Base de données leída: {len(df_base)} filas")

        # CORRECCIÓN: Extraer parte antes del punto del ticker
        # Ejemplo: "ECH.K" -> "ECH", "EDEN.O" -> "EDEN", "EFNL" -> "EFNL"
        def extraer_ticker_base(ticker):
            if pd.isna(ticker):
                return ""
            ticker_str = str(ticker).strip()
            # Tomar solo la parte antes del primer punto
            ticker_sin_sufijo = ticker_str.split('.')[0]
            return ticker_sin_sufijo.upper()

        # Normalizar tickers (sin sufijos después del punto)
        df_resultado['ticker_norm'] = df_resultado['etf_ticker'].apply(extraer_ticker_base)
        df_resultado['name_norm'] = df_resultado['etf_name'].apply(normalizar_nombre)

        df_base['ticker_norm'] = df_base['Country'].apply(normalizar_nombre)
        df_base['name_norm'] = df_base['Name'].apply(normalizar_nombre)

        logger.info(f"Ejemplos de normalización:")
        logger.info(f"  Resultado - ticker_norm: {df_resultado['ticker_norm'].head(3).tolist()}")
        logger.info(f"  Base - ticker_norm: {df_base['ticker_norm'].head(3).tolist()}")

        # Hacer merge OUTER para:
        # 1. Mantener todas las filas de Base de données (incluso sin match)
        # 2. Agregar filas del resultado ESG que NO están en Base de données
        df_actualizado = pd.merge(
            df_base,
            df_resultado[['ticker_norm', 'name_norm', 'year',
                         'esg_score_simple_avg', 'esg_score_weighted_avg']],
            on=['ticker_norm', 'name_norm', 'year'],
            how='outer'
        )

        # Para las filas nuevas (del resultado que no están en base),
        # completar Country y Name desde los valores normalizados
        mask_nuevas = df_actualizado['Country'].isna()
        if mask_nuevas.sum() > 0:
            # Obtener los valores originales del resultado para las filas nuevas
            df_resultado_para_nuevas = df_resultado[['ticker_norm', 'name_norm', 'year', 'etf_ticker', 'etf_name', 'country']].copy()

            # Extraer solo ticker sin sufijo para Country
            df_resultado_para_nuevas['Country'] = df_resultado_para_nuevas['etf_ticker'].apply(
                lambda x: str(x).split('.')[0] if pd.notna(x) else ''
            )
            df_resultado_para_nuevas['Name'] = df_resultado_para_nuevas['etf_name']

            # Hacer merge temporal para obtener los valores originales
            df_temp = pd.merge(
                df_actualizado[mask_nuevas][['ticker_norm', 'name_norm', 'year']],
                df_resultado_para_nuevas[['ticker_norm', 'name_norm', 'year', 'Country', 'Name']],
                on=['ticker_norm', 'name_norm', 'year'],
                how='left'
            )

            # Actualizar las filas nuevas
            df_actualizado.loc[mask_nuevas, 'Country'] = df_temp['Country'].values
            df_actualizado.loc[mask_nuevas, 'Name'] = df_temp['Name'].values

            logger.info(f"Agregadas {mask_nuevas.sum()} filas nuevas del resultado ESG a Base de données")

        # Renombrar columnas ESG según Base de données
        if 'esg_score_simple_avg' in df_actualizado.columns:
            df_actualizado.rename(columns={
                'esg_score_simple_avg': 'ESG moyen',
                'esg_score_weighted_avg': 'ESG pondéré'
            }, inplace=True)

        # Eliminar columnas temporales de normalización
        df_actualizado = df_actualizado.drop(columns=['ticker_norm', 'name_norm'], errors='ignore')

        # Exportar
        df_actualizado.to_excel(ARCHIVO_BASE_DONNEES_OUTPUT, index=False)

        matches = df_actualizado['ESG moyen'].notna().sum()
        total = len(df_actualizado)
        logger.info(f"Base de données actualizada: {ARCHIVO_BASE_DONNEES_OUTPUT} ({matches}/{total} matches)")

    except Exception as e:
        logger.error(f"Error actualizando Base de données: {e}", exc_info=True)


# ============================================================================
# FUNCIÓN PRINCIPAL
# ============================================================================

def main():
    """
    Función principal que orquesta el procesamiento completo.
    """
    logger = configurar_logging()

    try:
        logger.info("=" * 80)
        logger.info("=== INICIANDO PROCESAMIENTO ESG DE ETFs ===")
        logger.info("=" * 80)

        # 1. Validaciones iniciales
        validar_archivos_entrada()
        crear_carpetas_salida()

        # 2. Cargar masters (normalizar fechas a año)
        df_esg, df_cap = cargar_masters()

        # 3. Obtener lista de archivos ETF
        archivos_etf = [
            os.path.join(CARPETA_ETFS_INPUT, f)
            for f in os.listdir(CARPETA_ETFS_INPUT)
            if f.endswith('.xlsx') and not f.startswith('~$')
        ]

        logger.info(f"Procesando {len(archivos_etf)} archivos ETF...")

        # 4. Procesar cada ETF
        resultados = []
        for ruta_archivo in archivos_etf:
            resultado = procesar_etf(ruta_archivo, df_esg, df_cap)
            if resultado:
                resultados.append(resultado)

        # 5. Generar reportes
        logger.info("Generando reportes...")
        generar_reporte_trazabilidad(resultados)
        generar_reporte_etfs_omitidos(resultados)

        # 6. Exportar resultado final
        df_resultado = exportar_resultado_final(resultados)

        # 7. Actualizar Base de données
        if not df_resultado.empty:
            actualizar_base_donnees(df_resultado)

        # 8. Resumen final
        logger.info("=" * 80)
        logger.info("=== PROCESAMIENTO COMPLETADO EXITOSAMENTE ===")
        logger.info("=" * 80)

        procesados = len([r for r in resultados if not r.get('omitido', False)])
        omitidos = len([r for r in resultados if r.get('omitido', False)])

        logger.info(f"Total archivos procesados: {len(resultados)}")
        logger.info(f"ETFs con resultados ESG: {procesados}")
        logger.info(f"ETFs omitidos: {omitidos}")

    except Exception as e:
        logger.error(f"Error en el procesamiento: {e}", exc_info=True)
        raise


if __name__ == "__main__":
    main()

