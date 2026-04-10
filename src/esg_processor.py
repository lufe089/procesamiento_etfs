"""
ESG ETF Processor
=================
Construye una base final con el ESG por ETF usando:
  - Promedio simple de ESG Score
  - Promedio ponderado por Company Market Capitalization

Funciones principales:
    load_data(filepath)           -> pd.DataFrame
    compute_simple_average(df)    -> pd.DataFrame
    compute_weighted_average(df)  -> pd.DataFrame
    build_esg_base(df)            -> pd.DataFrame
    save_results(base, output)    -> None

Uso típico::

    from src.esg_processor import load_data, build_esg_base, save_results

    df   = load_data("data/etf_holdings.csv")
    base = build_esg_base(df)
    save_results(base, "output/esg_base.csv")
"""

import pandas as pd


def load_data(filepath: str) -> pd.DataFrame:
    """
    Carga el archivo de holdings de ETFs.

    Columnas esperadas:
        - ETF: identificador del fondo (e.g. 'SPY')
        - Company: nombre de la empresa
        - ESG_Score: puntuación ESG de la empresa (float)
        - Market_Cap: capitalización bursátil de la empresa en millones USD (float)

    Parameters
    ----------
    filepath : str
        Ruta al archivo CSV de entrada.

    Returns
    -------
    pd.DataFrame
        DataFrame con los datos cargados y tipos correctos.
    """
    df = pd.read_csv(filepath)

    required_columns = {"ETF", "Company", "ESG_Score", "Market_Cap"}
    missing = required_columns - set(df.columns)
    if missing:
        raise ValueError(f"El archivo no contiene las columnas requeridas: {missing}")

    df["ESG_Score"] = pd.to_numeric(df["ESG_Score"], errors="coerce")
    df["Market_Cap"] = pd.to_numeric(df["Market_Cap"], errors="coerce")

    return df


def compute_simple_average(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula el promedio simple del ESG Score por ETF.

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame con columnas ETF y ESG_Score.

    Returns
    -------
    pd.DataFrame
        DataFrame con columnas ETF y ESG_Promedio_Simple.
    """
    result = (
        df.groupby("ETF", sort=True)["ESG_Score"]
        .mean()
        .reset_index()
        .rename(columns={"ESG_Score": "ESG_Promedio_Simple"})
    )
    result["ESG_Promedio_Simple"] = result["ESG_Promedio_Simple"].round(4)
    return result


def compute_weighted_average(df: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula el promedio ponderado del ESG Score por ETF,
    usando Company Market Capitalization como ponderador.

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame con columnas ETF, ESG_Score y Market_Cap.

    Returns
    -------
    pd.DataFrame
        DataFrame con columnas ETF y ESG_Promedio_Ponderado.
    """

    def _weighted_mean(group: pd.DataFrame) -> float:
        weights = group["Market_Cap"]
        scores = group["ESG_Score"]
        total_weight = weights.sum()
        if total_weight == 0:
            return float("nan")
        return round((scores * weights).sum() / total_weight, 4)

    result = (
        df.groupby("ETF", sort=True)
        .apply(_weighted_mean, include_groups=False)
        .reset_index()
        .rename(columns={0: "ESG_Promedio_Ponderado"})
    )
    return result


def build_esg_base(df: pd.DataFrame) -> pd.DataFrame:
    """
    Construye la base final con el ESG por ETF combinando:
      - Promedio simple de ESG Score
      - Promedio ponderado por Company Market Capitalization

    Parameters
    ----------
    df : pd.DataFrame
        DataFrame con las holdings de los ETFs.

    Returns
    -------
    pd.DataFrame
        DataFrame final con columnas:
        ETF, ESG_Promedio_Simple, ESG_Promedio_Ponderado
    """
    simple = compute_simple_average(df)
    weighted = compute_weighted_average(df)
    base = simple.merge(weighted, on="ETF")
    return base


def save_results(base: pd.DataFrame, output_path: str) -> None:
    """
    Guarda la base final en un archivo CSV.

    Parameters
    ----------
    base : pd.DataFrame
        DataFrame con los resultados ESG por ETF.
    output_path : str
        Ruta del archivo CSV de salida.
    """
    base.to_csv(output_path, index=False)
    print(f"Resultados guardados en: {output_path}")
