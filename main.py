"""
main.py – Punto de entrada del procesador ESG por ETF.

Uso:
    python main.py [--input RUTA_CSV] [--output RUTA_SALIDA]

Ejemplo:
    python main.py --input data/etf_holdings.csv --output output/esg_base.csv
"""

import argparse
import os

import pandas as pd

from src.esg_processor import build_esg_base, load_data, save_results


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Construye la base final ESG por ETF."
    )
    parser.add_argument(
        "--input",
        default=os.path.join("data", "etf_holdings.csv"),
        help="Ruta al CSV de entrada con holdings (default: data/etf_holdings.csv)",
    )
    parser.add_argument(
        "--output",
        default=os.path.join("output", "esg_base.csv"),
        help="Ruta del CSV de salida (default: output/esg_base.csv)",
    )
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    print(f"Cargando datos desde: {args.input}")
    df = load_data(args.input)
    print(f"  → {len(df)} filas cargadas, {df['ETF'].nunique()} ETFs distintos.\n")

    print("Calculando ESG por ETF…")
    base = build_esg_base(df)

    print("\nBase final ESG por ETF:")
    print(base.to_string(index=False))

    os.makedirs(os.path.dirname(args.output) or ".", exist_ok=True)
    save_results(base, args.output)


if __name__ == "__main__":
    main()
