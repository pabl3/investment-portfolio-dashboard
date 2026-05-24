"""
Convierte portfolio_data.csv de separador decimal punto (.) a coma (,)
para compatibilidad con configuración regional de Windows en Argentina.

Uso:
    python src/convertir_csv.py
"""

import pandas as pd
import sys
sys.stdout.reconfigure(encoding="utf-8")

INPUT  = "data/processed/portfolio_data.csv"
OUTPUT = "data/processed/portfolio_data_ar.csv"

df = pd.read_csv(INPUT, index_col=0, parse_dates=True)

print(f"✓ CSV cargado: {df.shape[0]} filas × {df.shape[1]} columnas")
print(f"  Primer precio AAPL: {df['AAPL'].iloc[0]}")

# Exportar con separador decimal coma y separador de columnas punto y coma
df.to_csv(OUTPUT, sep=";", decimal=",", date_format="%Y-%m-%d")

print(f"✓ CSV exportado: {OUTPUT}")
print(f"  En Power BI usá delimitador: punto y coma (;)")
print(f"  Separador decimal: coma (,)")
