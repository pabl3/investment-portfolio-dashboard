"""
FASE 1 — Descarga y limpieza de datos de cartera de inversiones
================================================================
Herramientas: yfinance, pandas, numpy
Autor: [tu nombre]

Instalación:
    pip install yfinance pandas numpy openpyxl

Uso:
    python fase1_descarga_limpieza.py
"""

import sys
import yfinance as yf
import pandas as pd
import numpy as np
from datetime import datetime, timedelta
import os

sys.stdout.reconfigure(encoding="utf-8")


# ─────────────────────────────────────────────────────────
# CONFIGURACIÓN DE LA CARTERA DEMO
# Podés cambiar estos tickers y fechas según tu cartera real.
# ─────────────────────────────────────────────────────────

TICKERS = {
    "AAPL":  "Apple",
    "MSFT":  "Microsoft",
    "GOOGL": "Alphabet",
    "BRK-B": "Berkshire Hathaway",
    "JPM":   "JPMorgan Chase",
    "GLD":   "SPDR Gold ETF",
    "SPY":   "S&P 500 ETF (benchmark)",
}

FECHA_INICIO = "2021-01-01"
FECHA_FIN    = datetime.today().strftime("%Y-%m-%d")

OUTPUT_CSV   = "data/processed/portfolio_data.csv"
OUTPUT_EXCEL = "data/processed/portfolio_data.xlsx"


# ─────────────────────────────────────────────────────────
# 1. DESCARGA DE DATOS
# ─────────────────────────────────────────────────────────

def descargar_datos(tickers: dict, inicio: str, fin: str) -> pd.DataFrame:
    """
    Descarga precios de cierre ajustados para todos los tickers.
    Retorna un DataFrame con columnas: Date, Ticker, Nombre, Close.
    """
    print(f"\n{'='*55}")
    print("  Descargando datos desde Yahoo Finance...")
    print(f"  Período: {inicio} → {fin}")
    print(f"  Activos: {', '.join(tickers.keys())}")
    print(f"{'='*55}\n")

    lista_tickers = list(tickers.keys())

    # Descarga en bloque (más eficiente que de a uno)
    raw = yf.download(
        tickers=lista_tickers,
        start=inicio,
        end=fin,
        auto_adjust=True,   # Ajusta splits y dividendos automáticamente
        progress=True,
    )

    # Con múltiples tickers, yfinance devuelve MultiIndex
    # Nos quedamos solo con "Close"
    if isinstance(raw.columns, pd.MultiIndex):
        precios = raw["Close"].copy()
    else:
        # Caso de un solo ticker
        precios = raw[["Close"]].copy()
        precios.columns = lista_tickers

    print(f"\n✓ Descargados {len(precios)} días de datos")
    print(f"  Rango real: {precios.index[0].date()} → {precios.index[-1].date()}")

    return precios


# ─────────────────────────────────────────────────────────
# 2. LIMPIEZA DE DATOS
# ─────────────────────────────────────────────────────────

def limpiar_datos(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aplica todas las transformaciones de limpieza al DataFrame de precios.
    Pasos:
        1. Reporte de valores nulos
        2. Forward-fill para días sin cotización (feriados)
        3. Eliminación de filas completamente vacías
        4. Validación de fechas duplicadas
        5. Ordenamiento cronológico
    """
    print("\n" + "─"*55)
    print("  LIMPIEZA DE DATOS")
    print("─"*55)

    df = df.copy()
    df.index = pd.to_datetime(df.index)

    # ── 2.1 Reporte de nulos antes de limpiar ────────────
    nulos = df.isnull().sum()
    if nulos.sum() > 0:
        print("\n⚠ Valores nulos detectados:")
        for ticker, cantidad in nulos[nulos > 0].items():
            pct = (cantidad / len(df)) * 100
            print(f"   {ticker}: {cantidad} días ({pct:.1f}%)")
    else:
        print("\n✓ Sin valores nulos")

    # ── 2.2 Forward-fill (máx. 5 días consecutivos) ──────
    # Cubre feriados locales donde el mercado no opera.
    # Límite de 5 evita propagar datos durante suspensiones largas.
    df = df.ffill(limit=5)

    # ── 2.3 Eliminar filas donde TODOS los tickers son NaN ──
    filas_antes = len(df)
    df = df.dropna(how="all")
    filas_eliminadas = filas_antes - len(df)
    if filas_eliminadas > 0:
        print(f"\n⚠ Se eliminaron {filas_eliminadas} filas completamente vacías")
    else:
        print("✓ Sin filas vacías")

    # ── 2.4 Verificar duplicados de fecha ────────────────
    duplicados = df.index.duplicated().sum()
    if duplicados > 0:
        print(f"\n⚠ {duplicados} fechas duplicadas — se conserva la última")
        df = df[~df.index.duplicated(keep="last")]
    else:
        print("✓ Sin fechas duplicadas")

    # ── 2.5 Ordenar cronológicamente ─────────────────────
    df = df.sort_index()

    # ── 2.6 Nulos residuales (si existen al inicio) ──────
    nulos_residuales = df.isnull().sum().sum()
    if nulos_residuales > 0:
        print(f"\n⚠ {nulos_residuales} NaN residuales (probablemente al inicio del período)")
        print("  → Se rellenan con el primer precio disponible de cada ticker")
        df = df.bfill()

    print(f"\n✓ Dataset limpio: {len(df)} filas × {len(df.columns)} tickers")
    return df


# ─────────────────────────────────────────────────────────
# 3. TRANSFORMACIONES Y CÁLCULO DE RETORNOS
# ─────────────────────────────────────────────────────────

def calcular_retornos(precios: pd.DataFrame) -> dict:
    """
    Calcula tres tipos de retorno a partir de los precios de cierre:
        - Retorno diario simple (%)
        - Retorno acumulado desde el inicio (%)
        - Retorno anualizado (CAGR)

    Retorna un dict con tres DataFrames.
    """
    print("\n" + "─"*55)
    print("  CÁLCULO DE RETORNOS")
    print("─"*55)

    # ── 3.1 Retorno diario ────────────────────────────────
    retorno_diario = precios.pct_change(fill_method=None)
    retorno_diario.iloc[0] = 0  # Primer día sin retorno anterior

    # ── 3.2 Retorno acumulado ─────────────────────────────
    retorno_acumulado = (1 + retorno_diario).cumprod() - 1

    # ── 3.3 CAGR (Compound Annual Growth Rate) ───────────
    n_años = (precios.index[-1] - precios.index[0]).days / 365.25
    precio_final   = precios.iloc[-1]
    precio_inicial = precios.iloc[0]

    cagr = (precio_final / precio_inicial) ** (1 / n_años) - 1

    print("\n  Retorno acumulado total por activo:")
    print("  " + "─"*40)
    for ticker in precios.columns:
        ret_total = retorno_acumulado[ticker].iloc[-1] * 100
        cagr_pct  = cagr[ticker] * 100
        signo     = "+" if ret_total >= 0 else ""
        print(f"  {ticker:<8}  acum: {signo}{ret_total:6.1f}%   CAGR: {signo}{cagr_pct:.1f}%")

    print(f"\n✓ Retornos calculados para {len(precios.columns)} activos")
    print(f"  Período analizado: {n_años:.1f} años")

    return {
        "retorno_diario":     retorno_diario,
        "retorno_acumulado":  retorno_acumulado,
        "cagr":               cagr,
    }


# ─────────────────────────────────────────────────────────
# 4. ESTADÍSTICAS DE RESUMEN
# ─────────────────────────────────────────────────────────

def resumen_estadistico(precios: pd.DataFrame, retornos: dict) -> pd.DataFrame:
    """
    Genera una tabla resumen con métricas básicas por activo.
    Esta tabla es el input directo de la Fase 2 (métricas avanzadas).
    """
    rd = retornos["retorno_diario"]

    resumen = pd.DataFrame({
        "Primer precio (USD)":    precios.iloc[0].round(2),
        "Último precio (USD)":    precios.iloc[-1].round(2),
        "Retorno acumulado (%)":  (retornos["retorno_acumulado"].iloc[-1] * 100).round(2),
        "CAGR (%)":               (retornos["cagr"] * 100).round(2),
        "Volatilidad anual (%)":  (rd.std() * np.sqrt(252) * 100).round(2),
        "Mejor día (%)":          (rd.max() * 100).round(2),
        "Peor día (%)":           (rd.min() * 100).round(2),
        "Días analizados":        rd.count().astype(int),
    })

    print("\n" + "─"*55)
    print("  RESUMEN ESTADÍSTICO")
    print("─"*55)
    print(resumen.to_string())

    return resumen


# ─────────────────────────────────────────────────────────
# 5. EXPORTACIÓN
# ─────────────────────────────────────────────────────────

def exportar_datos(
    precios: pd.DataFrame,
    retornos: dict,
    resumen: pd.DataFrame,
    nombre_csv: str,
    nombre_excel: str,
):
    """
    Guarda los datos en dos formatos:
        - CSV: precios de cierre (para Power BI y Fase 2)
        - Excel: múltiples hojas con toda la información
    """
    print("\n" + "─"*55)
    print("  EXPORTACIÓN")
    print("─"*55)

    # ── 5.1 CSV de precios ───────────────────────────────
    precios.to_csv(nombre_csv, date_format="%Y-%m-%d")
    print(f"\n✓ CSV guardado: {nombre_csv}")

    # ── 5.2 Excel con múltiples hojas ────────────────────
    with pd.ExcelWriter(nombre_excel, engine="openpyxl") as writer:

        # Hoja 1: precios de cierre
        precios.to_excel(writer, sheet_name="Precios")

        # Hoja 2: retornos diarios
        retornos["retorno_diario"].mul(100).round(4).to_excel(
            writer, sheet_name="Retornos diarios (%)"
        )

        # Hoja 3: retorno acumulado
        retornos["retorno_acumulado"].mul(100).round(4).to_excel(
            writer, sheet_name="Retorno acumulado (%)"
        )

        # Hoja 4: resumen estadístico
        resumen.to_excel(writer, sheet_name="Resumen")

    print(f"✓ Excel guardado: {nombre_excel}")
    print("  Hojas: Precios | Retornos diarios | Retorno acumulado | Resumen")

    # ── 5.3 Tamaño de los archivos ───────────────────────
    for archivo in [nombre_csv, nombre_excel]:
        kb = os.path.getsize(archivo) / 1024
        print(f"  {archivo}: {kb:.0f} KB")


# ─────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────

def main():
    print("\n" + "="*55)
    print("  DASHBOARD DE CARTERA — FASE 1")
    print("  Extracción y limpieza de datos")
    print("="*55)

    # 1. Descargar
    precios_raw = descargar_datos(TICKERS, FECHA_INICIO, FECHA_FIN)

    # 2. Limpiar
    precios = limpiar_datos(precios_raw)

    # 3. Calcular retornos
    retornos = calcular_retornos(precios)

    # 4. Resumen estadístico
    resumen = resumen_estadistico(precios, retornos)

    # 5. Exportar
    exportar_datos(precios, retornos, resumen, OUTPUT_CSV, OUTPUT_EXCEL)

    print("\n" + "="*55)
    print("  ✓ FASE 1 COMPLETADA")
    print(f"  Archivos generados:")
    print(f"    → {OUTPUT_CSV}   (input para Power BI y Fase 2)")
    print(f"    → {OUTPUT_EXCEL}  (exploración y modelo Excel)")
    print("="*55 + "\n")


if __name__ == "__main__":
    main()
