"""
FASE 2 — Métricas financieras de riesgo
========================================
Herramientas: pandas, numpy
Input:  data/processed/portfolio_data.csv  (generado en Fase 1)
Output: data/processed/metrics_output.xlsx (input para Power BI)

Métricas calculadas:
    1. Sharpe Ratio por activo y cartera total
    2. Volatilidad rolling (30 y 90 días)
    3. Drawdown máximo
    4. Matriz de correlación

Uso:
    python src/fase2_metricas.py
"""

import sys
import pandas as pd
import numpy as np
import os

sys.stdout.reconfigure(encoding="utf-8")


# ─────────────────────────────────────────────────────────
# CONFIGURACIÓN
# ─────────────────────────────────────────────────────────

INPUT_CSV     = "data/processed/portfolio_data.csv"
OUTPUT_EXCEL  = "data/processed/metrics_output.xlsx"

# Tasa libre de riesgo anual (T-Bill EEUU aprox.)
# Se usa para calcular el Sharpe Ratio
TASA_LIBRE_RIESGO = 0.05  # 5% anual → 0.05 / 252 diario


# ─────────────────────────────────────────────────────────
# 0. CARGA DE DATOS
# ─────────────────────────────────────────────────────────

def cargar_datos(path: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Carga el CSV generado en Fase 1 y calcula retornos diarios.
    Retorna una tupla (precios, retornos_diarios).
    """
    print(f"\n{'='*55}")
    print("  DASHBOARD DE CARTERA — FASE 2")
    print("  Métricas financieras de riesgo")
    print(f"{'='*55}\n")

    if not os.path.exists(path):
        raise FileNotFoundError(
            f"No se encontró {path}\n"
            "Asegurate de haber ejecutado primero fase1_descarga_limpieza.py"
        )

    precios = pd.read_csv(path, index_col=0, parse_dates=True)

    # Eliminar columna Ticker si quedó del MultiIndex
    if "Ticker" in precios.columns:
        precios = precios.drop(columns=["Ticker"])

    retornos = precios.pct_change(fill_method=None).dropna()

    print(f"✓ Datos cargados: {len(precios)} días × {len(precios.columns)} activos")
    print(f"  Activos: {', '.join(precios.columns.tolist())}")
    print(f"  Período: {precios.index[0].date()} → {precios.index[-1].date()}")

    return precios, retornos


# ─────────────────────────────────────────────────────────
# 1. SHARPE RATIO
# ─────────────────────────────────────────────────────────

def calcular_sharpe(retornos: pd.DataFrame, tasa_anual: float = TASA_LIBRE_RIESGO) -> pd.DataFrame:
    """
    Calcula el Sharpe Ratio anualizado por activo y para la cartera total.

    Fórmula:
        Sharpe = (Retorno_anualizado - Tasa_libre_riesgo) / Volatilidad_anualizada

    Interpretación:
        < 1.0  → aceptable
        1.0 - 1.5 → bueno
        > 1.5  → muy bueno
        > 2.0  → excelente
    """
    print("\n" + "─"*55)
    print("  1. SHARPE RATIO")
    print("─"*55)

    tasa_diaria = tasa_anual / 252

    # Exceso de retorno diario sobre la tasa libre de riesgo
    exceso = retornos - tasa_diaria

    # Sharpe anualizado
    sharpe = (exceso.mean() / exceso.std()) * np.sqrt(252)

    # Cartera equiponderada (mismo peso para cada activo)
    n = len(retornos.columns)
    retorno_cartera   = retornos.mean(axis=1)
    exceso_cartera    = retorno_cartera - tasa_diaria
    sharpe_cartera    = (exceso_cartera.mean() / exceso_cartera.std()) * np.sqrt(252)

    # Tabla de resultados
    resultado = pd.DataFrame({
        "Sharpe Ratio":      sharpe.round(4),
        "Retorno anual (%)": (retornos.mean() * 252 * 100).round(2),
        "Volatilidad (%)":   (retornos.std() * np.sqrt(252) * 100).round(2),
        "Calificación":      sharpe.apply(_calificar_sharpe),
    })

    # Agregar fila de cartera total
    cartera_row = pd.DataFrame({
        "Sharpe Ratio":      [round(sharpe_cartera, 4)],
        "Retorno anual (%)": [round(retorno_cartera.mean() * 252 * 100, 2)],
        "Volatilidad (%)":   [round(retorno_cartera.std() * np.sqrt(252) * 100, 2)],
        "Calificación":      [_calificar_sharpe(sharpe_cartera)],
    }, index=["CARTERA TOTAL"])

    resultado = pd.concat([resultado, cartera_row])

    print(f"\n  Tasa libre de riesgo usada: {tasa_anual*100:.1f}% anual\n")
    print(resultado.to_string())

    return resultado


def _calificar_sharpe(valor: float) -> str:
    if valor >= 2.0:   return "Excelente"
    elif valor >= 1.5: return "Muy bueno"
    elif valor >= 1.0: return "Bueno"
    elif valor >= 0.5: return "Aceptable"
    else:              return "Bajo"


# ─────────────────────────────────────────────────────────
# 2. VOLATILIDAD ROLLING
# ─────────────────────────────────────────────────────────

def calcular_volatilidad_rolling(retornos: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Calcula la volatilidad anualizada en ventanas móviles de 30 y 90 días.
    Permite ver cómo cambió el riesgo a lo largo del tiempo.

    Resultado: dos DataFrames con la volatilidad rolling (%) por activo y fecha.
    """
    print("\n" + "─"*55)
    print("  2. VOLATILIDAD ROLLING")
    print("─"*55)

    # Volatilidad rolling anualizada (×√252)
    vol_30  = retornos.rolling(window=30).std()  * np.sqrt(252) * 100
    vol_90  = retornos.rolling(window=90).std()  * np.sqrt(252) * 100

    vol_30  = vol_30.dropna().round(4)
    vol_90  = vol_90.dropna().round(4)

    print(f"\n  Ventana 30 días  → {len(vol_30)} observaciones")
    print(f"  Ventana 90 días  → {len(vol_90)} observaciones")

    # Resumen: promedio y máximo de volatilidad por activo
    print("\n  Volatilidad promedio anualizada por activo (ventana 90d):")
    print("  " + "─"*40)
    for ticker in vol_90.columns:
        prom = vol_90[ticker].mean()
        maxi = vol_90[ticker].max()
        print(f"  {ticker:<8}  promedio: {prom:5.1f}%   máximo: {maxi:5.1f}%")

    return vol_30, vol_90


# ─────────────────────────────────────────────────────────
# 3. DRAWDOWN MÁXIMO
# ─────────────────────────────────────────────────────────

def calcular_drawdown(precios: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Calcula el drawdown diario y el drawdown máximo por activo.

    Drawdown = caída porcentual desde el pico más alto hasta el punto más bajo
    antes de la siguiente recuperación.

    Resultado:
        - drawdown_diario: serie temporal de drawdown (%) por activo
        - resumen_drawdown: tabla con el drawdown máximo y fecha por activo
    """
    print("\n" + "─"*55)
    print("  3. DRAWDOWN MÁXIMO")
    print("─"*55)

    # Precio máximo acumulado hasta cada fecha (running maximum)
    rolling_max   = precios.cummax()

    # Drawdown diario: qué tan lejos está el precio actual del máximo histórico
    drawdown_diario = ((precios - rolling_max) / rolling_max) * 100

    # Drawdown máximo por activo
    max_drawdown  = drawdown_diario.min()
    fecha_min     = drawdown_diario.idxmin()

    resumen = pd.DataFrame({
        "Drawdown máximo (%)": max_drawdown.round(2),
        "Fecha del mínimo":    fecha_min.dt.strftime("%Y-%m-%d"),
        "Precio en el mínimo": precios.loc[fecha_min.values, precios.columns].values.diagonal().round(2),
    }, index=precios.columns)

    print("\n  Drawdown máximo por activo:")
    print("  " + "─"*40)
    for ticker in precios.columns:
        dd   = resumen.loc[ticker, "Drawdown máximo (%)"]
        fecha = resumen.loc[ticker, "Fecha del mínimo"]
        print(f"  {ticker:<8}  {dd:7.2f}%   (peor momento: {fecha})")

    return drawdown_diario.round(4), resumen


# ─────────────────────────────────────────────────────────
# 4. MATRIZ DE CORRELACIÓN
# ─────────────────────────────────────────────────────────

def calcular_correlacion(retornos: pd.DataFrame) -> pd.DataFrame:
    """
    Calcula la matriz de correlación entre los retornos diarios de todos los activos.

    Interpretación:
         1.0 → se mueven exactamente igual
         0.0 → sin relación
        -1.0 → se mueven en direcciones opuestas

    Una cartera bien diversificada tiene correlaciones bajas entre activos.
    """
    print("\n" + "─"*55)
    print("  4. MATRIZ DE CORRELACIÓN")
    print("─"*55)

    correlacion = retornos.corr().round(4)

    print("\n" + correlacion.to_string())

    # Análisis de pares con alta correlación (posible sobreconcentración)
    print("\n  Pares con correlación > 0.80 (riesgo de concentración):")
    alta_correlacion = False
    tickers = correlacion.columns.tolist()
    for i in range(len(tickers)):
        for j in range(i+1, len(tickers)):
            val = correlacion.iloc[i, j]
            if val > 0.80:
                print(f"  {tickers[i]} ↔ {tickers[j]}: {val:.4f}")
                alta_correlacion = True
    if not alta_correlacion:
        print("  ✓ Ningún par supera 0.80 — buena diversificación")

    return correlacion


# ─────────────────────────────────────────────────────────
# 5. EXPORTACIÓN
# ─────────────────────────────────────────────────────────

def exportar_metricas(
    sharpe:          pd.DataFrame,
    vol_30:          pd.DataFrame,
    vol_90:          pd.DataFrame,
    drawdown_diario: pd.DataFrame,
    resumen_dd:      pd.DataFrame,
    correlacion:     pd.DataFrame,
    path:            str,
):
    """
    Exporta todas las métricas a un único Excel con 6 hojas.
    Este archivo es el input directo de Power BI (Fase 4).
    """
    print("\n" + "─"*55)
    print("  EXPORTACIÓN")
    print("─"*55)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        sharpe.to_excel(writer,          sheet_name="Sharpe Ratio")
        vol_30.to_excel(writer,          sheet_name="Volatilidad 30d")
        vol_90.to_excel(writer,          sheet_name="Volatilidad 90d")
        drawdown_diario.to_excel(writer, sheet_name="Drawdown diario")
        resumen_dd.to_excel(writer,      sheet_name="Drawdown resumen")
        correlacion.to_excel(writer,     sheet_name="Correlacion")

    kb = os.path.getsize(path) / 1024
    print(f"\n✓ Excel guardado: {path}")
    print(f"  Hojas: Sharpe Ratio | Volatilidad 30d | Volatilidad 90d |")
    print(f"         Drawdown diario | Drawdown resumen | Correlacion")
    print(f"  Tamaño: {kb:.0f} KB")


# ─────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────

def main():

    # 0. Cargar datos de Fase 1
    precios, retornos = cargar_datos(INPUT_CSV)

    # 1. Sharpe Ratio
    sharpe = calcular_sharpe(retornos)

    # 2. Volatilidad rolling
    vol_30, vol_90 = calcular_volatilidad_rolling(retornos)

    # 3. Drawdown máximo
    drawdown_diario, resumen_dd = calcular_drawdown(precios)

    # 4. Correlación
    correlacion = calcular_correlacion(retornos)

    # 5. Exportar todo
    exportar_metricas(
        sharpe, vol_30, vol_90,
        drawdown_diario, resumen_dd,
        correlacion, OUTPUT_EXCEL
    )

    print("\n" + "="*55)
    print("  ✓ FASE 2 COMPLETADA")
    print(f"  Archivo generado:")
    print(f"    → {OUTPUT_EXCEL}")
    print(f"  (input directo para Power BI — Fase 4)")
    print("="*55 + "\n")


if __name__ == "__main__":
    main()
