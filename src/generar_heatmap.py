"""
FASE 4 — Heatmap de correlación (imagen para Power BI)
=======================================================
Herramientas: pandas, seaborn, matplotlib
Input:  data/processed/portfolio_data.csv
Output: powerbi/screenshots/heatmap_correlacion.png

Uso:
    python src/generar_heatmap.py
"""

import sys
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
import seaborn as sns
import os

sys.stdout.reconfigure(encoding="utf-8")

INPUT_CSV  = "data/processed/portfolio_data.csv"
OUTPUT_IMG = "powerbi/screenshots/heatmap_correlacion.png"


def generar_heatmap():
    print(f"\n{'='*55}")
    print("  Generando heatmap de correlación...")
    print(f"{'='*55}\n")

    # Cargar datos y calcular retornos diarios
    precios   = pd.read_csv(INPUT_CSV, index_col=0, parse_dates=True)
    retornos  = precios.pct_change(fill_method=None).dropna()
    correlacion = retornos.corr().round(2)

    # ── Configuración visual ─────────────────────────────
    fig, ax = plt.subplots(figsize=(9, 7))
    fig.patch.set_facecolor("#0D1B2A")
    ax.set_facecolor("#0D1B2A")

    # Paleta: azul (baja correlación) → blanco → rojo (alta correlación)
    cmap = sns.diverging_palette(220, 10, as_cmap=True)

    # Dibujar heatmap
    sns.heatmap(
        correlacion,
        ax=ax,
        cmap=cmap,
        vmin=-1, vmax=1,
        center=0,
        annot=True,
        fmt=".2f",
        annot_kws={"size": 11, "weight": "bold", "color": "white"},
        linewidths=1.5,
        linecolor="#0D1B2A",
        square=True,
        cbar_kws={"shrink": 0.8, "label": "Correlación"},
    )

    # ── Estilo de ejes ───────────────────────────────────
    ax.set_title(
        "Matriz de Correlación — Retornos Diarios",
        color="white", fontsize=14, fontweight="bold", pad=16
    )
    ax.tick_params(colors="white", labelsize=11)
    ax.set_xticklabels(ax.get_xticklabels(), rotation=0,  color="white", fontsize=11)
    ax.set_yticklabels(ax.get_yticklabels(), rotation=0, color="white", fontsize=11)

    # Colorbar
    cbar = ax.collections[0].colorbar
    cbar.ax.yaxis.label.set_color("white")
    cbar.ax.tick_params(colors="white")
    plt.setp(plt.getp(cbar.ax, "yticklabels"), color="white")

    plt.tight_layout(pad=1.5)

    # ── Guardar ──────────────────────────────────────────
    os.makedirs("powerbi/screenshots", exist_ok=True)
    plt.savefig(OUTPUT_IMG, dpi=150, bbox_inches="tight",
                facecolor=fig.get_facecolor())
    plt.close()

    kb = os.path.getsize(OUTPUT_IMG) / 1024
    print(f"✓ Heatmap guardado: {OUTPUT_IMG}  ({kb:.0f} KB)")
    print("\n  Para insertarlo en Power BI:")
    print("  Insertar → Imagen → seleccioná heatmap_correlacion.png")


if __name__ == "__main__":
    generar_heatmap()
