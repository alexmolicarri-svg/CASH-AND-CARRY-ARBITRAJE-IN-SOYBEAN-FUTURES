"""
=============================================================================
EXTENSIÓN: GENERACIÓN DE FIGURAS PARA EL PAPER
=============================================================================
Módulo complementario a cash_carry_backtester.py. Lee el Excel de resultados
ya generado (hoja "Raw Data") y produce tres figuras académicas en PNG:

  Figure 1 — Spot proxy vs. target contract price over time
  Figure 2 — Implied net carry over time (contango/backwardation shading)
  Figure 3 — Distribution of trading signals (bar chart)

Uso:
    python generate_figures.py cash_carry_backtest_20260706_1632.xlsx

Requiere: pandas, matplotlib (pip install matplotlib si no está instalado).
=============================================================================
"""

import sys
from pathlib import Path

import pandas as pd
import matplotlib
matplotlib.use("Agg")  # backend sin ventana, seguro para scripts
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

# Paleta académica sobria, consistente con el resto del proyecto
COLOR_DARK = "#1a2332"
COLOR_ACCENT = "#8b1a1a"
COLOR_BLUE = "#2c5f8a"
COLOR_GREY = "#6b7280"


def load_data(excel_path: str) -> pd.DataFrame:
    """Carga la hoja 'Raw Data' del Excel de resultados del backtester."""
    df = pd.read_excel(excel_path, sheet_name="Raw Data")
    df["date"] = pd.to_datetime(df["date"])
    return df.sort_values("date")


def plot_prices(df: pd.DataFrame, output_dir: Path) -> None:
    """Figure 1 — Spot proxy vs. target contract price over time."""
    fig, ax = plt.subplots(figsize=(9, 4.3))
    ax.plot(df["date"], df["spot_bid"], color=COLOR_BLUE, lw=1.5,
            label=f"Spot proxy ({df['near_future_symbol'].iloc[0]})")
    ax.plot(df["date"], df["target_future_bid"], color=COLOR_ACCENT, lw=1.5,
            label=f"Target contract ({df['target_future_symbol'].iloc[0]})")
    ax.fill_between(df["date"], df["spot_bid"], df["target_future_bid"],
                     color=COLOR_ACCENT, alpha=0.08)
    ax.set_ylabel("USD / bushel", fontsize=10.5)
    ax.set_title("Figure 1 — Spot Proxy vs. Target Contract Price",
                 fontsize=11.5, weight="bold", color=COLOR_DARK)
    ax.legend(fontsize=9, frameon=False, loc="upper left")
    ax.grid(True, alpha=0.25)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b '%y"))
    ax.spines[["top", "right"]].set_visible(False)
    plt.tight_layout()
    plt.savefig(output_dir / "figure1_prices.png", dpi=200, bbox_inches="tight")
    plt.close()


def plot_carry(df: pd.DataFrame, output_dir: Path) -> None:
    """Figure 2 — Implied net carry over time, with contango/backwardation shading."""
    fig, ax = plt.subplots(figsize=(9, 3.8))
    carry_pct = df["implied_carry_rate"] * 100
    ax.plot(df["date"], carry_pct, color=COLOR_ACCENT, lw=1.3)
    ax.axhline(0, color=COLOR_DARK, lw=0.9)
    ax.fill_between(df["date"], carry_pct, 0, where=(carry_pct < 0),
                     color=COLOR_ACCENT, alpha=0.15, label="Backwardation (c_net < 0)")
    ax.fill_between(df["date"], carry_pct, 0, where=(carry_pct >= 0),
                     color=COLOR_BLUE, alpha=0.15, label="Contango (c_net > 0)")
    pair = df["carry_estimation_pair"].iloc[0]
    ax.set_ylabel("Annualized implied net carry (%)", fontsize=10.5)
    ax.set_title(f"Figure 2 — Implied Net Carry Over Time ({pair} calendar spread)",
                 fontsize=11.5, weight="bold", color=COLOR_DARK)
    ax.legend(fontsize=9, frameon=False, loc="upper right")
    ax.grid(True, alpha=0.25)
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b '%y"))
    ax.spines[["top", "right"]].set_visible(False)
    plt.tight_layout()
    plt.savefig(output_dir / "figure2_carry.png", dpi=200, bbox_inches="tight")
    plt.close()


def plot_signal_distribution(df: pd.DataFrame, output_dir: Path) -> None:
    """Figure 3 — Bar chart of trading signal distribution."""
    fig, ax = plt.subplots(figsize=(5.5, 4.0))
    labels_map = {
        "REVERSE_CASH_AND_CARRY_ARBITRAGE": "Reverse Cash-and-Carry\n(Backwardation)",
        "CASH_AND_CARRY_ARBITRAGE": "Cash-and-Carry\n(Contango)",
        "EQUILIBRIUM_NO_ARBITRAGE": "Equilibrium",
    }
    vc = df["signal"].value_counts()
    names = [labels_map.get(k, k) for k in vc.index]
    vals = vc.values
    colors_bar = [COLOR_ACCENT, COLOR_BLUE, COLOR_GREY][:len(vals)]
    bars = ax.bar(names, vals, color=colors_bar, width=0.55)
    n_total = len(df)
    for b, v in zip(bars, vals):
        ax.text(b.get_x() + b.get_width() / 2, v + 1.5, f"{v/n_total*100:.1f}%",
                ha="center", fontsize=10, weight="bold", color=COLOR_DARK)
    ax.set_ylabel("Number of trading days", fontsize=10.5)
    ax.set_title(f"Figure 3 — Signal Distribution (n={n_total} days)",
                 fontsize=11.5, weight="bold", color=COLOR_DARK)
    ax.grid(True, alpha=0.25, axis="y")
    ax.spines[["top", "right"]].set_visible(False)
    plt.tight_layout()
    plt.savefig(output_dir / "figure3_signals.png", dpi=200, bbox_inches="tight")
    plt.close()


def main() -> None:
    if len(sys.argv) < 2:
        print("Uso: python generate_figures.py <ruta_al_excel>.xlsx")
        sys.exit(1)

    excel_path = sys.argv[1]
    output_dir = Path(excel_path).resolve().parent
    df = load_data(excel_path)

    plot_prices(df, output_dir)
    plot_carry(df, output_dir)
    plot_signal_distribution(df, output_dir)

    print(f"3 figuras guardadas en: {output_dir}")
    print("  - figure1_prices.png")
    print("  - figure2_carry.png")
    print("  - figure3_signals.png")


if __name__ == "__main__":
    main()