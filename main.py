import tkinter as tk
from tkinter import filedialog
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
from matplotlib.ticker import FuncFormatter
from datetime import datetime
import os

# ==============================
# Einstellungen
# ==============================
trenner = ";"                                # CSV-Separator
haupt_datum = "WP_Bollinger_Baender.Datum"   # EINZIGE Datumsspalte

# Farben/Linienstile
CLR_PRICE = "#1f3a93"
CLR_MEAN  = "#6f42c1"
CLR_BUP   = "#ff8c00"
CLR_BLOW  = "#7fd1d1"
CLR_LIN   = "#6c757d"
CLR_STUP  = "#2e8b57"
CLR_STDN  = "#cc0000"
CLR_VOL_G = "#2ca02c"
CLR_VOL_R = "#d62728"

linienstile = {
    "Schlusskurs":     {"linestyle": "-",  "linewidth": 2.2, "color": CLR_PRICE},
    "Mean_20":         {"linestyle": "--", "linewidth": 1.6, "color": CLR_MEAN},
    "Bollinger_upper": {"linestyle": "-",  "linewidth": 1.2, "color": CLR_BUP},
    "Bollinger_lower": {"linestyle": "-",  "linewidth": 1.2, "color": CLR_BLOW},
    "Linie_6":         {"linestyle": "--", "linewidth": 1.0, "color": CLR_LIN},
    "Linie_8":         {"linestyle": "--", "linewidth": 1.0, "color": CLR_LIN},
    "Linie_50":        {"linestyle": ":",  "linewidth": 1.2, "color": CLR_LIN},
    "Linie_100":       {"linestyle": ":",  "linewidth": 1.2, "color": CLR_LIN},
    "Linie_200":       {"linestyle": "-.", "linewidth": 1.4, "color": CLR_LIN},
    "Supertrend_up":   {"linestyle": "-",  "linewidth": 1.5, "color": CLR_STUP},
    "Supertrend_down": {"linestyle": "-",  "linewidth": 1.5, "color": CLR_STDN},
}

dezimal_spalten = [
    "Schlusskurs","Mean_20","Variance_20","Bollinger_upper","Bollinger_lower",
    "Linie_6","Linie_8","Linie_50","Linie_100","Linie_200",
    "Unterstuetzungspunkte","Supertrend_up","Supertrend_down",
    "Linie_Volumen_gruen","Linie_Volumen_rot"
]

# ==============================
# Hilfsfunktionen
# ==============================
def datei_auswaehlen():
    root = tk.Tk(); root.withdraw()
    return filedialog.askopenfilename(
        title="CSV-Datei auswählen",
        filetypes=[("CSV-Dateien", "*.csv"), ("Alle Dateien", "*.*")]
    )

def plot_serie(ax, data, value_col, label=None, style=None):
    """Plottet eine Serie gegen die EINZIGE Datumsspalte und gibt y-Werte zurück (für min/max)."""
    if value_col not in data.columns:
        return None
    x = data[haupt_datum]
    y = data[value_col]
    m = x.notna() & y.notna()
    if not m.any():
        return None
    ax.plot(x[m], y[m], label=label or value_col, **(style or {}))
    return y[m]

# ==============================
# Datei wählen & CSV laden
# ==============================
csv_pfad = datei_auswaehlen()
if not csv_pfad:
    print("Keine Datei gewählt. Beende.")
    raise SystemExit

df = pd.read_csv(csv_pfad, sep=trenner)

# Dezimal-Komma -> Punkt -> float (nur vorhandene Spalten)
for c in [c for c in dezimal_spalten if c in df.columns]:
    df[c] = pd.to_numeric(df[c].astype(str).str.replace(",", ".", regex=False), errors="coerce")

# ==============================
# Nur EINE Datumsspalte parsen – ohne Warnungen
# ==============================
if haupt_datum not in df.columns:
    raise KeyError(f"Datumsspalte '{haupt_datum}' nicht gefunden.")

s = df[haupt_datum].astype(str).str.strip()
parsed = None
for fmt in ("%d.%m.%Y", "%d.%m.%y"):   # zuerst 4-stelliges Jahr versuchen, dann 2-stellig
    try:
        parsed = pd.to_datetime(s, format=fmt, errors="raise")
        break
    except Exception:
        pass
if parsed is None:
    # finaler Fallback, falls doch etwas „krummes“ drin ist
    parsed = pd.to_datetime(s, dayfirst=True, errors="coerce")
df[haupt_datum] = parsed

# ==============================
# Zielordner wählen
# ==============================
root = tk.Tk(); root.withdraw()
ziel_ordner = filedialog.askdirectory(title="Ordner zum Speichern der Charts auswählen")
if not ziel_ordner:
    heute = datetime.now().strftime("%Y-%m-%d")
    ziel_ordner = f"charts_{heute}"
    os.makedirs(ziel_ordner, exist_ok=True)

heute = datetime.now().strftime("%Y-%m-%d")

# ==============================
# Pro WKN Chart erstellen
# ==============================
for wkn, g in df.groupby("WKN"):
    g = g.sort_values(haupt_datum)

    fig = plt.figure(figsize=(13, 8))
    # Zwei Zeilen: oben Preis (6 Teile), unten Volumen (1 Teil)
    gs = fig.add_gridspec(nrows=2, ncols=1, height_ratios=[6, 1], hspace=0.04)
    ax  = fig.add_subplot(gs[0, 0])
    axv = fig.add_subplot(gs[1, 0], sharex=ax)

    fig.patch.set_facecolor("white")
    ax.set_facecolor("white")
    axv.set_facecolor("white")

    y_for_range = []

    # --- Preis & Indikatoren ---
    y = plot_serie(ax, g, "Schlusskurs", "Schlusskurs", linienstile["Schlusskurs"])
    if y is not None: y_for_range.append(y)

    if "Mean_20" in g.columns:
        y = plot_serie(ax, g, "Mean_20", "Mean 20", linienstile["Mean_20"])
        if y is not None: y_for_range.append(y)

    if {"Bollinger_upper","Bollinger_lower"}.issubset(g.columns):
        y = plot_serie(ax, g, "Bollinger_upper", "Bollinger upper", linienstile["Bollinger_upper"])
        if y is not None: y_for_range.append(y)
        y = plot_serie(ax, g, "Bollinger_lower", "Bollinger lower", linienstile["Bollinger_lower"])
        if y is not None: y_for_range.append(y)
        xb = g[haupt_datum]; bu = g["Bollinger_upper"]; bl = g["Bollinger_lower"]
        m = xb.notna() & bu.notna() & bl.notna()
        if m.any():
            ax.fill_between(xb[m], bl[m], bu[m], alpha=0.12, linewidth=0, color=CLR_BLOW)

    for n in [6, 8, 50, 100, 200]:
        col = f"Linie_{n}"
        if col in g.columns:
            y = plot_serie(ax, g, col, col, linienstile[col])
            if y is not None: y_for_range.append(y)

    if "Supertrend_up" in g.columns:
        y = plot_serie(ax, g, "Supertrend_up", "Supertrend up", linienstile["Supertrend_up"])
        if y is not None: y_for_range.append(y)
    if "Supertrend_down" in g.columns:
        y = plot_serie(ax, g, "Supertrend_down", "Supertrend down", linienstile["Supertrend_down"])
        if y is not None: y_for_range.append(y)

    # --- Unterstützungspunkte: vertikale Linien an den Datumspositionen ---
    if "Unterstuetzungspunkte" in g.columns:
        xs = g[haupt_datum]
        ys = g["Unterstuetzungspunkte"]
        m = xs.notna() & ys.notna()
        x_dates = pd.to_datetime(xs[m]).dropna().unique()
        for i, xdate in enumerate(sorted(x_dates)):
            ax.axvline(xdate, color="#777", linestyle="--", linewidth=0.9, alpha=0.6,
                       label="Unterstützung (Datum)" if i == 0 else None)
        if m.any():
            y_for_range.append(ys[m])

    # --- Volumen unten ---
    if "Linie_Volumen_gruen" in g.columns or "Linie_Volumen_rot" in g.columns:
        xv = g[haupt_datum]
        vg = g.get("Linie_Volumen_gruen", pd.Series([0]*len(g), index=g.index)).fillna(0)
        vr = g.get("Linie_Volumen_rot",   pd.Series([0]*len(g), index=g.index)).fillna(0)

        axv.bar(xv, vg, width=1.0, alpha=0.8, label="Volumen grün", color=CLR_VOL_G)
        axv.bar(xv, vr, width=1.0, alpha=0.8, label="Volumen rot",  color=CLR_VOL_R, bottom=vg)

        axv.axhline(0, linewidth=0.8, color="#666", alpha=0.6, linestyle="--")
        vmax = float((vg + vr).max())
        if vmax > 0:
            axv.set_ylim(0, vmax * 1.25)

        def _fmt_vol(x, _pos):
            axabs = abs(x)
            if axabs >= 1_000_000_000: return f"{x/1_000_000_000:.1f}B"
            if axabs >= 1_000_000:     return f"{x/1_000_000:.1f}M"
            if axabs >= 1_000:         return f"{x/1_000:.1f}k"
            return f"{x:.0f}"
        axv.yaxis.set_major_formatter(FuncFormatter(_fmt_vol))
        axv.set_ylabel("Vol.", fontsize=9, labelpad=4)

    # --- Achsenformatierung ---
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%d.%m.%Y"))
    axv.xaxis.set_major_formatter(mdates.DateFormatter("%d.%m.%Y"))
    plt.setp(ax.get_xticklabels(), visible=False)
    plt.setp(axv.get_xticklabels(), rotation=45)
    ax.grid(True, which="major", alpha=0.25)
    axv.grid(True, which="major", alpha=0.15)
    for spine in ["top", "right"]:
        ax.spines[spine].set_visible(False)
        axv.spines[spine].set_visible(False)

    ax.set_title(f"WKN {wkn} – Kurs & Indikatoren", fontsize=13, pad=8)
    ax.set_ylabel("Preis")

    # --- Y-Range korrekt bestimmen (alle relevanten Serien) ---
    if y_for_range:
        y_all = pd.concat(y_for_range, axis=0)
        y_all = y_all[np.isfinite(y_all)]
        if not y_all.empty:
            ymin, ymax = float(y_all.min()), float(y_all.max())
            if ymin == ymax:
                pad = 0.05 * (abs(ymin) if ymin != 0 else 1.0)
            else:
                pad = (ymax - ymin) * 0.07  # 7% Luft
            ax.set_ylim(ymin - pad, ymax + pad)

    # ===== Legende UNTER den Grafiken =====
    h1, l1 = ax.get_legend_handles_labels()
    h2, l2 = axv.get_legend_handles_labels()
    handles, labels = (h1 + h2, l1 + l2)

    # Platz für Legende unterhalb reservieren (kein tight/constrained layout)
    fig.subplots_adjust(left=0.08, right=0.98, top=0.92, bottom=0.26, hspace=0.04)

    if labels:
        fig.legend(
            handles, labels,
            loc="lower center",
            bbox_to_anchor=(0.5, 0.03),  # klar UNTER der x-Achse
            ncol=4, frameon=True
        )

    # --- Speichern ---
    dateiname = f"{heute}_WKN{wkn}.png"
    plt.savefig(os.path.join(ziel_ordner, dateiname), dpi=150)
    plt.close()

print(f"Charts gespeichert in: {ziel_ordner}")