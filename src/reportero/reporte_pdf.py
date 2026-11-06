"""Genera un resumen ejecutivo de una página en PDF (ideal para enviar por correo)."""
from __future__ import annotations

from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.backends.backend_pdf import PdfPages  # noqa: E402
from matplotlib.ticker import FuncFormatter  # noqa: E402

AZUL = "#1F4E79"
PALETA = ["#1F4E79", "#2E86C1", "#85C1E9", "#F39C12", "#7F8C8D"]


def _clp(valor: float) -> str:
    return "$" + f"{valor:,.0f}".replace(",", ".")


MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def _mes_legible(periodo: str) -> str:
    """'2026-07' -> 'Jul 2026'"""
    anio, mes = periodo.split("-")
    return f"{MESES[int(mes) - 1]} {anio}"


def _millones(x, _pos) -> str:
    return f"${x / 1_000_000:.1f}M".replace(".", ",")


def generar_pdf(ind: dict, destino: Path) -> None:
    fig = plt.figure(figsize=(8.27, 11.69))  # A4 vertical
    fig.text(0.07, 0.95, "Resumen ejecutivo de ventas", fontsize=20, weight="bold", color=AZUL)
    fig.text(0.07, 0.925, f"Período: {ind['periodo']}", fontsize=10, color="#595959")

    # Tarjetas de KPIs
    formatos = {"Ventas totales": _clp, "Ticket promedio": _clp}
    for i, (nombre, valor) in enumerate(ind["kpis"].items()):
        x = 0.07 + i * 0.225
        ax = fig.add_axes([x, 0.83, 0.2, 0.07])
        ax.set_facecolor("#DDEBF7")
        ax.set_xticks([]), ax.set_yticks([])
        for s in ax.spines.values():
            s.set_visible(False)
        texto = formatos.get(nombre, lambda v: f"{v:,.0f}".replace(",", "."))(valor)
        ax.text(0.5, 0.66, nombre, ha="center", fontsize=8.5, color="#595959")
        ax.text(0.5, 0.25, texto, ha="center", fontsize=13, weight="bold", color=AZUL)

    # Ventas por sucursal
    suc = ind["por_sucursal"]
    ax1 = fig.add_axes([0.1, 0.56, 0.38, 0.22])
    ax1.bar(suc["sucursal"], suc["ventas"], color=PALETA[: len(suc)])
    ax1.set_title("Ventas por sucursal", fontsize=11, loc="left", color=AZUL)
    ax1.yaxis.set_major_formatter(FuncFormatter(_millones))
    ax1.spines[["top", "right"]].set_visible(False)
    ax1.tick_params(labelsize=8)

    # Categorías
    cat = ind["categorias"]
    ax2 = fig.add_axes([0.56, 0.56, 0.36, 0.22])
    _, _, porcentajes = ax2.pie(cat["total"], labels=cat["categoria"], autopct="%1.0f%%", colors=PALETA,
            textprops={"fontsize": 8}, startangle=90, wedgeprops={"width": 0.45},
            pctdistance=0.78)
    for t in porcentajes:  # porcentajes en blanco sobre el anillo
        t.set_color("white")
        t.set_weight("bold")
    ax2.set_title("Participación por categoría", fontsize=11, loc="left", color=AZUL)

    # Evolución mensual
    mes = ind["por_mes"]
    etiquetas = [_mes_legible(m) for m in mes["mes"]]
    ax3 = fig.add_axes([0.1, 0.29, 0.82, 0.2])
    for i, col in enumerate(c for c in mes.columns if c not in ("mes", "Total")):
        ax3.plot(range(len(mes)), mes[col], marker="o", label=col, color=PALETA[i % len(PALETA)])
    ax3.set_xticks(range(len(mes)), etiquetas)
    ax3.set_title("Evolución mensual por sucursal", fontsize=11, loc="left", color=AZUL)
    ax3.yaxis.set_major_formatter(FuncFormatter(_millones))
    ax3.spines[["top", "right"]].set_visible(False)
    ax3.legend(fontsize=8, frameon=False)
    ax3.tick_params(labelsize=8)

    # Top 5 productos
    top = ind["productos"].head(5)
    ax4 = fig.add_axes([0.07, 0.06, 0.86, 0.17])
    ax4.axis("off")
    ax4.set_title("Top 5 productos por ventas", fontsize=11, loc="left", color=AZUL)
    tabla = ax4.table(
        cellText=[[p, c, f"{u:,}".replace(",", "."), _clp(v)]
                  for p, c, u, v in top[["producto", "categoria", "unidades", "ventas"]].itertuples(index=False)],
        colLabels=["Producto", "Categoría", "Unidades", "Ventas"],
        loc="upper center", cellLoc="left", colWidths=[0.42, 0.22, 0.14, 0.2],
    )
    tabla.auto_set_font_size(False)
    tabla.set_fontsize(9)
    tabla.scale(1, 1.4)
    for (fila, _), celda in tabla.get_celld().items():
        celda.set_edgecolor("#D9D9D9")
        if fila == 0:
            celda.set_facecolor(AZUL)
            celda.set_text_props(color="white", weight="bold")

    destino.parent.mkdir(parents=True, exist_ok=True)
    with PdfPages(destino) as pdf:
        pdf.savefig(fig)
    plt.close(fig)
