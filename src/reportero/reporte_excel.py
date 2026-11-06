"""Genera el reporte en Excel con formato, tablas y gráficos nativos."""
from __future__ import annotations

from pathlib import Path

import pandas as pd
from openpyxl import Workbook
from openpyxl.chart import BarChart, LineChart, PieChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo

AZUL = "1F4E79"
AZUL_CLARO = "DDEBF7"
GRIS = "595959"
FORMATO_CLP = '"$"#,##0'
BORDE = Border(bottom=Side(style="thin", color="BFBFBF"))


def calcular_indicadores(df: pd.DataFrame) -> dict[str, pd.DataFrame | dict]:
    kpis = {
        "Ventas totales": df["total"].sum(),
        "N° de ventas": len(df),
        "Unidades vendidas": int(df["cantidad"].sum()),
        "Ticket promedio": df["total"].mean(),
    }
    por_sucursal = (
        df.groupby("sucursal")
        .agg(ventas=("total", "sum"), transacciones=("total", "size"), unidades=("cantidad", "sum"))
        .sort_values("ventas", ascending=False)
        .reset_index()
    )
    por_sucursal["participacion"] = por_sucursal["ventas"] / por_sucursal["ventas"].sum()
    por_mes = df.pivot_table(index="mes", columns="sucursal", values="total", aggfunc="sum", fill_value=0)
    por_mes["Total"] = por_mes.sum(axis=1)
    por_mes = por_mes.reset_index()
    productos = (
        df.groupby(["producto", "categoria"])
        .agg(unidades=("cantidad", "sum"), ventas=("total", "sum"))
        .sort_values("ventas", ascending=False)
        .reset_index()
    )
    categorias = df.groupby("categoria")["total"].sum().sort_values(ascending=False).reset_index()
    vendedores = (
        df.groupby(["vendedor", "sucursal"])
        .agg(ventas=("total", "sum"), transacciones=("total", "size"))
        .sort_values("ventas", ascending=False)
        .reset_index()
    )
    vendedores["ticket_promedio"] = vendedores["ventas"] / vendedores["transacciones"]
    return {
        "kpis": kpis,
        "por_sucursal": por_sucursal,
        "por_mes": por_mes,
        "productos": productos,
        "categorias": categorias,
        "vendedores": vendedores,
    }


# ---------- utilidades de formato ----------
def _titulo(ws, texto: str, subtitulo: str = "") -> None:
    ws["A1"] = texto
    ws["A1"].font = Font(size=16, bold=True, color=AZUL)
    if subtitulo:
        ws["A2"] = subtitulo
        ws["A2"].font = Font(size=10, italic=True, color=GRIS)
    ws.sheet_view.showGridLines = False


def _escribir_tabla(ws, df: pd.DataFrame, fila: int, col: int, nombre: str,
                    formatos: dict[str, str] | None = None) -> tuple[int, int]:
    """Escribe un DataFrame como Tabla de Excel (con filtros). Devuelve (fila_fin, col_fin)."""
    formatos = formatos or {}
    encabezados = [str(c).replace("_", " ").capitalize() for c in df.columns]
    for j, h in enumerate(encabezados):
        ws.cell(row=fila, column=col + j, value=h)
    for i, registro in enumerate(df.itertuples(index=False), start=1):
        for j, valor in enumerate(registro):
            celda = ws.cell(row=fila + i, column=col + j, value=valor)
            fmt = formatos.get(df.columns[j])
            if fmt:
                celda.number_format = fmt
    fila_fin, col_fin = fila + len(df), col + len(df.columns) - 1
    rango = f"{get_column_letter(col)}{fila}:{get_column_letter(col_fin)}{fila_fin}"
    tabla = Table(displayName=nombre, ref=rango)
    tabla.tableStyleInfo = TableStyleInfo(name="TableStyleMedium2", showRowStripes=True)
    ws.add_table(tabla)
    return fila_fin, col_fin


def _anchos(ws, minimo: int = 10, maximo: int = 45) -> None:
    for columna in ws.columns:
        largo = max((len(str(c.value)) for c in columna if c.value is not None), default=0)
        ws.column_dimensions[get_column_letter(columna[0].column)].width = max(minimo, min(maximo, largo + 3))


# ---------- hojas ----------
def _hoja_resumen(ws, ind: dict, periodo: str) -> None:
    _titulo(ws, "Reporte de ventas consolidado", f"Período: {periodo}  ·  Generado automáticamente")
    formatos_kpi = {"Ventas totales": FORMATO_CLP, "Ticket promedio": FORMATO_CLP}
    for j, (nombre, valor) in enumerate(ind["kpis"].items()):
        col = 1 + j * 2
        ws.merge_cells(start_row=4, start_column=col, end_row=4, end_column=col + 1)
        ws.merge_cells(start_row=5, start_column=col, end_row=5, end_column=col + 1)
        etiqueta = ws.cell(row=4, column=col, value=nombre)
        etiqueta.font = Font(size=10, color="FFFFFF", bold=True)
        celda = ws.cell(row=5, column=col, value=valor)
        celda.font = Font(size=18, bold=True, color=AZUL)
        celda.number_format = formatos_kpi.get(nombre, "#,##0")
        for c in (col, col + 1):
            ws.cell(row=4, column=c).fill = PatternFill("solid", fgColor=AZUL)
            ws.cell(row=5, column=c).fill = PatternFill("solid", fgColor=AZUL_CLARO)
        etiqueta.alignment = celda.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[5].height = 34

    ws["A8"] = "Ventas por sucursal"
    ws["A8"].font = Font(size=12, bold=True, color=AZUL)
    fin, _ = _escribir_tabla(
        ws, ind["por_sucursal"], 9, 1, "TablaSucursales",
        {"ventas": FORMATO_CLP, "unidades": "#,##0", "participacion": "0.0%"},
    )
    grafico = BarChart()
    grafico.title = "Ventas por sucursal"
    grafico.y_axis.title = "CLP"
    grafico.y_axis.numFmt = '#,##0'
    grafico.legend = None
    grafico.x_axis.delete = grafico.y_axis.delete = False
    grafico.add_data(Reference(ws, min_col=2, min_row=9, max_row=fin), titles_from_data=True)
    grafico.set_categories(Reference(ws, min_col=1, min_row=10, max_row=fin))
    grafico.height, grafico.width = 7.5, 15
    ws.add_chart(grafico, "A15")

    ws["G8"] = "Ventas por categoría"
    ws["G8"].font = Font(size=12, bold=True, color=AZUL)
    fin_c, _ = _escribir_tabla(ws, ind["categorias"], 9, 7, "TablaCategorias", {"total": FORMATO_CLP})
    torta = PieChart()
    torta.title = "Participación por categoría"
    torta.add_data(Reference(ws, min_col=8, min_row=9, max_row=fin_c), titles_from_data=True)
    torta.set_categories(Reference(ws, min_col=7, min_row=10, max_row=fin_c))
    torta.height, torta.width = 7.5, 11
    ws.add_chart(torta, "G15")
    _anchos(ws, minimo=12)


def _hoja_mensual(ws, ind: dict) -> None:
    _titulo(ws, "Evolución mensual", "Ventas en CLP por sucursal y mes")
    df = ind["por_mes"]
    formatos = {c: FORMATO_CLP for c in df.columns if c != "mes"}
    fin, col_fin = _escribir_tabla(ws, df, 4, 1, "TablaMensual", formatos)
    linea = LineChart()
    linea.title = "Ventas mensuales por sucursal"
    linea.y_axis.numFmt = '#,##0'
    linea.x_axis.delete = linea.y_axis.delete = False
    linea.add_data(Reference(ws, min_col=2, max_col=col_fin - 1, min_row=4, max_row=fin), titles_from_data=True)
    linea.set_categories(Reference(ws, min_col=1, min_row=5, max_row=fin))
    linea.height, linea.width = 8, 18
    ws.add_chart(linea, f"A{fin + 3}")
    _anchos(ws, minimo=14)


def _hoja_simple(ws, titulo: str, sub: str, df: pd.DataFrame, nombre: str, formatos: dict) -> None:
    _titulo(ws, titulo, sub)
    _escribir_tabla(ws, df, 4, 1, nombre, formatos)
    ws.freeze_panes = "A5"
    _anchos(ws)


def generar_excel(df: pd.DataFrame, calidad: pd.DataFrame, destino: Path) -> dict:
    ind = calcular_indicadores(df)
    periodo = f"{df['fecha'].min():%d-%m-%Y} al {df['fecha'].max():%d-%m-%Y}"

    wb = Workbook()
    wb.active.title = "Resumen"  # antes de crear gráficos: sus referencias usan el nombre de la hoja
    _hoja_resumen(wb.active, ind, periodo)
    _hoja_mensual(wb.create_sheet("Mensual"), ind)
    _hoja_simple(wb.create_sheet("Productos"), "Ranking de productos", "Ordenado por ventas",
                 ind["productos"], "TablaProductos", {"ventas": FORMATO_CLP, "unidades": "#,##0"})
    _hoja_simple(wb.create_sheet("Vendedores"), "Desempeño de vendedores", "Ordenado por ventas",
                 ind["vendedores"], "TablaVendedores",
                 {"ventas": FORMATO_CLP, "ticket_promedio": FORMATO_CLP})
    detalle = df.copy()
    detalle["fecha"] = detalle["fecha"].dt.date
    _hoja_simple(wb.create_sheet("Detalle"), "Detalle de ventas (datos limpios)",
                 f"{len(df):,} registros consolidados".replace(",", "."), detalle, "TablaDetalle",
                 {"fecha": "DD-MM-YYYY", "precio_unitario": FORMATO_CLP, "total": FORMATO_CLP})
    _hoja_simple(wb.create_sheet("Calidad de datos"), "Calidad de datos",
                 "Correcciones aplicadas automáticamente a cada archivo de entrada",
                 calidad.fillna(0), "TablaCalidad", {})

    for ws in wb.worksheets:  # impresión: horizontal y ajustada al ancho de página
        ws.page_setup.orientation = "landscape"
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True

    destino.parent.mkdir(parents=True, exist_ok=True)
    wb.save(destino)
    return {**ind, "periodo": periodo}
