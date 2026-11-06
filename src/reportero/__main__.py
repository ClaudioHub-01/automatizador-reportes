"""
Automatizador de reportes de ventas.

Ejemplos:
    python -m reportero
    python -m reportero --entrada data/entrada --salida salida --pdf
    python -m reportero --pdf --enviar gerencia@empresa.cl
"""
from __future__ import annotations

import argparse
import logging
import sys
from datetime import datetime
from pathlib import Path

from .limpieza import RegistroCalidad, cargar_y_limpiar
from .reporte_excel import generar_excel

log = logging.getLogger("reportero")


def _argumentos() -> argparse.Namespace:
    p = argparse.ArgumentParser(prog="reportero", description="Consolida archivos de ventas y genera reportes.")
    p.add_argument("--entrada", type=Path, default=Path("data/entrada"), help="Carpeta con CSV/Excel de entrada")
    p.add_argument("--salida", type=Path, default=Path("salida"), help="Carpeta donde se guardan los reportes")
    p.add_argument("--pdf", action="store_true", help="Generar también un resumen ejecutivo en PDF")
    p.add_argument("--enviar", nargs="+", metavar="CORREO", help="Enviar el reporte a uno o más correos")
    return p.parse_args()


def main() -> int:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s  %(levelname)-7s %(message)s", datefmt="%H:%M:%S")
    args = _argumentos()
    marca = datetime.now().strftime("%Y%m%d_%H%M")

    try:
        registro = RegistroCalidad()
        log.info("Leyendo archivos desde %s", args.entrada)
        datos = cargar_y_limpiar(args.entrada, registro)
        calidad = registro.como_dataframe()
        for fila in calidad.to_dict("records"):
            log.info("  %-25s %4s filas leídas -> %4s válidas", fila["archivo"], fila["filas_leidas"], fila["filas_validas"])

        excel = args.salida / f"reporte_ventas_{marca}.xlsx"
        ind = generar_excel(datos, calidad, excel)
        log.info("Excel generado: %s", excel)
        adjuntos = [excel]

        if args.pdf:
            from .reporte_pdf import generar_pdf

            pdf = args.salida / f"resumen_ventas_{marca}.pdf"
            generar_pdf(ind, pdf)
            log.info("PDF generado:   %s", pdf)
            adjuntos.append(pdf)

        if args.enviar:
            from .correo import enviar_reporte

            k = ind["kpis"]
            resumen = (
                f"Período: {ind['periodo']}\n"
                f"Ventas totales: ${k['Ventas totales']:,.0f}\n".replace(",", ".")
                + f"N° de ventas: {k['N° de ventas']}"
            )
            enviar_reporte(args.enviar, adjuntos, resumen)
            log.info("Reporte enviado a: %s", ", ".join(args.enviar))
    except Exception as exc:  # noqa: BLE001 - se informa cualquier error al usuario
        log.error("%s", exc)
        return 1
    return 0


if __name__ == "__main__":
    sys.exit(main())
