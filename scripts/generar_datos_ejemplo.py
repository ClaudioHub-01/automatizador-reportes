"""
Genera datos de ventas FICTICIOS y deliberadamente "desordenados" para la demo.

Simula una distribuidora con 3 sucursales que envían sus ventas mensuales
cada una en un formato distinto (como pasa en la vida real):

- Santiago:    CSV con ';', decimales con coma, montos tipo "$12.990", fechas dd/mm/aaaa
- Valparaíso:  Excel con nombres de columna distintos y fechas aaaa-mm-dd
- Concepción:  CSV con ',' , columnas en MAYÚSCULAS, filas duplicadas y vacías

Uso:
    python scripts/generar_datos_ejemplo.py
"""
from __future__ import annotations

import random
from datetime import date, timedelta
from pathlib import Path

import pandas as pd

SEED = 42
RAIZ = Path(__file__).resolve().parents[1]
DESTINO = RAIZ / "data" / "entrada"

PRODUCTOS = {
    "Sábana 2 plazas": ("Ropa de cama", 18990),
    "Plumón 1,5 plazas": ("Ropa de cama", 34990),
    "Almohada viscoelástica": ("Ropa de cama", 15990),
    "Respaldo tapizado 2 plazas": ("Respaldos", 89990),
    "Respaldo madera 1,5 plazas": ("Respaldos", 64990),
    "Cubrecama quilt": ("Ropa de cama", 29990),
    "Toalla de baño": ("Baño", 7990),
    "Set de toallas x3": ("Baño", 19990),
}
VENDEDORES = {
    "Santiago": ["Ana Rojas", "Pedro Soto", "Camila Díaz"],
    "Valparaíso": ["Luis Muñoz", "Javiera Pérez"],
    "Concepción": ["Matías Vera", "Fernanda Silva"],
}


def _ventas(sucursal: str, n: int, inicio: date, dias: int) -> list[dict]:
    filas = []
    for i in range(n):
        prod = random.choice(list(PRODUCTOS))
        categoria, precio = PRODUCTOS[prod]
        filas.append(
            {
                "folio": f"{sucursal[:3].upper()}-{1000 + i}",
                "fecha": inicio + timedelta(days=random.randint(0, dias - 1)),
                "producto": prod,
                "categoria": categoria,
                "cantidad": random.choices([1, 2, 3, 4], weights=[60, 25, 10, 5])[0],
                "precio_unitario": precio,
                "vendedor": random.choice(VENDEDORES[sucursal]),
            }
        )
    return filas


def _monto_clp(valor: int) -> str:
    return "$" + f"{valor:,}".replace(",", ".")


def main() -> None:
    random.seed(SEED)
    DESTINO.mkdir(parents=True, exist_ok=True)
    inicio, dias = date(2026, 7, 1), 92  # julio a septiembre

    # --- Santiago: CSV ';' + montos "$12.990" + fechas dd/mm/aaaa ---
    stgo = pd.DataFrame(_ventas("Santiago", 420, inicio, dias))
    stgo["fecha"] = stgo["fecha"].map(lambda d: d.strftime("%d/%m/%Y"))
    stgo["precio_unitario"] = stgo["precio_unitario"].map(_monto_clp)
    stgo.columns = ["Folio", "Fecha", "Producto", "Categoría", "Cantidad", "Precio Unitario", "Vendedor"]
    stgo.to_csv(DESTINO / "ventas_santiago.csv", sep=";", index=False, encoding="utf-8-sig")

    # --- Valparaíso: Excel con otros nombres de columna ---
    valpo = pd.DataFrame(_ventas("Valparaíso", 260, inicio, dias))
    valpo = valpo.rename(
        columns={
            "folio": "N° Boleta",
            "fecha": "Fecha Venta",
            "producto": "Item",
            "categoria": "Familia",
            "cantidad": "Unidades",
            "precio_unitario": "Precio",
            "vendedor": "Ejecutivo",
        }
    )
    valpo["Fecha Venta"] = pd.to_datetime(valpo["Fecha Venta"])
    # algunos productos escritos con espacios/mayúsculas inconsistentes
    for idx in random.sample(range(len(valpo)), 25):
        valpo.loc[idx, "Item"] = "  " + valpo.loc[idx, "Item"].upper() + " "
    valpo.to_excel(DESTINO / "ventas_valparaiso.xlsx", index=False, sheet_name="Ventas")

    # --- Concepción: CSV ',' + MAYÚSCULAS + duplicados + filas vacías ---
    conce = pd.DataFrame(_ventas("Concepción", 230, inicio, dias))
    conce.columns = [c.upper() for c in conce.columns]
    conce["FECHA"] = conce["FECHA"].map(lambda d: d.strftime("%Y-%m-%d"))
    duplicados = conce.sample(12, random_state=SEED)
    vacias = pd.DataFrame([{c: None for c in conce.columns}] * 5)
    sin_cantidad = conce.sample(4, random_state=SEED + 1).copy()
    sin_cantidad["CANTIDAD"] = None
    sin_cantidad["FOLIO"] = sin_cantidad["FOLIO"] + "-X"
    conce = pd.concat([conce, duplicados, vacias, sin_cantidad]).sample(frac=1, random_state=SEED)
    conce.to_csv(DESTINO / "ventas_concepcion.csv", index=False, encoding="utf-8")

    print(f"Datos de ejemplo generados en: {DESTINO}")
    for f in sorted(DESTINO.iterdir()):
        print("  -", f.name)


if __name__ == "__main__":
    main()
