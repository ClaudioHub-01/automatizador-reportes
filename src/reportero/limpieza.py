"""
Lectura, normalización y limpieza de archivos de ventas heterogéneos.

Cada archivo puede venir con distinto separador, nombres de columna, formato
de fecha o de montos. Este módulo los lleva a un esquema único y registra
cada corrección para dejar trazabilidad (hoja "Calidad de datos").
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

# Esquema final y sinónimos aceptados para cada columna.
# Agregar un sinónimo aquí basta para soportar un nuevo formato de archivo.
SINONIMOS: dict[str, list[str]] = {
    "folio": ["folio", "n boleta", "boleta", "n documento", "documento", "id venta"],
    "fecha": ["fecha", "fecha venta", "fecha de venta", "dia"],
    "producto": ["producto", "item", "articulo", "descripcion"],
    "categoria": ["categoria", "familia", "linea", "rubro"],
    "cantidad": ["cantidad", "unidades", "cant", "qty"],
    "precio_unitario": ["precio unitario", "precio", "valor unitario", "p unitario"],
    "vendedor": ["vendedor", "ejecutivo", "vendedora", "responsable"],
}
OBLIGATORIAS = ["fecha", "producto", "cantidad", "precio_unitario"]
EXTENSIONES = {".csv", ".xlsx", ".xls"}


@dataclass
class RegistroCalidad:
    """Resumen de lo que se corrigió en cada archivo."""

    filas: list[dict] = field(default_factory=list)

    def agregar(self, archivo: str, **datos) -> None:
        self.filas.append({"archivo": archivo, **datos})

    def como_dataframe(self) -> pd.DataFrame:
        return pd.DataFrame(self.filas)


def _normalizar_texto(texto: str) -> str:
    """'N° Boleta ' -> 'n boleta'  |  'Categoría' -> 'categoria'"""
    texto = unicodedata.normalize("NFKD", str(texto)).encode("ascii", "ignore").decode()
    texto = re.sub(r"[^a-zA-Z0-9]+", " ", texto)
    return texto.strip().lower()


def _mapa_columnas(columnas: list[str]) -> dict[str, str]:
    indice = {s: canon for canon, sins in SINONIMOS.items() for s in sins}
    mapa = {}
    for col in columnas:
        canon = indice.get(_normalizar_texto(col)) or indice.get(_normalizar_texto(col).replace("_", " "))
        if canon and canon not in mapa.values():
            mapa[col] = canon
    return mapa


def _leer_archivo(ruta: Path) -> pd.DataFrame:
    if ruta.suffix.lower() in {".xlsx", ".xls"}:
        return pd.read_excel(ruta, dtype=str)
    # CSV: se detecta el separador automáticamente (',' ';' '\t' '|')
    return pd.read_csv(ruta, sep=None, engine="python", dtype=str, encoding="utf-8-sig")


def _a_numero(serie: pd.Series) -> pd.Series:
    """Convierte '$12.990', '12.990', '12990,5', '12990.0' a número."""

    def convertir(v):
        if pd.isna(v):
            return None
        s = re.sub(r"[^\d,.\-]", "", str(v))
        if not s:
            return None
        if "," in s and "." in s:  # 1.234,56 -> formato chileno
            s = s.replace(".", "").replace(",", ".")
        elif "," in s:  # 1234,56
            s = s.replace(",", ".")
        elif s.count(".") == 1 and len(s.split(".")[1]) == 3:  # 12.990 -> miles
            s = s.replace(".", "")
        elif s.count(".") > 1:  # 1.234.567
            s = s.replace(".", "")
        return float(s)

    return pd.to_numeric(serie.map(convertir), errors="coerce")


def _a_fecha(serie: pd.Series) -> pd.Series:
    """Acepta dd/mm/aaaa, aaaa-mm-dd y fechas con hora."""
    texto = serie.astype("string").str.strip()
    iso = texto.str.match(r"^\d{4}-\d{2}-\d{2}")
    resultado = pd.Series(pd.NaT, index=serie.index, dtype="datetime64[ns]")
    resultado[iso] = pd.to_datetime(texto[iso].str[:10], format="%Y-%m-%d", errors="coerce")
    resultado[~iso] = pd.to_datetime(texto[~iso], dayfirst=True, errors="coerce")
    return resultado


def _nombre_sucursal(ruta: Path) -> str:
    """'ventas_valparaiso.xlsx' -> 'Valparaiso'"""
    base = re.sub(r"^(ventas|reporte|datos)[_\- ]*", "", ruta.stem, flags=re.I)
    return base.replace("_", " ").strip().title() or ruta.stem


def cargar_y_limpiar(carpeta: Path, registro: RegistroCalidad) -> pd.DataFrame:
    """Lee todos los archivos de la carpeta y devuelve un único DataFrame limpio."""
    archivos = sorted(p for p in carpeta.iterdir() if p.suffix.lower() in EXTENSIONES)
    if not archivos:
        raise FileNotFoundError(f"No se encontraron archivos CSV/Excel en {carpeta}")

    partes = []
    for ruta in archivos:
        df = _leer_archivo(ruta)
        leidas = len(df)

        mapa = _mapa_columnas(list(df.columns))
        faltantes = [c for c in OBLIGATORIAS if c not in mapa.values()]
        if faltantes:
            registro.agregar(ruta.name, filas_leidas=leidas, filas_validas=0,
                             observacion=f"Omitido: faltan columnas {faltantes}")
            continue
        df = df.rename(columns=mapa)[list(mapa.values())]
        for col in SINONIMOS:
            if col not in df.columns:
                df[col] = None

        # 1) filas completamente vacías
        antes = len(df)
        df = df.dropna(how="all")
        vacias = antes - len(df)

        # 2) normalización de tipos y textos
        df["fecha"] = _a_fecha(df["fecha"])
        df["cantidad"] = _a_numero(df["cantidad"])
        df["precio_unitario"] = _a_numero(df["precio_unitario"])
        for col in ["producto", "categoria", "vendedor", "folio"]:
            df[col] = df[col].astype("string").str.strip().str.replace(r"\s+", " ", regex=True)
        textos_corregidos = int((df["producto"] != df["producto"].str.capitalize()).sum())
        df["producto"] = df["producto"].str.capitalize()
        df["categoria"] = df["categoria"].str.capitalize()
        df["vendedor"] = df["vendedor"].str.title()

        # 3) filas con datos obligatorios inválidos
        invalidas_mask = df[OBLIGATORIAS].isna().any(axis=1) | (df["cantidad"] <= 0)
        invalidas = int(invalidas_mask.sum())
        df = df[~invalidas_mask]

        # 4) duplicados exactos
        antes = len(df)
        df = df.drop_duplicates()
        duplicadas = antes - len(df)

        df["sucursal"] = _nombre_sucursal(ruta)
        df["total"] = df["cantidad"] * df["precio_unitario"]
        partes.append(df)

        registro.agregar(
            ruta.name,
            filas_leidas=leidas,
            filas_vacias=vacias,
            filas_invalidas=invalidas,
            duplicados=duplicadas,
            textos_normalizados=textos_corregidos,
            filas_validas=len(df),
            observacion="OK",
        )

    if not partes:
        raise ValueError("Ningún archivo tenía el formato mínimo requerido.")

    datos = pd.concat(partes, ignore_index=True)
    datos["cantidad"] = datos["cantidad"].astype(int)
    datos["mes"] = datos["fecha"].dt.to_period("M").astype(str)
    columnas = ["sucursal", "folio", "fecha", "mes", "producto", "categoria",
                "vendedor", "cantidad", "precio_unitario", "total"]
    return datos[columnas].sort_values(["fecha", "sucursal"]).reset_index(drop=True)
