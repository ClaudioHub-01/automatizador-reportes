import pandas as pd

from reportero.limpieza import RegistroCalidad, _a_fecha, _a_numero, _mapa_columnas, cargar_y_limpiar


def test_montos_en_distintos_formatos():
    serie = pd.Series(["$12.990", "12990", "1.234.567", "15,5", "$1.234,50", None, "abc"])
    resultado = _a_numero(serie).tolist()
    assert resultado[:5] == [12990, 12990, 1234567, 15.5, 1234.5]
    assert pd.isna(resultado[5]) and pd.isna(resultado[6])


def test_fechas_chilenas_e_iso():
    serie = pd.Series(["05/08/2026", "2026-08-05", "2026-08-05 00:00:00"])
    assert _a_fecha(serie).dt.strftime("%Y-%m-%d").tolist() == ["2026-08-05"] * 3


def test_sinonimos_de_columnas():
    mapa = _mapa_columnas(["N° Boleta", "Fecha Venta", "Item", "Unidades", "Precio", "EJECUTIVO"])
    assert mapa == {
        "N° Boleta": "folio",
        "Fecha Venta": "fecha",
        "Item": "producto",
        "Unidades": "cantidad",
        "Precio": "precio_unitario",
        "EJECUTIVO": "vendedor",
    }


def test_elimina_vacias_duplicadas_e_invalidas(tmp_path):
    (tmp_path / "ventas_prueba.csv").write_text(
        "fecha;producto;cantidad;precio\n"
        "01/07/2026;Sábana;2;$10.000\n"
        "01/07/2026;Sábana;2;$10.000\n"  # duplicado
        ";;;\n"  # vacía
        "02/07/2026;Toalla;;$5.000\n",  # sin cantidad
        encoding="utf-8",
    )
    registro = RegistroCalidad()
    df = cargar_y_limpiar(tmp_path, registro)
    assert len(df) == 1
    assert df.loc[0, "total"] == 20000
    assert df.loc[0, "sucursal"] == "Prueba"
    fila = registro.como_dataframe().iloc[0]
    assert (fila.duplicados, fila.filas_vacias, fila.filas_invalidas) == (1, 1, 1)
