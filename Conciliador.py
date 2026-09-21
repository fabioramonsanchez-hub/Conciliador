# ============================================================
# CONCILIADOR AUTOMÁTICO DE FACTURAS
# ============================================================
import pandas as pd
from rapidfuzz import fuzz


# ---------- 1. LEER ARCHIVOS ----------

def leer_extracto(archivo):
    """Lee el extracto bancario y lo normaliza."""
    df = pd.read_excel(archivo)
    print("Columnas del extracto:", df.columns.tolist())

    df.columns = df.columns.str.lower().str.strip()
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
    df["monto_abs"] = df["monto"].abs()
    df["descripcion_norm"] = (
        df["descripcion"]
        .astype(str)
        .str.lower()
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    df = df.dropna(subset=["fecha", "monto"])
    return df


def leer_facturas(archivo):
    """Lee las facturas y las normaliza."""
    df = pd.read_excel(archivo)
    print("Columnas de facturas:", df.columns.tolist())

    df.columns = df.columns.str.lower().str.strip()
    df["fecha"] = pd.to_datetime(df["fecha"], errors="coerce")
    df["monto_abs"] = df["monto"].abs()
    df["proveedor_norm"] = (
        df["proveedor"]
        .astype(str)
        .str.lower()
        .str.replace(r"\s+", " ", regex=True)
        .str.strip()
    )
    df = df.dropna(subset=["fecha", "monto"])
    return df


# ---------- 2. CONCILIAR ----------

def conciliar(extracto, facturas, tolerancia_dias=3, tolerancia_monto=1.0):
    """Cruza facturas con el extracto bancario."""
    resultados = []
    usados = set()

    for _, f in facturas.iterrows():
        mejor_idx = None
        mejor_score = 0
        mejor_detalle = {}

        for idx_e, e in extracto.iterrows():
            if idx_e in usados:
                continue

            # Monto
            dif_monto = abs(e["monto_abs"] - f["monto_abs"])
            if dif_monto > tolerancia_monto:
                continue

            # Fecha
            dif_dias = abs((e["fecha"] - f["fecha"]).days)
            if dif_dias > tolerancia_dias:
                continue

            # Texto similar
            score_texto = fuzz.partial_ratio(
                e["descripcion_norm"], f["proveedor_norm"]
            )
            score_fecha = max(0, 100 - dif_dias * 15)
            score_final = score_texto * 0.7 + score_fecha * 0.3

            if score_final > mejor_score:
                mejor_score = score_final
                mejor_idx = idx_e
                mejor_detalle = {"dias_dif": dif_dias}

        if mejor_idx is not None and mejor_score >= 70:
            usados.add(mejor_idx)
            estado = "conciliado" if mejor_score >= 85 else "dudoso"
            resultados.append({
                "factura_num": f.get("numero", ""),
                "proveedor": f["proveedor"],
                "monto_factura": f["monto_abs"],
                "fecha_factura": f["fecha"].date(),
                "match_banco": extracto.loc[mejor_idx, "descripcion"],
                "fecha_banco": extracto.loc[mejor_idx, "fecha"].date(),
                "score": round(mejor_score, 1),
                "dias_dif": mejor_detalle["dias_dif"],
                "estado": estado,
            })
        else:
            resultados.append({
                "factura_num": f.get("numero", ""),
                "proveedor": f["proveedor"],
                "monto_factura": f["monto_abs"],
                "fecha_factura": f["fecha"].date(),
                "match_banco": None,
                "fecha_banco": None,
                "score": 0,
                "dias_dif": None,
                "estado": "sin_match",
            })

    return pd.DataFrame(resultados)


# ---------- 3. GENERAR REPORTE ----------

def generar_reporte(resultados, salida="conciliacion.xlsx"):
    """Genera un Excel con 4 hojas."""
    with pd.ExcelWriter(salida, engine="openpyxl") as writer:
        resumen = pd.DataFrame({
            "Estado": ["Conciliados", "Dudosos", "Sin match", "TOTAL"],
            "Cantidad": [
                (resultados.estado == "conciliado").sum(),
                (resultados.estado == "dudoso").sum(),
                (resultados.estado == "sin_match").sum(),
                len(resultados),
            ],
        })
        resumen.to_excel(writer, sheet_name="Resumen", index=False)

        for estado, nombre in [
            ("conciliado", "OK"),
            ("dudoso", "Revisar"),
            ("sin_match", "Sin_match"),
        ]:
            subset = resultados[resultados.estado == estado]
            subset.to_excel(writer, sheet_name=nombre, index=False)

    print(f"Reporte generado: {salida}")


# ---------- 4. PROGRAMA PRINCIPAL ----------

if __name__ == "__main__":
    print("Leyendo archivos...")
    extracto = leer_extracto("extracto_banco.xlsx")
    facturas = leer_facturas("facturas.xlsx")

    print(f"  -> {len(extracto)} movimientos en extracto")
    print(f"  -> {len(facturas)} facturas")

    print("\nConciliando...")
    resultados = conciliar(extracto, facturas, tolerancia_dias=3)

    print("\nResultados:")
    print(resultados[["proveedor", "monto_factura", "match_banco", "score", "estado"]])

    print("\nResumen:")
    print(resultados["estado"].value_counts())

    generar_reporte(resultados)