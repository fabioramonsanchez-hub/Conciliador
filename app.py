# ============================================================
# CONCILIADOR WEB - Interfaz con Streamlit
# ============================================================
import streamlit as st
import pandas as pd
import io
from rapidfuzz import fuzz


# ============================================================
# CONFIGURACIÓN DE LA PÁGINA
# ============================================================
st.set_page_config(
    page_title="Conciliador de Facturas",
    page_icon="🧾",
    layout="wide"
)


# ============================================================
# FUNCIONES (las mismas del script, adaptadas para archivos subidos)
# ============================================================

def leer_extracto(archivo):
    """Lee el extracto bancario desde un archivo subido."""
    df = pd.read_excel(archivo)
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
    """Lee las facturas desde un archivo subido."""
    df = pd.read_excel(archivo)
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

            dif_monto = abs(e["monto_abs"] - f["monto_abs"])
            if dif_monto > tolerancia_monto:
                continue

            dif_dias = abs((e["fecha"] - f["fecha"]).days)
            if dif_dias > tolerancia_dias:
                continue

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


def generar_excel_bytes(resultados):
    """Genera el Excel en memoria para descargar."""
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
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

    return buffer.getvalue()


# ============================================================
# INTERFAZ DE USUARIO
# ============================================================

st.title("🧾 Conciliador Automático de Facturas")
st.markdown("Sube tu extracto bancario y tus facturas. La app las cruza automáticamente.")

st.divider()

# --- Zona de subida de archivos ---
col1, col2 = st.columns(2)

with col1:
    st.subheader("📄 Extracto bancario")
    extracto_file = st.file_uploader(
        "Sube el Excel del banco",
        type=["xlsx", "xls"],
        key="extracto"
    )

with col2:
    st.subheader("🧾 Facturas")
    facturas_file = st.file_uploader(
        "Sube el Excel de facturas",
        type=["xlsx", "xls"],
        key="facturas"
    )

st.divider()

# --- Opciones avanzadas ---
with st.expander("⚙️ Opciones avanzadas"):
    col_a, col_b = st.columns(2)
    with col_a:
        tolerancia_dias = st.slider(
            "Días de tolerancia entre factura y pago",
            min_value=0, max_value=15, value=3
        )
    with col_b:
        tolerancia_monto = st.number_input(
            "Tolerancia de monto (diferencia máxima)",
            min_value=0.0, value=1.0, step=0.5
        )

# --- Botón de procesar ---
if extracto_file and facturas_file:
    if st.button("🚀 Conciliar ahora", type="primary", width="stretch"):
        with st.spinner("Procesando archivos..."):
            try:
                extracto = leer_extracto(extracto_file)
                facturas = leer_facturas(facturas_file)

                resultados = conciliar(
                    extracto, facturas,
                    tolerancia_dias=tolerancia_dias,
                    tolerancia_monto=tolerancia_monto
                )

                st.success("✅ Conciliación completada")

                # --- Métricas ---
                conciliados = (resultados.estado == "conciliado").sum()
                dudosos = (resultados.estado == "dudoso").sum()
                sin_match = (resultados.estado == "sin_match").sum()

                m1, m2, m3, m4 = st.columns(4)
                m1.metric("Total facturas", len(resultados))
                m2.metric("✅ Conciliadas", conciliados)
                m3.metric("⚠️ Dudosas", dudosos)
                m4.metric("❌ Sin match", sin_match)

                st.divider()

                # --- Tabla de resultados ---
                st.subheader("📊 Detalle de la conciliación")

                filtro = st.radio(
                    "Ver:",
                    ["Todas", "Conciliadas", "Dudosas", "Sin match"],
                    horizontal=True
                )

                if filtro == "Conciliadas":
                    subset = resultados[resultados.estado == "conciliado"]
                elif filtro == "Dudosas":
                    subset = resultados[resultados.estado == "dudoso"]
                elif filtro == "Sin match":
                    subset = resultados[resultados.estado == "sin_match"]
                else:
                    subset = resultados

                st.dataframe(subset, width="stretch")

                # --- Descarga del Excel ---
                st.divider()
                excel_bytes = generar_excel_bytes(resultados)
                st.download_button(
                    label="📥 Descargar reporte Excel",
                    data=excel_bytes,
                    file_name="conciliacion.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    width="stretch",
                    type="primary"
                )

            except KeyError as e:
                st.error(f"❌ Falta una columna esperada en el archivo: {e}")
                st.info("Los archivos deben tener columnas: fecha, monto, descripcion (extracto) y fecha, proveedor, monto (facturas).")
            except Exception as e:
                st.error(f"❌ Error: {e}")
else:
    st.info("👆 Sube los dos archivos para empezar.")

# --- Footer ---
st.divider()
st.caption("Hecho con Python + Streamlit | Conciliador v1.0")