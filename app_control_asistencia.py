import io
import re
from datetime import time
from pathlib import Path

import pandas as pd
import streamlit as st


st.set_page_config(page_title="Control de asistencia", page_icon="🕒", layout="wide")


# ----------------------------
# Utilidades de procesamiento
# ----------------------------
def normalize_text(value: str) -> str:
    if pd.isna(value):
        return ""
    text = str(value).strip()
    text = re.sub(r"\s+", " ", text)
    return text


def canonical_estado(value: str) -> str:
    txt = normalize_text(value).lower().replace(".", "").replace("_", "").replace("-", "")
    txt = txt.replace(" ", "")

    entradas = {"m/ent", "ment", "entrada", "ingreso", "mentrada"}
    salidas = {"m/sal", "msal", "salida", "egreso", "msalida"}

    if txt in {x.replace("/", "") for x in entradas} or txt in entradas:
        return "M/Ent"
    if txt in {x.replace("/", "") for x in salidas} or txt in salidas:
        return "M/Sal"
    return normalize_text(value)


def detect_columns(df: pd.DataFrame):
    mapping = {}
    for col in df.columns:
        clean = normalize_text(col).lower()
        compact = clean.replace(" ", "")
        if compact in {"nombre", "nombres", "funcionario", "empleado", "personal"}:
            mapping["nombre"] = col
        elif compact in {"marc.", "marc", "marcado", "marcacion", "marcación", "fecha", "fechahora"}:
            mapping["marc"] = col
        elif compact in {"estado", "tipo", "tipoestado", "movimiento"}:
            mapping["estado"] = col
    return mapping


TURNOS = {
    "Mañana": time(7, 5),
    "Tarde": time(13, 5),
    "Noche": time(19, 5),
}


def infer_shift(dt: pd.Timestamp) -> str:
    if pd.isna(dt):
        return ""
    t = dt.time()
    minutes = t.hour * 60 + t.minute
    targets = {
        "Mañana": 7 * 60,
        "Tarde": 13 * 60,
        "Noche": 19 * 60,
    }
    closest = min(targets, key=lambda k: abs(minutes - targets[k]))
    return closest


def calc_delay_minutes(dt: pd.Timestamp, shift: str) -> int:
    if pd.isna(dt) or not shift:
        return 0
    cutoff = TURNOS[shift]
    cutoff_minutes = cutoff.hour * 60 + cutoff.minute
    current_minutes = dt.hour * 60 + dt.minute
    return max(0, current_minutes - cutoff_minutes)


def format_hours(delta) -> str:
    if pd.isna(delta):
        return ""
    total_minutes = int(delta.total_seconds() // 60)
    sign = "-" if total_minutes < 0 else ""
    total_minutes = abs(total_minutes)
    hours = total_minutes // 60
    minutes = total_minutes % 60
    return f"{sign}{hours:02d}:{minutes:02d}"


def parse_datetime_series(series: pd.Series) -> pd.Series:
    # Primero intento estándar día/mes/año
    parsed = pd.to_datetime(series, errors="coerce", dayfirst=True)
    # Si algún valor ya viene como número de Excel, esta línea ayuda en casos raros
    if parsed.isna().all() and pd.api.types.is_numeric_dtype(series):
        parsed = pd.to_datetime("1899-12-30") + pd.to_timedelta(series, unit="D")
    return parsed


def process_attendance(df: pd.DataFrame):
    cols = detect_columns(df)
    missing = [k for k in ["nombre", "marc", "estado"] if k not in cols]
    if missing:
        raise ValueError(
            "No encontré las columnas requeridas. La planilla debe incluir NOMBRE, MARC. y ESTADO."
        )

    work = df.copy()
    work = work.rename(columns={cols["nombre"]: "Nombre", cols["marc"]: "Marc.", cols["estado"]: "Estado"})
    work["Nombre"] = work["Nombre"].astype(str).str.strip()
    work["Marc. original"] = work["Marc."]
    work["Marc."] = parse_datetime_series(work["Marc."])
    work["Estado"] = work["Estado"].map(canonical_estado)
    work["Fecha"] = work["Marc."].dt.date
    work["Hora"] = work["Marc."].dt.strftime("%H:%M")
    work["Turno sugerido"] = work["Marc."].apply(infer_shift)
    work["Retraso (min)"] = work.apply(
        lambda r: calc_delay_minutes(r["Marc."], r["Turno sugerido"]) if r["Estado"] == "M/Ent" else 0,
        axis=1,
    )
    work["Tipo registro"] = work["Estado"].map({"M/Ent": "Ingreso", "M/Sal": "Salida"}).fillna("No reconocido")
    work = work.sort_values(["Nombre", "Marc.", "Estado"], kind="stable").reset_index(drop=True)

    pairs = []

    for nombre, group in work.groupby("Nombre", sort=False):
        open_entry = None
        for _, row in group.iterrows():
            estado = row["Estado"]
            marca = row["Marc."]
            if pd.isna(marca):
                pairs.append({
                    "Nombre": nombre,
                    "Fecha ingreso": "",
                    "Hora ingreso": "",
                    "Fecha salida": "",
                    "Hora salida": "",
                    "Turno": "",
                    "Retraso (min)": "",
                    "Horas trabajadas": "",
                    "Observación": "Marcación inválida o fecha/hora no reconocida",
                })
                continue

            if estado == "M/Ent":
                if open_entry is not None:
                    pairs.append({
                        "Nombre": nombre,
                        "Fecha ingreso": open_entry["Marc."].strftime("%d/%m/%Y"),
                        "Hora ingreso": open_entry["Marc."].strftime("%H:%M"),
                        "Fecha salida": "",
                        "Hora salida": "",
                        "Turno": infer_shift(open_entry["Marc."]),
                        "Retraso (min)": calc_delay_minutes(open_entry["Marc."], infer_shift(open_entry["Marc."])),
                        "Horas trabajadas": "",
                        "Observación": "Ingreso sin salida antes de un nuevo ingreso",
                    })
                open_entry = row

            elif estado == "M/Sal":
                if open_entry is None:
                    pairs.append({
                        "Nombre": nombre,
                        "Fecha ingreso": "",
                        "Hora ingreso": "",
                        "Fecha salida": row["Marc."].strftime("%d/%m/%Y"),
                        "Hora salida": row["Marc."].strftime("%H:%M"),
                        "Turno": infer_shift(row["Marc."]),
                        "Retraso (min)": "",
                        "Horas trabajadas": "",
                        "Observación": "Salida sin ingreso previo",
                    })
                else:
                    shift = infer_shift(open_entry["Marc."])
                    delay = calc_delay_minutes(open_entry["Marc."], shift)
                    delta = row["Marc."] - open_entry["Marc."]
                    obs = ""
                    if delta.total_seconds() < 0:
                        obs = "La salida es anterior al ingreso"
                    pairs.append({
                        "Nombre": nombre,
                        "Fecha ingreso": open_entry["Marc."].strftime("%d/%m/%Y"),
                        "Hora ingreso": open_entry["Marc."].strftime("%H:%M"),
                        "Fecha salida": row["Marc."].strftime("%d/%m/%Y"),
                        "Hora salida": row["Marc."].strftime("%H:%M"),
                        "Turno": shift,
                        "Retraso (min)": delay,
                        "Horas trabajadas": format_hours(delta),
                        "Observación": obs,
                    })
                    open_entry = None
            else:
                pairs.append({
                    "Nombre": nombre,
                    "Fecha ingreso": "",
                    "Hora ingreso": "",
                    "Fecha salida": "",
                    "Hora salida": "",
                    "Turno": "",
                    "Retraso (min)": "",
                    "Horas trabajadas": "",
                    "Observación": f"Estado no reconocido: {estado}",
                })

        if open_entry is not None:
            pairs.append({
                "Nombre": nombre,
                "Fecha ingreso": open_entry["Marc."].strftime("%d/%m/%Y"),
                "Hora ingreso": open_entry["Marc."].strftime("%H:%M"),
                "Fecha salida": "",
                "Hora salida": "",
                "Turno": infer_shift(open_entry["Marc."]),
                "Retraso (min)": calc_delay_minutes(open_entry["Marc."], infer_shift(open_entry["Marc."])),
                "Horas trabajadas": "",
                "Observación": "Ingreso sin salida",
            })

    pairs_df = pd.DataFrame(pairs)

    def to_minutes(hhmm):
        if not hhmm or pd.isna(hhmm):
            return 0
        sign = -1 if str(hhmm).startswith("-") else 1
        text = str(hhmm).replace("-", "")
        h, m = text.split(":")
        return sign * (int(h) * 60 + int(m))

    summary = pairs_df.copy()
    if not summary.empty:
        summary["Retraso_num"] = pd.to_numeric(summary["Retraso (min)"], errors="coerce").fillna(0)
        summary["Horas_min"] = summary["Horas trabajadas"].map(to_minutes)
        summary["Registros con observación"] = summary["Observación"].astype(str).str.strip().ne("").astype(int)
        summary_df = (
            summary.groupby("Nombre", as_index=False)
            .agg(
                Total_registros=("Nombre", "size"),
                Total_retraso_min=("Retraso_num", "sum"),
                Total_horas_min=("Horas_min", "sum"),
                Incidencias=("Registros con observación", "sum"),
            )
        )
        summary_df["Total horas trabajadas"] = summary_df["Total_horas_min"].apply(lambda x: f"{int(x // 60):02d}:{int(x % 60):02d}")
        summary_df = summary_df.drop(columns=["Total_horas_min"])
    else:
        summary_df = pd.DataFrame(columns=["Nombre", "Total_registros", "Total_retraso_min", "Incidencias", "Total horas trabajadas"])

    return work, pairs_df, summary_df


def export_excel(raw_df: pd.DataFrame, pairs_df: pd.DataFrame, summary_df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        raw_df.to_excel(writer, index=False, sheet_name="Marcaciones")
        pairs_df.to_excel(writer, index=False, sheet_name="Asistencia procesada")
        summary_df.to_excel(writer, index=False, sheet_name="Resumen")

        workbook = writer.book
        fmt_header = workbook.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
        fmt_time = workbook.add_format({"num_format": "hh:mm"})
        fmt_datetime = workbook.add_format({"num_format": "dd/mm/yyyy hh:mm"})

        for sheet_name, df in {
            "Marcaciones": raw_df,
            "Asistencia procesada": pairs_df,
            "Resumen": summary_df,
        }.items():
            ws = writer.sheets[sheet_name]
            for col_idx, col_name in enumerate(df.columns):
                width = max(len(str(col_name)), *(len(str(v)) for v in df[col_name].astype(str).head(200))) + 2
                ws.set_column(col_idx, col_idx, min(width, 28))
                ws.write(0, col_idx, col_name, fmt_header)

        ws = writer.sheets["Marcaciones"]
        if "Marc." in raw_df.columns:
            col_idx = list(raw_df.columns).index("Marc.")
            ws.set_column(col_idx, col_idx, 20, fmt_datetime)

    output.seek(0)
    return output.getvalue()


# ----------------------------
# Interfaz Streamlit
# ----------------------------
st.title("🕒 App de control de asistencia")
st.write(
    "Sube una planilla Excel con las columnas **NOMBRE**, **MARC.** y **ESTADO** para calcular retrasos y horas trabajadas."
)

with st.expander("Reglas de cálculo", expanded=False):
    st.markdown(
        """
- **M/Ent** se toma como ingreso.
- **M/Sal** se toma como salida.
- **Turno mañana:** tolerancia hasta **07:05**, retraso desde **07:06**.
- **Turno tarde:** tolerancia hasta **13:05**, retraso desde **13:06**.
- **Turno noche:** tolerancia hasta **19:05**, retraso desde **19:06**.
- Las **horas trabajadas** se calculan entre el ingreso y la salida emparejados.
        """
    )

archivo = st.file_uploader("Subir archivo Excel", type=["xlsx", "xls"])

if archivo is not None:
    try:
        xls = pd.ExcelFile(archivo)
        hoja = st.selectbox("Selecciona la hoja", xls.sheet_names, index=0)
        df = pd.read_excel(xls, sheet_name=hoja)

        st.subheader("Vista previa")
        st.dataframe(df.head(20), use_container_width=True)

        raw_df, pairs_df, summary_df = process_attendance(df)

        c1, c2, c3 = st.columns(3)
        with c1:
            st.metric("Registros procesados", len(pairs_df))
        with c2:
            total_delay = pd.to_numeric(pairs_df.get("Retraso (min)", pd.Series(dtype=float)), errors="coerce").fillna(0).sum()
            st.metric("Retraso total (min)", int(total_delay))
        with c3:
            st.metric("Funcionarios", int(summary_df["Nombre"].nunique() if not summary_df.empty else 0))

        tab1, tab2, tab3 = st.tabs(["Marcaciones", "Asistencia procesada", "Resumen"])
        with tab1:
            st.dataframe(raw_df, use_container_width=True)
        with tab2:
            st.dataframe(pairs_df, use_container_width=True)
        with tab3:
            st.dataframe(summary_df, use_container_width=True)

        excel_bytes = export_excel(raw_df, pairs_df, summary_df)
        out_name = f"procesado_{Path(archivo.name).stem}.xlsx"
        st.download_button(
            "📥 Descargar Excel procesado",
            data=excel_bytes,
            file_name=out_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        st.success("Archivo procesado correctamente.")

    except Exception as e:
        st.error(f"Ocurrió un error al procesar la planilla: {e}")
        st.info("Verifica que el archivo tenga las columnas NOMBRE, MARC. y ESTADO, y que MARC. incluya fecha y hora.")
else:
    st.info("Sube una planilla para empezar.")
