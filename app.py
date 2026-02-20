# app.py — Carga manual por grilla (sin Excel) · Multi-Base · Dashboard completo
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import BytesIO
from pathlib import Path
from datetime import datetime

# ==========================
# Configuración visual (tema oscuro pro)
# ==========================
st.set_page_config(page_title="Plan vs Real — Operación", layout="wide")
TEMPLATE = "plotly_dark"
FONT = "Inter, system-ui, Segoe UI, Roboto"

def stylize(fig, title=None, y_pct=False):
    fig.update_layout(
        template=TEMPLATE, title=title,
        font=dict(family=FONT, size=13, color="#E5E7EB"),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        legend_title_text="", margin=dict(t=45, r=10, b=30, l=10),
    )
    if y_pct:
        fig.update_yaxes(tickformat=".0%")
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="rgba(148,163,184,.25)")
    return fig

# ==========================
# Persistencia en /data
# ==========================
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)
MERGED_CSV = DATA_DIR / "merged.csv"   # todos los datos de todas las bases
BASES_DIR = DATA_DIR / "bases"
BASES_DIR.mkdir(exist_ok=True)

def save_csv(df: pd.DataFrame, path: Path):
    df.to_csv(path, index=False, encoding="utf-8")

def load_csv(path: Path) -> pd.DataFrame|None:
    return pd.read_csv(path, encoding="utf-8") if path.exists() else None

def to_excel_bytes(df: pd.DataFrame, sheet_name="datos", fname="reporte.xlsx") -> tuple[bytes,str]:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as wr:
        df.to_excel(wr, index=False, sheet_name=sheet_name)
    return buf.getvalue(), fname

# ==========================
# Estado de sesión
# ==========================
if "bases" not in st.session_state:
    # dict: base -> DataFrame normalizado (todas las columnas ya procesadas)
    st.session_state["bases"] = {}
if "merged" not in st.session_state:
    st.session_state["merged"] = pd.DataFrame()

# ==========================
# Utilidades de datos
# ==========================
GRID_COLS = ["HORA","SVC PROY","SVC REALES","MOV REQ","MOVILES X NOMINA","COEFICIENTE HS","DIF MOVILES"]

def default_grid():
    # 24 filas de 00:00 a 23:00
    hh = [f"{h:02d}:00" for h in range(24)]
    df = pd.DataFrame({"HORA": hh})
    for c in GRID_COLS[1:]:
        df[c] = np.nan
    return df

def _to_time(s):
    # Espera "HH:MM" -> time
    return pd.to_datetime(str(s), errors="coerce").time()

def _to_num(s):
    # "#¿NOMBRE?" u otros -> NaN
    return pd.to_numeric(s, errors="coerce")

def normalize_grid(df_grid: pd.DataFrame, fecha: str, base: str) -> pd.DataFrame:
    """
    Convierte la grilla editable a un DF normalizado.
    Columnas origen:
      HORA | SVC PROY | SVC REALES | MOV REQ | MOVILES X NOMINA | COEFICIENTE HS | DIF MOVILES
    """
    req = ["HORA","SVC PROY","SVC REALES","MOV REQ","MOVILES X NOMINA"]
    miss = [c for c in req if c not in df_grid.columns]
    if miss:
        raise ValueError(f"Faltan columnas en la grilla: {', '.join(miss)}")

    out = pd.DataFrame()
    out["Hora"] = df_grid["HORA"].map(_to_time)
    out["Servicios_Planificados"] = df_grid["SVC PROY"].map(_to_num)
    out["Servicios_Reales"]       = df_grid["SVC REALES"].map(_to_num)
    out["Moviles_Planificados"]   = df_grid["MOV REQ"].map(_to_num)
    out["Moviles_Reales"]         = df_grid["MOVILES X NOMINA"].map(_to_num)

    if "COEFICIENTE HS" in df_grid.columns:
        out["Coeficiente_HS"] = df_grid["COEFICIENTE HS"].map(_to_num)
    else:
        out["Coeficiente_HS"] = np.nan

    # Dif móviles: si no viene o hay NaN, lo recalculamos
    if "DIF MOVILES" in df_grid.columns:
        out["Dif_Moviles"] = df_grid["DIF MOVILES"].map(_to_num)
    else:
        out["Dif_Moviles"] = np.nan
    calc_dif = out["Moviles_Reales"] - out["Moviles_Planificados"]
    out["Dif_Moviles"] = np.where(out["Dif_Moviles"].notna(), out["Dif_Moviles"], calc_dif)

    # Metadatos de tiempo
    out["Fecha"] = pd.to_datetime(str(fecha)).date()
    out["Base"]  = str(base).strip().upper()

    # Derivados
    out["Fecha_dt"] = pd.to_datetime(out["Fecha"])
    iso = out["Fecha_dt"].dt.isocalendar()
    out["Año"]    = out["Fecha_dt"].dt.year
    out["Mes"]    = out["Fecha_dt"].dt.month
    out["Semana"] = iso.week  # Semana ISO
    out["Dia"]    = out["Fecha_dt"].dt.day
    out["HoraStr"]= pd.to_datetime(out["Hora"].astype(str)).dt.strftime("%H:%M")

    # Métricas sobre Servicios
    out["Dif_Servicios"] = out["Servicios_Reales"] - out["Servicios_Planificados"]
    out["Desvio_Servicios_%"] = np.where(out["Servicios_Planificados"]>0,
                                         out["Dif_Servicios"]/out["Servicios_Planificados"]*100, np.nan)
    out["Desvio_Moviles_%"] = np.where(out["Moviles_Planificados"]>0,
                                       out["Dif_Moviles"]/out["Moviles_Planificados"]*100, np.nan)
    out["Efectividad"] = np.where(out["Servicios_Planificados"]>0,
                                  1 - (out["Dif_Servicios"].abs()/out["Servicios_Planificados"]), np.nan)
    out["APE"]  = np.where(out["Servicios_Planificados"]>0,
                           (out["Servicios_Reales"] - out["Servicios_Planificados"]).abs()/out["Servicios_Planificados"], np.nan)
    out["AE"]   = (out["Servicios_Reales"] - out["Servicios_Planificados"]).abs()
    out["Bias"] = (out["Servicios_Planificados"] - out["Servicios_Reales"])

    # Clasificación y Status (si faltan Plan/Real en alguna hora)
    out["Status"] = np.select(
        [out["Servicios_Planificados"].notna() & out["Servicios_Reales"].isna(),
         out["Servicios_Planificados"].isna() & out["Servicios_Reales"].notna()],
        ["No ejecutado","No planificado"], default="OK"
    )
    out["Clasificacion"] = np.select(
        [out["Status"].eq("No ejecutado"),
         out["Status"].eq("No planificado"),
         out["Dif_Servicios"].fillna(0).eq(0),
         out["Dif_Servicios"].fillna(0)>0,
         out["Dif_Servicios"].fillna(0)<0],
        ["No ejecutado","No planificado","Exacto","Sobre planificado","Bajo planificado"], default="NA"
    )

    # Orden
    out = out.sort_values("HoraStr").reset_index(drop=True)
    return out

def merge_all_bases(bases_dict: dict[str,pd.DataFrame]) -> pd.DataFrame:
    if not bases_dict:
        return pd.DataFrame()
    dfs = []
    for b, df in bases_dict.items():
        if df is not None and not df.empty:
            dfs.append(df.copy())
    return pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()

def agg_error_metrics(df: pd.DataFrame) -> dict:
    d = df.copy()
    mape = d["APE"].mean()*100 if "APE" in d.columns and d["APE"].notna().any() else np.nan
    mae  = d["AE"].mean() if "AE" in d.columns and d["AE"].notna().any() else np.nan
    fbias = (d["Bias"].sum()/d["Servicios_Reales"].sum()*100) if "Bias" in d.columns and d["Servicios_Reales"].sum()!=0 else np.nan
    return {"MAPE_%":mape, "MAE":mae, "ForecastBias_%":fbias}

def apply_filters(df: pd.DataFrame, bases, fecha, semana, mes, horas):
    d = df.copy()
    if bases: d = d[d["Base"].isin(bases)]
    if fecha is not None: d = d[d["Fecha"].eq(pd.to_datetime(fecha).date())]
    if semana and semana>0: d = d[d["Semana"].eq(int(semana))]
    if mes:
        try:
            aa, mm = mes.split("-"); aa=int(aa); mm=int(mm)
            d = d[(d["Año"].eq(aa)) & (d["Mes"].eq(mm))]
        except Exception:
            pass
    if horas: d = d[d["HoraStr"].isin(horas)]
    return d

# ==========================
# Sidebar: Carga / edición por grilla
# ==========================
with st.sidebar:
    st.header("Carga / Edición de Base")
    # Selección de base existente o nueva
    bases_exist = sorted(list(st.session_state["bases"].keys()))
    base_sel = st.selectbox("Base", options=["(nueva)"] + bases_exist, index=0)
    if base_sel == "(nueva)":
        base_name = st.text_input("Nombre de la Base (ej.: PROY_6001)", value="")
    else:
        base_name = base_sel

    fecha_in = st.date_input("Fecha de trabajo", value=None)

    st.caption("Pegá/tipeá en la grilla (podés usar Ctrl+V desde Excel).")
    # Cargar grilla actual o default
    if base_name and base_name in st.session_state["bases"] and not st.session_state["bases"][base_name].empty and fecha_in:
        # si ya existe, muestro una grilla rellena con los valores de esa fecha y base
        df_exist = st.session_state["bases"][base_name]
        df_exist = df_exist[df_exist["Fecha"].eq(pd.to_datetime(fecha_in).date())]
        if not df_exist.empty:
            grid = pd.DataFrame({
                "HORA": df_exist["HoraStr"],
                "SVC PROY": df_exist["Servicios_Planificados"],
                "SVC REALES": df_exist["Servicios_Reales"],
                "MOV REQ": df_exist["Moviles_Planificados"],
                "MOVILES X NOMINA": df_exist["Moviles_Reales"],
                "COEFICIENTE HS": df_exist.get("Coeficiente_HS", np.nan),
                "DIF MOVILES": df_exist.get("Dif_Moviles", np.nan),
            })
        else:
            grid = default_grid()
    else:
        grid = default_grid()

    grid = st.data_editor(
        grid,
        num_rows="dynamic",
        use_container_width=True,
        hide_index=True,
        column_config={
            "HORA": st.column_config.TextColumn("HORA", help="Formato HH:MM (ej.: 07:00)"),
            "SVC PROY": st.column_config.NumberColumn("SVC PROY", min_value=0, step=1),
            "SVC REALES": st.column_config.NumberColumn("SVC REALES", min_value=0, step=1),
            "MOV REQ": st.column_config.NumberColumn("MOV REQ", min_value=0, step=1),
            "MOVILES X NOMINA": st.column_config.NumberColumn("MOVILES X NOMINA", min_value=0, step=1),
            "COEFICIENTE HS": st.column_config.NumberColumn("COEFICIENTE HS", min_value=0.0, step=0.01, format="%.2f"),
            "DIF MOVILES": st.column_config.NumberColumn("DIF MOVILES", step=1),
        },
        key="grid_editor",
        height=560
    )

    c_g1, c_g2 = st.columns(2)
    with c_g1:
        if st.button("💾 Guardar Base"):
            if not base_name:
                st.error("Ingresá el nombre de la Base.")
            elif not fecha_in:
                st.error("Elegí la fecha de trabajo.")
            else:
                try:
                    norm = normalize_grid(grid, fecha=str(fecha_in), base=base_name)
                    # agrego/actualizo en el dict (si ya existía para esa fecha, reemplazo esas filas)
                    df_prev = st.session_state["bases"].get(base_name, pd.DataFrame())
                    if not df_prev.empty:
                        df_prev = df_prev[~df_prev["Fecha"].eq(pd.to_datetime(fecha_in).date())]  # borro mismo día
                        norm = pd.concat([df_prev, norm], ignore_index=True)
                    st.session_state["bases"][base_name] = norm
                    # guardo por base
                    save_csv(norm, BASES_DIR / f"{base_name}.csv")
                    st.success(f"Base '{base_name}' guardada ({len(norm)} filas).")
                except Exception as e:
                    st.error(f"No se pudo guardar: {e}")

    with c_g2:
        if st.button("🗑️ Borrar Base del día"):
            if base_name and fecha_in and base_name in st.session_state["bases"]:
                df_prev = st.session_state["bases"][base_name]
                df_prev = df_prev[~df_prev["Fecha"].eq(pd.to_datetime(fecha_in).date())]
                st.session_state["bases"][base_name] = df_prev
                save_csv(df_prev, BASES_DIR / f"{base_name}.csv")
                st.warning(f"Eliminado el día {fecha_in} de la base {base_name}.")

    st.markdown("---")
    c_p1, c_p2 = st.columns(2)
    with c_p1:
        if st.button("💾 Guardar MERGED (/data)"):
            merged = merge_all_bases(st.session_state["bases"])
            st.session_state["merged"] = merged
            if not merged.empty:
                save_csv(merged, MERGED_CSV)
                st.success(f"MERGED guardado: {len(merged):,} filas.")
            else:
                st.info("No hay datos para guardar.")
    with c_p2:
        if st.button("🧹 Limpiar memoria"):
            st.session_state["bases"] = {}
            st.session_state["merged"] = pd.DataFrame()
            st.success("Memoria limpiada (no borra /data).")

# ==========================
# Contenido principal (tabs arriba)
# ==========================
st.title("Comparación — Planificación vs Realidad (carga manual)")
tabs = st.tabs(["Dashboard", "Análisis por Base", "Análisis Horario", "Auditoría Detallada"])

# Dataset unificado
merged = merge_all_bases(st.session_state["bases"])
st.session_state["merged"] = merged  # mantener sincronizado

# Barra de filtros en la parte superior (para todas las pestañas)
flt = st.container()
with flt:
    c1,c2,c3,c4,c5 = st.columns([1.3,1,1,1.2,1.6])
    with c1:
        bases_all = sorted(merged["Base"].unique().tolist()) if not merged.empty else []
        bases_fil = st.multiselect("Bases", options=bases_all, default=bases_all)
    with c2:
        fecha_fil = st.date_input("Día", value=None)
    with c3:
        semana_fil = st.number_input("Semana ISO", value=0, step=1, min_value=0)
    with c4:
        mes_fil = st.text_input("Mes (aaaa-mm)", value="")
    with c5:
        horas_all = sorted(merged["HoraStr"].unique().tolist()) if not merged.empty else []
        horas_fil = st.multiselect("Horas (HH:MM)", options=horas_all, default=horas_all)

df_f = apply_filters(merged, bases_fil, fecha_fil, semana_fil, mes_fil, horas_fil)

# ==========================
# TAB 1 — Dashboard
# ==========================
with tabs[0]:
    if df_f.empty:
        st.info("Cargá datos en la barra lateral y/o ajustá filtros.")
        st.stop()

    st.subheader("KPIs globales")
    tot_plan_m = df_f["Moviles_Planificados"].sum()
    tot_real_m = df_f["Moviles_Reales"].sum()
    tot_plan_s = df_f["Servicios_Planificados"].sum()
    tot_real_s = df_f["Servicios_Reales"].sum()

    desvio_m = (tot_real_m - tot_plan_m) / tot_plan_m * 100 if tot_plan_m>0 else np.nan
    desvio_s = (tot_real_s - tot_plan_s) / tot_plan_s * 100 if tot_plan_s>0 else np.nan
    efect    = 1 - (abs(tot_real_s - tot_plan_s) / tot_plan_s) if tot_plan_s>0 else np.nan

    m1,m2,m3 = st.columns(3)
    m1.metric("Móviles — % Desvío", f"{desvio_m:,.1f}%" if pd.notna(desvio_m) else "—")
    m2.metric("Servicios — % Desvío", f"{desvio_s:,.1f}%" if pd.notna(desvio_s) else "—")
    m3.metric("Efectividad", f"{efect:.1%}" if pd.notna(efect) else "—")

    # Semáforo por efectividad
    if pd.isna(efect): color, txt = ("#6B7280", "Sin datos")
    elif efect >= 0.92: color, txt = ("#059669", "OK (≥ 92%)")
    elif efect >= 0.89: color, txt = ("#F59E0B", "Atención (89–92%)")
    else:                color, txt = ("#DC2626", "Crítico (< 89%)")
    st.markdown(f"**Estado general:** <span style='color:{color}'>{txt}</span>", unsafe_allow_html=True)

    # Gráficos
    g = df_f.groupby(["Fecha","HoraStr"], as_index=False)[["Servicios_Planificados","Servicios_Reales"]].sum()
    fig1 = px.line(g, x="HoraStr", y=["Servicios_Planificados","Servicios_Reales"],
                   color_discrete_sequence=["#22D3EE","#10B981"])
    stylize(fig1, "Plan vs Real (Servicios por hora)")
    st.plotly_chart(fig1, use_container_width=True)

    g2 = df_f.groupby("HoraStr", as_index=False)["Dif_Servicios"].sum()
    g2p = df_f.groupby("HoraStr", as_index=False)["Servicios_Planificados"].sum()
    g2 = g2.merge(g2p, on="HoraStr", how="left")
    g2["Desvio_%"] = np.where(g2["Servicios_Planificados"]>0, g2["Dif_Servicios"]/g2["Servicios_Planificados"]*100, np.nan)
    fig2 = px.bar(g2, x="HoraStr", y="Desvio_%", color="Desvio_%", color_continuous_scale="RdYlGn")
    stylize(fig2, "Desvío % por hora (Servicios)")
    st.plotly_chart(fig2, use_container_width=True)

    piv = df_f.pivot_table(values="Dif_Servicios", index="Fecha", columns="HoraStr", aggfunc="sum").fillna(0)
    if not piv.empty:
        fig3 = px.imshow(piv, color_continuous_scale="RdYlGn", aspect="auto")
        stylize(fig3, "Heatmap — Desvío de servicios (Real - Plan)")
        st.plotly_chart(fig3, use_container_width=True)

    mets = agg_error_metrics(df_f)
    st.markdown(f"**MAPE:** {mets['MAPE_%']:.1f}% · **MAE:** {mets['MAE']:.2f} · **Forecast Bias:** {mets['ForecastBias_%']:.1f}%")

    # Detección automática
    g3 = df_f.groupby("HoraStr", as_index=False)["Dif_Servicios"].sum()
    sub = g3.nsmallest(5, "Dif_Servicios")
    sobre = g3.nlargest(5, "Dif_Servicios")
    wb = df_f.groupby("Base", as_index=False)["Dif_Servicios"].apply(lambda s: s.abs().sum()).rename(columns={"Dif_Servicios":"AbsDesvio"}) \
             .sort_values("AbsDesvio", ascending=False).head(1)

    c1,c2,c3 = st.columns(3)
    with c1:
        st.subheader("Top 5 Sub‑plan (horas)")
        st.dataframe(sub, use_container_width=True, hide_index=True)
    with c2:
        st.subheader("Top 5 Sobre‑plan (horas)")
        st.dataframe(sobre, use_container_width=True, hide_index=True)
    with c3:
        st.subheader("Base con mayor desvío")
        st.dataframe(wb, use_container_width=True, hide_index=True)

# ==========================
# TAB 2 — Análisis por Base
# ==========================
with tabs[1]:
    if df_f.empty:
        st.info("No hay datos para los filtros seleccionados.")
    else:
        st.subheader("Desvío por Base (Servicios)")
        g = df_f.groupby("Base", as_index=False)[["Servicios_Planificados","Servicios_Reales"]].sum()
        g["Desvio_%"] = np.where(g["Servicios_Planificados"]>0,
                                 (g["Servicios_Reales"]-g["Servicios_Planificados"])/g["Servicios_Planificados"]*100, np.nan)
        fig = px.bar(g, x="Base", y="Desvio_%", color="Desvio_%", color_continuous_scale="RdYlGn")
        stylize(fig, "Desvío % por Base")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(g, use_container_width=True, hide_index=True)

# ==========================
# TAB 3 — Análisis Horario
# ==========================
with tabs[2]:
    if df_f.empty:
        st.info("No hay datos para los filtros seleccionados.")
    else:
        st.subheader("Series por hora — Plan vs Real (Servicios)")
        g = df_f.groupby("HoraStr", as_index=False)[["Servicios_Planificados","Servicios_Reales"]].sum()
        fig = px.line(g, x="HoraStr", y=["Servicios_Planificados","Servicios_Reales"],
                      color_discrete_sequence=["#22D3EE","#10B981"])
        stylize(fig, "Plan vs Real por hora")
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("Desvío (Servicios)")
        g2 = df_f.groupby("HoraStr", as_index=False)["Dif_Servicios"].sum()
        fig2 = px.bar(g2, x="HoraStr", y="Dif_Servicios", color="Dif_Servicios", color_continuous_scale="RdYlGn")
        stylize(fig2, "Desvío (Real - Plan) por hora")
        st.plotly_chart(fig2, use_container_width=True)

        st.dataframe(
            df_f[["Fecha","HoraStr","Base","Servicios_Planificados","Servicios_Reales",
                  "Dif_Servicios","Desvio_Servicios_%","Clasificacion"]].sort_values(["Fecha","HoraStr","Base"]),
            use_container_width=True, hide_index=True
        )

# ==========================
# TAB 4 — Auditoría Detallada
# ==========================
with tabs[3]:
    if df_f.empty:
        st.info("No hay datos para los filtros seleccionados.")
    else:
        st.subheader("Tabla completa con clasificación")
        cols = ["Fecha","HoraStr","Base",
                "Moviles_Planificados","Moviles_Reales","Dif_Moviles","Desvio_Moviles_%",
                "Servicios_Planificados","Servicios_Reales","Dif_Servicios","Desvio_Servicios_%",
                "Efectividad","Clasificacion","Status","Semana","Mes","Año","Coeficiente_HS"]
        cols = [c for c in cols if c in df_f.columns]
        df_aud = df_f[cols].sort_values(["Fecha","HoraStr","Base"])
        st.dataframe(df_aud, use_container_width=True, hide_index=True)

        bytes_xls, fname = to_excel_bytes(df_aud, sheet_name="auditoria", fname="auditoria_plan_vs_real.xlsx")
        st.download_button("⬇️ Descargar Excel (auditoría)", data=bytes_xls, file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
