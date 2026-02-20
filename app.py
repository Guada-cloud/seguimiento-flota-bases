# app.py — Pegado directo (dos datasets: Planificación y Realidad) · Multi-Base · Dashboard completo
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from io import StringIO, BytesIO
from pathlib import Path

# =====================================
# Apariencia
# =====================================
st.set_page_config(page_title="Plan vs Real — Operación (Pegado directo)", layout="wide")
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

# =====================================
# Persistencia
# =====================================
DATA_DIR = Path("data"); DATA_DIR.mkdir(exist_ok=True)
BASES_DIR = DATA_DIR / "bases"; BASES_DIR.mkdir(exist_ok=True)
MERGED_CSV = DATA_DIR / "merged.csv"

def save_csv(df: pd.DataFrame, path: Path):
    df.to_csv(path, index=False, encoding="utf-8")

def load_csv(path: Path) -> pd.DataFrame|None:
    return pd.read_csv(path, encoding="utf-8") if path.exists() else None

def to_excel_bytes(df: pd.DataFrame, sheet="datos", name="reporte.xlsx"):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as wr:
        df.to_excel(wr, index=False, sheet_name=sheet)
    return buf.getvalue(), name

# =====================================
# Estado
# =====================================
if "bases" not in st.session_state:     # { base: DataFrame con múltiples fechas }
    st.session_state["bases"] = {}
if "merged" not in st.session_state:
    st.session_state["merged"] = pd.DataFrame()
if "_preview" not in st.session_state:  # vista previa del último pegado fusionado (plan+real) para la base/fecha seleccionada
    st.session_state["_preview"] = pd.DataFrame()

# =====================================
# Mapeo flexible de columnas
# (tolerante a acentos, mayúsculas/minúsculas y variaciones)
# =====================================
SYN = {
    "hora":      ["hora","hr","tiempo","h"],
    "svc_plan":  ["svc proy","servicios proy","svc plan","serv plan","serv proyectados","servicios proyectados","proyectado","planificado"],
    "svc_real":  ["svc reales","servicios reales","svc real","serv real","observado","observados","reales"],
    "mov_plan":  ["mov req","mov requeridos","mov plan","moviles plan","moviles requeridos","req moviles"],
    "mov_real":  ["moviles x nomina","mov x nomina","mov reales","mov real","nomina","nómina","moviles nómina"],
    "coef_hs":   ["coeficiente hs","coef hs","coef hs.","coeficiente horas"],
    "dif_mov":   ["dif moviles","dif mov","delta moviles","delta mov.","moviles delta"]
}

def _norm(s: str) -> str:
    s = str(s).strip().lower()
    rep = {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ñ":"n"}
    for a,b in rep.items(): s = s.replace(a,b)
    return " ".join(s.split())

def _find_col(cols, aliases):
    m = { _norm(c): c for c in cols }
    # exacto
    for a in aliases:
        if a in m: return m[a]
    # contiene
    for a in aliases:
        for k,v in m.items():
            if a in k: return v
    return None

# Conversión de números robusta (coma o punto decimal, miles, "#¿NOMBRE?")
def _to_num_series(s: pd.Series) -> pd.Series:
    s = s.astype(str).str.strip()
    s = s.replace({"": np.nan, "nan": np.nan, "None": np.nan, "#¿NOMBRE?": np.nan, "#¡NOMBRE?": np.nan}, regex=False)
    # Detecta decimal/coma: si hay coma y punto, usa la última como decimal y remueve el otro como miles
    def _fix_one(x:str):
        if x is np.nan or x is None: return np.nan
        txt = str(x)
        if txt.count(",") and txt.count("."):
            # si la última coma está después del último punto -> coma decimal
            if txt.rfind(",") > txt.rfind("."):
                txt = txt.replace(".", "").replace(",", ".")
            else:  # punto decimal
                txt = txt.replace(",", "")
        else:
            if "," in txt and "." not in txt:
                txt = txt.replace(",", ".")
            # si solo hay punto, lo toma como decimal directamente
        try:
            return float(txt)
        except Exception:
            return np.nan
    return s.map(_fix_one)

def _to_time_series(s: pd.Series) -> pd.Series:
    # Acepta "HH:MM" o "H:MM" o excel-like
    return pd.to_datetime(s.astype(str), errors="coerce").dt.time

# =====================================
# Parsers (Plan y Real) a partir de pegado
# =====================================
def detect_sep(text: str) -> str:
    sample = text[:1000]
    if "\t" in sample: return "\t"
    if ";" in sample:  return ";"
    return ","  # fallback

def parse_pasted_generic(text: str) -> pd.DataFrame:
    """
    Lee un pegado libre que puede contener Plan y/o Real en la misma tabla.
    Devuelve un DF con columnas estandarizadas: Hora, SvcPlan, SvcReal, MovPlan, MovReal, CoefHS, DifMov
    (lo que no esté en el pegado queda como NaN).
    """
    if not text or not text.strip():
        return pd.DataFrame()
    sep = detect_sep(text)
    df = pd.read_csv(StringIO(text), sep=sep, engine="python", dtype=str)

    hora_c  = _find_col(df.columns, SYN["hora"])
    sp_c    = _find_col(df.columns, SYN["svc_plan"])
    sr_c    = _find_col(df.columns, SYN["svc_real"])
    mp_c    = _find_col(df.columns, SYN["mov_plan"])
    mr_c    = _find_col(df.columns, SYN["mov_real"])
    coef_c  = _find_col(df.columns, SYN["coef_hs"])
    difm_c  = _find_col(df.columns, SYN["dif_mov"])

    if hora_c is None:
        raise ValueError("No se encontró columna de HORA en el pegado.")

    out = pd.DataFrame()
    out["Hora"]    = _to_time_series(df[hora_c])
    out["SvcPlan"] = _to_num_series(df[sp_c]) if sp_c else np.nan
    out["SvcReal"] = _to_num_series(df[sr_c]) if sr_c else np.nan
    out["MovPlan"] = _to_num_series(df[mp_c]) if mp_c else np.nan
    out["MovReal"] = _to_num_series(df[mr_c]) if mr_c else np.nan
    out["CoefHS"]  = _to_num_series(df[coef_c]) if coef_c else np.nan
    # Dif móviles: si viene, lo tomo; si no, lo calcularé luego
    out["DifMov_archivo"] = _to_num_series(df[difm_c]) if difm_c else np.nan

    out = out[out["Hora"].notna()]
    out["HoraStr"] = pd.to_datetime(out["Hora"].astype(str)).dt.strftime("%H:%M")
    return out.reset_index(drop=True)

def merge_plan_real_from_pastes(plan_df: pd.DataFrame, real_df: pd.DataFrame) -> pd.DataFrame:
    """
    Fusiona dos DF parciales (al menos Hora + {SvcPlan/MovPlan} y Hora + {SvcReal/MovReal}).
    Si alguno viene con las dos mitades, también se respeta (se usa el valor no nulo).
    """
    # Unificar por Hora
    key = ["HoraStr"]
    left  = plan_df[["Hora","HoraStr","SvcPlan","MovPlan","CoefHS","DifMov_archivo"]].copy()
    right = real_df[["HoraStr","SvcReal","MovReal"]].copy()

    m = pd.merge(left, right, on="HoraStr", how="outer")
    # Completar Hora si quedó vacía del lado real
    m["Hora"] = m["Hora"].fillna(pd.to_datetime(m["HoraStr"]).dt.time)

    # Recalcular DifMov si falta
    dif_calc = m["MovReal"] - m["MovPlan"]
    m["DifMov"] = np.where(m["DifMov_archivo"].notna(), m["DifMov_archivo"], dif_calc)

    # Renombrar a columnas estándar finales
    out = pd.DataFrame()
    out["Hora"]  = m["Hora"]
    out["HoraStr"] = m["HoraStr"]
    out["Servicios_Planificados"] = m["SvcPlan"]
    out["Servicios_Reales"]       = m["SvcReal"]
    out["Moviles_Planificados"]   = m["MovPlan"]
    out["Moviles_Reales"]         = m["MovReal"]
    out["Coeficiente_HS"]         = m["CoefHS"]
    out["Dif_Moviles"]            = m["DifMov"]
    return out.sort_values("HoraStr").reset_index(drop=True)

def enrich_with_time_and_metrics(df: pd.DataFrame, fecha, base) -> pd.DataFrame:
    out = df.copy()
    out["Fecha"] = pd.to_datetime(str(fecha)).date()
    out["Base"]  = str(base).strip().upper()

    out["Fecha_dt"] = pd.to_datetime(out["Fecha"])
    iso = out["Fecha_dt"].dt.isocalendar()
    out["Año"] = out["Fecha_dt"].dt.year
    out["Mes"] = out["Fecha_dt"].dt.month
    out["Semana"] = iso.week
    out["Dia"] = out["Fecha_dt"].dt.day

    # Métricas servicios y móviles
    out["Dif_Servicios"] = out["Servicios_Reales"] - out["Servicios_Planificados"]
    out["Desvio_Servicios_%"] = np.where(out["Servicios_Planificados"]>0,
                                         out["Dif_Servicios"]/out["Servicios_Planificados"]*100, np.nan)
    out["Desvio_Moviles_%"] = np.where(out["Moviles_Planificados"]>0,
                                       out["Dif_Moviles"]/out["Moviles_Planificados"]*100, np.nan)
    out["Efectividad"] = np.where(out["Servicios_Planificados"]>0,
                                  1 - (out["Dif_Servicios"].abs()/out["Servicios_Planificados"]), np.nan)
    out["APE"] = np.where(out["Servicios_Planificados"]>0,
                          (out["Servicios_Reales"] - out["Servicios_Planificados"]).abs()/out["Servicios_Planificados"], np.nan)
    out["AE"]  = (out["Servicios_Reales"] - out["Servicios_Planificados"]).abs()
    out["Bias"]= (out["Servicios_Planificados"] - out["Servicios_Reales"])

    # Estado/Clasificación
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
    return out

def agg_error_metrics(df: pd.DataFrame) -> dict:
    d = df.copy()
    mape = d["APE"].mean()*100 if "APE" in d and d["APE"].notna().any() else np.nan
    mae  = d["AE"].mean() if "AE" in d and d["AE"].notna().any() else np.nan
    fbias = (d["Bias"].sum()/d["Servicios_Reales"].sum()*100) if "Bias" in d and d["Servicios_Reales"].sum()!=0 else np.nan
    return {"MAPE_%":mape, "MAE":mae, "ForecastBias_%":fbias}

# =====================================
# Sidebar — Ingreso y guardado
# =====================================
with st.sidebar:
    st.header("Pegar Planificación y Realidad (sin Excel)")

    bases_exist = sorted(st.session_state["bases"].keys())
    base_sel = st.selectbox("Base", options=["(nueva)"] + bases_exist, index=0)
    base_name = st.text_input("Nombre de Base", value="" if base_sel=="(nueva)" else base_sel, help="Ej.: PROY_6001")
    fecha_in = st.date_input("Fecha", value=None)

    st.caption("Pegá la tabla copiada desde Excel. Acepta TAB/;/,; decimales con coma o punto; y variaciones de encabezados.")
    with st.expander("Planificación (contiene SVC PROY y/o MOV REQ)", expanded=True):
        txt_plan = st.text_area("Pegar Planificación", height=180, key="paste_plan")
    with st.expander("Realidad (contiene SVC REALES y/o MOVILES X NOMINA)", expanded=True):
        txt_real = st.text_area("Pegar Realidad", height=180, key="paste_real")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("🔎 Previsualizar (fusionar Plan + Real)"):
            if not base_name:
                st.error("Ingresá nombre de Base."); st.stop()
            if not fecha_in:
                st.error("Elegí la Fecha."); st.stop()

            try:
                # Parseos independientes (soporta que uno solo ya traiga todo)
                df_p = parse_pasted_generic(txt_plan) if txt_plan.strip() else pd.DataFrame()
                df_r = parse_pasted_generic(txt_real) if txt_real.strip() else pd.DataFrame()
                # Si alguno viene vacío, uso el otro; si ambos tienen info, hago merge
                if df_p.empty and df_r.empty:
                    st.error("No hay datos en Plan ni en Real."); st.stop()
                if df_p.empty:
                    df_p = df_r.copy()
                if df_r.empty:
                    df_r = df_p.copy()

                merged_hr = merge_plan_real_from_pastes(df_p, df_r)
                prev = enrich_with_time_and_metrics(merged_hr, fecha_in, base_name)
                st.session_state["_preview"] = prev
                st.success(f"Previsualización OK — filas: {len(prev)}")
                st.dataframe(prev.head(24), use_container_width=True)
            except Exception as e:
                st.error(f"No se pudo leer/fusionar: {e}")

    with c2:
        if st.button("💾 Guardar Base (día)"):
            if st.session_state["_preview"].empty:
                st.info("Primero presioná 'Previsualizar'.")
            else:
                df_prev = st.session_state["bases"].get(base_name, pd.DataFrame())
                # eliminar ese día para esa base (si existía) y agregar lo nuevo
                if not df_prev.empty:
                    df_prev = df_prev[~df_prev["Fecha"].eq(pd.to_datetime(fecha_in).date())]
                    df_new = pd.concat([df_prev, st.session_state["_preview"]], ignore_index=True)
                else:
                    df_new = st.session_state["_preview"].copy()
                st.session_state["bases"][base_name] = df_new
                save_csv(df_new, BASES_DIR / f"{base_name}.csv")
                st.success(f"Base '{base_name}' guardada ({len(df_new)} filas totales).")

    st.markdown("---")
    c3, c4 = st.columns(2)
    with c3:
        if st.button("💾 Guardar MERGED (/data)"):
            # unir todas las bases guardadas
            dfs = [df.copy() for df in st.session_state["bases"].values() if not df.empty]
            merged = pd.concat(dfs, ignore_index=True) if dfs else pd.DataFrame()
            st.session_state["merged"] = merged
            if not merged.empty:
                save_csv(merged, MERGED_CSV)
                st.success(f"MERGED guardado: {len(merged):,} filas.")
            else:
                st.info("No hay datos para guardar.")
    with c4:
        if st.button("🧹 Limpiar memoria"):
            st.session_state["bases"] = {}
            st.session_state["merged"] = pd.DataFrame()
            st.session_state["_preview"] = pd.DataFrame()
            st.success("Memoria limpiada (no borra /data).")

# =====================================
# Contenido principal — Tabs + Filtros
# =====================================
st.title("Comparación — Planificación vs Realidad (pegado directo)")
tabs = st.tabs(["Dashboard", "Análisis por Base", "Análisis Horario", "Auditoría Detallada", "Configuración"])

# Dataset unificado (en vivo)
dfs_live = [df.copy() for df in st.session_state["bases"].values() if not df.empty]
merged_live = pd.concat(dfs_live, ignore_index=True) if dfs_live else pd.DataFrame()
st.session_state["merged"] = merged_live

# Filtros globales
flt = st.container()
with flt:
    c1,c2,c3,c4,c5 = st.columns([1.3,1,1,1.2,1.6])
    with c1:
        bases_all = sorted(merged_live["Base"].unique().tolist()) if not merged_live.empty else []
        bases_fil = st.multiselect("Bases", options=bases_all, default=bases_all)
    with c2:
        fecha_fil = st.date_input("Día", value=None, key="dia_filter")
    with c3:
        semana_fil = st.number_input("Semana ISO", value=0, step=1, min_value=0)
    with c4:
        mes_fil = st.text_input("Mes (aaaa-mm)", value="")
    with c5:
        horas_all = sorted(merged_live["HoraStr"].unique().tolist()) if not merged_live.empty else []
        horas_fil = st.multiselect("Horas (HH:MM)", options=horas_all, default=horas_all)

def apply_filters(df: pd.DataFrame) -> pd.DataFrame:
    d = df.copy()
    if bases_fil: d = d[d["Base"].isin(bases_fil)]
    if fecha_fil is not None: d = d[d["Fecha"].eq(pd.to_datetime(fecha_fil).date())]
    if semana_fil and semana_fil>0: d = d[d["Semana"].eq(int(semana_fil))]
    if mes_fil:
        try:
            aa, mm = mes_fil.split("-"); aa=int(aa); mm=int(mm)
            d = d[(d["Año"].eq(aa)) & (d["Mes"].eq(mm))]
        except Exception:
            pass
    if horas_fil: d = d[d["HoraStr"].isin(horas_fil)]
    return d

# =====================================
# TAB 1 — Dashboard
# =====================================
with tabs[0]:
    df_f = apply_filters(merged_live)
    if df_f.empty:
        st.info("Pegá y guardá datos en el lateral, y/o ajustá filtros.")
    else:
        st.subheader("KPIs globales")
        tot_plan_m = df_f["Moviles_Planificados"].sum()
        tot_real_m = df_f["Moviles_Reales"].sum()
        tot_plan_s = df_f["Servicios_Planificados"].sum()
        tot_real_s = df_f["Servicios_Reales"].sum()

        desvio_m = (tot_real_m - tot_plan_m)/tot_plan_m*100 if tot_plan_m>0 else np.nan
        desvio_s = (tot_real_s - tot_plan_s)/tot_plan_s*100 if tot_plan_s>0 else np.nan
        efect    = 1 - (abs(tot_real_s - tot_plan_s)/tot_plan_s) if tot_plan_s>0 else np.nan

        m1,m2,m3 = st.columns(3)
        m1.metric("Móviles — % Desvío", f"{desvio_m:,.1f}%" if pd.notna(desvio_m) else "—")
        m2.metric("Servicios — % Desvío", f"{desvio_s:,.1f}%" if pd.notna(desvio_s) else "—")
        m3.metric("Efectividad", f"{efect:.1%}" if pd.notna(efect) else "—")

        # Semáforo
        if pd.isna(efect): color, txt = ("#6B7280", "Sin datos")
        elif efect >= 0.92: color, txt = ("#059669", "OK (≥ 92%)")
        elif efect >= 0.89: color, txt = ("#F59E0B", "Atención (89–92%)")
        else: color, txt = ("#DC2626", "Crítico (< 89%)")
        st.markdown(f"**Estado general:** <span style='color:{color}'>{txt}</span>", unsafe_allow_html=True)

        # Serie Plan vs Real (Servicios)
        g = df_f.groupby(["Fecha","HoraStr"], as_index=False)[["Servicios_Planificados","Servicios_Reales"]].sum()
        fig1 = px.line(g, x="HoraStr", y=["Servicios_Planificados","Servicios_Reales"],
                       color_discrete_sequence=["#22D3EE","#10B981"])
        stylize(fig1, "Plan vs Real (Servicios por hora)")
        st.plotly_chart(fig1, use_container_width=True)

        # Barras Desvío %
        g2 = df_f.groupby("HoraStr", as_index=False)["Dif_Servicios"].sum()
        g2p = df_f.groupby("HoraStr", as_index=False)["Servicios_Planificados"].sum()
        g2 = g2.merge(g2p, on="HoraStr", how="left")
        g2["Desvio_%"] = np.where(g2["Servicios_Planificados"]>0, g2["Dif_Servicios"]/g2["Servicios_Planificados"]*100, np.nan)
        fig2 = px.bar(g2, x="HoraStr", y="Desvio_%", color="Desvio_%", color_continuous_scale="RdYlGn")
        stylize(fig2, "Desvío % por hora (Servicios)")
        st.plotly_chart(fig2, use_container_width=True)

        # Heatmap
        piv = df_f.pivot_table(values="Dif_Servicios", index="Fecha", columns="HoraStr", aggfunc="sum").fillna(0)
        if not piv.empty:
            fig3 = px.imshow(piv, color_continuous_scale="RdYlGn", aspect="auto")
            stylize(fig3, "Heatmap — Desvío de servicios (Real - Plan)")
            st.plotly_chart(fig3, use_container_width=True)

        # Errores agregados
        m = {
            "MAPE_%": (df_f["APE"].mean()*100 if df_f["APE"].notna().any() else np.nan),
            "MAE":    (df_f["AE"].mean() if df_f["AE"].notna().any() else np.nan),
            "ForecastBias_%": ((df_f["Bias"].sum()/df_f["Servicios_Reales"].sum()*100) if df_f["Servicios_Reales"].sum()!=0 else np.nan)
        }
        st.markdown(f"**MAPE:** {m['MAPE_%']:.1f}% · **MAE:** {m['MAE']:.2f} · **Forecast Bias:** {m['ForecastBias_%']:.1f}%")

        # Detección
        g3 = df_f.groupby("HoraStr", as_index=False)["Dif_Servicios"].sum()
        sub = g3.nsmallest(5, "Dif_Servicios")
        sobre = g3.nlargest(5, "Dif_Servicios")
        wb = df_f.groupby("Base", as_index=False)["Dif_Servicios"].apply(lambda s: s.abs().sum()).rename(columns={"Dif_Servicios":"AbsDesvio"}) \
                 .sort_values("AbsDesvio", ascending=False).head(1)
        c1,c2,c3 = st.columns(3)
        with c1:
            st.subheader("Top 5 Sub‑plan (horas)"); st.dataframe(sub, use_container_width=True, hide_index=True)
        with c2:
            st.subheader("Top 5 Sobre‑plan (horas)"); st.dataframe(sobre, use_container_width=True, hide_index=True)
        with c3:
            st.subheader("Base con mayor desvío");   st.dataframe(wb, use_container_width=True, hide_index=True)

# =====================================
# TAB 2 — Análisis por Base
# =====================================
with tabs[1]:
    df_f = apply_filters(merged_live)
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

# =====================================
# TAB 3 — Análisis Horario
# =====================================
with tabs[2]:
    df_f = apply_filters(merged_live)
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

# =====================================
# TAB 4 — Auditoría Detallada (con descarga)
# =====================================
with tabs[3]:
    df_f = apply_filters(merged_live)
    if df_f.empty:
        st.info("No hay datos para los filtros seleccionados.")
    else:
        st.subheader("Auditoría (lo que ves)")
        cols = ["Fecha","HoraStr","Base",
                "Moviles_Planificados","Moviles_Reales","Dif_Moviles","Desvio_Moviles_%",
                "Servicios_Planificados","Servicios_Reales","Dif_Servicios","Desvio_Servicios_%",
                "Efectividad","Clasificacion","Status","Semana","Mes","Año","Coeficiente_HS"]
        cols = [c for c in cols if c in df_f.columns]
        df_aud = df_f[cols].sort_values(["Fecha","HoraStr","Base"])
        st.dataframe(df_aud, use_container_width=True, hide_index=True)

        xls, name = to_excel_bytes(df_aud, sheet="auditoria", name="auditoria_plan_vs_real.xlsx")
        st.download_button("⬇️ Descargar Excel (auditoría)", data=xls, file_name=name,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# =====================================
# TAB 5 — Configuración (opcional, para ver/recuperar guardados)
# =====================================
with tabs[4]:
    st.subheader("Configuración y persistencia")
    st.write("Directorios:", str(DATA_DIR), " / ", str(BASES_DIR))
    c1, c2 = st.columns(2)
    with c1:
        if st.button("📥 Cargar MERGED desde /data"):
            m = load_csv(MERGED_CSV)
            if m is not None:
                st.session_state["merged"] = m
                st.success(f"MERGED cargado: {len(m):,} filas.")
            else:
                st.info("No existe /data/merged.csv todavía.")
    with c2:
        if st.button("📥 Cargar todas las bases desde /data/bases"):
            loaded = 0
            for p in BASES_DIR.glob("*.csv"):
                try:
                    dfb = load_csv(p)
                    if dfb is not None:
                        st.session_state["bases"][p.stem] = dfb
                        loaded += 1
                except Exception:
                    pass
            st.success(f"Bases cargadas: {loaded}")
