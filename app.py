# app.py — Planificación vs Realidad Operativa 
import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
from pathlib import Path
from utils_ops import (
    EXPECTED_PLAN, EXPECTED_REAL, _guess_map, apply_map, enrich_time,
    merge_plan_real, compute_metrics, agg_error_metrics, add_time_keys,
    filter_df, top5_hours, worst_base,
    save_csv, load_csv, PLAN_CSV, REAL_CSV, MERG_CSV, to_excel_bytes
)

# ========= Config UI general =========
st.set_page_config(page_title="Plan vs Real — Operación", layout="wide")
TEMPLATE = "plotly_dark"
FONT = "Inter, system-ui, Segoe UI, Roboto"

def stylize(fig, title=None, y_pct=False):
    fig.update_layout(
        template=TEMPLATE,
        title=title,
        font=dict(family=FONT, size=13, color="#E5E7EB"),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        legend_title_text="", margin=dict(t=45, r=10, b=30, l=10),
    )
    if y_pct:
        fig.update_yaxes(tickformat=".0%")
    fig.update_xaxes(showgrid=False)
    fig.update_yaxes(gridcolor="rgba(148,163,184,.25)")
    return fig

# ========= Estado =========
if "plan_df" not in st.session_state: st.session_state["plan_df"] = None
if "real_df" not in st.session_state: st.session_state["real_df"] = None
if "merged"  not in st.session_state: st.session_state["merged"]  = None

# ========= Sidebar (menú y filtros) =========
with st.sidebar:
    st.header("Menú")
    page = st.radio("Navegación", ["Dashboard", "Análisis por Base", "Análisis Horario", "Auditoría Detallada", "Configuración"], index=0)

    st.markdown("---")
    st.header("Filtros")
    # Estos filtros se aplican a la tabla merged cuando exista
    fecha_sel = st.date_input("Día", value=None)
    semana_sel = st.number_input("Semana ISO", value=0, step=1, min_value=0)
    mes_text = st.text_input("Mes (aaaa-mm)", value="")

    base_sel = st.text_input("Base (dejar vacío = todas)", value="")
    horas_multi = st.text_input("Horas (HH:MM, separadas por coma)", value="")

    st.markdown("---")
    st.header("Persistencia")
    c1,c2 = st.columns(2)
    with c1:
        if st.button("💾 Guardar CSVs"):
            if st.session_state["plan_df"] is not None:
                save_csv(st.session_state["plan_df"], PLAN_CSV)
            if st.session_state["real_df"] is not None:
                save_csv(st.session_state["real_df"], REAL_CSV)
            if st.session_state["merged"] is not None:
                save_csv(st.session_state["merged"],  MERG_CSV)
            st.success("Datos guardados en /data.")
    with c2:
        if st.button("🧹 Limpiar memoria"):
            for k in ["plan_df","real_df","merged"]:
                st.session_state[k] = None
            st.success("Memoria limpiada (la carpeta /data no se toca).")

# ========= Encabezado =========
st.title("Comparación — Planificación vs Realidad Operativa")
st.caption("Móviles y Servicios · Filtros por Base, Hora, Día, Semana ISO y Mes · Persistencia en /data")

# ========= Utilidad de filtros a lista
def _hour_list(text: str) -> list[str]|None:
    t = [s.strip() for s in text.split(",") if s.strip()] if text else []
    return t if t else None

# ========= Carga inicial desde /data si existe
if st.session_state["plan_df"] is None and PLAN_CSV.exists():
    try:
        st.session_state["plan_df"] = load_csv(PLAN_CSV)
    except Exception:
        pass
if st.session_state["real_df"] is None and REAL_CSV.exists():
    try:
        st.session_state["real_df"] = load_csv(REAL_CSV)
    except Exception:
        pass
if st.session_state["merged"] is None and MERG_CSV.exists():
    try:
        st.session_state["merged"] = load_csv(MERG_CSV)
    except Exception:
        pass

# ========= Páginas =========
if page == "Configuración":
    st.subheader("Carga de Planificación")
    up_plan = st.file_uploader("Excel de Planificación", type=["xlsx","xlsm"], key="plan")
    if up_plan:
        dfp_raw = pd.read_excel(up_plan)
        st.write("Vista previa Planificación:", dfp_raw.head())

        # Mapeo de columnas
        suggest = _guess_map(dfp_raw, EXPECTED_PLAN)
        st.markdown("**Mapear columnas (Plan):**")
        m = {}
        for target in ["Fecha","Hora","Base","Moviles_Planificados","Servicios_Planificados"]:
            m[target] = st.selectbox(f"{target}", options=[""] + list(dfp_raw.columns),
                                     index=([""]+list(dfp_raw.columns)).index(suggest[target]) if suggest[target] in dfp_raw.columns else 0,
                                     key=f"map_plan_{target}")
        try:
            dfp = apply_map(dfp_raw, m, "plan")
            dfp = enrich_time(dfp)
            st.session_state["plan_df"] = dfp
            st.success("Planificación cargada y normalizada.")
        except Exception as e:
            st.error(f"Error mapeando Planificación: {e}")

    st.markdown("---")
    st.subheader("Carga de Realidad")
    up_real = st.file_uploader("Excel de Realidad", type=["xlsx","xlsm"], key="real")
    if up_real:
        dfr_raw = pd.read_excel(up_real)
        st.write("Vista previa Realidad:", dfr_raw.head())

        suggest_r = _guess_map(dfr_raw, EXPECTED_REAL)
        st.markdown("**Mapear columnas (Real):**")
        mr = {}
        for target in ["Fecha","Hora","Base","Moviles_Reales","Servicios_Reales"]:
            mr[target] = st.selectbox(f"{target}", options=[""] + list(dfr_raw.columns),
                                      index=([""]+list(dfr_raw.columns)).index(suggest_r[target]) if suggest_r[target] in dfr_raw.columns else 0,
                                      key=f"map_real_{target}")
        try:
            dfr = apply_map(dfr_raw, mr, "real")
            dfr = enrich_time(dfr)
            st.session_state["real_df"] = dfr
            st.success("Realidad cargada y normalizada.")
        except Exception as e:
            st.error(f"Error mapeando Realidad: {e}")

    st.markdown("---")
    if st.session_state["plan_df"] is not None and st.session_state["real_df"] is not None:
        st.subheader("Merge y cálculo de métricas")
        merged0 = merge_plan_real(st.session_state["plan_df"], st.session_state["real_df"])
        merged0 = add_time_keys(merged0)
        merged = compute_metrics(merged0)
        st.session_state["merged"] = merged
        st.success(f"Merge OK. Filas: {len(merged):,}")
        st.dataframe(merged.head(20), use_container_width=True)
    else:
        st.info("Cargá Planificación y Realidad para habilitar el merge.")

else:
    # Páginas analíticas requieren merged
    if st.session_state["merged"] is None:
        st.warning("Primero cargá y mapeá Planificación y Realidad en **Configuración**.")
        st.stop()

    df_all = st.session_state["merged"].copy()
    # Aplicar filtros
    bases_f = [b.strip() for b in base_sel.split(",") if b.strip()] if base_sel else None
    horas_f = _hour_list(horas_multi)
    df_f = filter_df(df_all, bases=bases_f, fecha=fecha_sel, semana=(int(semana_sel) if semana_sel>0 else None),
                     mes=(mes_text if mes_text else None), hora_sel=horas_f)

    if page == "Dashboard":
        st.subheader("KPIs globales")
        tot_plan_m = df_f["Moviles_Planificados"].sum()
        tot_real_m = df_f["Moviles_Reales"].sum()
        tot_plan_s = df_f["Servicios_Planificados"].sum()
        tot_real_s = df_f["Servicios_Reales"].sum()

        desvio_m = (tot_real_m - tot_plan_m) / tot_plan_m * 100 if tot_plan_m>0 else np.nan
        desvio_s = (tot_real_s - tot_plan_s) / tot_plan_s * 100 if tot_plan_s>0 else np.nan
        efect = 1 - (abs(tot_real_s - tot_plan_s) / tot_plan_s) if tot_plan_s>0 else np.nan

        m1,m2,m3 = st.columns(3)
        m1.metric("Móviles — % Desvío", f"{desvio_m:,.1f}%" if pd.notna(desvio_m) else "—")
        m2.metric("Servicios — % Desvío", f"{desvio_s:,.1f}%" if pd.notna(desvio_s) else "—")
        m3.metric("Efectividad", f"{efect:.1%}" if pd.notna(efect) else "—")

        # Estado general (semáforo por efectividad)
        if pd.isna(efect): color, txt = ("#6B7280", "Sin datos")
        elif efect >= 0.92: color, txt = ("#059669", "OK (≥ 92%)")
        elif efect >= 0.89: color, txt = ("#F59E0B", "Atención (89–92%)")
        else:                color, txt = ("#DC2626", "Crítico (< 89%)")
        st.markdown(f"**Estado general:** <span style='color:{color}'>{txt}</span>", unsafe_allow_html=True)

        # Gráfico línea Plan vs Real (Servicios agregados por Fecha+Hora)
        g = df_f.groupby(["Fecha","HoraStr"], as_index=False)[["Servicios_Planificados","Servicios_Reales"]].sum()
        fig1 = px.line(g, x="HoraStr", y=["Servicios_Planificados","Servicios_Reales"], color_discrete_sequence=["#22D3EE","#10B981"])
        stylize(fig1, "Plan vs Real (Servicios por hora)", y_pct=False); st.plotly_chart(fig1, use_container_width=True)

        # Barras de desvío %
        g2 = df_f.groupby("HoraStr", as_index=False)["Dif_Servicios"].sum()
        g2p = df_f.groupby("HoraStr", as_index=False)["Servicios_Planificados"].sum()
        g2 = g2.merge(g2p, on="HoraStr", how="left")
        g2["Desvio_%"] = np.where(g2["Servicios_Planificados"]>0, g2["Dif_Servicios"]/g2["Servicios_Planificados"]*100, np.nan)
        fig2 = px.bar(g2, x="HoraStr", y="Desvio_%", color="Desvio_%", color_continuous_scale="RdYlGn")
        stylize(fig2, "Desvío % por hora (Servicios)"); st.plotly_chart(fig2, use_container_width=True)

        # Heatmap por hora × día
        piv = df_f.pivot_table(values="Dif_Servicios", index="Fecha", columns="HoraStr", aggfunc="sum").fillna(0)
        fig3 = px.imshow(piv, color_continuous_scale="RdYlGn", aspect="auto")
        stylize(fig3, "Heatmap — Desvío de servicios (Real - Plan)"); st.plotly_chart(fig3, use_container_width=True)

        # Errores agregados
        mets = agg_error_metrics(df_f)
        st.markdown(f"**MAPE:** {mets['MAPE_%']:.1f}% · **MAE:** {mets['MAE']:.2f} · **Forecast Bias:** {mets['ForecastBias_%']:.1f}%")

        # Detección automática
        sub, sobre = top5_hours(df_f)
        wb = worst_base(df_f)
        c1,c2,c3 = st.columns([1,1,1])
        with c1:
            st.subheader("Top 5 Sub‑plan (horas)")
            st.dataframe(sub, use_container_width=True, hide_index=True)
        with c2:
            st.subheader("Top 5 Sobre‑plan (horas)")
            st.dataframe(sobre, use_container_width=True, hide_index=True)
        with c3:
            st.subheader("Base con mayor desvío")
            st.dataframe(wb, use_container_width=True, hide_index=True)

    elif page == "Análisis por Base":
        st.subheader("Desvío por Base (Servicios)")
        g = df_f.groupby("Base", as_index=False)[["Servicios_Planificados","Servicios_Reales"]].sum()
        g["Desvio_%"] = np.where(g["Servicios_Planificados"]>0, (g["Servicios_Reales"]-g["Servicios_Planificados"])/g["Servicios_Planificados"]*100, np.nan)
        fig = px.bar(g, x="Base", y="Desvio_%", color="Desvio_%", color_continuous_scale="RdYlGn")
        stylize(fig, "Desvío % por Base"); st.plotly_chart(fig, use_container_width=True)
        st.dataframe(g, use_container_width=True)

    elif page == "Análisis Horario":
        st.subheader("Series por hora — Plan vs Real (Servicios)")
        g = df_f.groupby("HoraStr", as_index=False)[["Servicios_Planificados","Servicios_Reales"]].sum()
        fig = px.line(g, x="HoraStr", y=["Servicios_Planificados","Servicios_Reales"], color_discrete_sequence=["#22D3EE","#10B981"])
        stylize(fig, "Plan vs Real por hora"); st.plotly_chart(fig, use_container_width=True)

        st.subheader("Distribución de desvío (Servicios)")
        g2 = df_f.groupby("HoraStr", as_index=False)["Dif_Servicios"].sum()
        fig2 = px.bar(g2, x="HoraStr", y="Dif_Servicios", color="Dif_Servicios", color_continuous_scale="RdYlGn")
        stylize(fig2, "Desvío (Real - Plan) por hora"); st.plotly_chart(fig2, use_container_width=True)
        st.dataframe(df_f[["Fecha","HoraStr","Base","Servicios_Planificados","Servicios_Reales","Dif_Servicios","Desvio_Servicios_%","Clasificacion"]].sort_values(["Fecha","HoraStr","Base"]),
                     use_container_width=True)

    elif page == "Auditoría Detallada":
        st.subheader("Tabla completa con clasificación")
        cols = ["Fecha","HoraStr","Base",
                "Moviles_Planificados","Moviles_Reales","Dif_Moviles","Desvio_Moviles_%",
                "Servicios_Planificados","Servicios_Reales","Dif_Servicios","Desvio_Servicios_%",
                "Efectividad","Clasificacion","Status","Semana","Mes","Año"]
        cols = [c for c in cols if c in df_f.columns]
        st.dataframe(df_f[cols].sort_values(["Fecha","HoraStr","Base"]), use_container_width=True)
        # Exportar a Excel
        bytes_xls, fname = to_excel_bytes(df_f[cols], sheet_name="auditoria", fname="auditoria_plan_vs_real.xlsx")
        st.download_button("⬇️ Descargar Excel (auditoría)", data=bytes_xls, file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
