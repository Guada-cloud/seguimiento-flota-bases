# app.py — Streamlit (ES) con Bases
import streamlit as st, pandas as pd, numpy as np, plotly.express as px
from utils_kpi import (
    normalizar_entrada, kpis_dia, kpis_semanal, kpis_mensual,
    BLOQUES_VALIDOS, FRANJAS, BASES_VALIDAS
)

st.set_page_config(page_title='Seguimiento de Precisión — Flota', layout='wide')
st.title('Seguimiento de Precisión — Flota')
st.caption('KPIs Diario (principal), Semana a semana y Mensual · Cortes por **Base**, **Bloque** y **Franja** · Sin macros')

with st.sidebar:
    st.header('Cargar datos')
    f = st.file_uploader(
        'Subí tu Excel (.xlsx). Puede ser 1 hoja por Base (PROY_6001, PROY_MECA, etc.) o un plano con columna Base.',
        type=['xlsx','xlsm']
    )
    demo = st.checkbox('Usar datos de muestra', value=not bool(f))
    st.markdown("**Tips**  \n- Si tu archivo tiene una hoja por Base, infiero la **Base** del **nombre de la hoja**.  \n- Columnas mínimas: **Fecha, Hora, Bloque, Proy, Real** (y opcional **Base**).")

@st.cache_data(show_spinner=False)
def _load_demo():
    rng = pd.date_range('2026-02-01','2026-02-10', freq='H')
    rows=[]; bls=BLOQUES_VALIDOS
    bases=['PROY_6001','PROY_MECA','PROY_10541','PROY_13305','DEMOTOS']
    for ts in rng:
        for bl in bls:
            for base in bases:
                proy = np.random.randint(40, 160)
                real = max(0, int(proy + np.random.normal(0, proy*0.15)))
                rows.append(dict(Fecha=ts.date(), Hora=ts.time().strftime('%H:%M'), Bloque=bl, Proy=proy, Real=real, Base=base))
    df = pd.DataFrame(rows)
    return normalizar_entrada(df)

@st.cache_data(show_spinner=False)
def _read_excel_to_long(file) -> pd.DataFrame:
    xls = pd.ExcelFile(file)
    frames = []
    for sheet in xls.sheet_names:
        d = pd.read_excel(xls, sheet_name=sheet)
        try:
            frames.append(normalizar_entrada(d, sheet_name=sheet))
        except Exception:
            # hojas no compatibles se ignoran
            pass
    if not frames:
        d = pd.read_excel(xls, sheet_name=0)
        frames = [normalizar_entrada(d, sheet_name=xls.sheet_names[0])]
    return pd.concat(frames, ignore_index=True)

df = _load_demo() if (demo and not f) else (_read_excel_to_long(f) if f else None)
if df is None:
    st.warning('Subí un archivo o activá **datos de muestra**.')
    st.stop()

st.success(f"Datos cargados: {len(df):,} filas · {df['Fecha'].min()} → {df['Fecha'].max()} · Bases: {', '.join(sorted(df['Base'].unique()))}")

# ===== Filtros =====
colf = st.columns([1.2,1,1,1.2,1.2,1.2])
with colf[0]:
    fecha_sel = st.date_input('Fecha (diario)', value=pd.to_datetime(df['Fecha'].max()).date())
with colf[1]:
    semana_sel_default = int(df.loc[df['Fecha']==fecha_sel,'Semana'].head(1).fillna(0).values[0]) if (df['Fecha']==fecha_sel).any() else int(df['Semana'].dropna().max())
    semana_sel = st.number_input('Semana (comparar)', value=int(semana_sel_default), step=1)
with colf[2]:
    mes_def = pd.to_datetime(fecha_sel).strftime('%Y-%m')
    mes_text = st.text_input('Mes (aaaa-mm)', value=mes_def)
    try:
        anio, mes = int(mes_text.split('-')[0]), int(mes_text.split('-')[1])
    except Exception:
        anio, mes = int(pd.to_datetime(df['Fecha'].max()).year), int(pd.to_datetime(df['Fecha'].max()).month)
with colf[3]:
    bases = sorted(df['Base'].unique().tolist())
    selected_bases = st.multiselect('Bases', options=bases, default=bases)
with colf[4]:
    bloque = st.selectbox('Bloque/Zona', options=['Todos']+BLOQUES_VALIDOS, index=0)
with colf[5]:
    franja = st.selectbox('Franja', options=['Todos']+FRANJAS, index=0)

st.divider()

TAB_TOTAL, TAB_BASES = st.tabs(["Visión Total","Detalle por Base"])

with TAB_TOTAL:
    by_hour, agg_total, prec_dia, wape_dia, sesgo_dia = kpis_dia(df, fecha_sel, selected_bases, bloque, franja)
    k1,k2,k3 = st.columns(3)
    k1.metric('Precisión (día)', f"{prec_dia:.2%}" if pd.notna(prec_dia) else '—')
    k2.metric('WAPE (día)', f"{wape_dia:.2%}" if pd.notna(wape_dia) else '—')
    k3.metric('Sesgo (día)', f"{sesgo_dia:,.0f}" if pd.notna(sesgo_dia) else '—')

    # Proy vs Real por hora (TOTAL)
    bh = by_hour.groupby('Hora', as_index=False).agg(Proy=('Proy','sum'), Real=('Real','sum'))
    fig1 = px.bar(bh, x='Hora', y=['Proy','Real'], barmode='group', title='Proyectado vs Real por hora (TOTAL)')
    fig1.update_layout(legend_title_text='', xaxis_title='Hora', yaxis_title='Servicios')
    st.plotly_chart(fig1, use_container_width=True)

    # Semanal (TOTAL)
    g = kpis_semanal(df, anio, mes, semana_sel, selected_bases, bloque, franja)
    g_total = g.groupby('Semana', as_index=False).agg(Precisión=('Precisión','mean'))
    fig2 = px.line(g_total, x='Semana', y='Precisión', markers=True, title='Precisión por semana (TOTAL)')
    fig2.update_layout(yaxis_tickformat='.0%')
    st.plotly_chart(fig2, use_container_width=True)

    # Mensual — WAPE por franja (TOTAL) y Precisión por bloque (TOTAL)
    fran_m, bloq_m = kpis_mensual(df, anio, mes, selected_bases, bloque if bloque!='Todos' else None, franja if franja!='Todos' else None)
    fm_total = fran_m.groupby('Valor', as_index=False).agg(WAPE=('WAPE','mean'))
    bm_total = bloq_m.groupby('Valor', as_index=False).agg(Precisión=('Precisión','mean'))

    cm1, cm2 = st.columns(2)
    with cm1:
        st.subheader('WAPE por franja (mensual) — TOTAL')
        fig3 = px.bar(fm_total, x='Valor', y='WAPE', text_auto='.1%')
        fig3.update_layout(yaxis_tickformat='.0%'); st.plotly_chart(fig3, use_container_width=True)
    with cm2:
        st.subheader('Precisión por bloque (mensual) — TOTAL')
        orden = ['TOTAL'] + [b for b in BLOQUES_VALIDOS]
        bm_total['Valor'] = pd.Categorical(bm_total['Valor'], categories=orden, ordered=True)
        bm_total = bm_total.sort_values('Valor')
        fig4 = px.bar(bm_total, x='Valor', y='Precisión', text_auto='.1%')
        fig4.update_layout(yaxis_tickformat='.0%'); st.plotly_chart(fig4, use_container_width=True)

with TAB_BASES:
    by_hour, agg_total, prec_dia, wape_dia, sesgo_dia = kpis_dia(df, fecha_sel, selected_bases, bloque, franja)
    st.subheader('KPIs por Base (día seleccionado)')
    st.dataframe(agg_total[['Base','Proy','Real','ErrorAbs','Precisión']])

    fig5 = px.bar(by_hour, x='Hora', y='Real', color='Base', barmode='group', title='Real por hora, desglosado por Base')
    fig5.update_layout(xaxis_title='Hora', yaxis_title='Servicios')
    st.plotly_chart(fig5, use_container_width=True)

    g = kpis_semanal(df, anio, mes, semana_sel, selected_bases, bloque, franja)
    fig6 = px.line(g, x='Semana', y='Precisión', color='Base', markers=True, title='Precisión por semana, por Base')
    fig6.update_layout(yaxis_tickformat='.0%'); st.plotly_chart(fig6, use_container_width=True)

    fran_m, bloq_m = kpis_mensual(df, anio, mes, selected_bases, bloque if bloque!='Todos' else None, franja if franja!='Todos' else None)
    cm1, cm2 = st.columns(2)
    with cm1:
        st.subheader('WAPE por franja (mensual) — por Base')
        fig7 = px.bar(fran_m, x='Valor', y='WAPE', color='Base', barmode='group', text_auto='.1%')
        fig7.update_layout(yaxis_tickformat='.0%'); st.plotly_chart(fig7, use_container_width=True)
    with cm2:
        st.subheader('Precisión por bloque (mensual) — por Base')
        orden = ['TOTAL'] + [b for b in BLOQUES_VALIDOS]
        bloq_m['Valor'] = pd.Categorical(bloq_m['Valor'], categories=orden, ordered=True)
        bloq_m = bloq_m.sort_values(['Base','Valor'])
        fig8 = px.bar(bloq_m, x='Valor', y='Precisión', color='Base', barmode='group', text_auto='.1%')
        fig8.update_layout(yaxis_tickformat='.0%'); st.plotly_chart(fig8, use_container_width=True)

with st.expander('Ver tablas (diario / semanal / mensual)'):
    st.write('Muestra de datos normalizados:')
    st.dataframe(df.head(30))
