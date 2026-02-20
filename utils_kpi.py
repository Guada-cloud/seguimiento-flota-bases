# utils_kpi.py — funciones de transformación y KPIs (es-AR) con dimensión Base
from __future__ import annotations
import pandas as pd
import numpy as np
from datetime import datetime
from typing import List, Optional

BLOQUES_VALIDOS = ["CABA","PLAT","N1","N2","N3","O1","S1","S2"]
FRANJAS = ["00-05","07-19","20-23"]
BASES_VALIDAS = ["PROY_6001","PROY_MECA","PROY_10541","PROY_13305","DEMOTOS"]

# === Equivalente a WEEKNUM(fecha,2) (Sistema 1; semana inicia Lunes) ===
def weeknum_excel_system1_monday(dt: pd.Timestamp) -> int:
    year = dt.year
    jan1 = pd.Timestamp(year,1,1)
    start_week1 = jan1 - pd.Timedelta(days=jan1.weekday())
    delta_days = (dt.normalize() - start_week1.normalize()).days
    return 1 if delta_days < 0 else int(delta_days // 7) + 1

def asignar_franja(hora) -> str:
    if isinstance(hora, str):
        try:
            h = int(hora.split(":")[0])
        except Exception:
            return ""
    elif isinstance(hora, pd.Timestamp):
        h = hora.hour
    elif isinstance(hora, pd.Timedelta):
        h = (datetime.min + hora).hour
    else:
        return ""
    if 7 <= h < 20:   return "07-19"
    elif 20 <= h <= 23: return "20-23"
    else:               return "00-05"

def _coerce_time_col(s: pd.Series) -> pd.Series:
    if np.issubdtype(s.dtype, np.number):
        return pd.to_timedelta((s % 1) * 24, unit='h')
    def parse_h(x):
        if pd.isna(x): return pd.NaT
        if isinstance(x, (pd.Timestamp, pd.Timedelta)): return x
        try:
            hh, mm = str(x).split(":")[:2]
            return pd.to_timedelta(int(hh), 'h') + pd.to_timedelta(int(mm), 'm')
        except Exception:
            return pd.NaT
    return s.map(parse_h)

def _infer_base_from_sheet(sheet_name: str) -> str:
    name = sheet_name.strip().upper()
    for b in BASES_VALIDAS:
        if b in name: return b
    return name.replace(" ", "_")[:20]

def normalizar_entrada(df: pd.DataFrame, sheet_name: Optional[str]=None) -> pd.DataFrame:
    def col_like(dfc, *names):
        for n in names:
            for c in dfc.columns:
                if c.strip().lower() == n.lower(): return c
        return None

    c_fecha = col_like(df, "fecha","date")
    c_hora  = col_like(df, "hora","time")
    c_bloq  = col_like(df, "bloque","zona","bloque/zona","block")
    c_proy  = col_like(df, "proy","proyectado","svc proy","proyectados","proyeccion","proyección")
    c_real  = col_like(df, "real","reales","svc reales","observado","observados")
    c_base  = col_like(df, "base","plataforma","sistema")

    required = [c_fecha,c_hora,c_bloq,c_proy,c_real]
    if any(x is None for x in required):
        faltan = [n for n,x in zip(['Fecha','Hora','Bloque','Proy','Real'], required) if x is None]
        raise ValueError(f"Faltan columnas obligatorias: {', '.join(faltan)}")

    out = df[[c_fecha,c_hora,c_bloq,c_proy,c_real]].copy()
    out.columns = ['Fecha','Hora','Bloque','Proy','Real']
    out['Fecha'] = pd.to_datetime(out['Fecha'], errors='coerce').dt.date
    out['Hora']  = _coerce_time_col(out['Hora'])

    out['Bloque'] = out['Bloque'].astype(str).str.strip().str.upper()
    out = out[out['Bloque'].isin(BLOQUES_VALIDOS)]
    out['Proy'] = pd.to_numeric(out['Proy'], errors='coerce')
    out['Real'] = pd.to_numeric(out['Real'], errors='coerce')

    # Base
    if c_base is not None:
        out['Base'] = df[c_base].astype(str).str.strip().str.upper()
    else:
        out['Base'] = _infer_base_from_sheet(sheet_name or 'SIN_BASE')

    # Derivados
    out['Semana'] = [weeknum_excel_system1_monday(pd.Timestamp(f)) if pd.notna(f) else np.nan for f in out['Fecha']]
    out['Día']    = [pd.Timestamp(f).strftime('%A') if pd.notna(f) else '' for f in out['Fecha']]
    out['Franja'] = out['Hora'].map(asignar_franja)
    out['Sesgo']  = out['Proy'] - out['Real']
    out['ErrorAbs'] = out['Sesgo'].abs()
    out['Precisión_fila'] = np.where(out['Real']>0, 1 - out['ErrorAbs']/out['Real'], np.nan)

    return out.dropna(subset=['Fecha','Hora'])

def kpis_dia(df: pd.DataFrame, fecha, bases: List[str], bloque: Optional[str], franja: Optional[str]):
    d = df[df['Fecha']==fecha]
    if bases: d = d[d['Base'].isin(bases)]
    if bloque and bloque!="Todos": d = d[d['Bloque']==bloque]
    if franja and franja!="Todos": d = d[d['Franja']==franja]

    by_hour = d.groupby(['Hora','Base'], as_index=False).agg(Proy=('Proy','sum'), Real=('Real','sum'))
    by_hour['ErrorAbs'] = (by_hour['Proy']-by_hour['Real']).abs()

    agg_total = by_hour.groupby('Base', as_index=False).agg(Proy=('Proy','sum'), Real=('Real','sum'), ErrorAbs=('ErrorAbs','sum'))
    agg_total['Precisión'] = np.where(agg_total['Real']>0, 1-agg_total['ErrorAbs']/agg_total['Real'], np.nan)

    t_proy, t_real, t_err = agg_total['Proy'].sum(), agg_total['Real'].sum(), agg_total['ErrorAbs'].sum()
    precision = (1 - t_err/t_real) if t_real>0 else np.nan
    wape = (t_err/t_real) if t_real>0 else np.nan
    sesgo = t_proy - t_real
    return by_hour, agg_total, precision, wape, sesgo

def kpis_semanal(df: pd.DataFrame, anio:int, mes:int, semana:int, bases: List[str], bloque: Optional[str], franja: Optional[str]):
    d = df[(pd.Series([f.year for f in df['Fecha']])==anio) & (pd.Series([f.month for f in df['Fecha']])==mes)]
    if bases: d = d[d['Base'].isin(bases)]
    if bloque and bloque!="Todos": d = d[d['Bloque']==bloque]
    if franja and franja!="Todos": d = d[d['Franja']==franja]

    g = d.groupby(['Semana','Base'], as_index=False).agg(Proy=('Proy','sum'), Real=('Real','sum'))
    g['ErrorAbs'] = (g['Proy']-g['Real']).abs()
    g['Precisión'] = np.where(g['Real']>0, 1-g['ErrorAbs']/g['Real'], np.nan)
    g['WAPE'] = np.where(g['Real']>0, g['ErrorAbs']/g['Real'], np.nan)

    sems = sorted(g['Semana'].dropna().unique().tolist())
    if semana and semana in sems:
        idx = sems.index(semana); sems = sems[idx:idx+6]
    else:
        sems = sems[:6]
    g = g[g['Semana'].isin(sems)].sort_values(['Base','Semana'])
    g['Precisión_sem_ant'] = g.groupby('Base')['Precisión'].shift(1)
    g['Delta_pp'] = g['Precisión'] - g['Precisión_sem_ant']
    return g

def kpis_mensual(df: pd.DataFrame, anio:int, mes:int, bases: List[str], bloque: Optional[str], franja: Optional[str]):
    d = df[(pd.Series([f.year for f in df['Fecha']])==anio) & (pd.Series([f.month for f in df['Fecha']])==mes)]
    if bases: d = d[d['Base'].isin(bases)]

    # Por franja y base (opcional: filtro de Bloque)
    df_fran = d if (not bloque or bloque=="Todos") else d[d['Bloque']==bloque]
    out_f=[]
    for (fr, base), dd in df_fran.groupby(['Franja','Base']):
        proy, real = dd['Proy'].sum(), dd['Real'].sum()
        err = (dd['Proy']-dd['Real']).abs().sum()
        prec = (1-err/real) if real>0 else np.nan
        wape = (err/real) if real>0 else np.nan
        out_f.append(dict(Corte='Franja', Valor=fr, Base=base, Proy=proy, Real=real, ErrorAbs=err, Precisión=prec, WAPE=wape))
    fran = pd.DataFrame(out_f)

    # Por bloque y base (opcional: filtro de Franja)
    df_b = d if (not franja or franja=="Todos") else d[d['Franja']==franja]
    out_b=[]
    for base, dd_base in df_b.groupby('Base'):
        proy, real = dd_base['Proy'].sum(), dd_base['Real'].sum()
        err = (dd_base['Proy']-dd_base['Real']).abs().sum()
        prec = (1-err/real) if real>0 else np.nan
        wape = (err/real) if real>0 else np.nan
        out_b.append(dict(Corte='Bloque', Valor='TOTAL', Base=base, Proy=proy, Real=real, ErrorAbs=err, Precisión=prec, WAPE=wape))
        for bl, dd in dd_base.groupby('Bloque'):
            proy, real = dd['Proy'].sum(), dd['Real'].sum()
            err = (dd['Proy']-dd['Real']).abs().sum()
            prec = (1-err/real) if real>0 else np.nan
            wape = (err/real) if real>0 else np.nan
            out_b.append(dict(Corte='Bloque', Valor=bl, Base=base, Proy=proy, Real=real, ErrorAbs=err, Precisión=prec, WAPE=wape))
    bloq = pd.DataFrame(out_b)

    return fran, bloq
