# utils_ops.py — utilidades de datos y métricas (Plan vs Real)
from __future__ import annotations
import pandas as pd
import numpy as np
from io import BytesIO
from pathlib import Path

# ==========================
# 1) Mapeo y normalización
# ==========================
EXPECTED_PLAN = {
    "Fecha": ["fecha", "date", "dia", "día"],
    "Hora":  ["hora", "time"],
    "Base":  ["base", "sede", "plataforma"],
    "Moviles_Planificados":  ["moviles_planificados","moviles_plan","mov_plan","mov_planif"],
    "Servicios_Planificados":["servicios_planificados","serv_plan","svc_plan","llamadas_plan"]
}
EXPECTED_REAL = {
    "Fecha": ["fecha","date","dia","día"],
    "Hora":  ["hora","time"],
    "Base":  ["base","sede","plataforma"],
    "Moviles_Reales":  ["moviles_reales","mov_real","mov_obs","moviles_obs"],
    "Servicios_Reales":["servicios_reales","svc_real","llamadas_real","llamadas_reales"]
}

def _guess_map(df: pd.DataFrame, expected: dict[str, list[str]]) -> dict[str, str]:
    """Sugerir mapeo por coincidencia insensible a mayúsculas y acentos simples."""
    def norm(s: str) -> str:
        tr = str(s).strip().lower()
        tr = (tr.replace("á","a").replace("é","e").replace("í","i")
                .replace("ó","o").replace("ú","u").replace("ñ","n"))
        return tr
    cols = {norm(c): c for c in df.columns}
    m = {}
    for target, aliases in expected.items():
        hit = None
        for a in aliases+[target]:
            na = norm(a)
            if na in cols:
                hit = cols[na]; break
        m[target] = hit if hit else ""
    return m

def apply_map(df: pd.DataFrame, mapping: dict[str,str], kind: str) -> pd.DataFrame:
    """Renombra columnas usando mapping y devuelve un DataFrame con sólo las esperadas."""
    if kind=="plan":
        want = ["Fecha","Hora","Base","Moviles_Planificados","Servicios_Planificados"]
    else:
        want = ["Fecha","Hora","Base","Moviles_Reales","Servicios_Reales"]
    miss = [k for k,v in mapping.items() if k in want and not v]
    if miss:
        raise ValueError(f"Faltan columnas en el mapeo: {', '.join(miss)}")
    df2 = df.rename(columns=mapping)[want].copy()
    return df2

def enrich_time(df: pd.DataFrame) -> pd.DataFrame:
    """Convierte Fecha/Hora, crea Año, Mes, Semana ISO, Día, HoraStr."""
    out = df.copy()
    out["Fecha"] = pd.to_datetime(out["Fecha"], errors="coerce").dt.date
    if np.issubdtype(pd.Series(out["Hora"]).dtype, np.number):
        out["Hora"] = pd.to_timedelta((out["Hora"]%1)*24, unit="h")
        out["Hora"] = (pd.Timestamp("1900-01-01")+out["Hora"]).dt.time
    else:
        out["Hora"] = pd.to_datetime(out["Hora"].astype(str), errors="coerce").dt.time
    out["Fecha_dt"] = pd.to_datetime(out["Fecha"])
    iso = out["Fecha_dt"].dt.isocalendar()
    out["Año"]   = out["Fecha_dt"].dt.year
    out["Mes"]   = out["Fecha_dt"].dt.month
    out["Semana"] = iso.week             # Semana ISO
    out["Dia"]   = out["Fecha_dt"].dt.day
    out["DiaNombre"] = out["Fecha_dt"].dt.day_name()  # en ES si el sistema tiene locale
    out["HoraStr"]   = pd.to_datetime(out["Hora"].astype(str)).dt.strftime("%H:%M")
    return out

# ==========================
# 2) Merge y clasificación
# ==========================
def merge_plan_real(plan: pd.DataFrame, real: pd.DataFrame) -> pd.DataFrame:
    """Merge outer por Fecha+Hora+Base y clasifica match vs no-match."""
    keys = ["Fecha","Hora","Base"]
    merged = pd.merge(
        plan, real, on=keys, how="outer", suffixes=("_Plan","_Real"), indicator=True
    )
    merged["Status"] = np.select(
        [
            merged["_merge"].eq("left_only"),
            merged["_merge"].eq("right_only"),
            merged["_merge"].eq("both"),
        ],
        ["No ejecutado","No planificado","OK"],
        default="Desconocido"
    )
    merged.drop(columns=["_merge"], inplace=True)
    return merged

def compute_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """Calcula dif., % desvío, clasificación, efectividad y errores (Servicios)."""
    out = df.copy()
    for c in ["Moviles_Planificados","Servicios_Planificados","Moviles_Reales","Servicios_Reales"]:
        if c not in out: out[c] = 0
    out[["Moviles_Planificados","Servicios_Planificados","Moviles_Reales","Servicios_Reales"]] = \
        out[["Moviles_Planificados","Servicios_Planificados","Moviles_Reales","Servicios_Reales"]].fillna(0)

    out["Dif_Moviles"]   = out["Moviles_Reales"]   - out["Moviles_Planificados"]
    out["Dif_Servicios"] = out["Servicios_Reales"] - out["Servicios_Planificados"]

    out["Desvio_Moviles_%"] = np.where(out["Moviles_Planificados"]>0,
                                       out["Dif_Moviles"]/out["Moviles_Planificados"]*100, np.nan)
    out["Desvio_Servicios_%"] = np.where(out["Servicios_Planificados"]>0,
                                         out["Dif_Servicios"]/out["Servicios_Planificados"]*100, np.nan)

    out["Clasificacion"] = np.select(
        [
            out["Status"].eq("No ejecutado"),
            out["Status"].eq("No planificado"),
            out["Dif_Servicios"].eq(0),
            out["Dif_Servicios"]>0,
            out["Dif_Servicios"]<0
        ],
        ["No ejecutado","No planificado","Exacto","Sobre planificado","Bajo planificado"],
        default="NA"
    )

    out["Efectividad"] = np.where(out["Servicios_Planificados"]>0,
                                  1 - (out["Dif_Servicios"].abs()/out["Servicios_Planificados"]), np.nan)
    out["APE"]  = np.where(out["Servicios_Planificados"]>0,
                           (out["Servicios_Reales"] - out["Servicios_Planificados"]).abs()/out["Servicios_Planificados"], np.nan)
    out["AE"]   = (out["Servicios_Reales"] - out["Servicios_Planificados"]).abs()
    out["Bias"] = (out["Servicios_Planificados"] - out["Servicios_Reales"])
    return out

def agg_error_metrics(df: pd.DataFrame) -> dict:
    """MAPE, MAE, Forecast Bias% (sobre Servicios)."""
    d = df.copy()
    mape = d["APE"].mean()*100 if d["APE"].notna().any() else np.nan
    mae  = d["AE"].mean() if d["AE"].notna().any() else np.nan
    fbias = (d["Bias"].sum()/d["Servicios_Reales"].sum()*100) if d["Servicios_Reales"].sum()!=0 else np.nan
    return {"MAPE_%":mape, "MAE":mae, "ForecastBias_%":fbias}

# ==========================
# 3) Filtros y top‑N
# ==========================
def add_time_keys(df: pd.DataFrame) -> pd.DataFrame:
    if "Fecha_dt" not in df.columns:
        df["Fecha_dt"] = pd.to_datetime(df["Fecha"])
    iso = df["Fecha_dt"].dt.isocalendar()
    df["Año"] = df["Fecha_dt"].dt.year
    df["Mes"] = df["Fecha_dt"].dt.month
    df["Semana"] = iso.week
    df["Dia"] = df["Fecha_dt"].dt.day
    df["HoraStr"] = pd.to_datetime(df["Hora"].astype(str)).dt.strftime("%H:%M")
    return df

def filter_df(df: pd.DataFrame, bases: list[str]|None=None,
              fecha: pd.Timestamp|None=None, semana:int|None=None,
              mes: str|None=None, hora_sel:list[str]|None=None) -> pd.DataFrame:
    d = df.copy()
    if bases: d = d[d["Base"].isin(bases)]
    if fecha is not None: d = d[d["Fecha"].eq(pd.to_datetime(fecha).date())]
    if semana: d = d[d["Semana"].eq(int(semana))]
    if mes:
        try:
            aa, mm = mes.split("-"); aa=int(aa); mm=int(mm)
            d = d[(d["Año"].eq(aa)) & (d["Mes"].eq(mm))]
        except Exception:
            pass
    if hora_sel: d = d[d["HoraStr"].isin(hora_sel)]
    return d

def top5_hours(df: pd.DataFrame):
    g = df.groupby("HoraStr", as_index=False)["Dif_Servicios"].sum()
    sub = g.nsmallest(5, "Dif_Servicios")
    sobre = g.nlargest(5, "Dif_Servicios")
    return sub, sobre

def worst_base(df: pd.DataFrame):
    g_abs = df.groupby("Base", as_index=False)["Dif_Servicios"].apply(lambda s: s.abs().sum()).rename(columns={"Dif_Servicios":"AbsDesvio"})
    g = df.groupby("Base", as_index=False)["Dif_Servicios"].sum()
    out = g.merge(g_abs, on="Base", how="left").sort_values("AbsDesvio", ascending=False)
    return out.head(1)

# ==========================
# 4) Persistencia CSV + export
# ==========================
DATA_DIR = Path("data")
DATA_DIR.mkdir(exist_ok=True)
PLAN_CSV = DATA_DIR/"plan.csv"
REAL_CSV = DATA_DIR/"real.csv"
MERG_CSV = DATA_DIR/"merged.csv"

def save_csv(df: pd.DataFrame, path: Path): df.to_csv(path, index=False, encoding="utf-8")
def load_csv(path: Path) -> pd.DataFrame|None:
    return pd.read_csv(path, encoding="utf-8") if path.exists() else None

def to_excel_bytes(df: pd.DataFrame, sheet_name="datos", fname="reporte.xlsx") -> tuple[bytes,str]:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as wr:
        df.to_excel(wr, index=False, sheet_name=sheet_name)
    return buf.getvalue(), fname
