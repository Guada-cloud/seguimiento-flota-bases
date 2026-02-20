# utils_ops.py — utilidades de datos y métricas (Plan vs Real)
from __future__ import annotations
import pandas as pd
import numpy as np
from io import BytesIO
from pathlib import Path
import re

# ==========================
# 1) Mapeo y normalización (Plan y Real)
# ==========================
EXPECTED_PLAN = {
    "Fecha": ["fecha", "date", "dia", "día"],
    "Hora":  ["hora", "time"],
    "Base":  ["base", "sede", "plataforma"],
    "Moviles_Planificados":  ["moviles_planificados","moviles_plan","mov_plan","mov_planif","mov req","mov requeridos"],
    "Servicios_Planificados":["servicios_planificados","serv_plan","svc_plan","llamadas_plan","svc proy","servicios proy"]
}
EXPECTED_REAL = {
    "Fecha": ["fecha","date","dia","día"],
    "Hora":  ["hora","time"],
    "Base":  ["base","sede","plataforma"],
    "Moviles_Reales":  ["moviles_reales","mov_real","mov_obs","moviles_obs","mov x nomina","moviles x nomina"],
    "Servicios_Reales":["servicios_reales","svc_real","llamadas_real","llamadas_reales","svc reales","servicios reales"]
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
    """Renombra columnas usando mapping y devuelve sólo las esperadas."""
    if kind=="plan":
        want = ["Fecha","Hora","Base","Moviles_Planificados","Servicios_Planificados"]
    else:
        want = ["Fecha","Hora","Base","Moviles_Reales","Servicios_Reales"]
    miss = [k for k,v in mapping.items() if k in want and not v]
    if miss:
        raise ValueError(f"Faltan columnas en el mapeo: {', '.join(miss)}")
    return df.rename(columns=mapping)[want].copy()

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
    out["Año"]    = out["Fecha_dt"].dt.year
    out["Mes"]    = out["Fecha_dt"].dt.month
    out["Semana"] = iso.week
    out["Dia"]    = out["Fecha_dt"].dt.day
    out["HoraStr"]= pd.to_datetime(out["Hora"].astype(str)).dt.strftime("%H:%M")
    return out

# ==========================
# 2) Merge y métricas
# ==========================
def merge_plan_real(plan: pd.DataFrame, real: pd.DataFrame) -> pd.DataFrame:
    """Merge outer por Fecha+Hora+Base y clasifica match vs no-match."""
    keys = ["Fecha","Hora","Base"]
    merged = pd.merge(plan, real, on=keys, how="outer", suffixes=("_Plan","_Real"), indicator=True)
    merged["Status"] = np.select(
        [merged["_merge"].eq("left_only"),
         merged["_merge"].eq("right_only"),
         merged["_merge"].eq("both")],
        ["No ejecutado","No planificado","OK"], default="Desconocido"
    )
    merged.drop(columns=["_merge"], inplace=True)
    return merged

def compute_metrics(df: pd.DataFrame) -> pd.DataFrame:
    """Diferencias, % desvío, clasificación, efectividad y errores (sobre Servicios)."""
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
        [out["Status"].eq("No ejecutado"),
         out["Status"].eq("No planificado"),
         out["Dif_Servicios"].eq(0),
         out["Dif_Servicios"]>0,
         out["Dif_Servicios"]<0],
        ["No ejecutado","No planificado","Exacto","Sobre planificado","Bajo planificado"], default="NA"
    )

    out["Efectividad"] = np.where(out["Servicios_Planificados"]>0,
                                  1 - (out["Dif_Servicios"].abs()/out["Servicios_Planificados"]), np.nan)
    out["APE"]  = np.where(out["Servicios_Planificados"]>0,
                           (out["Servicios_Reales"] - out["Servicios_Planificados"]).abs()/out["Servicios_Planificados"], np.nan)
    out["AE"]   = (out["Servicios_Reales"] - out["Servicios_Planificados"]).abs()
    out["Bias"] = (out["Servicios_Planificados"] - out["Servicios_Reales"])
    return out

def agg_error_metrics(df: pd.DataFrame) -> dict:
    """MAPE (%), MAE, Forecast Bias (%)."""
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

# ==========================
# 5) Formato horario simple (1 hoja)
# ==========================
def _coerce_num(s):
    """Convierte a número; '#¿NOMBRE?' y textos -> NaN."""
    return pd.to_numeric(s, errors="coerce")

def load_hourly_simple(df_raw: pd.DataFrame, fecha: str, base: str) -> tuple[pd.DataFrame, pd.DataFrame]:
    """
    Lee columnas:
      HORA | SVC PROY | SVC REALES | MOV REQ | MOVILES X NOMINA | [COEFICIENTE HS] | [DIF MOVILES]
    y devuelve plan_df y real_df normalizados.
    """
    df = df_raw.copy()

    def _find(colnames: list[str], *aliases) -> str:
        low = {str(c).strip().lower(): c for c in colnames}
        al = [a.lower() for a in aliases]
        for a in al:
            if a in low: return low[a]
        return ""

    cols = list(df.columns)
    c_hora   = _find(cols, "hora")
    c_svc_p  = _find(cols, "svc proy", "servicios proy", "svc plan", "serv plan")
    c_svc_r  = _find(cols, "svc reales", "servicios reales", "svc real", "serv real")
    c_mov_p  = _find(cols, "mov req", "mov requeridos", "moviles plan", "mov plan")
    c_mov_r  = _find(cols, "moviles x nomina", "mov x nomina", "mov reales", "mov real")
    c_coef   = _find(cols, "coeficiente hs", "coef hs")
    c_difm   = _find(cols, "dif moviles", "dif mov", "delta moviles")

    missing = [("HORA", c_hora), ("SVC PROY", c_svc_p), ("SVC REALES", c_svc_r),
               ("MOV REQ", c_mov_p), ("MOVILES X NOMINA", c_mov_r)]
    faltan = [k for k, v in missing if v == ""]
    if faltan:
        raise ValueError(f"Faltan columnas en el archivo: {', '.join(faltan)}")

    out = pd.DataFrame()
    out["Hora"] = pd.to_datetime(df[c_hora].astype(str), errors="coerce").dt.time
    out["Servicios_Planificados"] = _coerce_num(df[c_svc_p])
    out["Servicios_Reales"]       = _coerce_num(df[c_svc_r])
    out["Moviles_Planificados"]   = _coerce_num(df[c_mov_p])
    out["Moviles_Reales"]         = _coerce_num(df[c_mov_r])

    if c_coef: out["Coeficiente_HS"] = _coerce_num(df[c_coef])
    if c_difm: out["Dif_Moviles_Archivo"] = _coerce_num(df[c_difm])

    out["Fecha"] = pd.to_datetime(str(fecha)).date()
    out["Base"]  = str(base).strip().upper() if base else "TOTAL"

    plan = out[["Fecha","Hora","Base","Moviles_Planificados","Servicios_Planificados"]].copy()
    real = out[["Fecha","Hora","Base","Moviles_Reales","Servicios_Reales"]].copy()

    plan = enrich_time(plan)
    real = enrich_time(real)
    return plan, real

# ==========================
# 6) Libro multi-hojas (una hoja = una Base)
# ==========================
def _infer_base_from_text(texto: str) -> str:
    """Infiero Base desde el nombre de la hoja o un rótulo."""
    s = str(texto).upper()
    if "6001" in s:   return "PROY_6001"
    if "MECA" in s:   return "PROY_MECA"
    if "10541" in s:  return "PROY_10541"
    if "13305" in s:  return "PROY_13305"
    if "DEMOTOS" in s:return "DEMOTOS"
    s = re.sub(r"[^A-Z0-9_]+", "_", s)
    return s[:20] if s else "TOTAL"

def load_hourly_simple_book(xls_file, fecha: str):
    """
    Lee TODAS las hojas del libro como formato horario simple.
    Devuelve: (plan_concat, real_concat, reporte_list)
    """
    xl = pd.ExcelFile(xls_file)
    plan_all, real_all, reporte = [], [], []

    for sh in xl.sheet_names:
        base = _infer_base_from_text(sh)
        try:
            df_raw = pd.read_excel(xl, sheet_name=sh)
            try:
                top_left = str(df_raw.iloc[0,0])
                base2 = _infer_base_from_text(top_left)
                if base2 not in ("TOTAL", base): 
                    base = base2
            except Exception:
                pass

            plan_s, real_s = load_hourly_simple(df_raw, fecha=fecha, base=base)
            plan_all.append(plan_s); real_all.append(real_s)
            reporte.append({"sheet": sh, "base": base, "filas_plan": len(plan_s), "filas_real": len(real_s), "ok": True})
        except Exception as e:
            reporte.append({"sheet": sh, "base": base, "filas_plan": 0, "filas_real": 0, "ok": False, "error": str(e)})

    plan_concat = pd.concat(plan_all, ignore_index=True) if plan_all else pd.DataFrame()
    real_concat = pd.concat(real_all, ignore_index=True) if real_all else pd.DataFrame()
    return plan_concat, real_concat, reporte

# ==========================
# 7) Parámetros por Base (paneles)
# ==========================
def _norm_txt(s: str) -> str:
    t = str(s).strip().lower()
    for a,b in [("á","a"),("é","e"),("í","i"),("ó","o"),("ú","u"),("ñ","n")]:
        t = t.replace(a,b)
    return re.sub(r"\s+"," ", t)

def _find_label_value(df: pd.DataFrame, aliases: list[str]) -> float|None:
    """
    Busca rótulos (INTERVALO, TMO, etc.) en toda la hoja y devuelve el número
    a la derecha o debajo del rótulo. None si no encuentra.
    """
    nrows, ncols = df.shape
    for i in range(nrows):
        for j in range(ncols):
            v = df.iat[i,j]
            if pd.isna(v): 
                continue
            if not isinstance(v, (str,int,float,np.number)): 
                continue
            txt = _norm_txt(v)
            for a in aliases:
                if _norm_txt(a) == txt:
                    # derecha
                    for jj in range(j+1, min(j+4, ncols)):
                        val = pd.to_numeric(df.iat[i, jj], errors="coerce")
                        if pd.notna(val): return float(val)
                    # abajo
                    for ii in range(i+1, min(i+4, nrows)):
                        val = pd.to_numeric(df.iat[ii, j], errors="coerce")
                        if pd.notna(val): return float(val)
    return None

def parse_params_from_sheet(df_sheet: pd.DataFrame, base_guess: str, fecha: str) -> dict:
    """Extrae parámetros clave del panel de cada hoja."""
    p = {}
    # panel izquierdo
    p["Intervalo"]          = _find_label_value(df_sheet, ["intervalo"])
    p["TMO_Movil_seg"]      = _find_label_value(df_sheet, ["tmo del movil (en seg)","tmo movil (en seg)","tmo movil seg"])
    p["Dentro_SL_pct"]      = _find_label_value(df_sheet, ["% dentro de sl","dentro de sl"])
    p["Tiempo_Lleg_seg"]    = _find_label_value(df_sheet, ["tiempo lleg (seg)","tiempo llegada (seg)"])
    p["Ocupacion_Max_pct"]  = _find_label_value(df_sheet, ["ocupacion maxima","ocupacion max"])
    p["TMO_Grua_min"]       = _find_label_value(df_sheet, ["tmo grua en min","tmo grua (min)"])
    p["Tiempo_Llegada_min"] = _find_label_value(df_sheet, ["tiempo llegada en min","tiempo llegada (min)"])
    p["Reductores_pct"]     = _find_label_value(df_sheet, ["reductores","reductores %"])
    # panel derecho (requerimiento de cobertura)
    p["Servicios_Reales_total"]      = _find_label_value(df_sheet, ["servicios reales"])
    p["Servicios_Proyectados_total"] = _find_label_value(df_sheet, ["servicios proyectados","servicios proy"])
    p["Horas_Movil_Requeridas"]      = _find_label_value(df_sheet, ["horas movil requeridas","horas requeridas"])
    p["Objetivo"]                    = _find_label_value(df_sheet, ["objetivo"])
    p["Total_Horas_Op_Previstas"]    = _find_label_value(df_sheet, ["total hs operativas previstas de movil","total hs operativas"])
    p["Coeficiente_HS"]              = _find_label_value(df_sheet, ["coeficiente segun hs op previstas","coeficiente hs"])
    p["Servicios_Aprox_A_Derivar"]   = _find_label_value(df_sheet, ["servicios aprox a derivar para alcanzar coeficiente","servicios a derivar"])
    # metadatos
    p["Fecha"] = pd.to_datetime(str(fecha)).date()
    p["Base"]  = base_guess
    return p

def load_params_book(xls_file, fecha: str):
    """Extrae parámetros por Base de todas las hojas del libro."""
    xl = pd.ExcelFile(xls_file)
    rows = []
    for sh in xl.sheet_names:
        base = _infer_base_from_text(sh)
        df_raw = pd.read_excel(xl, sheet_name=sh, header=None)
        try:
            try:
                a1 = str(df_raw.iat[0,0])
                base2 = _infer_base_from_text(a1)
                if base2 and base2 != "TOTAL":
                    base = base2
            except Exception:
                pass
            rows.append(parse_params_from_sheet(df_raw, base, fecha))
        except Exception:
            rows.append({"Fecha": pd.to_datetime(str(fecha)).date(), "Base": base})
    return pd.DataFrame(rows)
