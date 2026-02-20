# ==========================
# 7) Parámetros por Base (panel izquierdo/derecho de cada hoja)
# ==========================
import re

def _norm_txt(s: str) -> str:
    """Normaliza texto: minúsculas sin acento/ñ, sin dobles espacios."""
    t = str(s).strip().lower()
    for a,b in [("á","a"),("é","e"),("í","i"),("ó","o"),("ú","u"),("ñ","n")]:
        t = t.replace(a,b)
    t = re.sub(r"\s+"," ", t)
    return t

def _find_label_value(df: pd.DataFrame, aliases: list[str]) -> float|None:
    """
    Busca un rótulo (INTERVALO, TMO DEL MOVIL (EN SEG), etc.) en toda la hoja.
    Si lo encuentra, devuelve el primer valor numérico que esté en:
      - la celda a la derecha (misma fila, siguiente(s) columna(s)), o
      - la celda de abajo (una o dos filas debajo, misma columna).
    Devuelve None si no lo encuentra.
    """
    # mapeo rápido string->posiciones
    nrows, ncols = df.shape
    for i in range(nrows):
        for j in range(ncols):
            v = df.iat[i,j]
            if pd.isna(v): 
                continue
            if not isinstance(v, str) and not isinstance(v, (int,float,np.number)): 
                continue
            txt = _norm_txt(v)
            for a in aliases:
                if _norm_txt(a) == txt:
                    # derecha
                    for jj in range(j+1, min(j+4, ncols)):
                        val = pd.to_numeric(df.iat[i, jj], errors="coerce")
                        if pd.notna(val): 
                            return float(val)
                    # abajo
                    for ii in range(i+1, min(i+4, nrows)):
                        val = pd.to_numeric(df.iat[ii, j], errors="coerce")
                        if pd.notna(val): 
                            return float(val)
    return None

def parse_params_from_sheet(df_sheet: pd.DataFrame, base_guess: str, fecha: str) -> dict:
    """
    Intenta extraer parámetros del panel de la hoja.
    Campos: Intervalo, TMO_Movil_seg, Dentro_SL_pct, Tiempo_Lleg_seg,
            Ocupacion_Max_pct, TMO_Grua_min, Tiempo_Llegada_min, Reductores_pct,
            Servicios_Reales_total, Servicios_Proyectados_total,
            Horas_Movil_Requeridas, Objetivo, Total_Horas_Op_Previstas,
            Coeficiente_HS, Servicios_Aprox_Derivar.
    """
    p = {}
    # valores panel izquierdo
    p["Intervalo"]          = _find_label_value(df_sheet, ["intervalo"])
    p["TMO_Movil_seg"]      = _find_label_value(df_sheet, ["tmo del movil (en seg)", "tmo movil (en seg)", "tmo movil seg"])
    p["Dentro_SL_pct"]      = _find_label_value(df_sheet, ["% dentro de sl", "dentro de sl"])
    p["Tiempo_Lleg_seg"]    = _find_label_value(df_sheet, ["tiempo lleg (seg)", "tiempo llegada (seg)"])
    p["Ocupacion_Max_pct"]  = _find_label_value(df_sheet, ["ocupacion maxima", "ocupacion max"])
    p["TMO_Grua_min"]       = _find_label_value(df_sheet, ["tmo grua en min","tmo grua (min)"])
    p["Tiempo_Llegada_min"] = _find_label_value(df_sheet, ["tiempo llegada en min","tiempo llegada (min)"])
    p["Reductores_pct"]     = _find_label_value(df_sheet, ["reductores","reductores %"])

    # panel derecho (requerimiento de cobertura)
    p["Servicios_Reales_total"]        = _find_label_value(df_sheet, ["servicios reales"])
    p["Servicios_Proyectados_total"]   = _find_label_value(df_sheet, ["servicios proyectados","servicios proy"])
    p["Horas_Movil_Requeridas"]        = _find_label_value(df_sheet, ["horas movil requeridas","horas requeridas"])
    p["Objetivo"]                      = _find_label_value(df_sheet, ["objetivo"])
    p["Total_Horas_Op_Previstas"]      = _find_label_value(df_sheet, ["total hs operativas previstas de movil","total hs operativas"])
    p["Coeficiente_HS"]                = _find_label_value(df_sheet, ["coeficiente segun hs op previstas","coeficiente hs"])
    p["Servicios_Aprox_A_Derivar"]     = _find_label_value(df_sheet, ["servicios aprox a derivar para alcanzar coeficiente","servicios a derivar"])

    # metadatos
    p["Fecha"] = pd.to_datetime(str(fecha)).date()
    p["Base"]  = base_guess
    return p

def load_params_book(xls_file, fecha: str):
    """
    Lee todas las hojas y extrae parámetros por Base.
    Devuelve DataFrame con una fila por Base (y fecha).
    """
    xl = pd.ExcelFile(xls_file)
    rows = []
    for sh in xl.sheet_names:
        base = _infer_base_from_text(sh)
        df_raw = pd.read_excel(xl, sheet_name=sh, header=None)  # sin encabezados, para buscar etiquetas en cualquier lado
        try:
            # si A1 tiene un rótulo informativo, ajusto la base
            try:
                a1 = str(df_raw.iat[0,0])
                base2 = _infer_base_from_text(a1)
                if base2 and base2 != "TOTAL":
                    base = base2
            except Exception:
                pass
            rows.append(parse_params_from_sheet(df_raw, base, fecha))
        except Exception:
            # si falla la hoja, igual dejo registro mínimo
            rows.append({"Fecha": pd.to_datetime(str(fecha)).date(),
                         "Base": base})
    return pd.DataFrame(rows), fname
