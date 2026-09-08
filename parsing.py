"""
Lógica de parseo para el dashboard TES (ALM).
Portado del dashboard HTML/JS a Python -- misma lógica, mismas reglas de
negocio, validada contra los mismos archivos reales.
"""
import re
import io
import datetime
import openpyxl
import xlrd


# ---------------- lectura cruda de archivos ----------------
def read_raw_rows(file_bytes, filename):
    """Devuelve una lista de filas (listas de valores crudos), soporta .xlsx y
    el formato legado .xls (usado por el boletín SEN)."""
    if filename.lower().endswith(".xls"):
        wb = xlrd.open_workbook(file_contents=file_bytes)
        sh = wb.sheet_by_index(0)
        rows = []
        for r in range(sh.nrows):
            rows.append(sh.row_values(r))
        return rows
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    ws = wb.worksheets[0]
    return [list(row) for row in ws.iter_rows(values_only=True)]


# ---------------- helpers de fecha/hora ----------------
def time_from_cell(v):
    if isinstance(v, datetime.datetime):
        return v.hour * 3600 + v.minute * 60 + v.second
    if isinstance(v, datetime.time):
        return v.hour * 3600 + v.minute * 60 + v.second
    if isinstance(v, (int, float)):
        return round(v * 86400)
    if isinstance(v, str) and v.strip():
        parts = v.strip().split(":")
        try:
            h = int(parts[0]) if len(parts) > 0 and parts[0] else 0
            m = int(parts[1]) if len(parts) > 1 and parts[1] else 0
            s = int(float(parts[2])) if len(parts) > 2 and parts[2] else 0
            return h * 3600 + m * 60 + s
        except (ValueError, IndexError):
            return None
    return None


def parse_date_from_filename(name):
    """DDMMAAAA en el nombre del archivo (formato OperDia)."""
    if not name:
        return None
    m = re.search(r"(\d{2})(\d{2})(\d{4})", name)
    if not m:
        return None
    dd, mm, yyyy = int(m.group(1)), int(m.group(2)), int(m.group(3))
    if mm < 1 or mm > 12 or dd < 1 or dd > 31 or yyyy < 2020:
        return None
    try:
        return datetime.date(yyyy, mm, dd)
    except ValueError:
        return None


def parse_date_from_sen_filename(name):
    """AAAA-MM-DD en el nombre del archivo (formato boletín SEN)."""
    if not name:
        return None
    m = re.search(r"(\d{4})-(\d{2})-(\d{2})", name)
    if not m:
        return None
    yyyy, mm, dd = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        return datetime.date(yyyy, mm, dd)
    except ValueError:
        return None


def parse_fecha_cierre(v):
    """Fecha Cierre del histórico mensual: un solo campo YYMMDD (ej. 260803
    = 03/08/2026), como número o como texto."""
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, datetime.date):
        return v
    s = str(v).strip()
    if not re.match(r"^\d{6}$", s):
        return None
    yy, mm, dd = int(s[0:2]), int(s[2:4]), int(s[4:6])
    if mm < 1 or mm > 12 or dd < 1 or dd > 31:
        return None
    try:
        return datetime.date(2000 + yy, mm, dd)
    except ValueError:
        return None


def parse_maturity_from_titulo(titulo):
    """El plazo/vencimiento de cada título se deriva de su propio código: los
    últimos 6 dígitos son la fecha de vencimiento en formato DDMMAA (ej.
    TFIT16281140 -> 28/11/2040)."""
    m = re.search(r"(\d{6})$", titulo)
    if not m:
        return None
    s = m.group(1)
    dd, mm, yy = int(s[0:2]), int(s[2:4]), int(s[4:6])
    if mm < 1 or mm > 12 or dd < 1 or dd > 31:
        return None
    try:
        return datetime.date(2000 + yy, mm, dd)
    except ValueError:
        return None


# ---------------- helpers de número ----------------
def parse_flexible_number(v):
    """Tolerante a los dos formatos vistos en los archivos de esta mesa:
    inglés (coma miles, punto decimal) y colombiano/español (punto miles,
    coma decimal)."""
    if isinstance(v, (int, float)):
        return float(v)
    if not isinstance(v, str):
        return None
    s = v.strip()
    if not s or s in ("–", "-", "—"):
        return None
    comma_count = s.count(",")
    dot_count = s.count(".")
    if comma_count and dot_count:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif comma_count > 1:
        s = s.replace(",", "")
    elif comma_count == 1:
        s = s.replace(",", ".")
    elif dot_count > 1:
        s = s.replace(".", "")
    try:
        return float(s)
    except ValueError:
        return None


def norm_decimal(v):
    """Precio y Tasa a veces llegan con el punto decimal perdido en el
    origen (100344 en vez de 100.344): los enteros grandes (>1000, algo que
    nunca pasa con un precio o tasa real de TES) se dividen entre 1000."""
    f = parse_flexible_number(v)
    if f is None:
        return None
    return f / 1000 if abs(f) > 1000 else f


def classify(titulo):
    if titulo.startswith("TUVT"):
        return "UVR"
    if titulo.startswith("TFIT"):
        return "Tasa Fija"
    return "Otro"


# ---------------- parseo de operaciones ----------------
def parse_operdia_format(rows, file_date):
    """Hora, Contraparte, Tipo, Título, Precio, Tasa, Nominal, Monto."""
    out = []
    for r in rows:
        if not r or len(r) < 8:
            continue
        tsec = time_from_cell(r[0])
        titulo = r[3]
        if tsec is None or not isinstance(titulo, str) or not titulo.strip():
            continue
        precio = norm_decimal(r[4])
        tasa = norm_decimal(r[5])
        nominal = parse_flexible_number(r[6])
        monto = parse_flexible_number(r[7])
        if precio is None or nominal is None:
            continue
        tipo_raw = str(r[2] or "")
        tl = tipo_raw.lower()
        if tl.startswith("compra"):
            tipo = "Compra"
        elif tl.startswith("venta"):
            tipo = "Venta"
        else:
            tipo = tipo_raw or None
        out.append({
            "date": file_date, "tsec": tsec, "tipo": tipo,
            "titulo": titulo.strip(), "precio": precio, "tasa": tasa,
            "nominal": nominal, "monto": monto,
        })
    return out


def parse_historico_mensual_format(rows):
    """Fecha Cierre, Hora Cierre, Sesión/Rueda, Plazo, Instrumento,
    Tasa/Precio, Tasa/Precio Equiv., VR. Nominal, Contravalor, ...
    Solo se toman las operaciones de rueda CONH."""
    out = []
    for r in rows:
        if not r or len(r) < 8:
            continue
        date = parse_fecha_cierre(r[0])
        if not date:
            continue
        rueda = str(r[2] or "").strip().upper()
        if rueda != "CONH":
            continue
        tsec = time_from_cell(r[1])
        titulo = r[4]
        if tsec is None or not isinstance(titulo, str) or not titulo.strip():
            continue
        v1 = norm_decimal(r[5])
        v2 = norm_decimal(r[6]) if len(r) > 6 else None
        if v1 is None:
            continue
        if v1 < 50 and v2 is not None and v2 >= 50:
            precio, tasa = v2, v1
        else:
            precio, tasa = v1, v2
        nominal = parse_flexible_number(r[7]) if len(r) > 7 else None
        monto = parse_flexible_number(r[8]) if len(r) > 8 else None
        if nominal is None:
            continue
        out.append({
            "date": date, "tsec": tsec, "tipo": None,
            "titulo": titulo.strip(), "precio": precio, "tasa": tasa,
            "nominal": nominal, "monto": monto,
        })
    return out


def looks_like_historico_mensual(rows):
    for r in rows:
        if not r or len(r) < 3:
            continue
        if parse_fecha_cierre(r[0]) and isinstance(r[2], str) and 3 <= len(r[2].strip()) <= 6:
            return True
    return False


def parse_workbook_trades(rows, file_date):
    if looks_like_historico_mensual(rows):
        return parse_historico_mensual_format(rows)
    return parse_operdia_format(rows, file_date)


# ---------------- parseo del boletín SEN ----------------
def parse_boletin_sen(rows, filename):
    """Boletín de cierres del sistema SEN. Estructura real: secciones
    marcadas con texto "RUEDA CONH" / "RUEDA SIML" (solo se toma CONH), cada
    una con sub-bloques "PESOS" (Tasa Fija) y "UVR".

    OJO -- bug real del boletín: en la sub-sección UVR, el encabezado rotula
    mal las columnas de tasa equivalente. Verificado con datos reales: la
    columna junto al PRECIO mínimo siempre trae la TASA máxima del día
    (precio bajo = tasa alta) en AMBAS secciones, así que se lee por
    posición de columna, no por el texto del encabezado.
    """
    date = parse_date_from_sen_filename(filename) or parse_date_from_filename(filename)
    out = []
    current_rueda = None
    for r in rows:
        if not r:
            continue
        joined = " ".join(str(c).upper() for c in r if isinstance(c, str))
        if re.search(r"RUEDA\s+CONH", joined):
            current_rueda = "CONH"
            continue
        if re.search(r"RUEDA\s+SIML", joined):
            current_rueda = "SIML"
            continue
        if current_rueda != "CONH":
            continue
        if len(r) < 16:
            continue
        raw_titulo = r[1]
        if not isinstance(raw_titulo, str) or not raw_titulo.strip():
            continue
        titulo = raw_titulo.strip().upper()
        if not re.match(r"^T[A-Z]{3}\d+", titulo):
            continue
        precio_cierre, tasa_cierre = r[14], r[15]
        if not isinstance(precio_cierre, (int, float)) or not isinstance(tasa_cierre, (int, float)):
            continue
        out.append({
            "date": date, "titulo": titulo,
            "nominal_trans": r[4] if isinstance(r[4], (int, float)) else None,
            "cant_cierres": r[5] if isinstance(r[5], (int, float)) else None,
            "precio_apertura": r[6], "tasa_apertura": r[7],
            "precio_min": r[8], "tasa_max": r[9],
            "precio_medio": r[10], "tasa_medio": r[11],
            "precio_max": r[12], "tasa_min": r[13],
            "precio_cierre": precio_cierre, "tasa_cierre": tasa_cierre,
        })
    return out


# ---------------- deduplicación ----------------
def dedupe_trades(trades):
    seen = {}
    for t in trades:
        key = (t["date"], t["tsec"], t["titulo"], t["precio"], t["nominal"])
        seen[key] = t
    return list(seen.values())


def dedupe_boletines(regs):
    seen = {}
    for b in regs:
        key = (b["date"], b["titulo"])
        seen[key] = b
    return list(seen.values())
