from fastapi.responses import JSONResponse, RedirectResponse, StreamingResponse
from fastapi import UploadFile, File, Form, Query, status
from database import SessionLocal, engine, Base
from fastapi.templating import Jinja2Templates
from fastapi import FastAPI, Request, Depends
from fastapi.staticfiles import StaticFiles
from fastapi import HTTPException
from sqlalchemy.orm import Session
from sqlalchemy import text  # SQL crudo
from datetime import datetime
from typing import Optional
from decimal import Decimal
from io import BytesIO
import pandas as pd
import models
import math
import crud
import re
import generador_cobros
from urllib.parse import parse_qs   # <--- para reconstruir filtros en /cliente/{id}

# ------------------- Inicialización -------------------
app = FastAPI()
templates = Jinja2Templates(directory="templates")

app.mount("/static", StaticFiles(directory="static"), name="static")
app.mount("/assets", StaticFiles(directory="assets"), name="assets")

# ------------------- Helpers -------------------
def fmt_money(value):
    """Filtro para mostrar números como moneda"""
    if value is None or value == "":
        return "-"
    try:
        v = Decimal(str(value))
        return f"{v:,.2f}"
    except Exception:
        return str(value)

templates.env.filters["fmt_money"] = fmt_money

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

def to_float(val):
    try:
        if pd.isna(val):
            return None
        val_str = str(val).replace(",", ".")
        val_str = re.sub(r"[^0-9.]", "", val_str)
        return float(val_str) if val_str else None
    except (ValueError, TypeError):
        return None

def clean_phone(val):
    """Normaliza teléfonos: quita .0, separadores y deja solo dígitos."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip()

    # 3046304674.0 -> 3046304674
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]

    # 3.0463E+09 -> 3046300000
    try:
        if re.fullmatch(r"\d+(\.\d+)?[eE][+-]?\d+", s):
            s = str(int(float(s)))
    except Exception:
        pass

    # 1234.00 -> 1234 si es entero exacto
    if re.fullmatch(r"\d+\.\d+", s):
        try:
            f = float(s)
            if f.is_integer():
                s = str(int(f))
        except Exception:
            pass

    # Dejar solo dígitos
    s = re.sub(r"[^\d]", "", s)
    return s or None

def clean_email(val):
    """Normaliza emails: recorta, minúsculas y valida forma simple."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    s = str(val).strip().lower()
    s = re.sub(r"\s+", "", s)
    if re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", s):
        return s
    return None

# ---- Null-safe parsers para MSSQL (evitan NaN/Timestamp) ----
def to_int_nullsafe(val):
    """Convierte a int o devuelve None si viene vacío/NaN."""
    try:
        if val is None:
            return None
        if isinstance(val, float) and math.isnan(val):  # NaN
            return None
        s = str(val).strip()
        if s == "":
            return None
        return int(float(s))  # admite "5.0"
    except Exception:
        return None

def to_dt_nullsafe(val):
    """Convierte a datetime (py) o None. Acepta pandas Timestamp, str, excel serial."""
    if val is None:
        return None
    try:
        if isinstance(val, float) and math.isnan(val):
            return None
    except Exception:
        pass
    ts = pd.to_datetime(val, errors="coerce")
    if pd.isna(ts):
        return None
    return ts.to_pydatetime()

# ---------- Helpers para borrado en SQL Server ----------
def _chunk(iterable, size):
    buf = []
    for x in iterable:
        buf.append(x)
        if len(buf) >= size:
            yield buf
            buf = []
    if buf:
        yield buf

def delete_by_pairs_mssql(db: Session, table_name: str, key_pairs, batch_size: int = 900):
    """
    MSSQL-safe:
    DELETE c
    FROM <tabla> AS c
    JOIN (VALUES (...), ...) AS v(nit_cliente, nro_docto_cruce)
      ON v.nit_cliente = c.nit_cliente
     AND v.nro_docto_cruce = c.nro_docto_cruce;
    """
    pairs = list(key_pairs)
    if not pairs:
        return

    for batch in _chunk(pairs, batch_size):
        values_clause = []
        params = {}
        for i, (nit, doc) in enumerate(batch):
            values_clause.append(f"(:nit{i}, :doc{i})")
            params[f"nit{i}"] = str(nit) if nit is not None else None
            params[f"doc{i}"] = str(doc) if doc is not None else None

        sql = f"""
        DELETE c
        FROM {table_name} AS c
        JOIN (VALUES {", ".join(values_clause)}) AS v(nit_cliente, nro_docto_cruce)
          ON v.nit_cliente = c.nit_cliente
         AND v.nro_docto_cruce = c.nro_docto_cruce;
        """
        db.execute(text(sql), params)

# ------------------- Exportar cartera a Excel -------------------
@app.get("/exportar_cartera.xlsx", name="exportar_cartera_xlsx")
def exportar_cartera_xlsx(
    db: Session = Depends(get_db),
    min_dias: Optional[str] = Query(None),
    max_dias: Optional[str] = Query(None),
    sort: Optional[str] = Query("dias_desc"),
    q: Optional[str] = Query(None),   # <--- filtro por nombre también en la exportación
):
    def to_int_or_none(val: Optional[str]):
        try:
            return int(val) if val not in (None, "") else None
        except (ValueError, TypeError):
            return None

    min_val = to_int_or_none(min_dias)
    max_val = to_int_or_none(max_dias)

    query = db.query(models.Cliente)
    if min_val is not None:
        query = query.filter(models.Cliente.dias_vencidos >= min_val)
    if max_val is not None:
        query = query.filter(models.Cliente.dias_vencidos <= max_val)

    # filtro por nombre
    if q:
        query = query.filter(models.Cliente.razon_social.ilike(f"%{q}%"))

    order_map = {
        "dias_desc": models.Cliente.dias_vencidos.desc(),
        "dias_asc": models.Cliente.dias_vencidos.asc(),
        "razon_asc": models.Cliente.razon_social.asc(),
        "nit_asc": models.Cliente.nit_cliente.asc(),
    }
    query = query.order_by(order_map.get(sort, order_map["dias_desc"]))
    clientes = query.all()

    rows = []
    for c in clientes:
        valor_docto = Decimal(str(c.valor_docto or 0))
        total_cop = Decimal(str(c.total_cop or 0))
        recaudo = Decimal(str(c.recaudo)) if c.recaudo is not None else (valor_docto - total_cop)

        # --- solo ÚLTIMA observación (si existe) ---
        obs_txt = ""
        if getattr(c, "observaciones", None):
            latest = None
            for o in c.observaciones:
                f = o.fecha_creacion or datetime.min
                if latest is None or f > (latest.fecha_creacion or datetime.min):
                    latest = o
            if latest:
                obs_txt = f"{(latest.fecha_creacion.strftime('%Y-%m-%d %H:%M') if latest.fecha_creacion else '')} - {latest.texto}"

        rows.append({
            "ID": c.id,
            "Razón social": c.razon_social,
            "NIT": c.nit_cliente,
            "Docto cruce": c.nro_docto_cruce,
            "Días vencidos": c.dias_vencidos,
            "Fecha docto": c.fecha_docto,
            "Fecha vcto": c.fecha_vcto,
            "Valor docto": float(valor_docto),
            "Total COP (saldo)": float(total_cop),
            "Recaudo": float(recaudo),
            "Teléfono": c.telefono,
            "Celular": c.celular,
            "Correo": c.correo,
            "Asesor": c.asesor,
            "Fecha gestión": c.fecha_gestion,
            "Tipo": c.tipo,
            "Observaciones": obs_txt,  # SOLO la última
        })

    df = pd.DataFrame(rows)

    for col in ["Fecha docto", "Fecha vcto", "Fecha gestión"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    if not df.empty:
        resumen = (
            df.groupby(["NIT", "Razón social"], dropna=False)
              .agg({
                  "Valor docto": "sum",
                  "Total COP (saldo)": "sum",
                  "Recaudo": "sum",
                  "Días vencidos": "max",
              })
              .rename(columns={
                  "Valor docto": "Total Valor Docto",
                  "Total COP (saldo)": "Saldo Total",
                  "Recaudo": "Recaudo Total",
                  "Días vencidos": "Max Días Vencidos",
              })
              .reset_index()
        )
        ref = (
            df.sort_values(["NIT"]).groupby(["NIT", "Razón social"], dropna=False)
              .agg({"Teléfono": "first", "Celular": "first", "Correo": "first", "Asesor": "first"})
              .reset_index()
        )
        resumen = resumen.merge(ref, on=["NIT", "Razón social"], how="left")

        facturas = (
            df.groupby(["NIT", "Razón social"], dropna=False)
              .size()
              .reset_index(name="# Facturas")
        )
        resumen = resumen.merge(facturas, on=["NIT", "Razón social"], how="left")
        resumen["# Facturas"] = resumen["# Facturas"].fillna(0).astype(int)
    else:
        resumen = pd.DataFrame(columns=[
            "NIT","Razón social","Total Valor Docto","Saldo Total","Recaudo Total",
            "Max Días Vencidos","Teléfono","Celular","Correo","Asesor","# Facturas"
        ])

    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter",
                        datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as writer:
        df.to_excel(writer, index=False, sheet_name="cartera")
        resumen.to_excel(writer, index=False, sheet_name="resumen")

        wb = writer.book
        ws1 = writer.sheets["cartera"]
        ws2 = writer.sheets["resumen"]

        for ws in (ws1, ws2):
            ws.freeze_panes(1, 0)

        fmt_header = wb.add_format({
            'bold': True, 'bg_color': '#D9E1F2', 'border': 1,
            'align': 'center', 'valign': 'vcenter'
        })
        fmt_money0 = wb.add_format({'num_format': '#,##0', 'border': 1})
        fmt_int = wb.add_format({'num_format': '#,##0', 'border': 1})
        fmt_date = wb.add_format({'num_format': 'yyyy-mm-dd', 'border': 1})
        fmt_text = wb.add_format({'text_wrap': True, 'border': 1, 'valign': 'top'})
        fmt_default = wb.add_format({'border': 1})

        def table_with_formats(ws, df_sheet, money_cols, date_cols, wrap_cols):
            rows, cols = df_sheet.shape
            headers = list(df_sheet.columns)

            per_col_format = {}
            for col in headers:
                if col in money_cols:
                    per_col_format[col] = fmt_money0
                elif col in date_cols:
                    per_col_format[col] = fmt_date
                elif col in wrap_cols:
                    per_col_format[col] = fmt_text
                elif df_sheet[col].dtype.kind in ("i", "u"):
                    per_col_format[col] = fmt_int
                else:
                    per_col_format[col] = fmt_default

            columns_def = [{'header': h, 'format': per_col_format[h]} for h in headers]

            if rows > 0:
                ws.add_table(0, 0, rows, cols - 1, {
                    'style': 'Table Style Medium 9',
                    'banded_rows': True,
                    'columns': columns_def
                })
            else:
                for c_idx, h in enumerate(headers):
                    ws.write(0, c_idx, h, fmt_header)

            col_widths = {}
            for c_idx, col in enumerate(headers):
                max_len = len(str(col)) + 2
                serie = df_sheet[col]

                if col in money_cols:
                    for v in serie.dropna():
                        try:
                            s = f"{int(round(float(v))):,}"
                            max_len = max(max_len, len(s))
                        except Exception:
                            pass
                elif col in date_cols:
                    max_len = max(max_len, 12)
                else:
                    for v in serie.dropna():
                        s = str(v).replace("\r", "")
                        max_len = max(max_len, max((len(seg) for seg in s.split("\n")), default=0))
                    if col in wrap_cols:
                        max_len = min(max(max_len, 40), 80)
                    else:
                        max_len = min(max_len, 40)

                width = max_len + 1
                col_widths[col] = width
                ws.set_column(c_idx, c_idx, width, per_col_format[col])

            return col_widths, per_col_format

        def autofit_row_heights(ws, df_sheet, wrap_cols, col_widths):
            if df_sheet.empty or not wrap_cols:
                return
            base_height = 15
            max_height = 300

            for r in range(1, len(df_sheet) + 1):
                max_lines = 1
                for col in wrap_cols:
                    if col not in df_sheet.columns:
                        continue
                    val = df_sheet.iloc[r - 1][col]
                    if pd.isna(val) or val is None:
                        continue
                    text = str(val).replace("\r", "")
                    if text == "":
                        continue

                    col_w = int(col_widths.get(col, 40))
                    chars_per_line = max(col_w - 2, 10)

                    total_lines = 0
                    for seg in text.split("\n"):
                        seg = seg.strip()
                        if seg == "":
                            total_lines += 1
                        else:
                            total_lines += math.ceil(len(seg) / chars_per_line)

                    max_lines = max(max_lines, total_lines)

                ws.set_row(r, min(base_height * max_lines + 4, max_height))

        money_cols_cartera = ["Valor docto", "Total COP (saldo)", "Recaudo"]
        date_cols_cartera = ["Fecha docto", "Fecha vcto", "Fecha gestión"]
        wrap_cols_cartera = ["Razón social", "Observaciones"]

        money_cols_resumen = ["Total Valor Docto", "Saldo Total", "Recaudo Total"]
        date_cols_resumen = []
        wrap_cols_resumen = ["Razón social"]

        widths_cartera, _ = table_with_formats(ws1, df, money_cols_cartera, date_cols_cartera, wrap_cols_cartera)
        widths_resumen, _ = table_with_formats(ws2, resumen, money_cols_resumen, date_cols_resumen, wrap_cols_resumen)

        autofit_row_heights(ws1, df, wrap_cols_cartera, widths_cartera)
        autofit_row_heights(ws2, resumen, wrap_cols_resumen, widths_resumen)

    output.seek(0)
    filename = f"cartera_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    return StreamingResponse(
        output,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f'attachment; filename=\"{filename}\"'}
    )

# ------------------- Importar cartera -------------------
@app.post("/importar_excel")
async def importar_excel(file: UploadFile = File(...), db: Session = Depends(get_db)):
    try:
        contents = await file.read()

        # --- Cargar Excel/CSV (ambos formatos) ---
        def _read_any(contents: bytes, filename: str):
            name = (filename or "").lower()
            if name.endswith(".csv"):
                df = pd.read_csv(BytesIO(contents), dtype=str, keep_default_na=False)
                return {"__csv__": df}
            return pd.read_excel(BytesIO(contents), sheet_name=None, dtype=str, keep_default_na=False)

        sheets = _read_any(contents, file.filename)

        # --- Normalizador simple de strings ---
        def norm(s: str) -> str:
            if s is None:
                return ""
            t = str(s).strip().lower()
            rep = {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ñ":"n","ä":"a","ë":"e","ï":"i","ö":"o","ü":"u","’":"'","´":"'","`":"'"}
            for a,b in rep.items(): t = t.replace(a,b)
            t = re.sub(r"\s+", " ", t)
            return t

        # --- Aliases (internos -> alias) ---
        alias_map = {
            "razon_social": {"razon social", "razon social vend. cliente", "razon social vend cliente"},
            "nit_cliente": {"nit", "nit cliente despacho", "nit cliente"},
            "nro_docto_cruce": {"docto cruce", "nro. docto. cruce", "nro docto cruce", "factura", "documento", "nro documento"},
            "fecha_docto": {"fecha docto", "fecha docto."},
            "fecha_vcto": {"fecha vcto", "fecha vcto."},
            "dias_vencidos": {"dias vencidos", "dias vencidos."},
            "valor_docto": {"valor docto", "valor docto."},
            "total_cop": {"total cop", "total cop (saldo)", "saldo", "total (cop)"},
            "telefono": {"telefono", "teléfono"},
            "celular": {"celular", "movil", "móvil"},
            "correo": {"correo", "email", "e-mail", "mail", "correo electronico", "correo electrónico"},
            "asesor": {"asesor", "vendedor", "razon social vend. cliente", "razon social vend cliente"},
            "obs_txt": {"observaciones"},
        }

        # --- Detectar si es export del proyecto (hoja 'cartera' con NIT/Docto) ---
        is_project_export = False
        df_project = None
        for sheet_name, df in sheets.items():
            if norm(sheet_name) == "cartera":
                df_project = df
                break
        if df_project is not None:
            cols_n = {norm(c) for c in df_project.columns}
            if "nit" in cols_n and ("docto cruce" in cols_n or "nro. docto. cruce" in cols_n or "nro docto cruce" in cols_n):
                is_project_export = True

        # --- Elegir hoja candidata ---
        if is_project_export:
            df_raw = df_project.copy()
        else:
            df_raw = None
            for _, df in sheets.items():
                cols_n = {norm(c) for c in df.columns}
                if any(x in cols_n for x in alias_map["nit_cliente"]) and \
                   any(x in cols_n for x in alias_map["nro_docto_cruce"]):
                    df_raw = df
                    break
            if df_raw is None:
                return RedirectResponse(
                    url="/?msg=No%20se%20encontro%20hoja%20compatible%20(ni%20cartera%20ni%20columnas%20clave)&msg_type=error",
                    status_code=303
                )

        # Guardamos copia de los encabezados originales por si necesitamos "rescatar" correo
        original_cols = list(df_raw.columns)

        df_raw.columns = [str(c).strip() for c in df_raw.columns]

        # --- Renombrado dinámico por alias ---
        rename_map_dynamic = {}
        for col in df_raw.columns:
            n = norm(col)
            for internal, aliases in alias_map.items():
                if n in aliases and internal not in rename_map_dynamic.values():
                    rename_map_dynamic[col] = internal
                    break

        # Encabezados exactos del export del proyecto
        project_head_map = {
            "Razón social": "razon_social",
            "NIT": "nit_cliente",
            "Docto cruce": "nro_docto_cruce",
            "Días vencidos": "dias_vencidos",
            "Fecha docto": "fecha_docto",
            "Fecha vcto": "fecha_vcto",
            "Valor docto": "valor_docto",
            "Total COP (saldo)": "total_cop",
            "Teléfono": "telefono",
            "Celular": "celular",
            "Correo": "correo",
            "Email": "correo",      # <---- soporte explícito
            "E-mail": "correo",     # <---- soporte explícito
            "Asesor": "asesor",
            "Observaciones": "obs_txt",
        }
        for col in df_raw.columns:
            if col in project_head_map and project_head_map[col] not in rename_map_dynamic.values():
                rename_map_dynamic[col] = project_head_map[col]

        df = df_raw.rename(columns=rename_map_dynamic)

        # --- Si no quedó 'correo' por renombrado, intentamos "rescatarlo" de columnas conocidas ---
        if "correo" not in df.columns:
            for cand in ["Correo", "Email", "E-mail", "email", "e-mail", "correo"]:
                if cand in original_cols:
                    df["correo"] = df_raw[cand]
                    break

        # Validar claves mínimas
        if "nit_cliente" not in df.columns or "nro_docto_cruce" not in df.columns:
            return RedirectResponse(
                url="/?msg=El%20archivo%20no%20tiene%20las%20columnas%20clave%20(NIT%20y%20Docto%20cruce)&msg_type=error",
                status_code=303
            )

        # --- Normalizaciones de campos ---
        if "telefono" in df.columns: df["telefono"] = df["telefono"].apply(clean_phone)
        if "celular" in df.columns: df["celular"] = df["celular"].apply(clean_phone)
        if "correo"  in df.columns: df["correo"]  = df["correo"].apply(clean_email)
        if "fecha_docto" in df.columns: df["fecha_docto"] = df["fecha_docto"].apply(to_dt_nullsafe)
        if "fecha_vcto"  in df.columns: df["fecha_vcto"]  = df["fecha_vcto"].apply(to_dt_nullsafe)
        if "dias_vencidos" in df.columns: df["dias_vencidos"] = df["dias_vencidos"].apply(to_int_nullsafe)

        def _to_money(x):
            v = to_float(x)
            return v if v is not None else 0.0

        if "valor_docto" in df.columns:
            df["valor_docto"] = df["valor_docto"].apply(_to_money)
        else:
            df["valor_docto"] = 0.0

        if "total_cop" in df.columns:
            df["total_cop"] = df["total_cop"].apply(_to_money)
        else:
            df["total_cop"] = df["valor_docto"]

        # Claves limpias
        df = df[(df["nit_cliente"].notna()) & (df["nro_docto_cruce"].notna())]
        df["nit_cliente"] = df["nit_cliente"].astype(str).str.strip()
        df["nro_docto_cruce"] = df["nro_docto_cruce"].astype(str).str.strip()
        df = df[(df["nit_cliente"] != "") & (df["nro_docto_cruce"] != "")]

        if df.empty:
            return RedirectResponse(url="/?msg=Archivo%20vacio%20o%20sin%20filas%20validas&msg_type=info", status_code=303)

        # --- Claves existentes en BD ---
        bd_clientes = db.query(models.Cliente).all()
        bd_claves = set((str(c.nit_cliente), str(c.nro_docto_cruce)) for c in bd_clientes)

        # --- Borrado condicional ---
        if not is_project_export:
            excel_claves = set(zip(df["nit_cliente"], df["nro_docto_cruce"]))
            claves_a_eliminar = bd_claves - excel_claves
            if claves_a_eliminar:
                delete_by_pairs_mssql(db, "cartera", claves_a_eliminar, batch_size=900)

        # Helpers de obs
        def _clean_obs_text(s: Optional[str]) -> Optional[str]:
            if not s:
                return None
            txt = str(s).strip()
            m = re.match(r"^\s*(\d{4}-\d{2}-\d{2}(?:\s+\d{2}:\d{2})?)\s*-\s*(.+)$", txt)
            if m:
                return m.group(2).strip() or None
            return txt or None

        def _latest_obs_text(cliente):
            latest = None
            for o in getattr(cliente, "observaciones", []) or []:
                f = getattr(o, "fecha_creacion", None) or datetime.min
                if latest is None or f > (getattr(latest, "fecha_creacion", None) or datetime.min):
                    latest = o
            return (latest.texto.strip() if latest and latest.texto else None)

        # --- Upsert + Observaciones (solo si viene del export) ---
        for _, row in df.iterrows():
            nit = row["nit_cliente"]
            docto = row["nro_docto_cruce"]

            valor_docto = row.get("valor_docto", 0.0) or 0.0
            total_excel = row.get("total_cop", valor_docto)
            recaudo_calc = round(max((valor_docto or 0.0) - (total_excel or 0.0), 0.0), 2)

            cliente = db.query(models.Cliente).filter_by(
                nit_cliente=nit,
                nro_docto_cruce=docto
            ).first()

            correo_val = clean_email(row.get("correo")) if "correo" in row else None

            if cliente:
                cliente.razon_social = row.get("razon_social")
                cliente.dias_vencidos = to_int_nullsafe(row.get("dias_vencidos"))
                cliente.fecha_docto = to_dt_nullsafe(row.get("fecha_docto"))
                cliente.fecha_vcto = to_dt_nullsafe(row.get("fecha_vcto"))
                cliente.valor_docto = valor_docto
                cliente.total_cop = total_excel
                cliente.recaudo = recaudo_calc
                cliente.telefono = clean_phone(row.get("telefono"))
                cliente.celular  = clean_phone(row.get("celular"))
                cliente.asesor = row.get("asesor")
                if correo_val:
                    cliente.correo = correo_val
            else:
                cliente = models.Cliente(
                    razon_social=row.get("razon_social"),
                    nit_cliente=nit,
                    nro_docto_cruce=docto,
                    dias_vencidos=to_int_nullsafe(row.get("dias_vencidos")),
                    fecha_docto=to_dt_nullsafe(row.get("fecha_docto")),
                    fecha_vcto=to_dt_nullsafe(row.get("fecha_vcto")),
                    valor_docto=valor_docto,
                    total_cop=total_excel,
                    recaudo=recaudo_calc,
                    telefono=clean_phone(row.get("telefono")),
                    celular=clean_phone(row.get("celular")),
                    asesor=row.get("asesor"),
                    correo=correo_val,
                )
                db.add(cliente)
                db.flush()

            if "obs_txt" in df.columns and is_project_export:
                new_obs_text = _clean_obs_text(row.get("obs_txt"))
                if new_obs_text:
                    last_text = _latest_obs_text(cliente)
                    if (last_text or "").strip() != new_obs_text.strip():
                        obs = models.Observacion(
                            cliente_id=cliente.id,
                            texto=new_obs_text,
                        )
                        db.add(obs)

        db.commit()
        return RedirectResponse(url="/?msg=Archivo%20subido%20correctamente&msg_type=success", status_code=303)

    except Exception as e:
        print("❌ Error importar_excel:", e)
        return RedirectResponse(url="/?msg=Error%20al%20subir%20el%20archivo&msg_type=error", status_code=303)

# ------------------- Observaciones -------------------
@app.post("/cliente/{cliente_id}/observacion")
def agregar_observacion(cliente_id: int, texto: str = Form(...), db: Session = Depends(get_db)):
    crud.add_observacion(db, cliente_id, texto)
    return RedirectResponse(
        url=f"/cliente/{cliente_id}?msg=Observaci%C3%B3n%20guardada&msg_type=success",
        status_code=303
    )

# ------------------- Actualizar cliente (editar recaudo) -------------------
@app.post("/cliente/{cliente_id}/update")
async def update_cliente(
    request: Request,
    cliente_id: int,
    db: Session = Depends(get_db)
):
    form_data = await request.form()

    cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
    if not cliente:
        return RedirectResponse(url="/?msg=Cliente%20no%20encontrado&msg_type=error", status_code=status.HTTP_303_SEE_OTHER)

    for campo in ["telefono", "celular", "correo", "fecha_gestion", "tipo", "asesor"]:
        val = form_data.get(campo)
        if val is not None:
            val_stripped = val.strip()
            
            # Si el usuario borró el texto, asignamos cadena vacía (o None para fechas)
            if val_stripped == "":
                if campo == "fecha_gestion":
                    setattr(cliente, campo, None)
                else:
                    setattr(cliente, campo, "")
            else:
                # Guardamos el dato aplicándole limpieza si es necesario
                if campo == "correo":
                    val_stripped = clean_email(val_stripped) or ""
                elif campo in ("telefono", "celular"):
                    val_stripped = clean_phone(val_stripped) or ""
                
                setattr(cliente, campo, val_stripped)

    valor_docto = float(cliente.valor_docto or 0.0)

    raw_recaudo = form_data.get("recaudo")
    raw_total_cop = form_data.get("total_cop")

    has_recaudo = raw_recaudo is not None and raw_recaudo.strip() != ""
    has_total_cop = raw_total_cop is not None and raw_total_cop.strip() != ""

    if has_recaudo:
        delta = to_float(raw_recaudo)
        if delta is not None:
            old_total = float(cliente.total_cop or 0.0)
            new_total = max(round(old_total - delta, 2), 0.0)
            cliente.total_cop = new_total
            cliente.recaudo = round(max(valor_docto - new_total, 0.0), 2)
    elif has_total_cop:
        nuevo_total = to_float(raw_total_cop)
        if nuevo_total is not None:
            nuevo_total = max(nuevo_total, 0.0)
            if nuevo_total > valor_docto:
                nuevo_total = valor_docto
            cliente.total_cop = round(nuevo_total, 2)
            cliente.recaudo = round(max(valor_docto - nuevo_total, 0.0), 2)

    nueva_obs = form_data.get("observaciones")
    if nueva_obs and nueva_obs.strip():
        obs = models.Observacion(
            cliente_id=cliente.id,
            texto=nueva_obs.strip(),
        )
        db.add(obs)

    db.commit()
    return RedirectResponse(
        url=f"/cliente/{cliente_id}?msg=Cambios%20guardados&msg_type=success",
        status_code=status.HTTP_303_SEE_OTHER
    )

# ------------------- historial cliente -------------------
@app.get("/cliente/{cliente_id}/historial")
def historial_cliente(cliente_id: int, db: Session = Depends(get_db)):
    cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
    if not cliente:
        return JSONResponse(content=[], status_code=200)

    historial = [
        {
            "texto": obs.texto,
            "fecha": obs.fecha_creacion.strftime("%d/%m/%Y %H:%M")
            if obs.fecha_creacion else None
        }
        for obs in cliente.observaciones
    ]

    return JSONResponse(content=historial, status_code=200)

# ------------------- Vista cliente -------------------
@app.get("/cliente/{cliente_id}")
def ver_cliente(cliente_id: int, request: Request, db: Session = Depends(get_db)):
    cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
    if not cliente:
        raise HTTPException(status_code=404, detail="Cliente no encontrado")

    # Persistir filtros al regresar
    ref_qs = request.query_params.get("ref")
    back_url = f"/?{ref_qs}" if ref_qs else "/"

    # === Navegación (prev/next) respetando filtros del index ===
    prev_id = None
    next_id = None

    try:
        if ref_qs:
            params_raw = parse_qs(ref_qs)
            params = {k: v[0] for k, v in params_raw.items() if v}

            view = params.get("view")
            min_dias = params.get("min_dias")
            max_dias = params.get("max_dias")
            sort = params.get("sort") or "dias_desc"
            q = params.get("q")

            def to_int_or_none(val: Optional[str]):
                try:
                    return int(val) if val not in (None, "") else None
                except (ValueError, TypeError):
                    return None

            min_val = to_int_or_none(min_dias)
            max_val = to_int_or_none(max_dias)

            facturas_query = db.query(models.Cliente)

            if min_val is not None:
                facturas_query = facturas_query.filter(models.Cliente.dias_vencidos >= min_val)
            if max_val is not None:
                facturas_query = facturas_query.filter(models.Cliente.dias_vencidos <= max_val)

            if q:
                facturas_query = facturas_query.filter(models.Cliente.razon_social.ilike(f"%{q}%"))

            order_map = {
                "dias_desc": models.Cliente.dias_vencidos.desc(),
                "dias_asc": models.Cliente.dias_vencidos.asc(),
                "razon_asc": models.Cliente.razon_social.asc(),
                "nit_asc": models.Cliente.nit_cliente.asc(),
            }
            facturas_query = facturas_query.order_by(order_map.get(sort, order_map["dias_desc"]))

            ids = [row[0] for row in facturas_query.with_entities(models.Cliente.id).all()]

            if cliente_id in ids:
                idx = ids.index(cliente_id)
                if idx > 0:
                    prev_id = ids[idx - 1]
                if idx < len(ids) - 1:
                    next_id = ids[idx + 1]
        else:
            # Sin filtros en ref: comportamiento “clásico” por id
            prev_cliente = (
                db.query(models.Cliente)
                .filter(models.Cliente.id < cliente_id)
                .order_by(models.Cliente.id.desc())
                .first()
            )
            next_cliente = (
                db.query(models.Cliente)
                .filter(models.Cliente.id > cliente_id)
                .order_by(models.Cliente.id.asc())
                .first()
            )
            prev_id = prev_cliente.id if prev_cliente else None
            next_id = next_cliente.id if next_cliente else None

    except Exception as _e:
        # Si algo falla en la reconstrucción de filtros, caemos al modo clásico
        prev_cliente = (
            db.query(models.Cliente)
            .filter(models.Cliente.id < cliente_id)
            .order_by(models.Cliente.id.desc())
            .first()
        )
        next_cliente = (
            db.query(models.Cliente)
            .filter(models.Cliente.id > cliente_id)
            .order_by(models.Cliente.id.asc())
            .first()
        )
        prev_id = prev_cliente.id if prev_cliente else None
        next_id = next_cliente.id if next_cliente else None

    return templates.TemplateResponse(
        "cliente.html",
        {
            "request": request,
            "cliente": cliente,
            "prev_id": prev_id,
            "next_id": next_id,
            "flash_msg": request.query_params.get("msg"),
            "flash_type": request.query_params.get("msg_type"),
            "back_url": back_url,
        }
    )

# ------------------- Index agrupado -------------------
@app.get("/")
def index(
    request: Request,
    db: Session = Depends(get_db),
    view: Optional[str] = Query(None),
    min_dias: Optional[str] = Query(None),
    max_dias: Optional[str] = Query(None),
    sort: Optional[str] = Query("dias_desc"),
    q: Optional[str] = Query(None),   # <--- buscador por nombre
):
    def to_int_or_none(val: Optional[str]):
        try:
            return int(val) if val not in (None, "") else None
        except (ValueError, TypeError):
            return None

    min_val = to_int_or_none(min_dias)
    max_val = to_int_or_none(max_dias)

    query = db.query(models.Cliente)

    if min_val is not None:
        query = query.filter(models.Cliente.dias_vencidos >= min_val)
    if max_val is not None:
        query = query.filter(models.Cliente.dias_vencidos <= max_val)

    # Filtro por nombre (razón social)
    if q:
        query = query.filter(models.Cliente.razon_social.ilike(f"%{q}%"))

    order_map = {
        "dias_desc": models.Cliente.dias_vencidos.desc(),
        "dias_asc": models.Cliente.dias_vencidos.asc(),
        "razon_asc": models.Cliente.razon_social.asc(),
        "nit_asc": models.Cliente.nit_cliente.asc(),
    }
    query = query.order_by(order_map.get(sort, order_map["dias_desc"]))

    clientes = query.all()

    # Flash + QS actual para persistir filtros en links
    flash_msg = request.query_params.get("msg")
    flash_type = request.query_params.get("msg_type")
    if (min_val is not None) or (max_val is not None) or q:
        if not flash_msg:
            flash_msg = "Filtros aplicados"
        if not flash_type:
            flash_type = "info"

    current_qs = request.url.query  # <- se pasa a plantillas para back & ver

    if view == "flat":
        filas = []
        for c in clientes:
            valor_docto = Decimal(str(c.valor_docto or 0))
            total_cop = Decimal(str(c.total_cop or 0))
            recaudo = valor_docto - total_cop

            filas.append({
                "id": c.id,
                "razon_social": c.razon_social,
                "nit_cliente": c.nit_cliente,
                "nro_docto_cruce": c.nro_docto_cruce,
                "dias_vencidos": c.dias_vencidos,
                "fecha_docto": c.fecha_docto,
                "fecha_vcto": c.fecha_vcto,
                "valor_docto": float(valor_docto),
                "total_cop": float(total_cop),
                "recaudo": float(recaudo),
                "asesor": c.asesor,
                "correo": c.correo,
            })

        return templates.TemplateResponse(
            "index.html",
            {
                "request": request,
                "view": "flat",
                "filas": filas,
                "flash_msg": flash_msg,
                "flash_type": flash_type,
                "current_qs": current_qs,
                "min_dias": min_dias,
                "max_dias": max_dias,
                "sort": sort,
                "q": q,
            }
        )

    # Vista agrupada
    agrupados = {}
    for c in clientes:
        if c.nit_cliente not in agrupados:
            agrupados[c.nit_cliente] = {
                "nit_cliente": c.nit_cliente,
                "razon_social": c.razon_social,
                "telefono": c.telefono,
                "celular": c.celular,
                "correo": c.correo,
                "asesor": c.asesor,
                "facturas": []
            }

        valor_docto = Decimal(str(c.valor_docto or 0))
        total_cop = Decimal(str(c.total_cop or 0))
        recaudo = valor_docto - total_cop

        agrupados[c.nit_cliente]["facturas"].append({
            "id": c.id,
            "nro_docto_cruce": c.nro_docto_cruce,
            "dias_vencidos": c.dias_vencidos,
            "fecha_docto": c.fecha_docto,
            "fecha_vcto": c.fecha_vcto,
            "valor_docto": float(valor_docto),
            "total_cop": float(total_cop),
            "recaudo": float(recaudo),
            "fecha_gestion": c.fecha_gestion,
            "tipo": c.tipo,
            "asesor": c.asesor,
            "observaciones": [obs.texto for obs in c.observaciones]
        })

    for cl in agrupados.values():
        dias = [f["dias_vencidos"] for f in cl["facturas"] if f["dias_vencidos"] is not None]
        cl["max_dias"] = max(dias) if dias else None

    return templates.TemplateResponse(
        "index.html",
        {
            "request": request,
            "clientes": list(agrupados.values()),
            "flash_msg": flash_msg,
            "flash_type": flash_type,
            "current_qs": current_qs,
            "min_dias": min_dias,
            "max_dias": max_dias,
            "sort": sort,
            "view": view,
            "q": q,
        }
    )




# from fastapi.responses import JSONResponse, RedirectResponse, StreamingResponse
# from fastapi import UploadFile, File, Form, Query, status
# from database import SessionLocal, engine, Base
# from fastapi.templating import Jinja2Templates
# from fastapi import FastAPI, Request, Depends
# from fastapi.staticfiles import StaticFiles
# from fastapi import HTTPException
# from sqlalchemy.orm import Session
# from sqlalchemy import text  # SQL crudo
# from datetime import datetime
# from typing import Optional
# from decimal import Decimal
# from io import BytesIO
# from urllib.parse import parse_qsl  # <-- agrega esto arriba junto a los demás imports
# import pandas as pd
# import models
# import math
# import crud
# import re

# # ------------------- Inicialización -------------------
# app = FastAPI()
# templates = Jinja2Templates(directory="templates")

# app.mount("/static", StaticFiles(directory="static"), name="static")
# app.mount("/assets", StaticFiles(directory="assets"), name="assets")

# # ------------------- Helpers -------------------
# def fmt_money(value):
#     """Filtro para mostrar números como moneda"""
#     if value is None or value == "":
#         return "-"
#     try:
#         v = Decimal(str(value))
#         return f"{v:,.2f}"
#     except Exception:
#         return str(value)

# templates.env.filters["fmt_money"] = fmt_money

# def get_db():
#     db = SessionLocal()
#     try:
#         yield db
#     finally:
#         db.close()

# def to_float(val):
#     try:
#         if pd.isna(val):
#             return None
#         val_str = str(val).replace(",", ".")
#         val_str = re.sub(r"[^0-9.]", "", val_str)
#         return float(val_str) if val_str else None
#     except (ValueError, TypeError):
#         return None

# def clean_phone(val):
#     """Normaliza teléfonos: quita .0, separadores y deja solo dígitos."""
#     if val is None or (isinstance(val, float) and pd.isna(val)):
#         return None
#     s = str(val).strip()

#     # 3046304674.0 -> 3046304674
#     if s.endswith(".0") and s[:-2].isdigit():
#         s = s[:-2]

#     # 3.0463E+09 -> 3046300000
#     try:
#         if re.fullmatch(r"\d+(\.\d+)?[eE][+-]?\d+", s):
#             s = str(int(float(s)))
#     except Exception:
#         pass

#     # 1234.00 -> 1234 si es entero exacto
#     if re.fullmatch(r"\d+\.\d+", s):
#         try:
#             f = float(s)
#             if f.is_integer():
#                 s = str(int(f))
#         except Exception:
#             pass

#     # Dejar solo dígitos
#     s = re.sub(r"[^\d]", "", s)
#     return s or None

# def clean_email(val):
#     """Normaliza emails: recorta, minúsculas y valida forma simple."""
#     if val is None or (isinstance(val, float) and pd.isna(val)):
#         return None
#     s = str(val).strip().lower()
#     s = re.sub(r"\s+", "", s)
#     if re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", s):
#         return s
#     return None

# # ---- Null-safe parsers para MSSQL (evitan NaN/Timestamp) ----
# def to_int_nullsafe(val):
#     """Convierte a int o devuelve None si viene vacío/NaN."""
#     try:
#         if val is None:
#             return None
#         if isinstance(val, float) and math.isnan(val):  # NaN
#             return None
#         s = str(val).strip()
#         if s == "":
#             return None
#         return int(float(s))  # admite "5.0"
#     except Exception:
#         return None

# def to_dt_nullsafe(val):
#     """Convierte a datetime (py) o None. Acepta pandas Timestamp, str, excel serial."""
#     if val is None:
#         return None
#     try:
#         if isinstance(val, float) and math.isnan(val):
#             return None
#     except Exception:
#         pass
#     ts = pd.to_datetime(val, errors="coerce")
#     if pd.isna(ts):
#         return None
#     return ts.to_pydatetime()

# # ---------- Helpers para borrado en SQL Server ----------
# def _chunk(iterable, size):
#     buf = []
#     for x in iterable:
#         buf.append(x)
#         if len(buf) >= size:
#             yield buf
#             buf = []
#     if buf:
#         yield buf

# def delete_by_pairs_mssql(db: Session, table_name: str, key_pairs, batch_size: int = 900):
#     """
#     MSSQL-safe:
#     DELETE c
#     FROM <tabla> AS c
#     JOIN (VALUES (...), ...) AS v(nit_cliente, nro_docto_cruce)
#       ON v.nit_cliente = c.nit_cliente
#      AND v.nro_docto_cruce = c.nro_docto_cruce;
#     """
#     pairs = list(key_pairs)
#     if not pairs:
#         return

#     for batch in _chunk(pairs, batch_size):
#         values_clause = []
#         params = {}
#         for i, (nit, doc) in enumerate(batch):
#             values_clause.append(f"(:nit{i}, :doc{i})")
#             params[f"nit{i}"] = str(nit) if nit is not None else None
#             params[f"doc{i}"] = str(doc) if doc is not None else None

#         sql = f"""
#         DELETE c
#         FROM {table_name} AS c
#         JOIN (VALUES {", ".join(values_clause)}) AS v(nit_cliente, nro_docto_cruce)
#           ON v.nit_cliente = c.nit_cliente
#          AND v.nro_docto_cruce = c.nro_docto_cruce;
#         """
#         db.execute(text(sql), params)

# # ------------------- Exportar cartera a Excel -------------------
# @app.get("/exportar_cartera.xlsx", name="exportar_cartera_xlsx")
# def exportar_cartera_xlsx(
#     db: Session = Depends(get_db),
#     min_dias: Optional[str] = Query(None),
#     max_dias: Optional[str] = Query(None),
#     sort: Optional[str] = Query("dias_desc"),
#     q: Optional[str] = Query(None),   # <-- BUSCADOR POR NOMBRE TAMBIÉN EN EXPORT
# ):
#     def to_int_or_none(val: Optional[str]):
#         try:
#             return int(val) if val not in (None, "") else None
#         except (ValueError, TypeError):
#             return None

#     min_val = to_int_or_none(min_dias)
#     max_val = to_int_or_none(max_dias)

#     query = db.query(models.Cliente)

#     # Filtro por nombre (razón social)
#     if q:
#         query = query.filter(models.Cliente.razon_social.ilike(f"%{q}%"))

#     if min_val is not None:
#         query = query.filter(models.Cliente.dias_vencidos >= min_val)
#     if max_val is not None:
#         query = query.filter(models.Cliente.dias_vencidos <= max_val)

#     # incluye razon_asc para coincidir con el template
#     order_map = {
#         "dias_desc": models.Cliente.dias_vencidos.desc(),
#         "dias_asc": models.Cliente.dias_vencidos.asc(),
#         "razon_asc": models.Cliente.razon_social.asc(),
#     }
#     query = query.order_by(order_map.get(sort, order_map["dias_desc"]))
#     clientes = query.all()

#     rows = []
#     for c in clientes:
#         valor_docto = Decimal(str(c.valor_docto or 0))
#         total_cop = Decimal(str(c.total_cop or 0))
#         recaudo = Decimal(str(c.recaudo)) if c.recaudo is not None else (valor_docto - total_cop)

#         # --- solo ÚLTIMA observación (si existe) ---
#         obs_txt = ""
#         if getattr(c, "observaciones", None):
#             latest = None
#             for o in c.observaciones:
#                 f = o.fecha_creacion or datetime.min
#                 if latest is None or f > (latest.fecha_creacion or datetime.min):
#                     latest = o
#             if latest:
#                 obs_txt = f"{(latest.fecha_creacion.strftime('%Y-%m-%d %H:%M') if latest.fecha_creacion else '')} - {latest.texto}"

#         rows.append({
#             "ID": c.id,
#             "Razón social": c.razon_social,
#             "NIT": c.nit_cliente,
#             "Docto cruce": c.nro_docto_cruce,
#             "Días vencidos": c.dias_vencidos,
#             "Fecha docto": c.fecha_docto,
#             "Fecha vcto": c.fecha_vcto,
#             "Valor docto": float(valor_docto),
#             "Total COP (saldo)": float(total_cop),
#             "Recaudo": float(recaudo),
#             "Teléfono": c.telefono,
#             "Celular": c.celular,
#             "Correo": c.correo,
#             "Asesor": c.asesor,
#             "Fecha gestión": c.fecha_gestion,
#             "Tipo": c.tipo,
#             "Observaciones": obs_txt,  # SOLO la última
#         })

#     df = pd.DataFrame(rows)

#     for col in ["Fecha docto", "Fecha vcto", "Fecha gestión"]:
#         if col in df.columns:
#             df[col] = pd.to_datetime(df[col], errors="coerce")

#     if not df.empty:
#         resumen = (
#             df.groupby(["NIT", "Razón social"], dropna=False)
#               .agg({
#                   "Valor docto": "sum",
#                   "Total COP (saldo)": "sum",
#                   "Recaudo": "sum",
#                   "Días vencidos": "max",
#               })
#               .rename(columns={
#                   "Valor docto": "Total Valor Docto",
#                   "Total COP (saldo)": "Saldo Total",
#                   "Recaudo": "Recaudo Total",
#                   "Días vencidos": "Max Días Vencidos",
#               })
#               .reset_index()
#         )
#         ref = (
#             df.sort_values(["NIT"]).groupby(["NIT", "Razón social"], dropna=False)
#               .agg({"Teléfono": "first", "Celular": "first", "Correo": "first", "Asesor": "first"})
#               .reset_index()
#         )
#         resumen = resumen.merge(ref, on=["NIT", "Razón social"], how="left")

#         facturas = (
#             df.groupby(["NIT", "Razón social"], dropna=False)
#               .size()
#               .reset_index(name="# Facturas")
#         )
#         resumen = resumen.merge(facturas, on=["NIT", "Razón social"], how="left")
#         resumen["# Facturas"] = resumen["# Facturas"].fillna(0).astype(int)
#     else:
#         resumen = pd.DataFrame(columns=[
#             "NIT","Razón social","Total Valor Docto","Saldo Total","Recaudo Total",
#             "Max Días Vencidos","Teléfono","Celular","Correo","Asesor","# Facturas"
#         ])

#     output = BytesIO()
#     with pd.ExcelWriter(output, engine="xlsxwriter",
#                         datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as writer:
#         df.to_excel(writer, index=False, sheet_name="cartera")
#         resumen.to_excel(writer, index=False, sheet_name="resumen")

#         wb = writer.book
#         ws1 = writer.sheets["cartera"]
#         ws2 = writer.sheets["resumen"]

#         for ws in (ws1, ws2):
#             ws.freeze_panes(1, 0)

#         fmt_header = wb.add_format({
#             'bold': True, 'bg_color': '#D9E1F2', 'border': 1,
#             'align': 'center', 'valign': 'vcenter'
#         })
#         fmt_money0 = wb.add_format({'num_format': '#,##0', 'border': 1})
#         fmt_int = wb.add_format({'num_format': '#,##0', 'border': 1})
#         fmt_date = wb.add_format({'num_format': 'yyyy-mm-dd', 'border': 1})
#         fmt_text = wb.add_format({'text_wrap': True, 'border': 1, 'valign': 'top'})
#         fmt_default = wb.add_format({'border': 1})

#         def table_with_formats(ws, df_sheet, money_cols, date_cols, wrap_cols):
#             rows, cols = df_sheet.shape
#             headers = list(df_sheet.columns)

#             per_col_format = {}
#             for col in headers:
#                 if col in money_cols:
#                     per_col_format[col] = fmt_money0
#                 elif col in date_cols:
#                     per_col_format[col] = fmt_date
#                 elif col in wrap_cols:
#                     per_col_format[col] = fmt_text
#                 elif df_sheet[col].dtype.kind in ("i", "u"):
#                     per_col_format[col] = fmt_int
#                 else:
#                     per_col_format[col] = fmt_default

#             columns_def = [{'header': h, 'format': per_col_format[h]} for h in headers]

#             if rows > 0:
#                 ws.add_table(0, 0, rows, cols - 1, {
#                     'style': 'Table Style Medium 9',
#                     'banded_rows': True,
#                     'columns': columns_def
#                 })
#             else:
#                 for c_idx, h in enumerate(headers):
#                     ws.write(0, c_idx, h, fmt_header)

#             col_widths = {}
#             for c_idx, col in enumerate(headers):
#                 max_len = len(str(col)) + 2
#                 serie = df_sheet[col]

#                 if col in money_cols:
#                     for v in serie.dropna():
#                         try:
#                             s = f"{int(round(float(v))):,}"
#                             max_len = max(max_len, len(s))
#                         except Exception:
#                             pass
#                 elif col in date_cols:
#                     max_len = max(max_len, 12)
#                 else:
#                     for v in serie.dropna():
#                         s = str(v).replace("\r", "")
#                         max_len = max(max_len, max((len(seg) for seg in s.split("\n")), default=0))
#                     if col in wrap_cols:
#                         max_len = min(max(max_len, 40), 80)
#                     else:
#                         max_len = min(max_len, 40)

#                 width = max_len + 1
#                 col_widths[col] = width
#                 ws.set_column(c_idx, c_idx, width, per_col_format[col])

#             return col_widths, per_col_format

#         def autofit_row_heights(ws, df_sheet, wrap_cols, col_widths):
#             if df_sheet.empty or not wrap_cols:
#                 return
#             base_height = 15
#             max_height = 300

#             for r in range(1, len(df_sheet) + 1):
#                 max_lines = 1
#                 for col in wrap_cols:
#                     if col not in df_sheet.columns:
#                         continue
#                     val = df_sheet.iloc[r - 1][col]
#                     if pd.isna(val) or val is None:
#                         continue
#                     text = str(val).replace("\r", "")
#                     if text == "":
#                         continue

#                     col_w = int(col_widths.get(col, 40))
#                     chars_per_line = max(col_w - 2, 10)

#                     total_lines = 0
#                     for seg in text.split("\n"):
#                         seg = seg.strip()
#                         if seg == "":
#                             total_lines += 1
#                         else:
#                             total_lines += math.ceil(len(seg) / chars_per_line)

#                     max_lines = max(max_lines, total_lines)

#                 ws.set_row(r, min(base_height * max_lines + 4, max_height))

#         money_cols_cartera = ["Valor docto", "Total COP (saldo)", "Recaudo"]
#         date_cols_cartera = ["Fecha docto", "Fecha vcto", "Fecha gestión"]
#         wrap_cols_cartera = ["Razón social", "Observaciones"]

#         money_cols_resumen = ["Total Valor Docto", "Saldo Total", "Recaudo Total"]
#         date_cols_resumen = []
#         wrap_cols_resumen = ["Razón social"]

#         widths_cartera, _ = table_with_formats(ws1, df, money_cols_cartera, date_cols_cartera, wrap_cols_cartera)
#         widths_resumen, _ = table_with_formats(ws2, resumen, money_cols_resumen, date_cols_resumen, wrap_cols_resumen)

#         autofit_row_heights(ws1, df, wrap_cols_cartera, widths_cartera)
#         autofit_row_heights(ws2, resumen, wrap_cols_resumen, widths_resumen)

#     output.seek(0)
#     filename = f"cartera_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
#     return StreamingResponse(
#         output,
#         media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#         headers={"Content-Disposition": f'attachment; filename="{filename}"'}
#     )

# # ------------------- Importar cartera -------------------
# @app.post("/importar_excel")
# async def importar_excel(file: UploadFile = File(...), db: Session = Depends(get_db)):
#     try:
#         contents = await file.read()

#         # --- Cargar Excel/CSV (ambos formatos) ---
#         def _read_any(contents: bytes, filename: str):
#             name = (filename or "").lower()
#             if name.endswith(".csv"):
#                 df = pd.read_csv(BytesIO(contents), dtype=str, keep_default_na=False)
#                 return {"__csv__": df}
#             return pd.read_excel(BytesIO(contents), sheet_name=None, dtype=str, keep_default_na=False)

#         sheets = _read_any(contents, file.filename)

#         # --- Normalizador simple de strings ---
#         def norm(s: str) -> str:
#             if s is None:
#                 return ""
#             t = str(s).strip().lower()
#             rep = {
#                 "á":"a","é":"e","í":"i","ó":"o","ú":"u","ñ":"n",
#                 "ä":"a","ë":"e","ï":"i","ö":"o","ü":"u",
#                 "’":"'","´":"'","`":"'"
#             }
#             for a,b in rep.items():
#                 t = t.replace(a,b)
#             t = re.sub(r"\s+", " ", t)
#             return t

#         # --- Aliases (internos -> alias) ---
#         alias_map = {
#             "razon_social": {"razon social", "razon social vend. cliente", "razon social vend cliente"},
#             "nit_cliente": {"nit", "nit cliente despacho", "nit cliente"},
#             "nro_docto_cruce": {"docto cruce", "nro. docto. cruce", "nro docto cruce", "factura", "documento", "nro documento"},
#             "fecha_docto": {"fecha docto", "fecha docto."},
#             "fecha_vcto": {"fecha vcto", "fecha vcto."},
#             "dias_vencidos": {"dias vencidos", "dias vencidos."},
#             "valor_docto": {"valor docto", "valor docto."},
#             "total_cop": {"total cop", "total cop (saldo)", "saldo", "total (cop)"},
#             "telefono": {"telefono", "teléfono"},
#             "celular": {"celular", "movil", "móvil"},
#             "correo": {"correo", "email", "e-mail", "mail", "correo electronico", "correo electrónico"},
#             "asesor": {"asesor", "vendedor", "razon social vend. cliente", "razon social vend cliente"},
#             "obs_txt": {"observaciones"},
#         }

#         # --- Detectar si es export del proyecto (hoja 'cartera' con NIT/Docto) ---
#         is_project_export = False
#         df_project = None
#         for sheet_name, df in sheets.items():
#             if norm(sheet_name) == "cartera":
#                 df_project = df
#                 break
#         if df_project is not None:
#             cols_n = {norm(c) for c in df_project.columns}
#             if "nit" in cols_n and ("docto cruce" in cols_n or "nro. docto. cruce" in cols_n or "nro docto cruce" in cols_n):
#                 is_project_export = True

#         # --- Elegir hoja candidata ---
#         if is_project_export:
#             df_raw = df_project.copy()
#         else:
#             df_raw = None
#             for _, df in sheets.items():
#                 cols_n = {norm(c) for c in df.columns}
#                 if any(x in cols_n for x in alias_map["nit_cliente"]) and \
#                    any(x in cols_n for x in alias_map["nro_docto_cruce"]):
#                     df_raw = df
#                     break
#             if df_raw is None:
#                 return RedirectResponse(
#                     url="/?msg=No%20se%20encontro%20hoja%20compatible%20(ni%20cartera%20ni%20columnas%20clave)&msg_type=error",
#                     status_code=303
#                 )

#         # Guardamos copia de los encabezados originales por si necesitamos "rescatar" correo
#         original_cols = list(df_raw.columns)

#         df_raw.columns = [str(c).strip() for c in df_raw.columns]

#         # --- Renombrado dinámico por alias ---
#         rename_map_dynamic = {}
#         for col in df_raw.columns:
#             n = norm(col)
#             for internal, aliases in alias_map.items():
#                 if n in aliases and internal not in rename_map_dynamic.values():
#                     rename_map_dynamic[col] = internal
#                     break

#         # Encabezados exactos del export del proyecto
#         project_head_map = {
#             "Razón social": "razon_social",
#             "NIT": "nit_cliente",
#             "Docto cruce": "nro_docto_cruce",
#             "Días vencidos": "dias_vencidos",
#             "Fecha docto": "fecha_docto",
#             "Fecha vcto": "fecha_vcto",
#             "Valor docto": "valor_docto",
#             "Total COP (saldo)": "total_cop",
#             "Teléfono": "telefono",
#             "Celular": "celular",
#             "Correo": "correo",
#             "Email": "correo",      # <---- soporte explícito
#             "E-mail": "correo",     # <---- soporte explícito
#             "Asesor": "asesor",
#             "Observaciones": "obs_txt",
#         }
#         for col in df_raw.columns:
#             if col in project_head_map and project_head_map[col] not in rename_map_dynamic.values():
#                 rename_map_dynamic[col] = project_head_map[col]

#         df = df_raw.rename(columns=rename_map_dynamic)

#         # --- Si no quedó 'correo' por renombrado, intentamos "rescatarlo" de columnas conocidas ---
#         if "correo" not in df.columns:
#             for cand in ["Correo", "Email", "E-mail", "email", "e-mail", "correo"]:
#                 if cand in original_cols:
#                     df["correo"] = df_raw[cand]
#                     break

#         # Validar claves mínimas
#         if "nit_cliente" not in df.columns or "nro_docto_cruce" not in df.columns:
#             return RedirectResponse(
#                 url="/?msg=El%20archivo%20no%20tiene%20las%20columnas%20clave%20(NIT%20y%20Docto%20cruce)&msg_type=error",
#                 status_code=303
#             )

#         # --- Normalizaciones de campos ---
#         if "telefono" in df.columns: df["telefono"] = df["telefono"].apply(clean_phone)
#         if "celular" in df.columns: df["celular"] = df["celular"].apply(clean_phone)
#         if "correo"  in df.columns: df["correo"]  = df["correo"].apply(clean_email)

#         if "fecha_docto" in df.columns: df["fecha_docto"] = df["fecha_docto"].apply(to_dt_nullsafe)
#         if "fecha_vcto"  in df.columns: df["fecha_vcto"]  = df["fecha_vcto"].apply(to_dt_nullsafe)
#         if "dias_vencidos" in df.columns: df["dias_vencidos"] = df["dias_vencidos"].apply(to_int_nullsafe)

#         def _to_money(x):
#             v = to_float(x)
#             return v if v is not None else 0.0

#         if "valor_docto" in df.columns:
#             df["valor_docto"] = df["valor_docto"].apply(_to_money)
#         else:
#             df["valor_docto"] = 0.0

#         if "total_cop" in df.columns:
#             df["total_cop"] = df["total_cop"].apply(_to_money)
#         else:
#             df["total_cop"] = df["valor_docto"]

#         # Claves limpias
#         df = df[(df["nit_cliente"].notna()) & (df["nro_docto_cruce"].notna())]
#         df["nit_cliente"] = df["nit_cliente"].astype(str).str.strip()
#         df["nro_docto_cruce"] = df["nro_docto_cruce"].astype(str).str.strip()
#         df = df[(df["nit_cliente"] != "") & (df["nro_docto_cruce"] != "")]

#         if df.empty:
#             return RedirectResponse(url="/?msg=Archivo%20vacio%20o%20sin%20filas%20validas&msg_type=info", status_code=303)

#         # --- Claves existentes en BD ---
#         bd_clientes = db.query(models.Cliente).all()
#         bd_claves = set((str(c.nit_cliente), str(c.nro_docto_cruce)) for c in bd_clientes)

#         # --- Borrado condicional ---
#         if not is_project_export:
#             excel_claves = set(zip(df["nit_cliente"], df["nro_docto_cruce"]))
#             claves_a_eliminar = bd_claves - excel_claves
#             if claves_a_eliminar:
#                 delete_by_pairs_mssql(db, "cartera", claves_a_eliminar, batch_size=900)

#         # Helpers de obs
#         def _clean_obs_text(s: Optional[str]) -> Optional[str]:
#             if not s:
#                 return None
#             txt = str(s).strip()
#             m = re.match(r"^\s*(\d{4}-\d{2}-\d{2}(?:\s+\d{2}:\d{2})?)\s*-\s*(.+)$", txt)
#             if m:
#                 return m.group(2).strip() or None
#             return txt or None

#         def _latest_obs_text(cliente):
#             latest = None
#             for o in getattr(cliente, "observaciones", []) or []:
#                 f = getattr(o, "fecha_creacion", None) or datetime.min
#                 if latest is None or f > (getattr(latest, "fecha_creacion", None) or datetime.min):
#                     latest = o
#             return (latest.texto.strip() if latest and latest.texto else None)

#         # --- Upsert + Observaciones (solo si viene del export) ---
#         for _, row in df.iterrows():
#             nit = row["nit_cliente"]
#             docto = row["nro_docto_cruce"]

#             valor_docto = row.get("valor_docto", 0.0) or 0.0
#             total_excel = row.get("total_cop", valor_docto)
#             recaudo_calc = round(max((valor_docto or 0.0) - (total_excel or 0.0), 0.0), 2)

#             cliente = db.query(models.Cliente).filter_by(
#                 nit_cliente=nit,
#                 nro_docto_cruce=docto
#             ).first()

#             correo_val = clean_email(row.get("correo")) if "correo" in row else None

#             if cliente:
#                 cliente.razon_social = row.get("razon_social")
#                 cliente.dias_vencidos = to_int_nullsafe(row.get("dias_vencidos"))
#                 cliente.fecha_docto = to_dt_nullsafe(row.get("fecha_docto"))
#                 cliente.fecha_vcto = to_dt_nullsafe(row.get("fecha_vcto"))
#                 cliente.valor_docto = valor_docto
#                 cliente.total_cop = total_excel
#                 cliente.recaudo = recaudo_calc
#                 cliente.telefono = clean_phone(row.get("telefono"))
#                 cliente.celular  = clean_phone(row.get("celular"))
#                 cliente.asesor = row.get("asesor")
#                 if correo_val:
#                     cliente.correo = correo_val
#             else:
#                 cliente = models.Cliente(
#                     razon_social=row.get("razon_social"),
#                     nit_cliente=nit,
#                     nro_docto_cruce=docto,
#                     dias_vencidos=to_int_nullsafe(row.get("dias_vencidos")),
#                     fecha_docto=to_dt_nullsafe(row.get("fecha_docto")),
#                     fecha_vcto=to_dt_nullsafe(row.get("fecha_vcto")),
#                     valor_docto=valor_docto,
#                     total_cop=total_excel,
#                     recaudo=recaudo_calc,
#                     telefono=clean_phone(row.get("telefono")),
#                     celular=clean_phone(row.get("celular")),
#                     asesor=row.get("asesor"),
#                     correo=correo_val,
#                 )
#                 db.add(cliente)
#                 db.flush()

#             if "obs_txt" in df.columns and is_project_export:
#                 new_obs_text = _clean_obs_text(row.get("obs_txt"))
#                 if new_obs_text:
#                     last_text = _latest_obs_text(cliente)
#                     if (last_text or "").strip() != new_obs_text.strip():
#                         obs = models.Observacion(
#                             cliente_id=cliente.id,
#                             texto=new_obs_text,
#                         )
#                         db.add(obs)

#         db.commit()
#         return RedirectResponse(url="/?msg=Archivo%20subido%20correctamente&msg_type=success", status_code=303)

#     except Exception as e:
#         print("❌ Error importar_excel:", e)
#         return RedirectResponse(url="/?msg=Error%20al%20subir%20el%20archivo&msg_type=error", status_code=303)


# # ------------------- Observaciones -------------------
# @app.post("/cliente/{cliente_id}/observacion")
# def agregar_observacion(cliente_id: int, texto: str = Form(...), db: Session = Depends(get_db)):
#     crud.add_observacion(db, cliente_id, texto)
#     return RedirectResponse(
#         url=f"/cliente/{cliente_id}?msg=Observaci%C3%B3n%20guardada&msg_type=success",
#         status_code=303
#     )

# # ------------------- Actualizar cliente (editar recaudo) -------------------
# @app.post("/cliente/{cliente_id}/update")
# async def update_cliente(
#     request: Request,
#     cliente_id: int,
#     db: Session = Depends(get_db)
# ):
#     form_data = await request.form()

#     cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
#     if not cliente:
#         return RedirectResponse(url="/?msg=Cliente%20no%20encontrado&msg_type=error", status_code=status.HTTP_303_SEE_OTHER)

#     for campo in ["telefono", "celular", "correo", "fecha_gestion", "tipo", "asesor"]:
#         val = form_data.get(campo)
#         if val is not None and val != "":
#             if campo == "correo":
#                 val = clean_email(val)
#             if campo in ("telefono", "celular"):
#                 val = clean_phone(val)
#             setattr(cliente, campo, val)

#     valor_docto = float(cliente.valor_docto or 0.0)

#     raw_recaudo = form_data.get("recaudo")
#     raw_total_cop = form_data.get("total_cop")

#     has_recaudo = raw_recaudo is not None and raw_recaudo.strip() != ""
#     has_total_cop = raw_total_cop is not None and raw_total_cop.strip() != ""

#     if has_recaudo:
#         delta = to_float(raw_recaudo)
#         if delta is not None:
#             old_total = float(cliente.total_cop or 0.0)
#             new_total = max(round(old_total - delta, 2), 0.0)
#             cliente.total_cop = new_total
#             cliente.recaudo = round(max(valor_docto - new_total, 0.0), 2)
#     elif has_total_cop:
#         nuevo_total = to_float(raw_total_cop)
#         if nuevo_total is not None:
#             nuevo_total = max(nuevo_total, 0.0)
#             if nuevo_total > valor_docto:
#                 nuevo_total = valor_docto
#             cliente.total_cop = round(nuevo_total, 2)
#             cliente.recaudo = round(max(valor_docto - nuevo_total, 0.0), 2)

#     nueva_obs = form_data.get("observaciones")
#     if nueva_obs and nueva_obs.strip():
#         obs = models.Observacion(
#             cliente_id=cliente.id,
#             texto=nueva_obs.strip(),
#         )
#         db.add(obs)

#     db.commit()
#     return RedirectResponse(
#         url=f"/cliente/{cliente_id}?msg=Cambios%20guardados&msg_type=success",
#         status_code=status.HTTP_303_SEE_OTHER
#     )

# # ------------------- historial cliente -------------------
# @app.get("/cliente/{cliente_id}/historial")
# def historial_cliente(cliente_id: int, db: Session = Depends(get_db)):
#     cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
#     if not cliente:
#         return JSONResponse(content=[], status_code=200)

#     historial = [
#         {
#             "texto": obs.texto,
#             "fecha": obs.fecha_creacion.strftime("%d/%m/%Y %H:%M")
#             if obs.fecha_creacion else None
#         }
#         for obs in cliente.observaciones
#     ]

#     return JSONResponse(content=historial, status_code=200)

# # ------------------- Vista cliente -------------------
# @app.get("/cliente/{cliente_id}")
# def ver_cliente(cliente_id: int, request: Request, db: Session = Depends(get_db)):
#     # 1) Cliente actual
#     cliente = db.query(models.Cliente).get(cliente_id)
#     if not cliente:
#         raise HTTPException(status_code=404, detail="Cliente no encontrado")

#     # 2) Recuperar filtros desde ref (lo que venía del index)
#     ref_qs = request.query_params.get("ref", "")
#     params = dict(parse_qsl(ref_qs)) if ref_qs else {}

#     min_dias = params.get("min_dias")
#     max_dias = params.get("max_dias")
#     sort = params.get("sort", "dias_desc")

#     def to_int_or_none(v):
#         try:
#             return int(v) if v not in (None, "") else None
#         except Exception:
#             return None

#     min_val = to_int_or_none(min_dias)
#     max_val = to_int_or_none(max_dias)

#     # 3) Base query: TODAS las facturas que cumplen filtros (como en el index)
#     q = db.query(models.Cliente)

#     if min_val is not None:
#         q = q.filter(models.Cliente.dias_vencidos >= min_val)
#     if max_val is not None:
#         q = q.filter(models.Cliente.dias_vencidos <= max_val)

#     # 4) Orden igual al index
#     order_map = {
#         "dias_desc": models.Cliente.dias_vencidos.desc(),
#         "dias_asc": models.Cliente.dias_vencidos.asc(),
#         "razon_asc": models.Cliente.razon_social.asc(),
#         "nit_asc": models.Cliente.nit_cliente.asc(),
#     }
#     q = q.order_by(order_map.get(sort, order_map["dias_desc"]))

#     # 5) Lista de IDs en el mismo orden que veías en el index
#     facturas_filtradas = q.all()
#     ids = [c.id for c in facturas_filtradas]

#     # Si por alguna razón el actual no está en la lista, lo agregamos al final
#     if cliente_id not in ids:
#         ids.append(cliente_id)

#     # 6) Calcular posición, prev y next dentro de esa lista
#     try:
#         pos = ids.index(cliente_id)
#     except ValueError:
#         pos = 0

#     prev_id = ids[pos - 1] if pos > 0 else None
#     next_id = ids[pos + 1] if pos < len(ids) - 1 else None

#     # 7) URL para "Regresar" con los filtros intactos
#     back_url = f"/?{ref_qs}" if ref_qs else "/"

#     return templates.TemplateResponse(
#         "cliente.html",
#         {
#             "request": request,
#             "cliente": cliente,
#             "prev_id": prev_id,
#             "next_id": next_id,
#             "flash_msg": request.query_params.get("msg"),
#             "flash_type": request.query_params.get("msg_type"),
#             "back_url": back_url,
#             "ref": ref_qs,  # para mantener ref en Prev / Next
#         }
#     )

# # ------------------- Index agrupado -------------------
# @app.get("/")
# def index(
#     request: Request,
#     db: Session = Depends(get_db),
#     view: Optional[str] = Query(None),
#     min_dias: Optional[str] = Query(None),
#     max_dias: Optional[str] = Query(None),
#     sort: Optional[str] = Query("dias_desc"),
#     q: Optional[str] = Query(None),   # <-- BUSCADOR POR NOMBRE
# ):
#     def to_int_or_none(val: Optional[str]):
#         try:
#             return int(val) if val not in (None, "") else None
#         except (ValueError, TypeError):
#             return None

#     min_val = to_int_or_none(min_dias)
#     max_val = to_int_or_none(max_dias)

#     query = db.query(models.Cliente)

#     # Filtro por nombre (razón social)
#     if q:
#         query = query.filter(models.Cliente.razon_social.ilike(f"%{q}%"))

#     if min_val is not None:
#         query = query.filter(models.Cliente.dias_vencidos >= min_val)
#     if max_val is not None:
#         query = query.filter(models.Cliente.dias_vencidos <= max_val)

#     order_map = {
#         "dias_desc": models.Cliente.dias_vencidos.desc(),
#         "dias_asc": models.Cliente.dias_vencidos.asc(),
#         "razon_asc": models.Cliente.razon_social.asc(),
#     }
#     query = query.order_by(order_map.get(sort, order_map["dias_desc"]))

#     clientes = query.all()

#     # Flash + QS actual para persistir filtros en links
#     flash_msg = request.query_params.get("msg")
#     flash_type = request.query_params.get("msg_type")
#     if (min_val is not None) or (max_val is not None) or q:
#         if not flash_msg:
#             flash_msg = "Filtros aplicados"
#         if not flash_type:
#             flash_type = "info"

#     current_qs = request.url.query  # <- pásalo a las plantillas

#     if view == "flat":
#         filas = []
#         for c in clientes:
#             valor_docto = Decimal(str(c.valor_docto or 0))
#             total_cop = Decimal(str(c.total_cop or 0))
#             recaudo = valor_docto - total_cop

#             filas.append({
#                 "id": c.id,
#                 "razon_social": c.razon_social,
#                 "nit_cliente": c.nit_cliente,
#                 "nro_docto_cruce": c.nro_docto_cruce,
#                 "dias_vencidos": c.dias_vencidos,
#                 "fecha_docto": c.fecha_docto,
#                 "fecha_vcto": c.fecha_vcto,
#                 "valor_docto": float(valor_docto),
#                 "total_cop": float(total_cop),
#                 "recaudo": float(recaudo),
#                 "asesor": c.asesor,
#                 "correo": c.correo,
#             })

#         return templates.TemplateResponse(
#             "index.html",
#             {
#                 "request": request,
#                 "view": "flat",
#                 "filas": filas,
#                 "flash_msg": flash_msg,
#                 "flash_type": flash_type,
#                 "current_qs": current_qs,  # <- usa esto en los href
#                 "q": q,
#                 "min_dias": min_dias,
#                 "max_dias": max_dias,
#                 "sort": sort,
#             }
#         )

#     agrupados = {}
#     for c in clientes:
#         if c.nit_cliente not in agrupados:
#             agrupados[c.nit_cliente] = {
#                 "nit_cliente": c.nit_cliente,
#                 "razon_social": c.razon_social,
#                 "telefono": c.telefono,
#                 "celular": c.celular,
#                 "correo": c.correo,
#                 "asesor": c.asesor,
#                 "facturas": []
#             }

#         valor_docto = Decimal(str(c.valor_docto or 0))
#         total_cop = Decimal(str(c.total_cop or 0))
#         recaudo = valor_docto - total_cop

#         agrupados[c.nit_cliente]["facturas"].append({
#             "id": c.id,
#             "nro_docto_cruce": c.nro_docto_cruce,
#             "dias_vencidos": c.dias_vencidos,
#             "fecha_docto": c.fecha_docto,
#             "fecha_vcto": c.fecha_vcto,
#             "valor_docto": float(valor_docto),
#             "total_cop": float(total_cop),
#             "recaudo": float(recaudo),
#             "fecha_gestion": c.fecha_gestion,
#             "tipo": c.tipo,
#             "asesor": c.asesor,
#             "observaciones": [obs.texto for obs in c.observaciones]
#         })

#     for cl in agrupados.values():
#         dias = [f["dias_vencidos"] for f in cl["facturas"] if f["dias_vencidos"] is not None]
#         cl["max_dias"] = max(dias) if dias else None

#     return templates.TemplateResponse(
#         "index.html",
#         {
#             "request": request,
#             "clientes": list(agrupados.values()),
#             "flash_msg": flash_msg,
#             "flash_type": flash_type,
#             "current_qs": current_qs,  # <- usa esto en los href
#             "q": q,
#             "min_dias": min_dias,
#             "max_dias": max_dias,
#             "sort": sort,
#         }
#     )



# ------------------- Imports y configuración inicial -------------------

# from fastapi.responses import JSONResponse, RedirectResponse, StreamingResponse
# from fastapi import UploadFile, File, Form, Query, status, FastAPI, Request, Depends, HTTPException
# from fastapi.templating import Jinja2Templates
# from fastapi.staticfiles import StaticFiles
# from sqlalchemy.orm import Session
# from sqlalchemy import text  # SQL crudo MSSQL-safe
# from datetime import datetime
# from typing import Optional, Tuple
# from decimal import Decimal
# from io import BytesIO
# from starlette.datastructures import QueryParams
# import pandas as pd
# import models
# import math
# import crud
# import re

# from database import SessionLocal, engine, Base  # tu módulo existente


# # ------------------- Inicialización -------------------

# app = FastAPI()
# templates = Jinja2Templates(directory="templates")

# app.mount("/static", StaticFiles(directory="static"), name="static")
# app.mount("/assets", StaticFiles(directory="assets"), name="assets")


# # ------------------- Helpers -------------------

# def fmt_money(value):
#     """Filtro para mostrar números como moneda"""
#     if value is None or value == "":
#         return "-"
#     try:
#         v = Decimal(str(value))
#         return f"{v:,.2f}"
#     except Exception:
#         return str(value)

# templates.env.filters["fmt_money"] = fmt_money


# def get_db():
#     db = SessionLocal()
#     try:
#         yield db
#     finally:
#         db.close()


# def to_float(val):
#     try:
#         if pd.isna(val):
#             return None
#         val_str = str(val).replace(",", ".")
#         val_str = re.sub(r"[^0-9.]", "", val_str)
#         return float(val_str) if val_str else None
#     except (ValueError, TypeError):
#         return None


# def clean_phone(val):
#     """Normaliza teléfonos: quita .0, exponentes, separadores; deja solo dígitos."""
#     if val is None or (isinstance(val, float) and pd.isna(val)):
#         return None
#     s = str(val).strip()

#     # 3046304674.0 -> 3046304674
#     if s.endswith(".0") and s[:-2].isdigit():
#         s = s[:-2]

#     # 3.0463E+09 -> 3046300000
#     try:
#         if re.fullmatch(r"\d+(\.\d+)?[eE][+-]?\d+", s):
#             s = str(int(float(s)))
#     except Exception:
#         pass

#     # 1234.00 -> 1234 si es entero exacto
#     if re.fullmatch(r"\d+\.\d+", s):
#         try:
#             f = float(s)
#             if f.is_integer():
#                 s = str(int(f))
#         except Exception:
#             pass

#     # Dejar solo dígitos
#     s = re.sub(r"[^\d]", "", s)
#     return s or None


# def clean_email(val):
#     """
#     Normaliza emails:
#     - Devuelve el PRIMER email válido encontrado (separadores ; , / espacio).
#     - Tolera mayúsculas/ruido.
#     """
#     if val is None or (isinstance(val, float) and pd.isna(val)):
#         return None
#     s = str(val).strip()
#     if not s or s.lower() == "nan":
#         return None

#     # Cortes por separadores comunes
#     candidates = re.split(r"[;,\n/\s]+", s)

#     email_re = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.IGNORECASE)

#     # Primero busca dentro de cada segmento
#     for part in candidates:
#         m = email_re.search(part or "")
#         if m:
#             return m.group(0).strip().lower()

#     # Si no, busca en todo el string
#     m = email_re.search(s)
#     if m:
#         return m.group(0).strip().lower()

#     return None


# # ---- Null-safe parsers para MSSQL (evitan NaN/Timestamp) ----
# def to_int_nullsafe(val):
#     """Convierte a int o devuelve None si viene vacío/NaN."""
#     try:
#         if val is None:
#             return None
#         if isinstance(val, float) and math.isnan(val):  # NaN
#             return None
#         s = str(val).strip()
#         if s == "":
#             return None
#         return int(float(s))  # admite "5.0"
#     except Exception:
#         return None


# def to_dt_nullsafe(val):
#     """Convierte a datetime (py) o None. Acepta pandas Timestamp, str, excel serial."""
#     if val is None:
#         return None
#     try:
#         if isinstance(val, float) and math.isnan(val):
#             return None
#     except Exception:
#         pass
#     ts = pd.to_datetime(val, errors="coerce")
#     if pd.isna(ts):
#         return None
#     return ts.to_pydatetime()


# # ---------- Helpers nuevos para borrado MSSQL-safe ----------
# def _chunk(iterable, size):
#     buf = []
#     for x in iterable:
#         buf.append(x)
#         if len(buf) >= size:
#             yield buf
#             buf = []
#     if buf:
#         yield buf


# def delete_by_pairs_mssql(db: Session, table_name: str, key_pairs, batch_size: int = 900):
#     """
#     MSSQL-safe:
#     DELETE c
#     FROM <tabla> AS c
#     JOIN (VALUES (...), ...) AS v(nit_cliente, nro_docto_cruce)
#       ON v.nit_cliente = c.nit_cliente
#      AND v.nro_docto_cruce = c.nro_docto_cruce;
#     """
#     pairs = list(key_pairs)
#     if not pairs:
#         return

#     for batch in _chunk(pairs, batch_size):
#         values_clause = []
#         params = {}
#         for i, (nit, doc) in enumerate(batch):
#             values_clause.append(f"(:nit{i}, :doc{i})")
#             params[f"nit{i}"] = str(nit) if nit is not None else None
#             params[f"doc{i}"] = str(doc) if doc is not None else None

#         sql = f"""
#         DELETE c
#         FROM {table_name} AS c
#         JOIN (VALUES {", ".join(values_clause)}) AS v(nit_cliente, nro_docto_cruce)
#           ON v.nit_cliente = c.nit_cliente
#          AND v.nro_docto_cruce = c.nro_docto_cruce;
#         """
#         db.execute(text(sql), params)


# # ------------------- Helpers de filtros/orden y navegación coherente -------------------

# def _to_int_or_none(val: Optional[str]):
#     try:
#         return int(val) if val not in (None, "") else None
#     except (ValueError, TypeError):
#         return None


# def _parse_filters(request_like) -> Tuple[Optional[int], Optional[int], str, Optional[str]]:
#     """
#     Extrae min_dias, max_dias, sort, view desde query_params (del request o de ref)
#     """
#     qp = request_like.query_params
#     min_val = _to_int_or_none(qp.get("min_dias"))
#     max_val = _to_int_or_none(qp.get("max_dias"))
#     sort = qp.get("sort", "dias_desc")
#     view = qp.get("view")
#     return min_val, max_val, sort, view


# def _apply_filters_and_order(db: Session, min_val: Optional[int], max_val: Optional[int], sort: str, view: Optional[str]):
#     """
#     Devuelve un Query con el mismo filtrado/orden que el index.
#     Si la vista es 'group' y no se especifica nit_asc/razon_asc, forzamos nit_asc
#     para que la navegación sea estable por NIT (coherente con el acordeón).
#     """
#     q = db.query(models.Cliente)

#     if min_val is not None:
#         q = q.filter(models.Cliente.dias_vencidos >= min_val)
#     if max_val is not None:
#         q = q.filter(models.Cliente.dias_vencidos <= max_val)

#     effective_sort = sort
#     if (view == "group") and (sort not in {"nit_asc", "razon_asc"}):
#         effective_sort = "nit_asc"

#     if effective_sort == "dias_asc":
#         q = q.order_by(
#             models.Cliente.dias_vencidos.asc(),
#             models.Cliente.nit_cliente.asc(),
#             models.Cliente.nro_docto_cruce.asc(),
#             models.Cliente.id.asc(),
#         )
#     elif effective_sort == "razon_asc":
#         q = q.order_by(
#             models.Cliente.razon_social.asc(),
#             models.Cliente.nit_cliente.asc(),
#             models.Cliente.nro_docto_cruce.asc(),
#             models.Cliente.id.asc(),
#         )
#     elif effective_sort == "nit_asc":
#         q = q.order_by(
#             models.Cliente.nit_cliente.asc(),
#             models.Cliente.nro_docto_cruce.asc(),
#             models.Cliente.fecha_docto.asc(),
#             models.Cliente.id.asc(),
#         )
#     else:  # dias_desc (default)
#         q = q.order_by(
#             models.Cliente.dias_vencidos.desc(),
#             models.Cliente.nit_cliente.asc(),
#             models.Cliente.nro_docto_cruce.asc(),
#             models.Cliente.id.asc(),
#         )
#     return q


# # ------------------- Exportar cartera a Excel -------------------
# @app.get("/exportar_cartera.xlsx", name="exportar_cartera_xlsx")
# def exportar_cartera_xlsx(
#     db: Session = Depends(get_db),
#     min_dias: Optional[str] = Query(None),
#     max_dias: Optional[str] = Query(None),
#     sort: Optional[str] = Query("dias_desc"),
# ):
#     def to_int_or_none(val: Optional[str]):
#         try:
#             return int(val) if val not in (None, "") else None
#         except (ValueError, TypeError):
#             return None

#     min_val = to_int_or_none(min_dias)
#     max_val = to_int_or_none(max_dias)

#     query = db.query(models.Cliente)
#     if min_val is not None:
#         query = query.filter(models.Cliente.dias_vencidos >= min_val)
#     if max_val is not None:
#         query = query.filter(models.Cliente.dias_vencidos <= max_val)

#     order_map = {
#         "dias_desc": models.Cliente.dias_vencidos.desc(),
#         "dias_asc": models.Cliente.dias_vencidos.asc(),
#         "razon_asc": models.Cliente.razon_social.asc(),
#     }
#     query = query.order_by(order_map.get(sort, order_map["dias_desc"]))
#     clientes = query.all()

#     rows = []
#     for c in clientes:
#         valor_docto = Decimal(str(c.valor_docto or 0))
#         total_cop = Decimal(str(c.total_cop or 0))
#         recaudo = Decimal(str(c.recaudo)) if c.recaudo is not None else (valor_docto - total_cop)

#         # ÚLTIMA observación (si existe) para export
#         obs_txt = ""
#         if getattr(c, "observaciones", None):
#             latest = None
#             for o in c.observaciones:
#                 f = o.fecha_creacion or datetime.min
#                 if latest is None or f > (latest.fecha_creacion or datetime.min):
#                     latest = o
#             if latest:
#                 obs_txt = f"{(latest.fecha_creacion.strftime('%Y-%m-%d %H:%M') if latest.fecha_creacion else '')} - {latest.texto}"

#         rows.append({
#             "ID": c.id,
#             "Razón social": c.razon_social,
#             "NIT": c.nit_cliente,
#             "Docto cruce": c.nro_docto_cruce,
#             "Días vencidos": c.dias_vencidos,
#             "Fecha docto": c.fecha_docto,
#             "Fecha vcto": c.fecha_vcto,
#             "Valor docto": float(valor_docto),
#             "Total COP (saldo)": float(total_cop),
#             "Recaudo": float(recaudo),
#             "Teléfono": c.telefono,
#             "Celular": c.celular,
#             "Correo": c.correo,
#             "Asesor": c.asesor,
#             "Fecha gestión": c.fecha_gestion,
#             "Tipo": c.tipo,
#             "Observaciones": obs_txt,
#         })

#     df = pd.DataFrame(rows)

#     # Normaliza fechas a datetime (Excel-friendly)
#     for col in ["Fecha docto", "Fecha vcto", "Fecha gestión"]:
#         if col in df.columns:
#             df[col] = pd.to_datetime(df[col], errors="coerce")

#     # --- Resumen por cliente ---
#     if not df.empty:
#         resumen = (
#             df.groupby(["NIT", "Razón social"], dropna=False)
#               .agg({
#                   "Valor docto": "sum",
#                   "Total COP (saldo)": "sum",
#                   "Recaudo": "sum",
#                   "Días vencidos": "max",
#               })
#               .rename(columns={
#                   "Valor docto": "Total Valor Docto",
#                   "Total COP (saldo)": "Saldo Total",
#                   "Recaudo": "Recaudo Total",
#                   "Días vencidos": "Max Días Vencidos",
#               })
#               .reset_index()
#         )
#         ref = (
#             df.sort_values(["NIT"]).groupby(["NIT", "Razón social"], dropna=False)
#               .agg({"Teléfono": "first", "Celular": "first", "Correo": "first", "Asesor": "first"})
#               .reset_index()
#         )
#         resumen = resumen.merge(ref, on=["NIT", "Razón social"], how="left")

#         facturas = (
#             df.groupby(["NIT", "Razón social"], dropna=False)
#               .size()
#               .reset_index(name="# Facturas")
#         )
#         resumen = resumen.merge(facturas, on=["NIT", "Razón social"], how="left")
#         resumen["# Facturas"] = resumen["# Facturas"].fillna(0).astype(int)
#     else:
#         resumen = pd.DataFrame(columns=[
#             "NIT","Razón social","Total Valor Docto","Saldo Total","Recaudo Total",
#             "Max Días Vencidos","Teléfono","Celular","Correo","Asesor","# Facturas"
#         ])

#     # --- Excel en memoria con formato ---
#     output = BytesIO()
#     with pd.ExcelWriter(output, engine="xlsxwriter",
#                         datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as writer:
#         df.to_excel(writer, index=False, sheet_name="cartera")
#         resumen.to_excel(writer, index=False, sheet_name="resumen")

#         wb = writer.book
#         ws1 = writer.sheets["cartera"]
#         ws2 = writer.sheets["resumen"]

#         for ws in (ws1, ws2):
#             ws.freeze_panes(1, 0)

#         fmt_header = wb.add_format({
#             'bold': True, 'bg_color': '#D9E1F2', 'border': 1,
#             'align': 'center', 'valign': 'vcenter'
#         })
#         fmt_money0 = wb.add_format({'num_format': '#,##0', 'border': 1})
#         fmt_int = wb.add_format({'num_format': '#,##0', 'border': 1})
#         fmt_date = wb.add_format({'num_format': 'yyyy-mm-dd', 'border': 1})
#         fmt_text = wb.add_format({'text_wrap': True, 'border': 1, 'valign': 'top'})
#         fmt_default = wb.add_format({'border': 1})

#         def table_with_formats(ws, df_sheet, money_cols, date_cols, wrap_cols):
#             """Crea tabla con formatos por columna y calcula anchos."""
#             rows, cols = df_sheet.shape
#             headers = list(df_sheet.columns)

#             # Mapear formato por columna
#             per_col_format = {}
#             for col in headers:
#                 if col in money_cols:
#                     per_col_format[col] = fmt_money0
#                 elif col in date_cols:
#                     per_col_format[col] = fmt_date
#                 elif col in wrap_cols:
#                     per_col_format[col] = fmt_text
#                 elif df_sheet[col].dtype.kind in ("i", "u"):  # enteros
#                     per_col_format[col] = fmt_int
#                 else:
#                     per_col_format[col] = fmt_default

#             # Definir columnas de la tabla con formato (esto sí aplica dentro de la tabla)
#             columns_def = [{'header': h, 'format': per_col_format[h]} for h in headers]

#             # Crear tabla (si hay datos). Si no, solo pintamos encabezados.
#             if rows > 0:
#                 ws.add_table(0, 0, rows, cols - 1, {
#                     'style': 'Table Style Medium 9',
#                     'banded_rows': True,
#                     'columns': columns_def
#                 })
#             else:
#                 for c_idx, h in enumerate(headers):
#                     ws.write(0, c_idx, h, fmt_header)

#             # Auto–ancho aproximado por columna
#             col_widths = {}
#             for c_idx, col in enumerate(headers):
#                 max_len = len(str(col)) + 2
#                 serie = df_sheet[col]

#                 if col in money_cols:
#                     # medir como miles sin decimales
#                     for v in serie.dropna():
#                         try:
#                             s = f"{int(round(float(v))):,}"
#                             max_len = max(max_len, len(s))
#                         except Exception:
#                             pass
#                 elif col in date_cols:
#                     max_len = max(max_len, 12)
#                 else:
#                     # considerar saltos de línea (tomar la línea más larga)
#                     for v in serie.dropna():
#                         s = str(v).replace("\r", "")
#                         max_len = max(max_len, max((len(seg) for seg in s.split("\n")), default=0))
#                     # límite práctico
#                     if col in wrap_cols:
#                         max_len = min(max(max_len, 40), 80)
#                     else:
#                         max_len = min(max_len, 40)

#                 width = max_len + 1
#                 col_widths[col] = width
#                 ws.set_column(c_idx, c_idx, width, per_col_format[col])

#             return col_widths, per_col_format

#         def autofit_row_heights(ws, df_sheet, wrap_cols, col_widths):
#             """Ajusta alto de fila según # líneas estimadas en columnas wrap."""
#             if df_sheet.empty or not wrap_cols:
#                 return
#             base_height = 15  # pts aprox por línea
#             max_height = 300

#             for r in range(1, len(df_sheet) + 1):  # +1 por header
#                 max_lines = 1
#                 for col in wrap_cols:
#                     if col not in df_sheet.columns:
#                         continue
#                     val = df_sheet.iloc[r - 1][col]
#                     if pd.isna(val) or val is None:
#                         continue
#                     text = str(val).replace("\r", "")
#                     if text == "":
#                         continue

#                     # Estimar caracteres por línea a partir del ancho de columna
#                     col_w = int(col_widths.get(col, 40))
#                     chars_per_line = max(col_w - 2, 10)

#                     total_lines = 0
#                     for seg in text.split("\n"):
#                         seg = seg.strip()
#                         if seg == "":
#                             total_lines += 1
#                         else:
#                             total_lines += math.ceil(len(seg) / chars_per_line)

#                     max_lines = max(max_lines, total_lines)

#                 ws.set_row(r, min(base_height * max_lines + 4, max_height))

#         # Columnas por tipo (cartera)
#         money_cols_cartera = ["Valor docto", "Total COP (saldo)", "Recaudo"]
#         date_cols_cartera = ["Fecha docto", "Fecha vcto", "Fecha gestión"]
#         wrap_cols_cartera = ["Razón social", "Observaciones"]  # ⇦ ajustar ancho/alto por texto

#         # Columnas por tipo (resumen)
#         money_cols_resumen = ["Total Valor Docto", "Saldo Total", "Recaudo Total"]
#         date_cols_resumen = []
#         wrap_cols_resumen = ["Razón social"]

#         # Aplicar tabla + formatos + anchos
#         widths_cartera, _ = table_with_formats(ws1, df, money_cols_cartera, date_cols_cartera, wrap_cols_cartera)
#         widths_resumen, _ = table_with_formats(ws2, resumen, money_cols_resumen, date_cols_resumen, wrap_cols_resumen)

#         # Ajustar alturas (solo donde hay wrap)
#         autofit_row_heights(ws1, df, wrap_cols_cartera, widths_cartera)
#         autofit_row_heights(ws2, resumen, wrap_cols_resumen, widths_resumen)

#     output.seek(0)
#     filename = f"cartera_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
#     return StreamingResponse(
#         output,
#         media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#         headers={"Content-Disposition": f'attachment; filename="{filename}"'}
#     )


# # ------------------- Importar cartera -------------------
# @app.post("/importar_excel")
# async def importar_excel(file: UploadFile = File(...), db: Session = Depends(get_db)):
#     try:
#         contents = await file.read()

#         # --- Cargar Excel/CSV (ambos formatos) ---
#         def _read_any(contents: bytes, filename: str):
#             name = (filename or "").lower()
#             if name.endswith(".csv"):
#                 df = pd.read_csv(BytesIO(contents), dtype=str, keep_default_na=False)
#                 return {"__csv__": df}
#             return pd.read_excel(BytesIO(contents), sheet_name=None, dtype=str, keep_default_na=False)

#         sheets = _read_any(contents, file.filename)

#         # --- Normalizador simple de strings ---
#         def norm(s: str) -> str:
#             if s is None:
#                 return ""
#             t = str(s).strip().lower()
#             rep = {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ñ":"n","ä":"a","ë":"e","ï":"i","ö":"o","ü":"u","’":"'","´":"'","`":"'"}
#             for a,b in rep.items(): t = t.replace(a,b)
#             t = re.sub(r"\s+", " ", t)
#             return t

#         # --- Aliases (internos -> alias) ---
#         alias_map = {
#             "razon_social": {"razon social", "razon social vend. cliente", "razon social vend cliente"},
#             "nit_cliente": {"nit", "nit cliente despacho", "nit cliente"},
#             "nro_docto_cruce": {"docto cruce", "nro. docto. cruce", "nro docto cruce", "factura", "documento", "nro documento"},
#             "fecha_docto": {"fecha docto", "fecha docto."},
#             "fecha_vcto": {"fecha vcto", "fecha vcto."},
#             "dias_vencidos": {"dias vencidos", "dias vencidos."},
#             "valor_docto": {"valor docto", "valor docto."},
#             "total_cop": {"total cop", "total cop (saldo)", "saldo", "total (cop)"},
#             "telefono": {"telefono", "teléfono"},
#             "celular": {"celular", "movil", "móvil"},
#             "correo": {"correo", "email", "e-mail", "mail", "correo electronico", "correo electrónico"},
#             "asesor": {"asesor", "vendedor", "razon social vend. cliente", "razon social vend cliente"},
#             "obs_txt": {"observaciones"},
#         }

#         # --- Detectar si es export del proyecto (hoja 'cartera' con NIT/Docto) ---
#         is_project_export = False
#         df_project = None
#         for sheet_name, df in sheets.items():
#             if norm(sheet_name) == "cartera":
#                 df_project = df
#                 break
#         if df_project is not None:
#             cols_n = {norm(c) for c in df_project.columns}
#             if "nit" in cols_n and ("docto cruce" in cols_n or "nro. docto. cruce" in cols_n or "nro docto cruce" in cols_n):
#                 is_project_export = True

#         # --- Elegir hoja candidata ---
#         if is_project_export:
#             df_raw = df_project.copy()
#         else:
#             df_raw = None
#             for _, df in sheets.items():
#                 cols_n = {norm(c) for c in df.columns}
#                 if any(x in cols_n for x in alias_map["nit_cliente"]) and \
#                    any(x in cols_n for x in alias_map["nro_docto_cruce"]):
#                     df_raw = df
#                     break
#             if df_raw is None:
#                 return RedirectResponse(
#                     url="/?msg=No%20se%20encontro%20hoja%20compatible%20(ni%20cartera%20ni%20columnas%20clave)&msg_type=error",
#                     status_code=303
#                 )

#         # Copia de encabezados originales por si necesitamos "rescatar" correo
#         original_cols = list(df_raw.columns)

#         df_raw.columns = [str(c).strip() for c in df_raw.columns]

#         # --- Renombrado dinámico por alias ---
#         rename_map_dynamic = {}
#         for col in df_raw.columns:
#             n = norm(col)
#             for internal, aliases in alias_map.items():
#                 if n in aliases and internal not in rename_map_dynamic.values():
#                     rename_map_dynamic[col] = internal
#                     break

#         # Encabezados exactos del export del proyecto
#         project_head_map = {
#             "Razón social": "razon_social",
#             "NIT": "nit_cliente",
#             "Docto cruce": "nro_docto_cruce",
#             "Días vencidos": "dias_vencidos",
#             "Fecha docto": "fecha_docto",
#             "Fecha vcto": "fecha_vcto",
#             "Valor docto": "valor_docto",
#             "Total COP (saldo)": "total_cop",
#             "Teléfono": "telefono",
#             "Celular": "celular",
#             "Correo": "correo",
#             "Email": "correo",
#             "E-mail": "correo",
#             "Asesor": "asesor",
#             "Observaciones": "obs_txt",
#         }
#         for col in df_raw.columns:
#             if col in project_head_map and project_head_map[col] not in rename_map_dynamic.values():
#                 rename_map_dynamic[col] = project_head_map[col]

#         df = df_raw.rename(columns=rename_map_dynamic)

#         # --- Asegurar/Coalescer columna 'correo' desde cualquier variante ---
#         def _norm2(s: str) -> str:
#             if s is None:
#                 return ""
#             t = str(s).strip().lower()
#             rep = {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ñ":"n"}
#             for a,b in rep.items(): t = t.replace(a,b)
#             t = re.sub(r"\s+", " ", t)
#             return t

#         email_cols_candidates = []
#         for col in df_raw.columns:
#             n = _norm2(col)
#             if re.fullmatch(r"(correo( electronico)?)|(e[\s-]?mail)|email|mail", n):
#                 email_cols_candidates.append(col)

#         if "correo" not in df.columns:
#             if email_cols_candidates:
#                 correo_series = None
#                 for col in email_cols_candidates:
#                     s = df_raw[col].astype(str)
#                     s = s.where(~s.str.fullmatch(r"(?i)\s*nan\s*"), "")
#                     if correo_series is None:
#                         correo_series = s
#                     else:
#                         correo_series = correo_series.mask(
#                             (correo_series.astype(str).str.strip() == "") & (s.astype(str).str.strip() != ""),
#                             s
#                         )
#                 df["correo"] = correo_series
#             else:
#                 df["correo"] = None

#         # Normalizar emails al final
#         df["correo"] = df["correo"].apply(clean_email)

#         # Validar claves mínimas
#         if "nit_cliente" not in df.columns or "nro_docto_cruce" not in df.columns:
#             return RedirectResponse(
#                 url="/?msg=El%20archivo%20no%20tiene%20las%20columnas%20clave%20(NIT%20y%20Docto%20cruce)&msg_type=error",
#                 status_code=303
#             )

#         # --- Normalizaciones de campos ---
#         if "telefono" in df.columns: df["telefono"] = df["telefono"].apply(clean_phone)
#         if "celular" in df.columns: df["celular"] = df["celular"].apply(clean_phone)

#         if "fecha_docto" in df.columns: df["fecha_docto"] = df["fecha_docto"].apply(to_dt_nullsafe)
#         if "fecha_vcto"  in df.columns: df["fecha_vcto"]  = df["fecha_vcto"].apply(to_dt_nullsafe)
#         if "dias_vencidos" in df.columns: df["dias_vencidos"] = df["dias_vencidos"].apply(to_int_nullsafe)

#         def _to_money(x):
#             v = to_float(x)
#             return v if v is not None else 0.0

#         if "valor_docto" in df.columns:
#             df["valor_docto"] = df["valor_docto"].apply(_to_money)
#         else:
#             df["valor_docto"] = 0.0

#         if "total_cop" in df.columns:
#             df["total_cop"] = df["total_cop"].apply(_to_money)
#         else:
#             df["total_cop"] = df["valor_docto"]

#         # Claves limpias
#         df = df[(df["nit_cliente"].notna()) & (df["nro_docto_cruce"].notna())]
#         df["nit_cliente"] = df["nit_cliente"].astype(str).str.strip()
#         df["nro_docto_cruce"] = df["nro_docto_cruce"].astype(str).str.strip()
#         df = df[(df["nit_cliente"] != "") & (df["nro_docto_cruce"] != "")]

#         if df.empty:
#             return RedirectResponse(url="/?msg=Archivo%20vacio%20o%20sin%20filas%20validas&msg_type=info", status_code=303)

#         # --- Claves existentes en BD ---
#         bd_clientes = db.query(models.Cliente).all()
#         bd_claves = set((str(c.nit_cliente), str(c.nro_docto_cruce)) for c in bd_clientes)

#         # --- Borrado condicional ---
#         if not is_project_export:
#             excel_claves = set(zip(df["nit_cliente"], df["nro_docto_cruce"]))
#             claves_a_eliminar = bd_claves - excel_claves
#             if claves_a_eliminar:
#                 delete_by_pairs_mssql(db, "cartera", claves_a_eliminar, batch_size=900)

#         # Helpers de obs
#         def _clean_obs_text(s: Optional[str]) -> Optional[str]:
#             if not s:
#                 return None
#             txt = str(s).strip()
#             m = re.match(r"^\s*(\d{4}-\d{2}-\d{2}(?:\s+\d{2}:\d{2})?)\s*-\s*(.+)$", txt)
#             if m:
#                 return m.group(2).strip() or None
#             return txt or None

#         def _latest_obs_text(cliente):
#             latest = None
#             for o in getattr(cliente, "observaciones", []) or []:
#                 f = getattr(o, "fecha_creacion", None) or datetime.min
#                 if latest is None or f > (getattr(latest, "fecha_creacion", None) or datetime.min):
#                     latest = o
#             return (latest.texto.strip() if latest and latest.texto else None)

#         # --- Upsert + Observaciones (solo si viene del export) ---
#         for _, row in df.iterrows():
#             nit = row["nit_cliente"]
#             docto = row["nro_docto_cruce"]

#             valor_docto = row.get("valor_docto", 0.0) or 0.0
#             total_excel = row.get("total_cop", valor_docto)
#             recaudo_calc = round(max((valor_docto or 0.0) - (total_excel or 0.0), 0.0), 2)

#             cliente = db.query(models.Cliente).filter_by(
#                 nit_cliente=nit,
#                 nro_docto_cruce=docto
#             ).first()

#             correo_val = clean_email(row.get("correo")) if "correo" in row else None

#             if cliente:
#                 cliente.razon_social = row.get("razon_social")
#                 cliente.dias_vencidos = to_int_nullsafe(row.get("dias_vencidos"))
#                 cliente.fecha_docto = to_dt_nullsafe(row.get("fecha_docto"))
#                 cliente.fecha_vcto = to_dt_nullsafe(row.get("fecha_vcto"))
#                 cliente.valor_docto = valor_docto
#                 cliente.total_cop = total_excel
#                 cliente.recaudo = recaudo_calc
#                 cliente.telefono = clean_phone(row.get("telefono"))
#                 cliente.celular  = clean_phone(row.get("celular"))
#                 cliente.asesor = row.get("asesor")
#                 if correo_val and (not cliente.correo or cliente.correo != correo_val):
#                     cliente.correo = correo_val
#             else:
#                 cliente = models.Cliente(
#                     razon_social=row.get("razon_social"),
#                     nit_cliente=nit,
#                     nro_docto_cruce=docto,
#                     dias_vencidos=to_int_nullsafe(row.get("dias_vencidos")),
#                     fecha_docto=to_dt_nullsafe(row.get("fecha_docto")),
#                     fecha_vcto=to_dt_nullsafe(row.get("fecha_vcto")),
#                     valor_docto=valor_docto,
#                     total_cop=total_excel,
#                     recaudo=recaudo_calc,
#                     telefono=clean_phone(row.get("telefono")),
#                     celular=clean_phone(row.get("celular")),
#                     asesor=row.get("asesor"),
#                     correo=correo_val,
#                 )
#                 db.add(cliente)
#                 db.flush()

#             if "obs_txt" in df.columns and is_project_export:
#                 new_obs_text = _clean_obs_text(row.get("obs_txt"))
#                 if new_obs_text:
#                     last_text = _latest_obs_text(cliente)
#                     if (last_text or "").strip() != new_obs_text.strip():
#                         obs = models.Observacion(
#                             cliente_id=cliente.id,
#                             texto=new_obs_text,
#                         )
#                         db.add(obs)

#         db.commit()
#         return RedirectResponse(url="/?msg=Archivo%20subido%20correctamente&msg_type=success", status_code=303)

#     except Exception as e:
#         print("❌ Error importar_excel:", e)
#         return RedirectResponse(url="/?msg=Error%20al%20subir%20el%20archivo&msg_type=error", status_code=303)


# # ------------------- Observaciones -------------------
# @app.post("/cliente/{cliente_id}/observacion")
# def agregar_observacion(cliente_id: int, texto: str = Form(...), db: Session = Depends(get_db)):
#     crud.add_observacion(db, cliente_id, texto)
#     return RedirectResponse(
#         url=f"/cliente/{cliente_id}?msg=Observaci%C3%B3n%20guardada&msg_type=success",
#         status_code=303
#     )


# # ------------------- Actualizar cliente (editar recaudo) -------------------
# @app.post("/cliente/{cliente_id}/update")
# async def update_cliente(
#     request: Request,
#     cliente_id: int,
#     db: Session = Depends(get_db)
# ):
#     form_data = await request.form()

#     cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
#     if not cliente:
#         return RedirectResponse(url="/?msg=Cliente%20no%20encontrado&msg_type=error", status_code=status.HTTP_303_SEE_OTHER)

#     for campo in ["telefono", "celular", "correo", "fecha_gestion", "tipo", "asesor"]:
#         val = form_data.get(campo)
#         if val is not None and val != "":
#             if campo == "correo":
#                 val = clean_email(val)
#             if campo in ("telefono", "celular"):
#                 val = clean_phone(val)
#             setattr(cliente, campo, val)

#     valor_docto = float(cliente.valor_docto or 0.0)

#     raw_recaudo = form_data.get("recaudo")
#     raw_total_cop = form_data.get("total_cop")

#     has_recaudo = raw_recaudo is not None and raw_recaudo.strip() != ""
#     has_total_cop = raw_total_cop is not None and raw_total_cop.strip() != ""

#     if has_recaudo:
#         # Interpreta recaudo como ABONO (delta)
#         delta = to_float(raw_recaudo)
#         if delta is not None:
#             old_total = float(cliente.total_cop or 0.0)
#             new_total = max(round(old_total - delta, 2), 0.0)
#             cliente.total_cop = new_total
#             cliente.recaudo = round(max(valor_docto - new_total, 0.0), 2)
#     elif has_total_cop:
#         # Ajuste explícito de saldo total_cop => recalcular recaudo
#         nuevo_total = to_float(raw_total_cop)
#         if nuevo_total is not None:
#             nuevo_total = max(nuevo_total, 0.0)
#             # No dejar que el saldo supere el valor del documento
#             if nuevo_total > valor_docto:
#                 nuevo_total = valor_docto
#             cliente.total_cop = round(nuevo_total, 2)
#             cliente.recaudo = round(max(valor_docto - nuevo_total, 0.0), 2)

#     nueva_obs = form_data.get("observaciones")
#     if nueva_obs and nueva_obs.strip():
#         obs = models.Observacion(
#             cliente_id=cliente.id,
#             texto=nueva_obs.strip(),
#         )
#         db.add(obs)

#     db.commit()
#     return RedirectResponse(
#         url=f"/cliente/{cliente_id}?msg=Cambios%20guardados&msg_type=success",
#         status_code=status.HTTP_303_SEE_OTHER
#     )


# # ------------------- historial cliente -------------------
# @app.get("/cliente/{cliente_id}/historial")
# def historial_cliente(cliente_id: int, db: Session = Depends(get_db)):
#     cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
#     if not cliente:
#         return JSONResponse(content=[], status_code=200)

#     historial = [
#         {
#             "texto": obs.texto,
#             "fecha": obs.fecha_creacion.strftime("%d/%m/%Y %H:%M")
#             if obs.fecha_creacion else None
#         }
#         for obs in cliente.observaciones
#     ]

#     return JSONResponse(content=historial, status_code=200)


# # ------------------- Vista cliente -------------------
# @app.get("/cliente/{cliente_id}")
# def ver_cliente(cliente_id: int, request: Request, db: Session = Depends(get_db)):
#     cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
#     if not cliente:
#         raise HTTPException(status_code=404, detail="Cliente no encontrado")

#     # Reconstruye filtros desde ref (si existe) para navegar en el MISMO universo y orden
#     ref_qs = request.query_params.get("ref")
#     req_like = request
#     if ref_qs:
#         # Creamos un objeto "request-like" solo con query_params = QueryParams(ref_qs)
#         req_like = type("R", (), {"query_params": QueryParams(ref_qs)})

#     min_val, max_val, sort, view = _parse_filters(req_like)

#     # Lista ordenada de ids con la MISMA lógica del index
#     ordered_ids = [
#         row.id for row in _apply_filters_and_order(db, min_val, max_val, sort, view)
#                     .with_entities(models.Cliente.id).all()
#     ]

#     # Ubica la posición actual y calcula prev/sig basados en la lista ordenada
#     prev_id = next_id = None
#     try:
#         idx = ordered_ids.index(cliente_id)
#         if idx > 0:
#             prev_id = ordered_ids[idx - 1]
#         if idx < len(ordered_ids) - 1:
#             next_id = ordered_ids[idx + 1]
#     except ValueError:
#         # Si el id no está en la lista filtrada (raro), caemos a navegación por id
#         prev_cliente = db.query(models.Cliente).filter(models.Cliente.id < cliente_id).order_by(models.Cliente.id.desc()).first()
#         next_cliente = db.query(models.Cliente).filter(models.Cliente.id > cliente_id).order_by(models.Cliente.id.asc()).first()
#         prev_id = prev_cliente.id if prev_cliente else None
#         next_id = next_cliente.id if next_cliente else None

#     back_url = f"/?{ref_qs}" if ref_qs else "/"

#     return templates.TemplateResponse(
#         "cliente.html",
#         {
#             "request": request,
#             "cliente": cliente,
#             "prev_id": prev_id,
#             "next_id": next_id,
#             "flash_msg": request.query_params.get("msg"),
#             "flash_type": request.query_params.get("msg_type"),
#             "back_url": back_url,
#         }
#     )


# # ------------------- Index agrupado/plano -------------------
# @app.get("/")
# def index(
#     request: Request,
#     db: Session = Depends(get_db),
#     view: Optional[str] = Query(None),
#     min_dias: Optional[str] = Query(None),
#     max_dias: Optional[str] = Query(None),
#     sort: Optional[str] = Query("dias_desc"),
# ):
#     # Reconstruye filtros y aplica MISMO orden que usaremos en ver_cliente()
#     min_val = _to_int_or_none(min_dias)
#     max_val = _to_int_or_none(max_dias)

#     # NOTA: pasamos "view" al motor de orden para que 'group' fuerce nit_asc por defecto
#     query = _apply_filters_and_order(db, min_val, max_val, sort, view)
#     clientes = query.all()

#     # Flash
#     flash_msg = request.query_params.get("msg")
#     flash_type = request.query_params.get("msg_type")
#     if (min_val is not None) or (max_val is not None):
#         if not flash_msg:
#             flash_msg = "Filtros aplicados"
#         if not flash_type:
#             flash_type = "info"

#     current_qs = request.url.query  # para mantener filtros en hrefs

#     # Vista plana
#     if view == "flat":
#         filas = []
#         for c in clientes:
#             valor_docto = Decimal(str(c.valor_docto or 0))
#             total_cop = Decimal(str(c.total_cop or 0))
#             recaudo = valor_docto - total_cop
#             filas.append({
#                 "id": c.id,
#                 "razon_social": c.razon_social,
#                 "nit_cliente": c.nit_cliente,
#                 "nro_docto_cruce": c.nro_docto_cruce,
#                 "dias_vencidos": c.dias_vencidos,
#                 "fecha_docto": c.fecha_docto,
#                 "fecha_vcto": c.fecha_vcto,
#                 "valor_docto": float(valor_docto),
#                 "total_cop": float(total_cop),
#                 "recaudo": float(recaudo),
#                 "asesor": c.asesor,
#                 "correo": c.correo,
#             })
#         return templates.TemplateResponse(
#             "index.html",
#             {
#                 "request": request,
#                 "view": "flat",
#                 "filas": filas,
#                 "flash_msg": flash_msg,
#                 "flash_type": flash_type,
#                 "current_qs": current_qs,
#                 "min_dias": min_dias,
#                 "max_dias": max_dias,
#                 "sort": sort,
#             }
#         )

    # # Vista agrupada (orden estable por NIT/Factura para coherencia visual)
    # agrupados = {}
    # for c in clientes:
    #     if c.nit_cliente not in agrupados:
    #         agrupados[c.nit_cliente] = {
    #             "nit_cliente": c.nit_cliente,
    #             "razon_social": c.razon_social,
    #             "telefono": c.telefono,
    #             "celular": c.celular,
    #             "correo": c.correo,
    #             "asesor": c.asesor,
    #             "facturas": []
    #         }

    #     valor_docto = Decimal(str(c.valor_docto or 0))
    #     total_cop = Decimal(str(c.total_cop or 0))
    #     recaudo = valor_docto - total_cop

    #     agrupados[c.nit_cliente]["facturas"].append({
    #         "id": c.id,
    #         "nro_docto_cruce": c.nro_docto_cruce,
    #         "dias_vencidos": c.dias_vencidos,
    #         "fecha_docto": c.fecha_docto,
    #         "fecha_vcto": c.fecha_vcto,
    #         "valor_docto": float(valor_docto),
    #         "total_cop": float(total_cop),
    #         "recaudo": float(recaudo),
    #         "fecha_gestion": c.fecha_gestion,
    #         "tipo": c.tipo,
    #         "asesor": c.asesor,
    #         "observaciones": [obs.texto for obs in c.observaciones]
    #     })

    # for cl in agrupados.values():
    #     dias = [f["dias_vencidos"] for f in cl["facturas"] if f["dias_vencidos"] is not None]
    #     cl["max_dias"] = max(dias) if dias else None

    # return templates.TemplateResponse(
    #     "index.html",
    #     {
    #         "request": request,
    #         "clientes": list(agrupados.values()),
    #         "flash_msg": flash_msg,
    #         "flash_type": flash_type,
    #         "current_qs": current_qs,
    #         "min_dias": min_dias,
    #         "max_dias": max_dias,
    #         "sort": sort,
    #         "view": view,
    #     }
    # )





# # ------------------- Imports y configuración inicial -------------------

# from fastapi.responses import JSONResponse, RedirectResponse, StreamingResponse
# from fastapi import UploadFile, File, Form, Query, status, FastAPI, Request, Depends, HTTPException
# from fastapi.templating import Jinja2Templates
# from fastapi.staticfiles import StaticFiles
# from sqlalchemy.orm import Session
# from sqlalchemy import text  # SQL crudo MSSQL-safe
# from datetime import datetime
# from typing import Optional, Tuple
# from decimal import Decimal
# from io import BytesIO
# from starlette.datastructures import QueryParams
# import pandas as pd
# import models
# import math
# import crud
# import re

# from database import SessionLocal, engine, Base  # tu módulo existente


# # ------------------- Inicialización -------------------

# app = FastAPI()
# templates = Jinja2Templates(directory="templates")

# app.mount("/static", StaticFiles(directory="static"), name="static")
# app.mount("/assets", StaticFiles(directory="assets"), name="assets")


# # ------------------- Helpers -------------------

# def fmt_money(value):
#     """Filtro para mostrar números como moneda"""
#     if value is None or value == "":
#         return "-"
#     try:
#         v = Decimal(str(value))
#         return f"{v:,.2f}"
#     except Exception:
#         return str(value)

# templates.env.filters["fmt_money"] = fmt_money


# def get_db():
#     db = SessionLocal()
#     try:
#         yield db
#     finally:
#         db.close()


# def to_float(val):
#     try:
#         if pd.isna(val):
#             return None
#         val_str = str(val).replace(",", ".")
#         val_str = re.sub(r"[^0-9.]", "", val_str)
#         return float(val_str) if val_str else None
#     except (ValueError, TypeError):
#         return None


# def clean_phone(val):
#     """Normaliza teléfonos: quita .0, exponentes, separadores; deja solo dígitos."""
#     if val is None or (isinstance(val, float) and pd.isna(val)):
#         return None
#     s = str(val).strip()

#     # 3046304674.0 -> 3046304674
#     if s.endswith(".0") and s[:-2].isdigit():
#         s = s[:-2]

#     # 3.0463E+09 -> 3046300000
#     try:
#         if re.fullmatch(r"\d+(\.\d+)?[eE][+-]?\d+", s):
#             s = str(int(float(s)))
#     except Exception:
#         pass

#     # 1234.00 -> 1234 si es entero exacto
#     if re.fullmatch(r"\d+\.\d+", s):
#         try:
#             f = float(s)
#             if f.is_integer():
#                 s = str(int(f))
#         except Exception:
#             pass

#     # Dejar solo dígitos
#     s = re.sub(r"[^\d]", "", s)
#     return s or None


# def clean_email(val):
#     """
#     Normaliza emails:
#     - Devuelve el PRIMER email válido encontrado (separadores ; , / espacio).
#     - Tolera mayúsculas/ruido.
#     """
#     if val is None or (isinstance(val, float) and pd.isna(val)):
#         return None
#     s = str(val).strip()
#     if not s or s.lower() == "nan":
#         return None

#     # Cortes por separadores comunes
#     candidates = re.split(r"[;,\n/\s]+", s)

#     email_re = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.IGNORECASE)

#     # Primero busca dentro de cada segmento
#     for part in candidates:
#         m = email_re.search(part or "")
#         if m:
#             return m.group(0).strip().lower()

#     # Si no, busca en todo el string
#     m = email_re.search(s)
#     if m:
#         return m.group(0).strip().lower()

#     return None


# # ---- Null-safe parsers para MSSQL (evitan NaN/Timestamp) ----
# def to_int_nullsafe(val):
#     """Convierte a int o devuelve None si viene vacío/NaN."""
#     try:
#         if val is None:
#             return None
#         if isinstance(val, float) and math.isnan(val):  # NaN
#             return None
#         s = str(val).strip()
#         if s == "":
#             return None
#         return int(float(s))  # admite "5.0"
#     except Exception:
#         return None


# def to_dt_nullsafe(val):
#     """Convierte a datetime (py) o None. Acepta pandas Timestamp, str, excel serial."""
#     if val is None:
#         return None
#     try:
#         if isinstance(val, float) and math.isnan(val):
#             return None
#     except Exception:
#         pass
#     ts = pd.to_datetime(val, errors="coerce")
#     if pd.isna(ts):
#         return None
#     return ts.to_pydatetime()


# # ---------- Helpers nuevos para borrado MSSQL-safe ----------
# def _chunk(iterable, size):
#     buf = []
#     for x in iterable:
#         buf.append(x)
#         if len(buf) >= size:
#             yield buf
#             buf = []
#     if buf:
#         yield buf


# def delete_by_pairs_mssql(db: Session, table_name: str, key_pairs, batch_size: int = 900):
#     """
#     MSSQL-safe:
#     DELETE c
#     FROM <tabla> AS c
#     JOIN (VALUES (...), ...) AS v(nit_cliente, nro_docto_cruce)
#       ON v.nit_cliente = c.nit_cliente
#      AND v.nro_docto_cruce = c.nro_docto_cruce;
#     """
#     pairs = list(key_pairs)
#     if not pairs:
#         return

#     for batch in _chunk(pairs, batch_size):
#         values_clause = []
#         params = {}
#         for i, (nit, doc) in enumerate(batch):
#             values_clause.append(f"(:nit{i}, :doc{i})")
#             params[f"nit{i}"] = str(nit) if nit is not None else None
#             params[f"doc{i}"] = str(doc) if doc is not None else None

#         sql = f"""
#         DELETE c
#         FROM {table_name} AS c
#         JOIN (VALUES {", ".join(values_clause)}) AS v(nit_cliente, nro_docto_cruce)
#           ON v.nit_cliente = c.nit_cliente
#          AND v.nro_docto_cruce = c.nro_docto_cruce;
#         """
#         db.execute(text(sql), params)


# # ------------------- Helpers de filtros/orden y navegación coherente -------------------

# def _to_int_or_none(val: Optional[str]):
#     try:
#         return int(val) if val not in (None, "") else None
#     except (ValueError, TypeError):
#         return None


# def _parse_filters(request_like) -> Tuple[Optional[int], Optional[int], str, Optional[str]]:
#     """
#     Extrae min_dias, max_dias, sort, view desde query_params (del request o de ref)
#     """
#     qp = request_like.query_params
#     min_val = _to_int_or_none(qp.get("min_dias"))
#     max_val = _to_int_or_none(qp.get("max_dias"))
#     sort = qp.get("sort", "dias_desc")
#     view = qp.get("view")
#     return min_val, max_val, sort, view


# def _apply_filters_and_order(db: Session, min_val: Optional[int], max_val: Optional[int], sort: str, view: Optional[str]):
#     """
#     Devuelve un Query con el mismo filtrado/orden que el index.
#     Si la vista es 'group' y no se especifica nit_asc/razon_asc, forzamos nit_asc
#     para que la navegación sea estable por NIT (coherente con el acordeón).
#     """
#     q = db.query(models.Cliente)

#     if min_val is not None:
#         q = q.filter(models.Cliente.dias_vencidos >= min_val)
#     if max_val is not None:
#         q = q.filter(models.Cliente.dias_vencidos <= max_val)

#     effective_sort = sort
#     if (view == "group") and (sort not in {"nit_asc", "razon_asc"}):
#         effective_sort = "nit_asc"

#     if effective_sort == "dias_asc":
#         q = q.order_by(
#             models.Cliente.dias_vencidos.asc(),
#             models.Cliente.nit_cliente.asc(),
#             models.Cliente.nro_docto_cruce.asc(),
#             models.Cliente.id.asc(),
#         )
#     elif effective_sort == "razon_asc":
#         q = q.order_by(
#             models.Cliente.razon_social.asc(),
#             models.Cliente.nit_cliente.asc(),
#             models.Cliente.nro_docto_cruce.asc(),
#             models.Cliente.id.asc(),
#         )
#     elif effective_sort == "nit_asc":
#         q = q.order_by(
#             models.Cliente.nit_cliente.asc(),
#             models.Cliente.nro_docto_cruce.asc(),
#             models.Cliente.fecha_docto.asc(),
#             models.Cliente.id.asc(),
#         )
#     else:  # dias_desc (default)
#         q = q.order_by(
#             models.Cliente.dias_vencidos.desc(),
#             models.Cliente.nit_cliente.asc(),
#             models.Cliente.nro_docto_cruce.asc(),
#             models.Cliente.id.asc(),
#         )
#     return q


# # ------------------- Exportar cartera a Excel -------------------
# @app.get("/exportar_cartera.xlsx", name="exportar_cartera_xlsx")
# def exportar_cartera_xlsx(
#     db: Session = Depends(get_db),
#     min_dias: Optional[str] = Query(None),
#     max_dias: Optional[str] = Query(None),
#     sort: Optional[str] = Query("dias_desc"),
# ):
#     def to_int_or_none(val: Optional[str]):
#         try:
#             return int(val) if val not in (None, "") else None
#         except (ValueError, TypeError):
#             return None

#     min_val = to_int_or_none(min_dias)
#     max_val = to_int_or_none(max_dias)

#     query = db.query(models.Cliente)
#     if min_val is not None:
#         query = query.filter(models.Cliente.dias_vencidos >= min_val)
#     if max_val is not None:
#         query = query.filter(models.Cliente.dias_vencidos <= max_val)

#     order_map = {
#         "dias_desc": models.Cliente.dias_vencidos.desc(),
#         "dias_asc": models.Cliente.dias_vencidos.asc(),
#         "razon_asc": models.Cliente.razon_social.asc(),
#     }
#     query = query.order_by(order_map.get(sort, order_map["dias_desc"]))
#     clientes = query.all()

#     rows = []
#     for c in clientes:
#         valor_docto = Decimal(str(c.valor_docto or 0))
#         total_cop = Decimal(str(c.total_cop or 0))
#         recaudo = Decimal(str(c.recaudo)) if c.recaudo is not None else (valor_docto - total_cop)

#         # ÚLTIMA observación (si existe) para export
#         obs_txt = ""
#         if getattr(c, "observaciones", None):
#             latest = None
#             for o in c.observaciones:
#                 f = o.fecha_creacion or datetime.min
#                 if latest is None or f > (latest.fecha_creacion or datetime.min):
#                     latest = o
#             if latest:
#                 obs_txt = f"{(latest.fecha_creacion.strftime('%Y-%m-%d %H:%M') if latest.fecha_creacion else '')} - {latest.texto}"

#         rows.append({
#             "ID": c.id,
#             "Razón social": c.razon_social,
#             "NIT": c.nit_cliente,
#             "Docto cruce": c.nro_docto_cruce,
#             "Días vencidos": c.dias_vencidos,
#             "Fecha docto": c.fecha_docto,
#             "Fecha vcto": c.fecha_vcto,
#             "Valor docto": float(valor_docto),
#             "Total COP (saldo)": float(total_cop),
#             "Recaudo": float(recaudo),
#             "Teléfono": c.telefono,
#             "Celular": c.celular,
#             "Correo": c.correo,
#             "Asesor": c.asesor,
#             "Fecha gestión": c.fecha_gestion,
#             "Tipo": c.tipo,
#             "Observaciones": obs_txt,
#         })

#     df = pd.DataFrame(rows)

#     # Normaliza fechas a datetime (Excel-friendly)
#     for col in ["Fecha docto", "Fecha vcto", "Fecha gestión"]:
#         if col in df.columns:
#             df[col] = pd.to_datetime(df[col], errors="coerce")

#     # --- Resumen por cliente ---
#     if not df.empty:
#         resumen = (
#             df.groupby(["NIT", "Razón social"], dropna=False)
#               .agg({
#                   "Valor docto": "sum",
#                   "Total COP (saldo)": "sum",
#                   "Recaudo": "sum",
#                   "Días vencidos": "max",
#               })
#               .rename(columns={
#                   "Valor docto": "Total Valor Docto",
#                   "Total COP (saldo)": "Saldo Total",
#                   "Recaudo": "Recaudo Total",
#                   "Días vencidos": "Max Días Vencidos",
#               })
#               .reset_index()
#         )
#         ref = (
#             df.sort_values(["NIT"]).groupby(["NIT", "Razón social"], dropna=False)
#               .agg({"Teléfono": "first", "Celular": "first", "Correo": "first", "Asesor": "first"})
#               .reset_index()
#         )
#         resumen = resumen.merge(ref, on=["NIT", "Razón social"], how="left")

#         facturas = (
#             df.groupby(["NIT", "Razón social"], dropna=False)
#               .size()
#               .reset_index(name="# Facturas")
#         )
#         resumen = resumen.merge(facturas, on=["NIT", "Razón social"], how="left")
#         resumen["# Facturas"] = resumen["# Facturas"].fillna(0).astype(int)
#     else:
#         resumen = pd.DataFrame(columns=[
#             "NIT","Razón social","Total Valor Docto","Saldo Total","Recaudo Total",
#             "Max Días Vencidos","Teléfono","Celular","Correo","Asesor","# Facturas"
#         ])

#     # --- Excel en memoria con formato ---
#     output = BytesIO()
#     with pd.ExcelWriter(output, engine="xlsxwriter",
#                         datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as writer:
#         df.to_excel(writer, index=False, sheet_name="cartera")
#         resumen.to_excel(writer, index=False, sheet_name="resumen")

#         wb = writer.book
#         ws1 = writer.sheets["cartera"]
#         ws2 = writer.sheets["resumen"]

#         for ws in (ws1, ws2):
#             ws.freeze_panes(1, 0)

#         fmt_header = wb.add_format({
#             'bold': True, 'bg_color': '#D9E1F2', 'border': 1,
#             'align': 'center', 'valign': 'vcenter'
#         })
#         fmt_money0 = wb.add_format({'num_format': '#,##0', 'border': 1})
#         fmt_int = wb.add_format({'num_format': '#,##0', 'border': 1})
#         fmt_date = wb.add_format({'num_format': 'yyyy-mm-dd', 'border': 1})
#         fmt_text = wb.add_format({'text_wrap': True, 'border': 1, 'valign': 'top'})
#         fmt_default = wb.add_format({'border': 1})

#         def table_with_formats(ws, df_sheet, money_cols, date_cols, wrap_cols):
#             """Crea tabla con formatos por columna y calcula anchos."""
#             rows, cols = df_sheet.shape
#             headers = list(df_sheet.columns)

#             # Mapear formato por columna
#             per_col_format = {}
#             for col in headers:
#                 if col in money_cols:
#                     per_col_format[col] = fmt_money0
#                 elif col in date_cols:
#                     per_col_format[col] = fmt_date
#                 elif col in wrap_cols:
#                     per_col_format[col] = fmt_text
#                 elif df_sheet[col].dtype.kind in ("i", "u"):  # enteros
#                     per_col_format[col] = fmt_int
#                 else:
#                     per_col_format[col] = fmt_default

#             # Definir columnas de la tabla con formato (esto sí aplica dentro de la tabla)
#             columns_def = [{'header': h, 'format': per_col_format[h]} for h in headers]

#             # Crear tabla (si hay datos). Si no, solo pintamos encabezados.
#             if rows > 0:
#                 ws.add_table(0, 0, rows, cols - 1, {
#                     'style': 'Table Style Medium 9',
#                     'banded_rows': True,
#                     'columns': columns_def
#                 })
#             else:
#                 for c_idx, h in enumerate(headers):
#                     ws.write(0, c_idx, h, fmt_header)

#             # Auto–ancho aproximado por columna
#             col_widths = {}
#             for c_idx, col in enumerate(headers):
#                 max_len = len(str(col)) + 2
#                 serie = df_sheet[col]

#                 if col in money_cols:
#                     # medir como miles sin decimales
#                     for v in serie.dropna():
#                         try:
#                             s = f"{int(round(float(v))):,}"
#                             max_len = max(max_len, len(s))
#                         except Exception:
#                             pass
#                 elif col in date_cols:
#                     max_len = max(max_len, 12)
#                 else:
#                     # considerar saltos de línea (tomar la línea más larga)
#                     for v in serie.dropna():
#                         s = str(v).replace("\r", "")
#                         max_len = max(max_len, max((len(seg) for seg in s.split("\n")), default=0))
#                     # límite práctico
#                     if col in wrap_cols:
#                         max_len = min(max(max_len, 40), 80)
#                     else:
#                         max_len = min(max_len, 40)

#                 width = max_len + 1
#                 col_widths[col] = width
#                 ws.set_column(c_idx, c_idx, width, per_col_format[col])

#             return col_widths, per_col_format

#         def autofit_row_heights(ws, df_sheet, wrap_cols, col_widths):
#             """Ajusta alto de fila según # líneas estimadas en columnas wrap."""
#             if df_sheet.empty or not wrap_cols:
#                 return
#             base_height = 15  # pts aprox por línea
#             max_height = 300

#             for r in range(1, len(df_sheet) + 1):  # +1 por header
#                 max_lines = 1
#                 for col in wrap_cols:
#                     if col not in df_sheet.columns:
#                         continue
#                     val = df_sheet.iloc[r - 1][col]
#                     if pd.isna(val) or val is None:
#                         continue
#                     text = str(val).replace("\r", "")
#                     if text == "":
#                         continue

#                     # Estimar caracteres por línea a partir del ancho de columna
#                     col_w = int(col_widths.get(col, 40))
#                     chars_per_line = max(col_w - 2, 10)

#                     total_lines = 0
#                     for seg in text.split("\n"):
#                         seg = seg.strip()
#                         if seg == "":
#                             total_lines += 1
#                         else:
#                             total_lines += math.ceil(len(seg) / chars_per_line)

#                     max_lines = max(max_lines, total_lines)

#                 ws.set_row(r, min(base_height * max_lines + 4, max_height))

#         # Columnas por tipo (cartera)
#         money_cols_cartera = ["Valor docto", "Total COP (saldo)", "Recaudo"]
#         date_cols_cartera = ["Fecha docto", "Fecha vcto", "Fecha gestión"]
#         wrap_cols_cartera = ["Razón social", "Observaciones"]  # ⇦ ajustar ancho/alto por texto

#         # Columnas por tipo (resumen)
#         money_cols_resumen = ["Total Valor Docto", "Saldo Total", "Recaudo Total"]
#         date_cols_resumen = []
#         wrap_cols_resumen = ["Razón social"]

#         # Aplicar tabla + formatos + anchos
#         widths_cartera, _ = table_with_formats(ws1, df, money_cols_cartera, date_cols_cartera, wrap_cols_cartera)
#         widths_resumen, _ = table_with_formats(ws2, resumen, money_cols_resumen, date_cols_resumen, wrap_cols_resumen)

#         # Ajustar alturas (solo donde hay wrap)
#         autofit_row_heights(ws1, df, wrap_cols_cartera, widths_cartera)
#         autofit_row_heights(ws2, resumen, wrap_cols_resumen, widths_resumen)

#     output.seek(0)
#     filename = f"cartera_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
#     return StreamingResponse(
#         output,
#         media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#         headers={"Content-Disposition": f'attachment; filename="{filename}"'}
#     )


# # ------------------- Importar cartera -------------------
# @app.post("/importar_excel")
# async def importar_excel(file: UploadFile = File(...), db: Session = Depends(get_db)):
#     try:
#         contents = await file.read()

#         # --- Cargar Excel/CSV (ambos formatos) ---
#         def _read_any(contents: bytes, filename: str):
#             name = (filename or "").lower()
#             if name.endswith(".csv"):
#                 df = pd.read_csv(BytesIO(contents), dtype=str, keep_default_na=False)
#                 return {"__csv__": df}
#             return pd.read_excel(BytesIO(contents), sheet_name=None, dtype=str, keep_default_na=False)

#         sheets = _read_any(contents, file.filename)

#         # --- Normalizador simple de strings ---
#         def norm(s: str) -> str:
#             if s is None:
#                 return ""
#             t = str(s).strip().lower()
#             rep = {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ñ":"n","ä":"a","ë":"e","ï":"i","ö":"o","ü":"u","’":"'","´":"'","`":"'"}
#             for a,b in rep.items(): t = t.replace(a,b)
#             t = re.sub(r"\s+", " ", t)
#             return t

#         # --- Aliases (internos -> alias) ---
#         alias_map = {
#             "razon_social": {"razon social", "razon social vend. cliente", "razon social vend cliente"},
#             "nit_cliente": {"nit", "nit cliente despacho", "nit cliente"},
#             "nro_docto_cruce": {"docto cruce", "nro. docto. cruce", "nro docto cruce", "factura", "documento", "nro documento"},
#             "fecha_docto": {"fecha docto", "fecha docto."},
#             "fecha_vcto": {"fecha vcto", "fecha vcto."},
#             "dias_vencidos": {"dias vencidos", "dias vencidos."},
#             "valor_docto": {"valor docto", "valor docto."},
#             "total_cop": {"total cop", "total cop (saldo)", "saldo", "total (cop)"},
#             "telefono": {"telefono", "teléfono"},
#             "celular": {"celular", "movil", "móvil"},
#             "correo": {"correo", "email", "e-mail", "mail", "correo electronico", "correo electrónico"},
#             "asesor": {"asesor", "vendedor", "razon social vend. cliente", "razon social vend cliente"},
#             "obs_txt": {"observaciones"},
#         }

#         # --- Detectar si es export del proyecto (hoja 'cartera' con NIT/Docto) ---
#         is_project_export = False
#         df_project = None
#         for sheet_name, df in sheets.items():
#             if norm(sheet_name) == "cartera":
#                 df_project = df
#                 break
#         if df_project is not None:
#             cols_n = {norm(c) for c in df_project.columns}
#             if "nit" in cols_n and ("docto cruce" in cols_n or "nro. docto. cruce" in cols_n or "nro docto cruce" in cols_n):
#                 is_project_export = True

#         # --- Elegir hoja candidata ---
#         if is_project_export:
#             df_raw = df_project.copy()
#         else:
#             df_raw = None
#             for _, df in sheets.items():
#                 cols_n = {norm(c) for c in df.columns}
#                 if any(x in cols_n for x in alias_map["nit_cliente"]) and \
#                    any(x in cols_n for x in alias_map["nro_docto_cruce"]):
#                     df_raw = df
#                     break
#             if df_raw is None:
#                 return RedirectResponse(
#                     url="/?msg=No%20se%20encontro%20hoja%20compatible%20(ni%20cartera%20ni%20columnas%20clave)&msg_type=error",
#                     status_code=303
#                 )

#         # Copia de encabezados originales por si necesitamos "rescatar" correo
#         original_cols = list(df_raw.columns)

#         df_raw.columns = [str(c).strip() for c in df_raw.columns]

#         # --- Renombrado dinámico por alias ---
#         rename_map_dynamic = {}
#         for col in df_raw.columns:
#             n = norm(col)
#             for internal, aliases in alias_map.items():
#                 if n in aliases and internal not in rename_map_dynamic.values():
#                     rename_map_dynamic[col] = internal
#                     break

#         # Encabezados exactos del export del proyecto
#         project_head_map = {
#             "Razón social": "razon_social",
#             "NIT": "nit_cliente",
#             "Docto cruce": "nro_docto_cruce",
#             "Días vencidos": "dias_vencidos",
#             "Fecha docto": "fecha_docto",
#             "Fecha vcto": "fecha_vcto",
#             "Valor docto": "valor_docto",
#             "Total COP (saldo)": "total_cop",
#             "Teléfono": "telefono",
#             "Celular": "celular",
#             "Correo": "correo",
#             "Email": "correo",
#             "E-mail": "correo",
#             "Asesor": "asesor",
#             "Observaciones": "obs_txt",
#         }
#         for col in df_raw.columns:
#             if col in project_head_map and project_head_map[col] not in rename_map_dynamic.values():
#                 rename_map_dynamic[col] = project_head_map[col]

#         df = df_raw.rename(columns=rename_map_dynamic)

#         # --- Asegurar/Coalescer columna 'correo' desde cualquier variante ---
#         def _norm2(s: str) -> str:
#             if s is None:
#                 return ""
#             t = str(s).strip().lower()
#             rep = {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ñ":"n"}
#             for a,b in rep.items(): t = t.replace(a,b)
#             t = re.sub(r"\s+", " ", t)
#             return t

#         email_cols_candidates = []
#         for col in df_raw.columns:
#             n = _norm2(col)
#             if re.fullmatch(r"(correo( electronico)?)|(e[\s-]?mail)|email|mail", n):
#                 email_cols_candidates.append(col)

#         if "correo" not in df.columns:
#             if email_cols_candidates:
#                 correo_series = None
#                 for col in email_cols_candidates:
#                     s = df_raw[col].astype(str)
#                     s = s.where(~s.str.fullmatch(r"(?i)\s*nan\s*"), "")
#                     if correo_series is None:
#                         correo_series = s
#                     else:
#                         correo_series = correo_series.mask(
#                             (correo_series.astype(str).str.strip() == "") & (s.astype(str).str.strip() != ""),
#                             s
#                         )
#                 df["correo"] = correo_series
#             else:
#                 df["correo"] = None

#         # Normalizar emails al final
#         df["correo"] = df["correo"].apply(clean_email)

#         # Validar claves mínimas
#         if "nit_cliente" not in df.columns or "nro_docto_cruce" not in df.columns:
#             return RedirectResponse(
#                 url="/?msg=El%20archivo%20no%20tiene%20las%20columnas%20clave%20(NIT%20y%20Docto%20cruce)&msg_type=error",
#                 status_code=303
#             )

#         # --- Normalizaciones de campos ---
#         if "telefono" in df.columns: df["telefono"] = df["telefono"].apply(clean_phone)
#         if "celular" in df.columns: df["celular"] = df["celular"].apply(clean_phone)

#         if "fecha_docto" in df.columns: df["fecha_docto"] = df["fecha_docto"].apply(to_dt_nullsafe)
#         if "fecha_vcto"  in df.columns: df["fecha_vcto"]  = df["fecha_vcto"].apply(to_dt_nullsafe)
#         if "dias_vencidos" in df.columns: df["dias_vencidos"] = df["dias_vencidos"].apply(to_int_nullsafe)

#         def _to_money(x):
#             v = to_float(x)
#             return v if v is not None else 0.0

#         if "valor_docto" in df.columns:
#             df["valor_docto"] = df["valor_docto"].apply(_to_money)
#         else:
#             df["valor_docto"] = 0.0

#         if "total_cop" in df.columns:
#             df["total_cop"] = df["total_cop"].apply(_to_money)
#         else:
#             df["total_cop"] = df["valor_docto"]

#         # Claves limpias
#         df = df[(df["nit_cliente"].notna()) & (df["nro_docto_cruce"].notna())]
#         df["nit_cliente"] = df["nit_cliente"].astype(str).str.strip()
#         df["nro_docto_cruce"] = df["nro_docto_cruce"].astype(str).str.strip()
#         df = df[(df["nit_cliente"] != "") & (df["nro_docto_cruce"] != "")]

#         if df.empty:
#             return RedirectResponse(url="/?msg=Archivo%20vacio%20o%20sin%20filas%20validas&msg_type=info", status_code=303)

#         # --- Claves existentes en BD ---
#         bd_clientes = db.query(models.Cliente).all()
#         bd_claves = set((str(c.nit_cliente), str(c.nro_docto_cruce)) for c in bd_clientes)

#         # --- Borrado condicional ---
#         if not is_project_export:
#             excel_claves = set(zip(df["nit_cliente"], df["nro_docto_cruce"]))
#             claves_a_eliminar = bd_claves - excel_claves
#             if claves_a_eliminar:
#                 delete_by_pairs_mssql(db, "cartera", claves_a_eliminar, batch_size=900)

#         # Helpers de obs
#         def _clean_obs_text(s: Optional[str]) -> Optional[str]:
#             if not s:
#                 return None
#             txt = str(s).strip()
#             m = re.match(r"^\s*(\d{4}-\d{2}-\d{2}(?:\s+\d{2}:\d{2})?)\s*-\s*(.+)$", txt)
#             if m:
#                 return m.group(2).strip() or None
#             return txt or None

#         def _latest_obs_text(cliente):
#             latest = None
#             for o in getattr(cliente, "observaciones", []) or []:
#                 f = getattr(o, "fecha_creacion", None) or datetime.min
#                 if latest is None or f > (getattr(latest, "fecha_creacion", None) or datetime.min):
#                     latest = o
#             return (latest.texto.strip() if latest and latest.texto else None)

#         # --- Upsert + Observaciones (solo si viene del export) ---
#         for _, row in df.iterrows():
#             nit = row["nit_cliente"]
#             docto = row["nro_docto_cruce"]

#             valor_docto = row.get("valor_docto", 0.0) or 0.0
#             total_excel = row.get("total_cop", valor_docto)
#             recaudo_calc = round(max((valor_docto or 0.0) - (total_excel or 0.0), 0.0), 2)

#             cliente = db.query(models.Cliente).filter_by(
#                 nit_cliente=nit,
#                 nro_docto_cruce=docto
#             ).first()

#             correo_val = clean_email(row.get("correo")) if "correo" in row else None

#             if cliente:
#                 cliente.razon_social = row.get("razon_social")
#                 cliente.dias_vencidos = to_int_nullsafe(row.get("dias_vencidos"))
#                 cliente.fecha_docto = to_dt_nullsafe(row.get("fecha_docto"))
#                 cliente.fecha_vcto = to_dt_nullsafe(row.get("fecha_vcto"))
#                 cliente.valor_docto = valor_docto
#                 cliente.total_cop = total_excel
#                 cliente.recaudo = recaudo_calc
#                 cliente.telefono = clean_phone(row.get("telefono"))
#                 cliente.celular  = clean_phone(row.get("celular"))
#                 cliente.asesor = row.get("asesor")
#                 if correo_val and (not cliente.correo or cliente.correo != correo_val):
#                     cliente.correo = correo_val
#             else:
#                 cliente = models.Cliente(
#                     razon_social=row.get("razon_social"),
#                     nit_cliente=nit,
#                     nro_docto_cruce=docto,
#                     dias_vencidos=to_int_nullsafe(row.get("dias_vencidos")),
#                     fecha_docto=to_dt_nullsafe(row.get("fecha_docto")),
#                     fecha_vcto=to_dt_nullsafe(row.get("fecha_vcto")),
#                     valor_docto=valor_docto,
#                     total_cop=total_excel,
#                     recaudo=recaudo_calc,
#                     telefono=clean_phone(row.get("telefono")),
#                     celular=clean_phone(row.get("celular")),
#                     asesor=row.get("asesor"),
#                     correo=correo_val,
#                 )
#                 db.add(cliente)
#                 db.flush()

#             if "obs_txt" in df.columns and is_project_export:
#                 new_obs_text = _clean_obs_text(row.get("obs_txt"))
#                 if new_obs_text:
#                     last_text = _latest_obs_text(cliente)
#                     if (last_text or "").strip() != new_obs_text.strip():
#                         obs = models.Observacion(
#                             cliente_id=cliente.id,
#                             texto=new_obs_text,
#                         )
#                         db.add(obs)

#         db.commit()
#         return RedirectResponse(url="/?msg=Archivo%20subido%20correctamente&msg_type=success", status_code=303)

#     except Exception as e:
#         print("❌ Error importar_excel:", e)
#         return RedirectResponse(url="/?msg=Error%20al%20subir%20el%20archivo&msg_type=error", status_code=303)


# # ------------------- Observaciones -------------------
# @app.post("/cliente/{cliente_id}/observacion")
# def agregar_observacion(cliente_id: int, texto: str = Form(...), db: Session = Depends(get_db)):
#     crud.add_observacion(db, cliente_id, texto)
#     return RedirectResponse(
#         url=f"/cliente/{cliente_id}?msg=Observaci%C3%B3n%20guardada&msg_type=success",
#         status_code=303
#     )


# # ------------------- Actualizar cliente (editar recaudo) -------------------
# @app.post("/cliente/{cliente_id}/update")
# async def update_cliente(
#     request: Request,
#     cliente_id: int,
#     db: Session = Depends(get_db)
# ):
#     form_data = await request.form()

#     cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
#     if not cliente:
#         return RedirectResponse(url="/?msg=Cliente%20no%20encontrado&msg_type=error", status_code=status.HTTP_303_SEE_OTHER)

#     for campo in ["telefono", "celular", "correo", "fecha_gestion", "tipo", "asesor"]:
#         val = form_data.get(campo)
#         if val is not None and val != "":
#             if campo == "correo":
#                 val = clean_email(val)
#             if campo in ("telefono", "celular"):
#                 val = clean_phone(val)
#             setattr(cliente, campo, val)

#     valor_docto = float(cliente.valor_docto or 0.0)

#     raw_recaudo = form_data.get("recaudo")
#     raw_total_cop = form_data.get("total_cop")

#     has_recaudo = raw_recaudo is not None and raw_recaudo.strip() != ""
#     has_total_cop = raw_total_cop is not None and raw_total_cop.strip() != ""

#     if has_recaudo:
#         # Interpreta recaudo como ABONO (delta)
#         delta = to_float(raw_recaudo)
#         if delta is not None:
#             old_total = float(cliente.total_cop or 0.0)
#             new_total = max(round(old_total - delta, 2), 0.0)
#             cliente.total_cop = new_total
#             cliente.recaudo = round(max(valor_docto - new_total, 0.0), 2)
#     elif has_total_cop:
#         # Ajuste explícito de saldo total_cop => recalcular recaudo
#         nuevo_total = to_float(raw_total_cop)
#         if nuevo_total is not None:
#             nuevo_total = max(nuevo_total, 0.0)
#             # No dejar que el saldo supere el valor del documento
#             if nuevo_total > valor_docto:
#                 nuevo_total = valor_docto
#             cliente.total_cop = round(nuevo_total, 2)
#             cliente.recaudo = round(max(valor_docto - nuevo_total, 0.0), 2)

#     nueva_obs = form_data.get("observaciones")
#     if nueva_obs and nueva_obs.strip():
#         obs = models.Observacion(
#             cliente_id=cliente.id,
#             texto=nueva_obs.strip(),
#         )
#         db.add(obs)

#     db.commit()
#     return RedirectResponse(
#         url=f"/cliente/{cliente_id}?msg=Cambios%20guardados&msg_type=success",
#         status_code=status.HTTP_303_SEE_OTHER
#     )


# # ------------------- historial cliente -------------------
# @app.get("/cliente/{cliente_id}/historial")
# def historial_cliente(cliente_id: int, db: Session = Depends(get_db)):
#     cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
#     if not cliente:
#         return JSONResponse(content=[], status_code=200)

#     historial = [
#         {
#             "texto": obs.texto,
#             "fecha": obs.fecha_creacion.strftime("%d/%m/%Y %H:%M")
#             if obs.fecha_creacion else None
#         }
#         for obs in cliente.observaciones
#     ]

#     return JSONResponse(content=historial, status_code=200)


# # ------------------- Vista cliente -------------------
# @app.get("/cliente/{cliente_id}")
# def ver_cliente(cliente_id: int, request: Request, db: Session = Depends(get_db)):
#     cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
#     if not cliente:
#         raise HTTPException(status_code=404, detail="Cliente no encontrado")

#     # Reconstruye filtros desde ref (si existe) para navegar en el MISMO universo y orden
#     ref_qs = request.query_params.get("ref")
#     req_like = request
#     if ref_qs:
#         # Creamos un objeto "request-like" solo con query_params = QueryParams(ref_qs)
#         req_like = type("R", (), {"query_params": QueryParams(ref_qs)})

#     min_val, max_val, sort, view = _parse_filters(req_like)

#     # Lista ordenada de ids con la MISMA lógica del index
#     ordered_ids = [
#         row.id for row in _apply_filters_and_order(db, min_val, max_val, sort, view)
#                     .with_entities(models.Cliente.id).all()
#     ]

#     # Ubica la posición actual y calcula prev/sig basados en la lista ordenada
#     prev_id = next_id = None
#     try:
#         idx = ordered_ids.index(cliente_id)
#         if idx > 0:
#             prev_id = ordered_ids[idx - 1]
#         if idx < len(ordered_ids) - 1:
#             next_id = ordered_ids[idx + 1]
#     except ValueError:
#         # Si el id no está en la lista filtrada (raro), caemos a navegación por id
#         prev_cliente = db.query(models.Cliente).filter(models.Cliente.id < cliente_id).order_by(models.Cliente.id.desc()).first()
#         next_cliente = db.query(models.Cliente).filter(models.Cliente.id > cliente_id).order_by(models.Cliente.id.asc()).first()
#         prev_id = prev_cliente.id if prev_cliente else None
#         next_id = next_cliente.id if next_cliente else None

#     back_url = f"/?{ref_qs}" if ref_qs else "/"

#     return templates.TemplateResponse(
#         "cliente.html",
#         {
#             "request": request,
#             "cliente": cliente,
#             "prev_id": prev_id,
#             "next_id": next_id,
#             "flash_msg": request.query_params.get("msg"),
#             "flash_type": request.query_params.get("msg_type"),
#             "back_url": back_url,
#         }
#     )


# # ------------------- Index agrupado/plano -------------------
# @app.get("/")
# def index(
#     request: Request,
#     db: Session = Depends(get_db),
#     view: Optional[str] = Query(None),
#     min_dias: Optional[str] = Query(None),
#     max_dias: Optional[str] = Query(None),
#     sort: Optional[str] = Query("dias_desc"),
# ):
#     # Reconstruye filtros y aplica MISMO orden que usaremos en ver_cliente()
#     min_val = _to_int_or_none(min_dias)
#     max_val = _to_int_or_none(max_dias)

#     # NOTA: pasamos "view" al motor de orden para que 'group' fuerce nit_asc por defecto
#     query = _apply_filters_and_order(db, min_val, max_val, sort, view)
#     clientes = query.all()

#     # Flash
#     flash_msg = request.query_params.get("msg")
#     flash_type = request.query_params.get("msg_type")
#     if (min_val is not None) or (max_val is not None):
#         if not flash_msg:
#             flash_msg = "Filtros aplicados"
#         if not flash_type:
#             flash_type = "info"

#     current_qs = request.url.query  # para mantener filtros en hrefs

#     # Vista plana
#     if view == "flat":
#         filas = []
#         for c in clientes:
#             valor_docto = Decimal(str(c.valor_docto or 0))
#             total_cop = Decimal(str(c.total_cop or 0))
#             recaudo = valor_docto - total_cop
#             filas.append({
#                 "id": c.id,
#                 "razon_social": c.razon_social,
#                 "nit_cliente": c.nit_cliente,
#                 "nro_docto_cruce": c.nro_docto_cruce,
#                 "dias_vencidos": c.dias_vencidos,
#                 "fecha_docto": c.fecha_docto,
#                 "fecha_vcto": c.fecha_vcto,
#                 "valor_docto": float(valor_docto),
#                 "total_cop": float(total_cop),
#                 "recaudo": float(recaudo),
#                 "asesor": c.asesor,
#                 "correo": c.correo,
#             })
#         return templates.TemplateResponse(
#             "index.html",
#             {
#                 "request": request,
#                 "view": "flat",
#                 "filas": filas,
#                 "flash_msg": flash_msg,
#                 "flash_type": flash_type,
#                 "current_qs": current_qs,
#                 "min_dias": min_dias,
#                 "max_dias": max_dias,
#                 "sort": sort,
#             }
#         )

#     # Vista agrupada (orden estable por NIT/Factura para coherencia visual)
#     agrupados = {}
#     for c in clientes:
#         if c.nit_cliente not in agrupados:
#             agrupados[c.nit_cliente] = {
#                 "nit_cliente": c.nit_cliente,
#                 "razon_social": c.razon_social,
#                 "telefono": c.telefono,
#                 "celular": c.celular,
#                 "correo": c.correo,
#                 "asesor": c.asesor,
#                 "facturas": []
#             }

#         valor_docto = Decimal(str(c.valor_docto or 0))
#         total_cop = Decimal(str(c.total_cop or 0))
#         recaudo = valor_docto - total_cop

#         agrupados[c.nit_cliente]["facturas"].append({
#             "id": c.id,
#             "nro_docto_cruce": c.nro_docto_cruce,
#             "dias_vencidos": c.dias_vencidos,
#             "fecha_docto": c.fecha_docto,
#             "fecha_vcto": c.fecha_vcto,
#             "valor_docto": float(valor_docto),
#             "total_cop": float(total_cop),
#             "recaudo": float(recaudo),
#             "fecha_gestion": c.fecha_gestion,
#             "tipo": c.tipo,
#             "asesor": c.asesor,
#             "observaciones": [obs.texto for obs in c.observaciones]
#         })

#     for cl in agrupados.values():
#         dias = [f["dias_vencidos"] for f in cl["facturas"] if f["dias_vencidos"] is not None]
#         cl["max_dias"] = max(dias) if dias else None

#     return templates.TemplateResponse(
#         "index.html",
#         {
#             "request": request,
#             "clientes": list(agrupados.values()),
#             "flash_msg": flash_msg,
#             "flash_type": flash_type,
#             "current_qs": current_qs,
#             "min_dias": min_dias,
#             "max_dias": max_dias,
#             "sort": sort,
#             "view": view,
#         }
#     )



# # ------------------- Imports y configuración 4 -------------------

# from fastapi.responses import JSONResponse, RedirectResponse, StreamingResponse
# from fastapi import UploadFile, File, Form, Query, status, FastAPI, Request, Depends, HTTPException
# from fastapi.templating import Jinja2Templates
# from fastapi.staticfiles import StaticFiles
# from sqlalchemy.orm import Session
# from sqlalchemy import text
# from database import SessionLocal, engine, Base
# from datetime import datetime
# from typing import Optional
# from decimal import Decimal
# from io import BytesIO
# from starlette.datastructures import QueryParams
# import pandas as pd
# import models
# import crud
# import math
# import re

# # ------------------- Inicialización -------------------

# app = FastAPI()
# templates = Jinja2Templates(directory="templates")

# app.mount("/static", StaticFiles(directory="static"), name="static")
# app.mount("/assets", StaticFiles(directory="assets"), name="assets")

# # ------------------- Helpers -------------------

# def fmt_money(value):
#     if value is None or value == "":
#         return "-"
#     try:
#         v = Decimal(str(value))
#         return f"{v:,.2f}"
#     except Exception:
#         return str(value)

# templates.env.filters["fmt_money"] = fmt_money


# def get_db():
#     db = SessionLocal()
#     try:
#         yield db
#     finally:
#         db.close()


# def to_float(val):
#     try:
#         if pd.isna(val):
#             return None
#         val_str = str(val).replace(",", ".")
#         val_str = re.sub(r"[^0-9.]", "", val_str)
#         return float(val_str) if val_str else None
#     except (ValueError, TypeError):
#         return None


# def clean_phone(val):
#     if val is None or (isinstance(val, float) and pd.isna(val)):
#         return None
#     s = str(val).strip()

#     if s.endswith(".0") and s[:-2].isdigit():
#         s = s[:-2]
#     try:
#         if re.fullmatch(r"\d+(\.\d+)?[eE][+-]?\d+", s):
#             s = str(int(float(s)))
#     except Exception:
#         pass
#     if re.fullmatch(r"\d+\.\d+", s):
#         try:
#             f = float(s)
#             if f.is_integer():
#                 s = str(int(f))
#         except Exception:
#             pass

#     s = re.sub(r"[^\d]", "", s)
#     return s or None


# # ---------- Helpers nuevos para borrado MSSQL-safe ----------
# def _chunk(iterable, size):
#     buf = []
#     for x in iterable:
#         buf.append(x)
#         if len(buf) >= size:
#             yield buf
#             buf = []
#     if buf:
#         yield buf


# def delete_by_pairs_mssql(db: Session, table_name: str, key_pairs, batch_size: int = 900):
#     pairs = list(key_pairs)
#     if not pairs:
#         return
#     for batch in _chunk(pairs, batch_size):
#         values_clause = []
#         params = {}
#         for i, (nit, doc) in enumerate(batch):
#             values_clause.append(f"(:nit{i}, :doc{i})")
#             params[f"nit{i}"] = str(nit) if nit is not None else None
#             params[f"doc{i}"] = str(doc) if doc is not None else None

#         sql = f"""
#         DELETE c
#         FROM {table_name} AS c
#         JOIN (VALUES {", ".join(values_clause)}) AS v(nit_cliente, nro_docto_cruce)
#           ON v.nit_cliente = c.nit_cliente
#          AND v.nro_docto_cruce = c.nro_docto_cruce;
#         """
#         db.execute(text(sql), params)


# # ------------------- Helpers de filtros y orden -------------------

# def _to_int_or_none(val: Optional[str]):
#     try:
#         return int(val) if val not in (None, "") else None
#     except (ValueError, TypeError):
#         return None


# def _parse_filters(request_like) -> tuple[Optional[int], Optional[int], str, Optional[str]]:
#     qp = request_like.query_params
#     min_val = _to_int_or_none(qp.get("min_dias"))
#     max_val = _to_int_or_none(qp.get("max_dias"))
#     sort = qp.get("sort", "dias_desc")
#     view = qp.get("view")
#     return min_val, max_val, sort, view


# def _apply_filters_and_order(db: Session, min_val: Optional[int], max_val: Optional[int], sort: str, view: Optional[str]):
#     q = db.query(models.Cliente)
#     if min_val is not None:
#         q = q.filter(models.Cliente.dias_vencidos >= min_val)
#     if max_val is not None:
#         q = q.filter(models.Cliente.dias_vencidos <= max_val)

#     effective_sort = sort
#     if (view == "group") and (sort not in {"nit_asc", "razon_asc"}):
#         effective_sort = "nit_asc"

#     if effective_sort == "dias_asc":
#         q = q.order_by(models.Cliente.dias_vencidos.asc(),
#                        models.Cliente.nit_cliente.asc(),
#                        models.Cliente.nro_docto_cruce.asc(),
#                        models.Cliente.id.asc())
#     elif effective_sort == "razon_asc":
#         q = q.order_by(models.Cliente.razon_social.asc(),
#                        models.Cliente.nit_cliente.asc(),
#                        models.Cliente.nro_docto_cruce.asc(),
#                        models.Cliente.id.asc())
#     elif effective_sort == "nit_asc":
#         q = q.order_by(models.Cliente.nit_cliente.asc(),
#                        models.Cliente.nro_docto_cruce.asc(),
#                        models.Cliente.fecha_docto.asc().nulls_last()
#                        if hasattr(models.Cliente.fecha_docto, "nulls_last") else models.Cliente.fecha_docto.asc(),
#                        models.Cliente.id.asc())
#     else:  # dias_desc
#         q = q.order_by(models.Cliente.dias_vencidos.desc(),
#                        models.Cliente.nit_cliente.asc(),
#                        models.Cliente.nro_docto_cruce.asc(),
#                        models.Cliente.id.asc())
#     return q


# # ------------------- Exportar cartera -------------------

# @app.get("/exportar_cartera.xlsx", name="exportar_cartera_xlsx")
# def exportar_cartera_xlsx(
#     db: Session = Depends(get_db),
#     min_dias: Optional[str] = Query(None),
#     max_dias: Optional[str] = Query(None),
#     sort: Optional[str] = Query("dias_desc"),
# ):
#     min_val = _to_int_or_none(min_dias)
#     max_val = _to_int_or_none(max_dias)

#     query = _apply_filters_and_order(db, min_val, max_val, sort, None)
#     clientes = query.all()

#     rows = []
#     for c in clientes:
#         valor_docto = Decimal(str(c.valor_docto or 0))
#         total_cop = Decimal(str(c.total_cop or 0))
#         recaudo = Decimal(str(c.recaudo)) if c.recaudo is not None else (valor_docto - total_cop)
#         obs_txt = ""
#         if getattr(c, "observaciones", None):
#             obs_txt = "\n".join(
#                 f"{(o.fecha_creacion.strftime('%Y-%m-%d %H:%M') if o.fecha_creacion else '')} - {o.texto}"
#                 for o in c.observaciones
#             )
#         rows.append({
#             "ID": c.id,
#             "Razón social": c.razon_social,
#             "NIT": c.nit_cliente,
#             "Docto cruce": c.nro_docto_cruce,
#             "Días vencidos": c.dias_vencidos,
#             "Fecha docto": c.fecha_docto,
#             "Fecha vcto": c.fecha_vcto,
#             "Valor docto": float(valor_docto),
#             "Total COP (saldo)": float(total_cop),
#             "Recaudo": float(recaudo),
#             "Teléfono": c.telefono,
#             "Celular": c.celular,
#             "Asesor": c.asesor,
#             "Fecha gestión": c.fecha_gestion,
#             "Tipo": c.tipo,
#             "Observaciones": obs_txt,
#         })

#     df = pd.DataFrame(rows)
#     output = BytesIO()
#     with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
#         df.to_excel(writer, index=False, sheet_name="cartera")

#     output.seek(0)
#     filename = f"cartera_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
#     return StreamingResponse(
#         output,
#         media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#         headers={"Content-Disposition": f'attachment; filename="{filename}"'}
#     )


# # ------------------- Importar cartera -------------------

# @app.post("/importar_excel")
# async def importar_excel(file: UploadFile = File(...), db: Session = Depends(get_db)):
#     try:
#         contents = await file.read()
#         df = pd.read_excel(BytesIO(contents))
#         rename_map = {
#             "Razón social": "razon_social",
#             "Nit cliente despacho": "nit_cliente",
#             "Nro. docto. cruce": "nro_docto_cruce",
#             "Fecha docto.": "fecha_docto",
#             "Dias vencidos": "dias_vencidos",
#             "Valor docto": "valor_docto",
#             "Total COP": "total_cop",
#             "Fecha vcto.": "fecha_vcto",
#             "Celular": "celular",
#             "Teléfono": "telefono",
#             "Razón social vend. cliente": "asesor"
#         }
#         df.rename(columns=rename_map, inplace=True)

#         for col in ("telefono", "celular"):
#             if col in df.columns:
#                 df[col] = df[col].apply(clean_phone)

#         if df.empty:
#             return RedirectResponse(url="/?msg=Archivo%20vac%C3%ADo&msg_type=info", status_code=303)

#         excel_claves = set(
#             (str(row["nit_cliente"]).strip(), str(row["nro_docto_cruce"]).strip())
#             for _, row in df.iterrows()
#             if row.get("nit_cliente") and row.get("nro_docto_cruce")
#         )

#         bd_clientes = db.query(models.Cliente).all()
#         bd_claves = set((str(c.nit_cliente), str(c.nro_docto_cruce)) for c in bd_clientes)

#         claves_a_eliminar = bd_claves - excel_claves
#         if claves_a_eliminar:
#             delete_by_pairs_mssql(db, "cartera", claves_a_eliminar)

#         for _, row in df.iterrows():
#             nit = str(row.get("nit_cliente")).strip() if row.get("nit_cliente") else None
#             docto = str(row.get("nro_docto_cruce")).strip() if row.get("nro_docto_cruce") else None
#             if not nit or not docto:
#                 continue
#             valor_docto = to_float(row.get("valor_docto")) or 0.0
#             total_excel = to_float(row.get("total_cop")) if row.get("total_cop") else valor_docto
#             recaudo_calc = round(max(valor_docto - total_excel, 0.0), 2)

#             cliente = db.query(models.Cliente).filter_by(nit_cliente=nit, nro_docto_cruce=docto).first()
#             if cliente:
#                 cliente.razon_social = row.get("razon_social")
#                 cliente.dias_vencidos = row.get("dias_vencidos")
#                 cliente.fecha_docto = row.get("fecha_docto")
#                 cliente.fecha_vcto = row.get("fecha_vcto")
#                 cliente.valor_docto = valor_docto
#                 cliente.total_cop = total_excel
#                 cliente.recaudo = recaudo_calc
#                 cliente.telefono = clean_phone(row.get("telefono"))
#                 cliente.celular = clean_phone(row.get("celular"))
#                 cliente.asesor = row.get("asesor")
#             else:
#                 nuevo = models.Cliente(
#                     razon_social=row.get("razon_social"),
#                     nit_cliente=nit,
#                     nro_docto_cruce=docto,
#                     dias_vencidos=row.get("dias_vencidos"),
#                     fecha_docto=row.get("fecha_docto"),
#                     fecha_vcto=row.get("fecha_vcto"),
#                     valor_docto=valor_docto,
#                     total_cop=total_excel,
#                     recaudo=recaudo_calc,
#                     telefono=clean_phone(row.get("telefono")),
#                     celular=clean_phone(row.get("celular")),
#                     asesor=row.get("asesor"),
#                 )
#                 db.add(nuevo)
#         db.commit()
#         return RedirectResponse(url="/?msg=Archivo%20subido%20correctamente&msg_type=success", status_code=303)
#     except Exception as e:
#         print("❌ Error importar_excel:", e)
#         return RedirectResponse(url="/?msg=Error%20al%20subir%20archivo&msg_type=error", status_code=303)


# # ------------------- Observaciones -------------------

# @app.post("/cliente/{cliente_id}/observacion")
# def agregar_observacion(cliente_id: int, texto: str = Form(...), db: Session = Depends(get_db)):
#     crud.add_observacion(db, cliente_id, texto)
#     return RedirectResponse(
#         url=f"/cliente/{cliente_id}?msg=Observaci%C3%B3n%20guardada&msg_type=success",
#         status_code=303
#     )


# # ------------------- Actualizar cliente (editar recaudo) -------------------

# @app.post("/cliente/{cliente_id}/update")
# async def update_cliente(
#     request: Request,
#     cliente_id: int,
#     db: Session = Depends(get_db)
# ):
#     form_data = await request.form()

#     cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
#     if not cliente:
#         return RedirectResponse(url="/?msg=Cliente%20no%20encontrado&msg_type=error", status_code=status.HTTP_303_SEE_OTHER)

#     for campo in ["telefono", "celular", "fecha_gestion", "tipo", "asesor"]:
#         val = form_data.get(campo)
#         if val is not None and val != "":
#             setattr(cliente, campo, val)

#     valor_docto = float(cliente.valor_docto or 0.0)
#     raw_recaudo = form_data.get("recaudo")
#     raw_total_cop = form_data.get("total_cop")

#     has_recaudo = raw_recaudo is not None and raw_recaudo.strip() != ""
#     has_total_cop = raw_total_cop is not None and raw_total_cop.strip() != ""

#     if has_recaudo:
#         delta = to_float(raw_recaudo)
#         if delta is not None:
#             old_total = float(cliente.total_cop or 0.0)
#             new_total = max(round(old_total - delta, 2), 0.0)
#             cliente.total_cop = new_total
#             cliente.recaudo = round(max(valor_docto - new_total, 0.0), 2)
#     elif has_total_cop:
#         nuevo_total = to_float(raw_total_cop)
#         if nuevo_total is not None:
#             nuevo_total = max(nuevo_total, 0.0)
#             if nuevo_total > valor_docto:
#                 nuevo_total = valor_docto
#             cliente.total_cop = round(nuevo_total, 2)
#             cliente.recaudo = round(max(valor_docto - nuevo_total, 0.0), 2)

#     nueva_obs = form_data.get("observaciones")
#     if nueva_obs and nueva_obs.strip():
#         obs = models.Observacion(
#             cliente_id=cliente.id,
#             texto=nueva_obs.strip(),
#         )
#         db.add(obs)

#     db.commit()
#     return RedirectResponse(
#         url=f"/cliente/{cliente_id}?msg=Cambios%20guardados&msg_type=success",
#         status_code=status.HTTP_303_SEE_OTHER
#     )


# # ------------------- historial cliente -------------------

# @app.get("/cliente/{cliente_id}/historial")
# def historial_cliente(cliente_id: int, db: Session = Depends(get_db)):
#     cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
#     if not cliente:
#         return JSONResponse(content=[], status_code=200)

#     historial = [
#         {
#             "texto": obs.texto,
#             "fecha": obs.fecha_creacion.strftime("%d/%m/%Y %H:%M")
#             if obs.fecha_creacion else None
#         }
#         for obs in cliente.observaciones
#     ]

#     return JSONResponse(content=historial, status_code=200)


# # ------------------- Vista cliente -------------------

# @app.get("/cliente/{cliente_id}")
# def ver_cliente(cliente_id: int, request: Request, db: Session = Depends(get_db)):
#     cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
#     if not cliente:
#         raise HTTPException(status_code=404, detail="Cliente no encontrado")

#     # Reconstruimos los filtros/orden desde `ref`
#     ref_qs = request.query_params.get("ref")
#     req_like = request
#     if ref_qs:
#         req_like = type("R", (), {"query_params": QueryParams(ref_qs)})

#     min_val, max_val, sort, view = _parse_filters(req_like)

#     # Usamos el MISMO orden del index para calcular prev/next dentro del mismo universo filtrado
#     ordered_ids = [
#         row.id for row in _apply_filters_and_order(db, min_val, max_val, sort, view)
#                     .with_entities(models.Cliente.id).all()
#     ]

#     prev_id = next_id = None
#     try:
#         idx = ordered_ids.index(cliente_id)
#         if idx > 0:
#             prev_id = ordered_ids[idx - 1]
#         if idx < len(ordered_ids) - 1:
#             next_id = ordered_ids[idx + 1]
#     except ValueError:
#         # si por alguna razón el id no está en el conjunto filtrado, no rompemos
#         pass

#     back_url = f"/?{ref_qs}" if ref_qs else "/"

#     return templates.TemplateResponse(
#         "cliente.html",
#         {
#             "request": request,
#             "cliente": cliente,
#             "prev_id": prev_id,
#             "next_id": next_id,
#             "flash_msg": request.query_params.get("msg"),
#             "flash_type": request.query_params.get("msg_type"),
#             "back_url": back_url,
#         }
#     )


# # ------------------- Index agrupado y plano -------------------

# @app.get("/")
# def index(
#     request: Request,
#     db: Session = Depends(get_db),
#     view: Optional[str] = Query(None),
#     min_dias: Optional[str] = Query(None),
#     max_dias: Optional[str] = Query(None),
#     sort: Optional[str] = Query("dias_desc"),
# ):
#     min_val = _to_int_or_none(min_dias)
#     max_val = _to_int_or_none(max_dias)

#     # Misma lógica de orden que usa ver_cliente()
#     query = _apply_filters_and_order(db, min_val, max_val, sort, view)
#     clientes = query.all()

#     # Flash
#     flash_msg = request.query_params.get("msg")
#     flash_type = request.query_params.get("msg_type")
#     if (min_val is not None) or (max_val is not None):
#         if not flash_msg:
#             flash_msg = "Filtros aplicados"
#         if not flash_type:
#             flash_type = "info"

#     current_qs = request.url.query  # para preservar filtros en hrefs

#     # Vista plana
#     if view == "flat":
#         filas = []
#         for c in clientes:
#             valor_docto = Decimal(str(c.valor_docto or 0))
#             total_cop = Decimal(str(c.total_cop or 0))
#             recaudo = valor_docto - total_cop

#             filas.append({
#                 "id": c.id,
#                 "razon_social": c.razon_social,
#                 "nit_cliente": c.nit_cliente,
#                 "nro_docto_cruce": c.nro_docto_cruce,
#                 "dias_vencidos": c.dias_vencidos,
#                 "fecha_docto": c.fecha_docto,
#                 "fecha_vcto": c.fecha_vcto,
#                 "valor_docto": float(valor_docto),
#                 "total_cop": float(total_cop),
#                 "recaudo": float(recaudo),
#                 "asesor": c.asesor,
#                 "correo": c.correo,
#             })

#         return templates.TemplateResponse(
#             "index.html",
#             {
#                 "request": request,
#                 "view": "flat",
#                 "filas": filas,
#                 "flash_msg": flash_msg,
#                 "flash_type": flash_type,
#                 "current_qs": current_qs,
#                 "min_dias": min_dias,
#                 "max_dias": max_dias,
#                 "sort": sort,
#             }
#         )

#     # Vista agrupada
#     agrupados = {}
#     for c in clientes:
#         if c.nit_cliente not in agrupados:
#             agrupados[c.nit_cliente] = {
#                 "nit_cliente": c.nit_cliente,
#                 "razon_social": c.razon_social,
#                 "telefono": c.telefono,
#                 "celular": c.celular,
#                 "correo": c.correo,
#                 "asesor": c.asesor,
#                 "facturas": []
#             }

#         valor_docto = Decimal(str(c.valor_docto or 0))
#         total_cop = Decimal(str(c.total_cop or 0))
#         recaudo = valor_docto - total_cop

#         agrupados[c.nit_cliente]["facturas"].append({
#             "id": c.id,
#             "nro_docto_cruce": c.nro_docto_cruce,
#             "dias_vencidos": c.dias_vencidos,
#             "fecha_docto": c.fecha_docto,
#             "fecha_vcto": c.fecha_vcto,
#             "valor_docto": float(valor_docto),
#             "total_cop": float(total_cop),
#             "recaudo": float(recaudo),
#             "fecha_gestion": c.fecha_gestion,
#             "tipo": c.tipo,
#             "asesor": c.asesor,
#             "observaciones": [obs.texto for obs in c.observaciones]
#         })

#     for cl in agrupados.values():
#         dias = [f["dias_vencidos"] for f in cl["facturas"] if f["dias_vencidos"] is not None]
#         cl["max_dias"] = max(dias) if dias else None

#     return templates.TemplateResponse(
#         "index.html",
#         {
#             "request": request,
#             "clientes": list(agrupados.values()),
#             "flash_msg": flash_msg,
#             "flash_type": flash_type,
#             "current_qs": current_qs,
#             "min_dias": min_dias,
#             "max_dias": max_dias,
#             "sort": sort,
#             "view": view,
#         }
#     )




#tercer vercion del codigo


# from fastapi.responses import JSONResponse, RedirectResponse, StreamingResponse
# from fastapi import UploadFile, File, Form, Query, status
# from database import SessionLocal, engine, Base
# from fastapi.templating import Jinja2Templates
# from fastapi import FastAPI, Request, Depends
# from fastapi.staticfiles import StaticFiles
# from fastapi import HTTPException
# from sqlalchemy.orm import Session
# from sqlalchemy import text
# from datetime import datetime
# from typing import Optional
# from decimal import Decimal
# from io import BytesIO
# import pandas as pd
# import models
# import math
# import crud
# import re

# # ------------------- Inicialización -------------------
# app = FastAPI()
# templates = Jinja2Templates(directory="templates")

# app.mount("/static", StaticFiles(directory="static"), name="static")
# app.mount("/assets", StaticFiles(directory="assets"), name="assets")

# # ------------------- Helpers -------------------
# def fmt_money(value):
#     if value is None or value == "":
#         return "-"
#     try:
#         v = Decimal(str(value))
#         return f"{v:,.2f}"
#     except Exception:
#         return str(value)

# templates.env.filters["fmt_money"] = fmt_money

# def get_db():
#     db = SessionLocal()
#     try:
#         yield db
#     finally:
#         db.close()

# def to_float(val):
#     try:
#         if pd.isna(val):
#             return None
#         val_str = str(val).replace(",", ".")
#         val_str = re.sub(r"[^0-9.]", "", val_str)
#         return float(val_str) if val_str else None
#     except (ValueError, TypeError):
#         return None

# def clean_phone(val):
#     if val is None or (isinstance(val, float) and pd.isna(val)):
#         return None
#     s = str(val).strip()
#     if s.endswith(".0") and s[:-2].isdigit():
#         s = s[:-2]
#     try:
#         if re.fullmatch(r"\d+(\.\d+)?[eE][+-]?\d+", s):
#             s = str(int(float(s)))
#     except Exception:
#         pass
#     if re.fullmatch(r"\d+\.\d+", s):
#         try:
#             f = float(s)
#             if f.is_integer():
#                 s = str(int(f))
#         except Exception:
#             pass
#     s = re.sub(r"[^\d]", "", s)
#     return s or None

# def clean_email(val):
#     if val is None or (isinstance(val, float) and pd.isna(val)):
#         return None
#     s = str(val).strip()
#     if not s or s.lower() == "nan":
#         return None
#     candidates = re.split(r"[;,\n/]+", s)
#     email_re = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.IGNORECASE)
#     for part in candidates:
#         m = email_re.search(part)
#         if m:
#             return m.group(0).strip().lower()
#     m = email_re.search(s)
#     if m:
#         return m.group(0).strip().lower()
#     return None

# def to_int_nullsafe(val):
#     try:
#         if val is None:
#             return None
#         if isinstance(val, float) and math.isnan(val):
#             return None
#         s = str(val).strip()
#         if s == "":
#             return None
#         return int(float(s))
#     except Exception:
#         return None

# def to_dt_nullsafe(val):
#     if val is None:
#         return None
#     try:
#         if isinstance(val, float) and math.isnan(val):
#             return None
#     except Exception:
#         pass
#     ts = pd.to_datetime(val, errors="coerce")
#     if pd.isna(ts):
#         return None
#     return ts.to_pydatetime()

# def _chunk(iterable, size):
#     buf = []
#     for x in iterable:
#         buf.append(x)
#         if len(buf) >= size:
#             yield buf
#             buf = []
#     if buf:
#         yield buf

# def delete_by_pairs_mssql(db: Session, table_name: str, key_pairs, batch_size: int = 900):
#     pairs = list(key_pairs)
#     if not pairs:
#         return
#     for batch in _chunk(pairs, batch_size):
#         values_clause = []
#         params = {}
#         for i, (nit, doc) in enumerate(batch):
#             values_clause.append(f"(:nit{i}, :doc{i})")
#             params[f"nit{i}"] = str(nit) if nit is not None else None
#             params[f"doc{i}"] = str(doc) if doc is not None else None
#         sql = f"""
#         DELETE c
#         FROM {table_name} AS c
#         JOIN (VALUES {", ".join(values_clause)}) AS v(nit_cliente, nro_docto_cruce)
#           ON v.nit_cliente = c.nit_cliente
#          AND v.nro_docto_cruce = c.nro_docto_cruce;
#         """
#         db.execute(text(sql), params)

# # ------------------- Exportar cartera a Excel -------------------
# @app.get("/exportar_cartera.xlsx", name="exportar_cartera_xlsx")
# def exportar_cartera_xlsx(
#     db: Session = Depends(get_db),
#     min_dias: Optional[str] = Query(None),
#     max_dias: Optional[str] = Query(None),
#     sort: Optional[str] = Query("dias_desc"),
# ):
#     def to_int_or_none(val: Optional[str]):
#         try:
#             return int(val) if val not in (None, "") else None
#         except (ValueError, TypeError):
#             return None

#     min_val = to_int_or_none(min_dias)
#     max_val = to_int_or_none(max_dias)

#     query = db.query(models.Cliente)
#     if min_val is not None:
#         query = query.filter(models.Cliente.dias_vencidos >= min_val)
#     if max_val is not None:
#         query = query.filter(models.Cliente.dias_vencidos <= max_val)

#     order_map = {
#         "dias_desc": models.Cliente.dias_vencidos.desc(),
#         "dias_asc": models.Cliente.dias_vencidos.asc(),
#         "razon_asc": models.Cliente.razon_social.asc(),
#     }
#     query = query.order_by(order_map.get(sort, order_map["dias_desc"]))
#     clientes = query.all()

#     rows = []
#     for c in clientes:
#         valor_docto = Decimal(str(c.valor_docto or 0))
#         total_cop = Decimal(str(c.total_cop or 0))
#         recaudo = Decimal(str(c.recaudo)) if c.recaudo is not None else (valor_docto - total_cop)
#         obs_txt = ""
#         if getattr(c, "observaciones", None):
#             latest = None
#             for o in c.observaciones:
#                 f = o.fecha_creacion or datetime.min
#                 if latest is None or f > (latest.fecha_creacion or datetime.min):
#                     latest = o
#             if latest:
#                 obs_txt = f"{(latest.fecha_creacion.strftime('%Y-%m-%d %H:%M') if latest.fecha_creacion else '')} - {latest.texto}"
#         rows.append({
#             "ID": c.id,
#             "Razón social": c.razon_social,
#             "NIT": c.nit_cliente,
#             "Docto cruce": c.nro_docto_cruce,
#             "Días vencidos": c.dias_vencidos,
#             "Fecha docto": c.fecha_docto,
#             "Fecha vcto": c.fecha_vcto,
#             "Valor docto": float(valor_docto),
#             "Total COP (saldo)": float(total_cop),
#             "Recaudo": float(recaudo),
#             "Teléfono": c.telefono,
#             "Celular": c.celular,
#             "Correo": c.correo,
#             "Asesor": c.asesor,
#             "Fecha gestión": c.fecha_gestion,
#             "Tipo": c.tipo,
#             "Observaciones": obs_txt,
#         })
#     df = pd.DataFrame(rows)

#     for col in ["Fecha docto", "Fecha vcto", "Fecha gestión"]:
#         if col in df.columns:
#             df[col] = pd.to_datetime(df[col], errors="coerce")

#     if not df.empty:
#         resumen = (
#             df.groupby(["NIT", "Razón social"], dropna=False)
#               .agg({
#                   "Valor docto": "sum",
#                   "Total COP (saldo)": "sum",
#                   "Recaudo": "sum",
#                   "Días vencidos": "max",
#               })
#               .rename(columns={
#                   "Valor docto": "Total Valor Docto",
#                   "Total COP (saldo)": "Saldo Total",
#                   "Recaudo": "Recaudo Total",
#                   "Días vencidos": "Max Días Vencidos",
#               })
#               .reset_index()
#         )
#         ref = (
#             df.sort_values(["NIT"]).groupby(["NIT", "Razón social"], dropna=False)
#               .agg({"Teléfono": "first", "Celular": "first", "Correo": "first", "Asesor": "first"})
#               .reset_index()
#         )
#         resumen = resumen.merge(ref, on=["NIT", "Razón social"], how="left")

#         facturas = (
#             df.groupby(["NIT", "Razón social"], dropna=False)
#               .size()
#               .reset_index(name="# Facturas")
#         )
#         resumen = resumen.merge(facturas, on=["NIT", "Razón social"], how="left")
#         resumen["# Facturas"] = resumen["# Facturas"].fillna(0).astype(int)
#     else:
#         resumen = pd.DataFrame(columns=[
#             "NIT","Razón social","Total Valor Docto","Saldo Total","Recaudo Total",
#             "Max Días Vencidos","Teléfono","Celular","Correo","Asesor","# Facturas"
#         ])

#     output = BytesIO()
#     with pd.ExcelWriter(output, engine="xlsxwriter",
#                         datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as writer:
#         df.to_excel(writer, index=False, sheet_name="cartera")
#         resumen.to_excel(writer, index=False, sheet_name="resumen")

#         wb = writer.book
#         ws1 = writer.sheets["cartera"]
#         ws2 = writer.sheets["resumen"]

#         for ws in (ws1, ws2):
#             ws.freeze_panes(1, 0)

#         fmt_header = wb.add_format({
#             'bold': True, 'bg_color': '#D9E1F2', 'border': 1,
#             'align': 'center', 'valign': 'vcenter'
#         })
#         fmt_money0 = wb.add_format({'num_format': '#,##0', 'border': 1})
#         fmt_int = wb.add_format({'num_format': '#,##0', 'border': 1})
#         fmt_date = wb.add_format({'num_format': 'yyyy-mm-dd', 'border': 1})
#         fmt_text = wb.add_format({'text_wrap': True, 'border': 1, 'valign': 'top'})
#         fmt_default = wb.add_format({'border': 1})

#         def table_with_formats(ws, df_sheet, money_cols, date_cols, wrap_cols):
#             rows, cols = df_sheet.shape
#             headers = list(df_sheet.columns)
#             per_col_format = {}
#             for col in headers:
#                 if col in money_cols:
#                     per_col_format[col] = fmt_money0
#                 elif col in date_cols:
#                     per_col_format[col] = fmt_date
#                 elif col in wrap_cols:
#                     per_col_format[col] = fmt_text
#                 elif df_sheet[col].dtype.kind in ("i", "u"):
#                     per_col_format[col] = fmt_int
#                 else:
#                     per_col_format[col] = fmt_default

#             columns_def = [{'header': h, 'format': per_col_format[h]} for h in headers]

#             if rows > 0:
#                 ws.add_table(0, 0, rows, cols - 1, {
#                     'style': 'Table Style Medium 9',
#                     'banded_rows': True,
#                     'columns': columns_def
#                 })
#             else:
#                 for c_idx, h in enumerate(headers):
#                     ws.write(0, c_idx, h, fmt_header)

#             col_widths = {}
#             for c_idx, col in enumerate(headers):
#                 max_len = len(str(col)) + 2
#                 serie = df_sheet[col]
#                 if col in money_cols:
#                     for v in serie.dropna():
#                         try:
#                             s = f"{int(round(float(v))):,}"
#                             max_len = max(max_len, len(s))
#                         except Exception:
#                             pass
#                 elif col in date_cols:
#                     max_len = max(max_len, 12)
#                 else:
#                     for v in serie.dropna():
#                         s = str(v).replace("\r", "")
#                         max_len = max(max_len, max((len(seg) for seg in s.split("\n")), default=0))
#                     if col in wrap_cols:
#                         max_len = min(max(max_len, 40), 80)
#                     else:
#                         max_len = min(max_len, 40)
#                 width = max_len + 1
#                 col_widths[col] = width
#                 ws.set_column(c_idx, c_idx, width, per_col_format[col])
#             return col_widths, per_col_format

#         def autofit_row_heights(ws, df_sheet, wrap_cols, col_widths):
#             if df_sheet.empty or not wrap_cols:
#                 return
#             base_height = 15
#             max_height = 300
#             for r in range(1, len(df_sheet) + 1):
#                 max_lines = 1
#                 for col in wrap_cols:
#                     if col not in df_sheet.columns:
#                         continue
#                     val = df_sheet.iloc[r - 1][col]
#                     if pd.isna(val) or val is None:
#                         continue
#                     text = str(val).replace("\r", "")
#                     if text == "":
#                         continue
#                     col_w = int(col_widths.get(col, 40))
#                     chars_per_line = max(col_w - 2, 10)
#                     total_lines = 0
#                     for seg in text.split("\n"):
#                         seg = seg.strip()
#                         if seg == "":
#                             total_lines += 1
#                         else:
#                             total_lines += math.ceil(len(seg) / chars_per_line)
#                     max_lines = max(max_lines, total_lines)
#                 ws.set_row(r, min(base_height * max_lines + 4, max_height))

#         money_cols_cartera = ["Valor docto", "Total COP (saldo)", "Recaudo"]
#         date_cols_cartera = ["Fecha docto", "Fecha vcto", "Fecha gestión"]
#         wrap_cols_cartera = ["Razón social", "Observaciones"]
#         money_cols_resumen = ["Total Valor Docto", "Saldo Total", "Recaudo Total"]
#         date_cols_resumen = []
#         wrap_cols_resumen = ["Razón social"]

#         widths_cartera, _ = table_with_formats(ws1, df, money_cols_cartera, date_cols_cartera, wrap_cols_cartera)
#         widths_resumen, _ = table_with_formats(ws2, resumen, money_cols_resumen, date_cols_resumen, wrap_cols_resumen)
#         autofit_row_heights(ws1, df, wrap_cols_cartera, widths_cartera)
#         autofit_row_heights(ws2, resumen, wrap_cols_resumen, widths_resumen)

#     output.seek(0)
#     filename = f"cartera_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
#     return StreamingResponse(
#         output,
#         media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#         headers={"Content-Disposition": f'attachment; filename="{filename}"'}
#     )

# # ------------------- Importar cartera -------------------
# @app.post("/importar_excel")
# async def importar_excel(file: UploadFile = File(...), db: Session = Depends(get_db)):
#     try:
#         contents = await file.read()
#         def _read_any(contents: bytes, filename: str):
#             name = (filename or "").lower()
#             if name.endswith(".csv"):
#                 df = pd.read_csv(BytesIO(contents), dtype=str, keep_default_na=False)
#                 return {"__csv__": df}
#             return pd.read_excel(BytesIO(contents), sheet_name=None, dtype=str, keep_default_na=False)
#         sheets = _read_any(contents, file.filename)
#         # --- Normalizador simple de strings ---
#         def norm(s: str) -> str:
#             if s is None:
#                 return ""
#             t = str(s).strip().lower()
#             rep = {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ñ":"n","ä":"a","ë":"e","ï":"i","ö":"o","ü":"u","’":"'","´":"'","`":"'"}
#             for a,b in rep.items(): t = t.replace(a,b)
#             t = re.sub(r"\s+", " ", t)
#             return t

#         # --- Aliases (internos -> alias) ---
#         alias_map = {
#             "razon_social": {"razon social", "razon social vend. cliente", "razon social vend cliente"},
#             "nit_cliente": {"nit", "nit cliente despacho", "nit cliente"},
#             "nro_docto_cruce": {"docto cruce", "nro. docto. cruce", "nro docto cruce", "factura", "documento", "nro documento"},
#             "fecha_docto": {"fecha docto", "fecha docto."},
#             "fecha_vcto": {"fecha vcto", "fecha vcto."},
#             "dias_vencidos": {"dias vencidos", "dias vencidos."},
#             "valor_docto": {"valor docto", "valor docto."},
#             "total_cop": {"total cop", "total cop (saldo)", "saldo", "total (cop)"},
#             "telefono": {"telefono", "teléfono"},
#             "celular": {"celular", "movil", "móvil"},
#             "correo": {"correo", "email", "e-mail", "mail", "correo electronico", "correo electrónico"},
#             "asesor": {"asesor", "vendedor"},
#             "obs_txt": {"observaciones"},
#         }

#         # --- Detectar si es export del proyecto ---
#         is_project_export = False
#         df_project = None
#         for sheet_name, df in sheets.items():
#             if norm(sheet_name) == "cartera":
#                 df_project = df
#                 break
#         if df_project is not None:
#             cols_n = {norm(c) for c in df_project.columns}
#             if "nit" in cols_n and ("docto cruce" in cols_n or "nro docto cruce" in cols_n):
#                 is_project_export = True

#         # --- Elegir hoja candidata ---
#         if is_project_export:
#             df_raw = df_project.copy()
#         else:
#             df_raw = None
#             for _, df in sheets.items():
#                 cols_n = {norm(c) for c in df.columns}
#                 if any(x in cols_n for x in alias_map["nit_cliente"]) and \
#                    any(x in cols_n for x in alias_map["nro_docto_cruce"]):
#                     df_raw = df
#                     break
#             if df_raw is None:
#                 return RedirectResponse(
#                     url="/?msg=No%20se%20encontro%20hoja%20compatible&msg_type=error",
#                     status_code=303
#                 )

#         original_cols = list(df_raw.columns)
#         df_raw.columns = [str(c).strip() for c in df_raw.columns]
#         rename_map_dynamic = {}
#         for col in df_raw.columns:
#             n = norm(col)
#             for internal, aliases in alias_map.items():
#                 if n in aliases and internal not in rename_map_dynamic.values():
#                     rename_map_dynamic[col] = internal
#                     break

#         project_head_map = {
#             "Razón social": "razon_social",
#             "NIT": "nit_cliente",
#             "Docto cruce": "nro_docto_cruce",
#             "Días vencidos": "dias_vencidos",
#             "Fecha docto": "fecha_docto",
#             "Fecha vcto": "fecha_vcto",
#             "Valor docto": "valor_docto",
#             "Total COP (saldo)": "total_cop",
#             "Teléfono": "telefono",
#             "Celular": "celular",
#             "Correo": "correo",
#             "Asesor": "asesor",
#             "Observaciones": "obs_txt",
#         }
#         for col in df_raw.columns:
#             if col in project_head_map and project_head_map[col] not in rename_map_dynamic.values():
#                 rename_map_dynamic[col] = project_head_map[col]

#         df = df_raw.rename(columns=rename_map_dynamic)
#         if "correo" not in df.columns:
#             df["correo"] = None
#         df["correo"] = df["correo"].apply(clean_email)

#         if "nit_cliente" not in df.columns or "nro_docto_cruce" not in df.columns:
#             return RedirectResponse(
#                 url="/?msg=Faltan%20columnas%20clave&msg_type=error",
#                 status_code=303
#             )

#         if "telefono" in df.columns: df["telefono"] = df["telefono"].apply(clean_phone)
#         if "celular" in df.columns: df["celular"] = df["celular"].apply(clean_phone)
#         if "fecha_docto" in df.columns: df["fecha_docto"] = df["fecha_docto"].apply(to_dt_nullsafe)
#         if "fecha_vcto" in df.columns: df["fecha_vcto"] = df["fecha_vcto"].apply(to_dt_nullsafe)
#         if "dias_vencidos" in df.columns: df["dias_vencidos"] = df["dias_vencidos"].apply(to_int_nullsafe)

#         def _to_money(x):
#             v = to_float(x)
#             return v if v is not None else 0.0
#         if "valor_docto" in df.columns: df["valor_docto"] = df["valor_docto"].apply(_to_money)
#         else: df["valor_docto"] = 0.0
#         if "total_cop" in df.columns: df["total_cop"] = df["total_cop"].apply(_to_money)
#         else: df["total_cop"] = df["valor_docto"]

#         df = df[(df["nit_cliente"].notna()) & (df["nro_docto_cruce"].notna())]
#         df["nit_cliente"] = df["nit_cliente"].astype(str).str.strip()
#         df["nro_docto_cruce"] = df["nro_docto_cruce"].astype(str).str.strip()
#         df = df[(df["nit_cliente"] != "") & (df["nro_docto_cruce"] != "")]
#         if df.empty:
#             return RedirectResponse(url="/?msg=Archivo%20vacio&msg_type=info", status_code=303)

#         bd_clientes = db.query(models.Cliente).all()
#         bd_claves = set((str(c.nit_cliente), str(c.nro_docto_cruce)) for c in bd_clientes)

#         if not is_project_export:
#             excel_claves = set(zip(df["nit_cliente"], df["nro_docto_cruce"]))
#             claves_a_eliminar = bd_claves - excel_claves
#             if claves_a_eliminar:
#                 delete_by_pairs_mssql(db, "cartera", claves_a_eliminar)

#         for _, row in df.iterrows():
#             nit = row["nit_cliente"]
#             docto = row["nro_docto_cruce"]
#             valor_docto = row.get("valor_docto", 0.0) or 0.0
#             total_excel = row.get("total_cop", valor_docto)
#             recaudo_calc = round(max((valor_docto or 0.0) - (total_excel or 0.0), 0.0), 2)
#             cliente = db.query(models.Cliente).filter_by(nit_cliente=nit, nro_docto_cruce=docto).first()
#             correo_val = clean_email(row.get("correo")) if "correo" in row else None
#             if cliente:
#                 cliente.razon_social = row.get("razon_social")
#                 cliente.dias_vencidos = to_int_nullsafe(row.get("dias_vencidos"))
#                 cliente.fecha_docto = to_dt_nullsafe(row.get("fecha_docto"))
#                 cliente.fecha_vcto = to_dt_nullsafe(row.get("fecha_vcto"))
#                 cliente.valor_docto = valor_docto
#                 cliente.total_cop = total_excel
#                 cliente.recaudo = recaudo_calc
#                 cliente.telefono = clean_phone(row.get("telefono"))
#                 cliente.celular = clean_phone(row.get("celular"))
#                 cliente.asesor = row.get("asesor")
#                 if correo_val and (not cliente.correo or cliente.correo != correo_val):
#                     cliente.correo = correo_val
#             else:
#                 cliente = models.Cliente(
#                     razon_social=row.get("razon_social"),
#                     nit_cliente=nit,
#                     nro_docto_cruce=docto,
#                     dias_vencidos=to_int_nullsafe(row.get("dias_vencidos")),
#                     fecha_docto=to_dt_nullsafe(row.get("fecha_docto")),
#                     fecha_vcto=to_dt_nullsafe(row.get("fecha_vcto")),
#                     valor_docto=valor_docto,
#                     total_cop=total_excel,
#                     recaudo=recaudo_calc,
#                     telefono=clean_phone(row.get("telefono")),
#                     celular=clean_phone(row.get("celular")),
#                     asesor=row.get("asesor"),
#                     correo=correo_val,
#                 )
#                 db.add(cliente)
#                 db.flush()

#         db.commit()
#         return RedirectResponse(url="/?msg=Archivo%20subido%20correctamente&msg_type=success", status_code=303)

#     except Exception as e:
#         print("❌ Error importar_excel:", e)
#         return RedirectResponse(url="/?msg=Error%20al%20subir%20el%20archivo&msg_type=error", status_code=303)


# # ------------------- Observaciones -------------------
# @app.post("/cliente/{cliente_id}/observacion")
# def agregar_observacion(cliente_id: int, texto: str = Form(...), db: Session = Depends(get_db)):
#     crud.add_observacion(db, cliente_id, texto)
#     return RedirectResponse(
#         url=f"/cliente/{cliente_id}?msg=Observacion%20guardada&msg_type=success",
#         status_code=303
#     )


# # ------------------- Actualizar cliente -------------------
# @app.post("/cliente/{cliente_id}/update")
# async def update_cliente(request: Request, cliente_id: int, db: Session = Depends(get_db)):
#     form_data = await request.form()
#     cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
#     if not cliente:
#         return RedirectResponse(url="/?msg=Cliente%20no%20encontrado&msg_type=error", status_code=status.HTTP_303_SEE_OTHER)
#     for campo in ["telefono", "celular", "correo", "fecha_gestion", "tipo", "asesor"]:
#         val = form_data.get(campo)
#         if val is not None and val != "":
#             if campo == "correo": val = clean_email(val)
#             if campo in ("telefono", "celular"): val = clean_phone(val)
#             setattr(cliente, campo, val)
#     valor_docto = float(cliente.valor_docto or 0.0)
#     raw_recaudo = form_data.get("recaudo")
#     raw_total_cop = form_data.get("total_cop")
#     if raw_recaudo:
#         delta = to_float(raw_recaudo)
#         if delta is not None:
#             old_total = float(cliente.total_cop or 0.0)
#             new_total = max(round(old_total - delta, 2), 0.0)
#             cliente.total_cop = new_total
#             cliente.recaudo = round(max(valor_docto - new_total, 0.0), 2)
#     elif raw_total_cop:
#         nuevo_total = to_float(raw_total_cop)
#         if nuevo_total is not None:
#             nuevo_total = max(nuevo_total, 0.0)
#             if nuevo_total > valor_docto:
#                 nuevo_total = valor_docto
#             cliente.total_cop = round(nuevo_total, 2)
#             cliente.recaudo = round(max(valor_docto - nuevo_total, 0.0), 2)

#     nueva_obs = form_data.get("observaciones")
#     if nueva_obs and nueva_obs.strip():
#         obs = models.Observacion(cliente_id=cliente.id, texto=nueva_obs.strip())
#         db.add(obs)
#     db.commit()
#     return RedirectResponse(
#         url=f"/cliente/{cliente_id}?msg=Cambios%20guardados&msg_type=success",
#         status_code=status.HTTP_303_SEE_OTHER
#     )


# # ------------------- historial cliente -------------------
# @app.get("/cliente/{cliente_id}/historial")
# def historial_cliente(cliente_id: int, db: Session = Depends(get_db)):
#     cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
#     if not cliente:
#         return JSONResponse(content=[], status_code=200)
#     historial = [
#         {"texto": obs.texto, "fecha": obs.fecha_creacion.strftime("%d/%m/%Y %H:%M") if obs.fecha_creacion else None}
#         for obs in cliente.observaciones
#     ]
#     return JSONResponse(content=historial, status_code=200)


# # ------------------- Vista cliente -------------------
# @app.get("/cliente/{cliente_id}")
# def ver_cliente(cliente_id: int, request: Request, db: Session = Depends(get_db)):
#     cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
#     if not cliente:
#         raise HTTPException(status_code=404, detail="Cliente no encontrado")

#     # Orden estable por id
#     prev_cliente = db.query(models.Cliente).filter(models.Cliente.id < cliente_id).order_by(models.Cliente.id.desc()).first()
#     next_cliente = db.query(models.Cliente).filter(models.Cliente.id > cliente_id).order_by(models.Cliente.id.asc()).first()
#     prev_id = prev_cliente.id if prev_cliente else None
#     next_id = next_cliente.id if next_cliente else None

#     ref_qs = request.query_params.get("ref")
#     back_url = f"/?{ref_qs}" if ref_qs else "/"

#     return templates.TemplateResponse(
#         "cliente.html",
#         {
#             "request": request,
#             "cliente": cliente,
#             "prev_id": prev_id,
#             "next_id": next_id,
#             "flash_msg": request.query_params.get("msg"),
#             "flash_type": request.query_params.get("msg_type"),
#             "back_url": back_url,
#         }
#     )


# # ------------------- Index agrupado y plano -------------------
# @app.get("/")
# def index(request: Request, db: Session = Depends(get_db),
#           view: Optional[str] = Query(None),
#           min_dias: Optional[str] = Query(None),
#           max_dias: Optional[str] = Query(None),
#           sort: Optional[str] = Query("dias_desc")):

#     def to_int_or_none(val: Optional[str]):
#         try: return int(val) if val not in (None, "") else None
#         except (ValueError, TypeError): return None

#     min_val, max_val = to_int_or_none(min_dias), to_int_or_none(max_dias)
#     query = db.query(models.Cliente)
#     if min_val is not None: query = query.filter(models.Cliente.dias_vencidos >= min_val)
#     if max_val is not None: query = query.filter(models.Cliente.dias_vencidos <= max_val)

#     order_map = {
#         "dias_desc": models.Cliente.dias_vencidos.desc(),
#         "dias_asc": models.Cliente.dias_vencidos.asc(),
#         "razon_asc": models.Cliente.razon_social.asc(),
#     }
#     query = query.order_by(order_map.get(sort, order_map["dias_desc"]))
#     clientes = query.all()

#     flash_msg, flash_type = request.query_params.get("msg"), request.query_params.get("msg_type")
#     if (min_val is not None) or (max_val is not None):
#         if not flash_msg: flash_msg = "Filtros aplicados"
#         if not flash_type: flash_type = "info"

#     current_qs = request.url.query

#     # Vista plana
#     if view == "flat":
#         filas = []
#         for c in clientes:
#             valor_docto = Decimal(str(c.valor_docto or 0))
#             total_cop = Decimal(str(c.total_cop or 0))
#             recaudo = valor_docto - total_cop
#             filas.append({
#                 "id": c.id,
#                 "razon_social": c.razon_social,
#                 "nit_cliente": c.nit_cliente,
#                 "nro_docto_cruce": c.nro_docto_cruce,
#                 "dias_vencidos": c.dias_vencidos,
#                 "fecha_docto": c.fecha_docto,
#                 "fecha_vcto": c.fecha_vcto,
#                 "valor_docto": float(valor_docto),
#                 "total_cop": float(total_cop),
#                 "recaudo": float(recaudo),
#                 "asesor": c.asesor,
#                 "correo": c.correo,
#             })
#         return templates.TemplateResponse("index.html", {
#             "request": request,
#             "view": "flat",
#             "filas": filas,
#             "flash_msg": flash_msg,
#             "flash_type": flash_type,
#             "current_qs": current_qs,
#         })

#     # Vista agrupada
#     agrupados = {}
#     for c in sorted(clientes, key=lambda x: (x.nit_cliente or "", x.nro_docto_cruce or "")):
#         if c.nit_cliente not in agrupados:
#             agrupados[c.nit_cliente] = {
#                 "nit_cliente": c.nit_cliente,
#                 "razon_social": c.razon_social,
#                 "telefono": c.telefono,
#                 "celular": c.celular,
#                 "correo": c.correo,
#                 "asesor": c.asesor,
#                 "facturas": []
#             }
#         valor_docto = Decimal(str(c.valor_docto or 0))
#         total_cop = Decimal(str(c.total_cop or 0))
#         recaudo = valor_docto - total_cop
#         agrupados[c.nit_cliente]["facturas"].append({
#             "id": c.id,
#             "nro_docto_cruce": c.nro_docto_cruce,
#             "dias_vencidos": c.dias_vencidos,
#             "fecha_docto": c.fecha_docto,
#             "fecha_vcto": c.fecha_vcto,
#             "valor_docto": float(valor_docto),
#             "total_cop": float(total_cop),
#             "recaudo": float(recaudo),
#             "fecha_gestion": c.fecha_gestion,
#             "tipo": c.tipo,
#             "asesor": c.asesor,
#             "observaciones": [obs.texto for obs in c.observaciones]
#         })

#     # Calcular el máximo de días vencidos por cliente
#     for cl in agrupados.values():
#         dias = [f["dias_vencidos"] for f in cl["facturas"] if f["dias_vencidos"] is not None]
#         cl["max_dias"] = max(dias) if dias else None

#     # Renderizado final
#     return templates.TemplateResponse(
#         "index.html",
#         {
#             "request": request,
#             "clientes": list(agrupados.values()),
#             "flash_msg": flash_msg,
#             "flash_type": flash_type,
#             "current_qs": current_qs,  # <- mantiene los filtros en los href
#         }
#     )









# segundo plano: app.py


# from fastapi.responses import JSONResponse, RedirectResponse, StreamingResponse
# from fastapi import UploadFile, File, Form, Query, status
# from database import SessionLocal, engine, Base
# from fastapi.templating import Jinja2Templates
# from fastapi import FastAPI, Request, Depends
# from fastapi.staticfiles import StaticFiles
# from fastapi import HTTPException
# from sqlalchemy.orm import Session
# from sqlalchemy import text  # SQL crudo
# from datetime import datetime
# from typing import Optional
# from decimal import Decimal
# from io import BytesIO
# import pandas as pd
# import models
# import math
# import crud
# import re

# # ------------------- Inicialización -------------------
# app = FastAPI()
# templates = Jinja2Templates(directory="templates")

# app.mount("/static", StaticFiles(directory="static"), name="static")
# app.mount("/assets", StaticFiles(directory="assets"), name="assets")

# # ------------------- Helpers -------------------
# def fmt_money(value):
#     """Filtro para mostrar números como moneda"""
#     if value is None or value == "":
#         return "-"
#     try:
#         v = Decimal(str(value))
#         return f"{v:,.2f}"
#     except Exception:
#         return str(value)

# templates.env.filters["fmt_money"] = fmt_money

# def get_db():
#     db = SessionLocal()
#     try:
#         yield db
#     finally:
#         db.close()

# def to_float(val):
#     try:
#         if pd.isna(val):
#             return None
#         val_str = str(val).replace(",", ".")
#         val_str = re.sub(r"[^0-9.]", "", val_str)
#         return float(val_str) if val_str else None
#     except (ValueError, TypeError):
#         return None

# def clean_phone(val):
#     """Normaliza teléfonos: quita .0, separadores y deja solo dígitos."""
#     if val is None or (isinstance(val, float) and pd.isna(val)):
#         return None
#     s = str(val).strip()

#     # 3046304674.0 -> 3046304674
#     if s.endswith(".0") and s[:-2].isdigit():
#         s = s[:-2]

#     # 3.0463E+09 -> 3046300000
#     try:
#         if re.fullmatch(r"\d+(\.\d+)?[eE][+-]?\d+", s):
#             s = str(int(float(s)))
#     except Exception:
#         pass

#     # 1234.00 -> 1234 si es entero exacto
#     if re.fullmatch(r"\d+\.\d+", s):
#         try:
#             f = float(s)
#             if f.is_integer():
#                 s = str(int(f))
#         except Exception:
#             pass

#     # Dejar solo dígitos
#     s = re.sub(r"[^\d]", "", s)
#     return s or None

# def clean_email(val):
#     """
#     Normaliza emails:
#     - Busca el PRIMER email válido, aunque vengan varios (separados por ; , / o espacios).
#     - Tolera mayúsculas, espacios y texto ruidoso (e.g. 'EMAIL: pep@ACME.com / otro@x.com').
#     """
#     if val is None or (isinstance(val, float) and pd.isna(val)):
#         return None
#     s = str(val).strip()
#     if not s or s.lower() == "nan":
#         return None

#     # Cortes rápidos por separadores comunes
#     candidates = re.split(r"[;,\n/]+", s)

#     email_re = re.compile(r"[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}", re.IGNORECASE)

#     # Primero busca dentro de cada segmento
#     for part in candidates:
#         m = email_re.search(part)
#         if m:
#             return m.group(0).strip().lower()

#     # Si no, busca en todo el string
#     m = email_re.search(s)
#     if m:
#         return m.group(0).strip().lower()

#     return None

# # ---- Null-safe parsers para MSSQL (evitan NaN/Timestamp) ----
# def to_int_nullsafe(val):
#     """Convierte a int o devuelve None si viene vacío/NaN."""
#     try:
#         if val is None:
#             return None
#         if isinstance(val, float) and math.isnan(val):  # NaN
#             return None
#         s = str(val).strip()
#         if s == "":
#             return None
#         return int(float(s))  # admite "5.0"
#     except Exception:
#         return None

# def to_dt_nullsafe(val):
#     """Convierte a datetime (py) o None. Acepta pandas Timestamp, str, excel serial."""
#     if val is None:
#         return None
#     try:
#         if isinstance(val, float) and math.isnan(val):
#             return None
#     except Exception:
#         pass
#     ts = pd.to_datetime(val, errors="coerce")
#     if pd.isna(ts):
#         return None
#     return ts.to_pydatetime()

# # ---------- Helpers para borrado en SQL Server ----------
# def _chunk(iterable, size):
#     buf = []
#     for x in iterable:
#         buf.append(x)
#         if len(buf) >= size:
#             yield buf
#             buf = []
#     if buf:
#         yield buf

# def delete_by_pairs_mssql(db: Session, table_name: str, key_pairs, batch_size: int = 900):
#     """
#     MSSQL-safe:
#     DELETE c
#     FROM <tabla> AS c
#     JOIN (VALUES (...), ...) AS v(nit_cliente, nro_docto_cruce)
#       ON v.nit_cliente = c.nit_cliente
#      AND v.nro_docto_cruce = c.nro_docto_cruce;
#     """
#     pairs = list(key_pairs)
#     if not pairs:
#         return

#     for batch in _chunk(pairs, batch_size):
#         values_clause = []
#         params = {}
#         for i, (nit, doc) in enumerate(batch):
#             values_clause.append(f"(:nit{i}, :doc{i})")
#             params[f"nit{i}"] = str(nit) if nit is not None else None
#             params[f"doc{i}"] = str(doc) if doc is not None else None

#         sql = f"""
#         DELETE c
#         FROM {table_name} AS c
#         JOIN (VALUES {", ".join(values_clause)}) AS v(nit_cliente, nro_docto_cruce)
#           ON v.nit_cliente = c.nit_cliente
#          AND v.nro_docto_cruce = c.nro_docto_cruce;
#         """
#         db.execute(text(sql), params)

# # ================== MEJORAS: FILTROS/ORDEN ESTABLES ==================
# def _to_int_or_none(val):
#     try:
#         return int(val) if val not in (None, "") else None
#     except (ValueError, TypeError):
#         return None

# def _parse_filters(request: Request):
#     min_val = _to_int_or_none(request.query_params.get("min_dias"))
#     max_val = _to_int_or_none(request.query_params.get("max_dias"))
#     sort = request.query_params.get("sort", "dias_desc")
#     return min_val, max_val, sort

# def _apply_filters_and_order(db: Session, min_val, max_val, sort: str):
#     q = db.query(models.Cliente)
#     if min_val is not None:
#         q = q.filter(models.Cliente.dias_vencidos >= min_val)
#     if max_val is not None:
#         q = q.filter(models.Cliente.dias_vencidos <= max_val)

#     # Orden determinístico con tie-breakers
#     if sort == "dias_asc":
#         q = q.order_by(
#             models.Cliente.dias_vencidos.asc(),
#             models.Cliente.nit_cliente.asc(),
#             models.Cliente.nro_docto_cruce.asc(),
#             models.Cliente.id.asc(),
#         )
#     elif sort == "razon_asc":
#         q = q.order_by(
#             models.Cliente.razon_social.asc(),
#             models.Cliente.nit_cliente.asc(),
#             models.Cliente.nro_docto_cruce.asc(),
#             models.Cliente.id.asc(),
#         )
#     elif sort == "nit_asc":
#         q = q.order_by(
#             models.Cliente.nit_cliente.asc(),
#             models.Cliente.nro_docto_cruce.asc(),
#             models.Cliente.fecha_docto.asc().nulls_last(),
#             models.Cliente.id.asc(),
#         )
#     else:  # dias_desc por defecto
#         q = q.order_by(
#             models.Cliente.dias_vencidos.desc(),
#             models.Cliente.nit_cliente.asc(),
#             models.Cliente.nro_docto_cruce.asc(),
#             models.Cliente.id.asc(),
#         )
#     return q
# # ================== FIN MEJORAS ==================

# # ------------------- Exportar cartera a Excel -------------------
# @app.get("/exportar_cartera.xlsx", name="exportar_cartera_xlsx")
# def exportar_cartera_xlsx(
#     db: Session = Depends(get_db),
#     min_dias: Optional[str] = Query(None),
#     max_dias: Optional[str] = Query(None),
#     sort: Optional[str] = Query("dias_desc"),
# ):
#     def to_int_or_none(val: Optional[str]):
#         try:
#             return int(val) if val not in (None, "") else None
#         except (ValueError, TypeError):
#             return None

#     min_val = to_int_or_none(min_dias)
#     max_val = to_int_or_none(max_dias)

#     query = db.query(models.Cliente)
#     if min_val is not None:
#         query = query.filter(models.Cliente.dias_vencidos >= min_val)
#     if max_val is not None:
#         query = query.filter(models.Cliente.dias_vencidos <= max_val)

#     # incluye razon_asc para coincidir con el template
#     order_map = {
#         "dias_desc": models.Cliente.dias_vencidos.desc(),
#         "dias_asc": models.Cliente.dias_vencidos.asc(),
#         "razon_asc": models.Cliente.razon_social.asc(),
#     }
#     query = query.order_by(order_map.get(sort, order_map["dias_desc"]))
#     clientes = query.all()

#     rows = []
#     for c in clientes:
#         valor_docto = Decimal(str(c.valor_docto or 0))
#         total_cop = Decimal(str(c.total_cop or 0))
#         recaudo = Decimal(str(c.recaudo)) if c.recaudo is not None else (valor_docto - total_cop)

#         # --- solo ÚLTIMA observación (si existe) ---
#         obs_txt = ""
#         if getattr(c, "observaciones", None):
#             latest = None
#             for o in c.observaciones:
#                 f = o.fecha_creacion or datetime.min
#                 if latest is None or f > (latest.fecha_creacion or datetime.min):
#                     latest = o
#             if latest:
#                 obs_txt = f"{(latest.fecha_creacion.strftime('%Y-%m-%d %H:%M') if latest.fecha_creacion else '')} - {latest.texto}"

#         rows.append({
#             "ID": c.id,
#             "Razón social": c.razon_social,
#             "NIT": c.nit_cliente,
#             "Docto cruce": c.nro_docto_cruce,
#             "Días vencidos": c.dias_vencidos,
#             "Fecha docto": c.fecha_docto,
#             "Fecha vcto": c.fecha_vcto,
#             "Valor docto": float(valor_docto),
#             "Total COP (saldo)": float(total_cop),
#             "Recaudo": float(recaudo),
#             "Teléfono": c.telefono,
#             "Celular": c.celular,
#             "Correo": c.correo,
#             "Asesor": c.asesor,
#             "Fecha gestión": c.fecha_gestion,
#             "Tipo": c.tipo,
#             "Observaciones": obs_txt,  # SOLO la última
#         })

#     df = pd.DataFrame(rows)

#     for col in ["Fecha docto", "Fecha vcto", "Fecha gestión"]:
#         if col in df.columns:
#             df[col] = pd.to_datetime(df[col], errors="coerce")

#     if not df.empty:
#         resumen = (
#             df.groupby(["NIT", "Razón social"], dropna=False)
#               .agg({
#                   "Valor docto": "sum",
#                   "Total COP (saldo)": "sum",
#                   "Recaudo": "sum",
#                   "Días vencidos": "max",
#               })
#               .rename(columns={
#                   "Valor docto": "Total Valor Docto",
#                   "Total COP (saldo)": "Saldo Total",
#                   "Recaudo": "Recaudo Total",
#                   "Días vencidos": "Max Días Vencidos",
#               })
#               .reset_index()
#         )
#         ref = (
#             df.sort_values(["NIT"]).groupby(["NIT", "Razón social"], dropna=False)
#               .agg({"Teléfono": "first", "Celular": "first", "Correo": "first", "Asesor": "first"})
#               .reset_index()
#         )
#         resumen = resumen.merge(ref, on=["NIT", "Razón social"], how="left")

#         facturas = (
#             df.groupby(["NIT", "Razón social"], dropna=False)
#               .size()
#               .reset_index(name="# Facturas")
#         )
#         resumen = resumen.merge(facturas, on=["NIT", "Razón social"], how="left")
#         resumen["# Facturas"] = resumen["# Facturas"].fillna(0).astype(int)
#     else:
#         resumen = pd.DataFrame(columns=[
#             "NIT","Razón social","Total Valor Docto","Saldo Total","Recaudo Total",
#             "Max Días Vencidos","Teléfono","Celular","Correo","Asesor","# Facturas"
#         ])

#     output = BytesIO()
#     with pd.ExcelWriter(output, engine="xlsxwriter",
#                         datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as writer:
#         df.to_excel(writer, index=False, sheet_name="cartera")
#         resumen.to_excel(writer, index=False, sheet_name="resumen")

#         wb = writer.book
#         ws1 = writer.sheets["cartera"]
#         ws2 = writer.sheets["resumen"]

#         for ws in (ws1, ws2):
#             ws.freeze_panes(1, 0)

#         fmt_header = wb.add_format({
#             'bold': True, 'bg_color': '#D9E1F2', 'border': 1,
#             'align': 'center', 'valign': 'vcenter'
#         })
#         fmt_money0 = wb.add_format({'num_format': '#,##0', 'border': 1})
#         fmt_int = wb.add_format({'num_format': '#,##0', 'border': 1})
#         fmt_date = wb.add_format({'num_format': 'yyyy-mm-dd', 'border': 1})
#         fmt_text = wb.add_format({'text_wrap': True, 'border': 1, 'valign': 'top'})
#         fmt_default = wb.add_format({'border': 1})

#         def table_with_formats(ws, df_sheet, money_cols, date_cols, wrap_cols):
#             rows, cols = df_sheet.shape
#             headers = list(df_sheet.columns)

#             per_col_format = {}
#             for col in headers:
#                 if col in money_cols:
#                     per_col_format[col] = fmt_money0
#                 elif col in date_cols:
#                     per_col_format[col] = fmt_date
#                 elif col in wrap_cols:
#                     per_col_format[col] = fmt_text
#                 elif df_sheet[col].dtype.kind in ("i", "u"):
#                     per_col_format[col] = fmt_int
#                 else:
#                     per_col_format[col] = fmt_default

#             columns_def = [{'header': h, 'format': per_col_format[h]} for h in headers]

#             if rows > 0:
#                 ws.add_table(0, 0, rows, cols - 1, {
#                     'style': 'Table Style Medium 9',
#                     'banded_rows': True,
#                     'columns': columns_def
#                 })
#             else:
#                 for c_idx, h in enumerate(headers):
#                     ws.write(0, c_idx, h, fmt_header)

#             col_widths = {}
#             for c_idx, col in enumerate(headers):
#                 max_len = len(str(col)) + 2
#                 serie = df_sheet[col]

#                 if col in money_cols:
#                     for v in serie.dropna():
#                         try:
#                             s = f"{int(round(float(v))):,}"
#                             max_len = max(max_len, len(s))
#                         except Exception:
#                             pass
#                 elif col in date_cols:
#                     max_len = max(max_len, 12)
#                 else:
#                     for v in serie.dropna():
#                         s = str(v).replace("\r", "")
#                         max_len = max(max_len, max((len(seg) for seg in s.split("\n")), default=0))
#                     if col in wrap_cols:
#                         max_len = min(max(max_len, 40), 80)
#                     else:
#                         max_len = min(max_len, 40)

#                 width = max_len + 1
#                 col_widths[col] = width
#                 ws.set_column(c_idx, c_idx, width, per_col_format[col])

#             return col_widths, per_col_format

#         def autofit_row_heights(ws, df_sheet, wrap_cols, col_widths):
#             if df_sheet.empty or not wrap_cols:
#                 return
#             base_height = 15
#             max_height = 300

#             for r in range(1, len(df_sheet) + 1):
#                 max_lines = 1
#                 for col in wrap_cols:
#                     if col not in df_sheet.columns:
#                         continue
#                     val = df_sheet.iloc[r - 1][col]
#                     if pd.isna(val) or val is None:
#                         continue
#                     text = str(val).replace("\r", "")
#                     if text == "":
#                         continue

#                     col_w = int(col_widths.get(col, 40))
#                     chars_per_line = max(col_w - 2, 10)

#                     total_lines = 0
#                     for seg in text.split("\n"):
#                         seg = seg.strip()
#                         if seg == "":
#                             total_lines += 1
#                         else:
#                             total_lines += math.ceil(len(seg) / chars_per_line)

#                     max_lines = max(max_lines, total_lines)

#                 ws.set_row(r, min(base_height * max_lines + 4, max_height))

#         money_cols_cartera = ["Valor docto", "Total COP (saldo)", "Recaudo"]
#         date_cols_cartera = ["Fecha docto", "Fecha vcto", "Fecha gestión"]
#         wrap_cols_cartera = ["Razón social", "Observaciones"]

#         money_cols_resumen = ["Total Valor Docto", "Saldo Total", "Recaudo Total"]
#         date_cols_resumen = []
#         wrap_cols_resumen = ["Razón social"]

#         widths_cartera, _ = table_with_formats(ws1, df, money_cols_cartera, date_cols_cartera, wrap_cols_cartera)
#         widths_resumen, _ = table_with_formats(ws2, resumen, money_cols_resumen, date_cols_resumen, wrap_cols_resumen)

#         autofit_row_heights(ws1, df, wrap_cols_cartera, widths_cartera)
#         autofit_row_heights(ws2, resumen, wrap_cols_resumen, widths_resumen)

#     output.seek(0)
#     filename = f"cartera_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
#     return StreamingResponse(
#         output,
#         media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#         headers={"Content-Disposition": f'attachment; filename="{filename}"'}
#     )

# # ------------------- Importar cartera -------------------
# @app.post("/importar_excel")
# async def importar_excel(file: UploadFile = File(...), db: Session = Depends(get_db)):
#     try:
#         contents = await file.read()

#         # --- Cargar Excel/CSV (ambos formatos) ---
#         def _read_any(contents: bytes, filename: str):
#             name = (filename or "").lower()
#             if name.endswith(".csv"):
#                 df = pd.read_csv(BytesIO(contents), dtype=str, keep_default_na=False)
#                 return {"__csv__": df}
#             return pd.read_excel(BytesIO(contents), sheet_name=None, dtype=str, keep_default_na=False)

#         sheets = _read_any(contents, file.filename)

#         # --- Normalizador simple de strings ---
#         def norm(s: str) -> str:
#             if s is None:
#                 return ""
#             t = str(s).strip().lower()
#             rep = {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ñ":"n","ä":"a","ë":"e","ï":"i","ö":"o","ü":"u","’":"'","´":"'","`":"'"}
#             for a,b in rep.items(): t = t.replace(a,b)
#             t = re.sub(r"\s+", " ", t)
#             return t

#         # --- Aliases (internos -> alias) ---
#         alias_map = {
#             "razon_social": {"razon social", "razon social vend. cliente", "razon social vend cliente"},
#             "nit_cliente": {"nit", "nit cliente despacho", "nit cliente"},
#             "nro_docto_cruce": {"docto cruce", "nro. docto. cruce", "nro docto cruce", "factura", "documento", "nro documento"},
#             "fecha_docto": {"fecha docto", "fecha docto."},
#             "fecha_vcto": {"fecha vcto", "fecha vcto."},
#             "dias_vencidos": {"dias vencidos", "dias vencidos."},
#             "valor_docto": {"valor docto", "valor docto."},
#             "total_cop": {"total cop", "total cop (saldo)", "saldo", "total (cop)"},
#             "telefono": {"telefono", "teléfono"},
#             "celular": {"celular", "movil", "móvil"},
#             "correo": {"correo", "email", "e-mail", "mail", "correo electronico", "correo electrónico"},
#             "asesor": {"asesor", "vendedor", "razon social vend. cliente", "razon social vend cliente"},
#             "obs_txt": {"observaciones"},
#         }

#         # --- Detectar si es export del proyecto (hoja 'cartera' con NIT/Docto) ---
#         is_project_export = False
#         df_project = None
#         for sheet_name, df in sheets.items():
#             if norm(sheet_name) == "cartera":
#                 df_project = df
#                 break
#         if df_project is not None:
#             cols_n = {norm(c) for c in df_project.columns}
#             if "nit" in cols_n and ("docto cruce" in cols_n or "nro. docto. cruce" in cols_n or "nro docto cruce" in cols_n):
#                 is_project_export = True

#         # --- Elegir hoja candidata ---
#         if is_project_export:
#             df_raw = df_project.copy()
#         else:
#             df_raw = None
#             for _, df in sheets.items():
#                 cols_n = {norm(c) for c in df.columns}
#                 if any(x in cols_n for x in alias_map["nit_cliente"]) and \
#                    any(x in cols_n for x in alias_map["nro_docto_cruce"]):
#                     df_raw = df
#                     break
#             if df_raw is None:
#                 return RedirectResponse(
#                     url="/?msg=No%20se%20encontro%20hoja%20compatible%20(ni%20cartera%20ni%20columnas%20clave)&msg_type=error",
#                     status_code=303
#                 )

#         # Copia de encabezados originales por si necesitamos "rescatar" correo
#         original_cols = list(df_raw.columns)

#         df_raw.columns = [str(c).strip() for c in df_raw.columns]

#         # --- Renombrado dinámico por alias ---
#         rename_map_dynamic = {}
#         for col in df_raw.columns:
#             n = norm(col)
#             for internal, aliases in alias_map.items():
#                 if n in aliases and internal not in rename_map_dynamic.values():
#                     rename_map_dynamic[col] = internal
#                     break

#         # Encabezados exactos del export del proyecto
#         project_head_map = {
#             "Razón social": "razon_social",
#             "NIT": "nit_cliente",
#             "Docto cruce": "nro_docto_cruce",
#             "Días vencidos": "dias_vencidos",
#             "Fecha docto": "fecha_docto",
#             "Fecha vcto": "fecha_vcto",
#             "Valor docto": "valor_docto",
#             "Total COP (saldo)": "total_cop",
#             "Teléfono": "telefono",
#             "Celular": "celular",
#             "Correo": "correo",
#             "Email": "correo",      # soporte explícito
#             "E-mail": "correo",     # soporte explícito
#             "Asesor": "asesor",
#             "Observaciones": "obs_txt",
#         }
#         for col in df_raw.columns:
#             if col in project_head_map and project_head_map[col] not in rename_map_dynamic.values():
#                 rename_map_dynamic[col] = project_head_map[col]

#         df = df_raw.rename(columns=rename_map_dynamic)

#         # --- Asegurar/Coalescer columna 'correo' desde cualquier variante ---
#         def _norm2(s: str) -> str:
#             if s is None:
#                 return ""
#             t = str(s).strip().lower()
#             rep = {"á":"a","é":"e","í":"i","ó":"o","ú":"u","ñ":"n"}
#             for a,b in rep.items(): t = t.replace(a,b)
#             t = re.sub(r"\s+", " ", t)
#             return t

#         email_cols_candidates = []
#         for col in df_raw.columns:
#             n = _norm2(col)
#             if re.fullmatch(r"(correo( electronico)?)|(e[\s-]?mail)|email|mail", n):
#                 email_cols_candidates.append(col)

#         if "correo" not in df.columns:
#             if email_cols_candidates:
#                 correo_series = None
#                 for col in email_cols_candidates:
#                     s = df_raw[col].astype(str)
#                     s = s.where(~s.str.fullmatch(r"(?i)\s*nan\s*"), "")
#                     if correo_series is None:
#                         correo_series = s
#                     else:
#                         correo_series = correo_series.mask(
#                             (correo_series.astype(str).str.strip() == "") & (s.astype(str).str.strip() != ""),
#                             s
#                         )
#                 df["correo"] = correo_series
#             else:
#                 df["correo"] = None

#         # Normalizar emails al final
#         df["correo"] = df["correo"].apply(clean_email)

#         # Validar claves mínimas
#         if "nit_cliente" not in df.columns or "nro_docto_cruce" not in df.columns:
#             return RedirectResponse(
#                 url="/?msg=El%20archivo%20no%20tiene%20las%20columnas%20clave%20(NIT%20y%20Docto%20cruce)&msg_type=error",
#                 status_code=303
#             )

#         # --- Normalizaciones de campos ---
#         if "telefono" in df.columns: df["telefono"] = df["telefono"].apply(clean_phone)
#         if "celular" in df.columns: df["celular"] = df["celular"].apply(clean_phone)

#         if "fecha_docto" in df.columns: df["fecha_docto"] = df["fecha_docto"].apply(to_dt_nullsafe)
#         if "fecha_vcto"  in df.columns: df["fecha_vcto"]  = df["fecha_vcto"].apply(to_dt_nullsafe)
#         if "dias_vencidos" in df.columns: df["dias_vencidos"] = df["dias_vencidos"].apply(to_int_nullsafe)

#         def _to_money(x):
#             v = to_float(x)
#             return v if v is not None else 0.0

#         if "valor_docto" in df.columns:
#             df["valor_docto"] = df["valor_docto"].apply(_to_money)
#         else:
#             df["valor_docto"] = 0.0

#         if "total_cop" in df.columns:
#             df["total_cop"] = df["total_cop"].apply(_to_money)
#         else:
#             df["total_cop"] = df["valor_docto"]

#         # Claves limpias
#         df = df[(df["nit_cliente"].notna()) & (df["nro_docto_cruce"].notna())]
#         df["nit_cliente"] = df["nit_cliente"].astype(str).str.strip()
#         df["nro_docto_cruce"] = df["nro_docto_cruce"].astype(str).str.strip()
#         df = df[(df["nit_cliente"] != "") & (df["nro_docto_cruce"] != "")]

#         if df.empty:
#             return RedirectResponse(url="/?msg=Archivo%20vacio%20o%20sin%20filas%20validas&msg_type=info", status_code=303)

#         # --- Claves existentes en BD ---
#         bd_clientes = db.query(models.Cliente).all()
#         bd_claves = set((str(c.nit_cliente), str(c.nro_docto_cruce)) for c in bd_clientes)

#         # --- Borrado condicional ---
#         if not is_project_export:
#             excel_claves = set(zip(df["nit_cliente"], df["nro_docto_cruce"]))
#             claves_a_eliminar = bd_claves - excel_claves
#             if claves_a_eliminar:
#                 delete_by_pairs_mssql(db, "cartera", claves_a_eliminar, batch_size=900)

#         # Helpers de obs
#         def _clean_obs_text(s: Optional[str]) -> Optional[str]:
#             if not s:
#                 return None
#             txt = str(s).strip()
#             m = re.match(r"^\s*(\d{4}-\d{2}-\d{2}(?:\s+\d{2}:\d{2})?)\s*-\s*(.+)$", txt)
#             if m:
#                 return m.group(2).strip() or None
#             return txt or None

#         def _latest_obs_text(cliente):
#             latest = None
#             for o in getattr(cliente, "observaciones", []) or []:
#                 f = getattr(o, "fecha_creacion", None) or datetime.min
#                 if latest is None or f > (getattr(latest, "fecha_creacion", None) or datetime.min):
#                     latest = o
#             return (latest.texto.strip() if latest and latest.texto else None)

#         # --- Upsert + Observaciones (solo si viene del export) ---
#         for _, row in df.iterrows():
#             nit = row["nit_cliente"]
#             docto = row["nro_docto_cruce"]

#             valor_docto = row.get("valor_docto", 0.0) or 0.0
#             total_excel = row.get("total_cop", valor_docto)
#             recaudo_calc = round(max((valor_docto or 0.0) - (total_excel or 0.0), 0.0), 2)

#             cliente = db.query(models.Cliente).filter_by(
#                 nit_cliente=nit,
#                 nro_docto_cruce=docto
#             ).first()

#             correo_val = clean_email(row.get("correo")) if "correo" in row else None

#             if cliente:
#                 cliente.razon_social = row.get("razon_social")
#                 cliente.dias_vencidos = to_int_nullsafe(row.get("dias_vencidos"))
#                 cliente.fecha_docto = to_dt_nullsafe(row.get("fecha_docto"))
#                 cliente.fecha_vcto = to_dt_nullsafe(row.get("fecha_vcto"))
#                 cliente.valor_docto = valor_docto
#                 cliente.total_cop = total_excel
#                 cliente.recaudo = recaudo_calc
#                 cliente.telefono = clean_phone(row.get("telefono"))
#                 cliente.celular  = clean_phone(row.get("celular"))
#                 cliente.asesor = row.get("asesor")
#                 if correo_val and (not cliente.correo or cliente.correo != correo_val):
#                     cliente.correo = correo_val
#             else:
#                 cliente = models.Cliente(
#                     razon_social=row.get("razon_social"),
#                     nit_cliente=nit,
#                     nro_docto_cruce=docto,
#                     dias_vencidos=to_int_nullsafe(row.get("dias_vencidos")),
#                     fecha_docto=to_dt_nullsafe(row.get("fecha_docto")),
#                     fecha_vcto=to_dt_nullsafe(row.get("fecha_vcto")),
#                     valor_docto=valor_docto,
#                     total_cop=total_excel,
#                     recaudo=recaudo_calc,
#                     telefono=clean_phone(row.get("telefono")),
#                     celular=clean_phone(row.get("celular")),
#                     asesor=row.get("asesor"),
#                     correo=correo_val,
#                 )
#                 db.add(cliente)
#                 db.flush()

#             if "obs_txt" in df.columns and is_project_export:
#                 new_obs_text = _clean_obs_text(row.get("obs_txt"))
#                 if new_obs_text:
#                     last_text = _latest_obs_text(cliente)
#                     if (last_text or "").strip() != new_obs_text.strip():
#                         obs = models.Observacion(
#                             cliente_id=cliente.id,
#                             texto=new_obs_text,
#                         )
#                         db.add(obs)

#         db.commit()
#         return RedirectResponse(url="/?msg=Archivo%20subido%20correctamente&msg_type=success", status_code=303)

#     except Exception as e:
#         print("❌ Error importar_excel:", e)
#         return RedirectResponse(url="/?msg=Error%20al%20subir%20el%20archivo&msg_type=error", status_code=303)


# # ------------------- Observaciones -------------------
# @app.post("/cliente/{cliente_id}/observacion")
# def agregar_observacion(cliente_id: int, texto: str = Form(...), db: Session = Depends(get_db)):
#     crud.add_observacion(db, cliente_id, texto)
#     return RedirectResponse(
#         url=f"/cliente/{cliente_id}?msg=Observaci%C3%B3n%20guardada&msg_type=success",
#         status_code=303
#     )

# # ------------------- Actualizar cliente (editar recaudo) -------------------
# @app.post("/cliente/{cliente_id}/update")
# async def update_cliente(
#     request: Request,
#     cliente_id: int,
#     db: Session = Depends(get_db)
# ):
#     form_data = await request.form()

#     cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
#     if not cliente:
#         return RedirectResponse(url="/?msg=Cliente%20no%20encontrado&msg_type=error", status_code=status.HTTP_303_SEE_OTHER)

#     for campo in ["telefono", "celular", "correo", "fecha_gestion", "tipo", "asesor"]:
#         val = form_data.get(campo)
#         if val is not None and val != "":
#             if campo == "correo":
#                 val = clean_email(val)
#             if campo in ("telefono", "celular"):
#                 val = clean_phone(val)
#             setattr(cliente, campo, val)

#     valor_docto = float(cliente.valor_docto or 0.0)

#     raw_recaudo = form_data.get("recaudo")
#     raw_total_cop = form_data.get("total_cop")

#     has_recaudo = raw_recaudo is not None and raw_recaudo.strip() != ""
#     has_total_cop = raw_total_cop is not None and raw_total_cop.strip() != ""

#     if has_recaudo:
#         delta = to_float(raw_recaudo)
#         if delta is not None:
#             old_total = float(cliente.total_cop or 0.0)
#             new_total = max(round(old_total - delta, 2), 0.0)
#             cliente.total_cop = new_total
#             cliente.recaudo = round(max(valor_docto - new_total, 0.0), 2)
#     elif has_total_cop:
#         nuevo_total = to_float(raw_total_cop)
#         if nuevo_total is not None:
#             nuevo_total = max(nuevo_total, 0.0)
#             if nuevo_total > valor_docto:
#                 nuevo_total = valor_docto
#             cliente.total_cop = round(nuevo_total, 2)
#             cliente.recaudo = round(max(valor_docto - nuevo_total, 0.0), 2)

#     nueva_obs = form_data.get("observaciones")
#     if nueva_obs and nueva_obs.strip():
#         obs = models.Observacion(
#             cliente_id=cliente.id,
#             texto=nueva_obs.strip(),
#         )
#         db.add(obs)

#     db.commit()
#     return RedirectResponse(
#         url=f"/cliente/{cliente_id}?msg=Cambios%20guardados&msg_type=success",
#         status_code=status.HTTP_303_SEE_OTHER
#     )

# # ------------------- historial cliente -------------------
# @app.get("/cliente/{cliente_id}/historial")
# def historial_cliente(cliente_id: int, db: Session = Depends(get_db)):
#     cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
#     if not cliente:
#         return JSONResponse(content=[], status_code=200)

#     historial = [
#         {
#             "texto": obs.texto,
#             "fecha": obs.fecha_creacion.strftime("%d/%m/%Y %H:%M")
#             if obs.fecha_creacion else None
#         }
#         for obs in cliente.observaciones
#     ]

#     return JSONResponse(content=historial, status_code=200)

# # ------------------- Vista cliente (MEJORADA: navegación + filtros) -------------------
# @app.get("/cliente/{cliente_id}")
# def ver_cliente(cliente_id: int, request: Request, db: Session = Depends(get_db)):
#     cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
#     if not cliente:
#         raise HTTPException(status_code=404, detail="Cliente no encontrado")

#     # Persistir filtros al regresar + reconstruir filtros para secuencia Prev/Next
#     ref_qs = request.query_params.get("ref")

#     # Construye un objeto con query_params = ref_qs para reutilizar _parse_filters()
#     from starlette.datastructures import QueryParams
#     req_like = request
#     if ref_qs:
#         req_like = type("R", (), {"query_params": QueryParams(ref_qs)})

#     min_val, max_val, sort = _parse_filters(req_like)

#     # MISMA lista (mismos filtros/orden) para calcular prev/next
#     ordered_ids = [row.id for row in _apply_filters_and_order(db, min_val, max_val, sort)
#                                       .with_entities(models.Cliente.id).all()]

#     prev_id = next_id = None
#     try:
#         idx = ordered_ids.index(cliente_id)
#         if idx > 0:
#             prev_id = ordered_ids[idx - 1]
#         if idx < len(ordered_ids) - 1:
#             next_id = ordered_ids[idx + 1]
#     except ValueError:
#         pass

#     back_url = f"/?{ref_qs}" if ref_qs else "/"

#     return templates.TemplateResponse(
#         "cliente.html",
#         {
#             "request": request,
#             "cliente": cliente,
#             "prev_id": prev_id,
#             "next_id": next_id,
#             "flash_msg": request.query_params.get("msg"),
#             "flash_type": request.query_params.get("msg_type"),
#             "back_url": back_url,  # para el botón "Regresar"
#         }
#     )

# # ------------------- Index agrupado (MEJORADO: orden estable) -------------------
# @app.get("/")
# def index(
#     request: Request,
#     db: Session = Depends(get_db),
#     view: Optional[str] = Query(None),
#     min_dias: Optional[str] = Query(None),
#     max_dias: Optional[str] = Query(None),
#     sort: Optional[str] = Query("dias_desc"),
# ):
#     # USAR SIEMPRE la misma función para filtros/orden
#     min_val, max_val, sort = _parse_filters(request)
#     query = _apply_filters_and_order(db, min_val, max_val, sort)
#     clientes = query.all()

#     # Flash + QS actual para persistir filtros en links
#     flash_msg = request.query_params.get("msg")
#     flash_type = request.query_params.get("msg_type")
#     if (min_val is not None) or (max_val is not None):
#         if not flash_msg:
#             flash_msg = "Filtros aplicados"
#         if not flash_type:
#             flash_type = "info"

#     current_qs = request.url.query  # <- pásalo a las plantillas

#     if view == "flat":
#         filas = []
#         for c in clientes:
#             valor_docto = Decimal(str(c.valor_docto or 0))
#             total_cop = Decimal(str(c.total_cop or 0))
#             recaudo = valor_docto - total_cop

#             filas.append({
#                 "id": c.id,
#                 "razon_social": c.razon_social,
#                 "nit_cliente": c.nit_cliente,
#                 "nro_docto_cruce": c.nro_docto_cruce,
#                 "dias_vencidos": c.dias_vencidos,
#                 "fecha_docto": c.fecha_docto,
#                 "fecha_vcto": c.fecha_vcto,
#                 "valor_docto": float(valor_docto),
#                 "total_cop": float(total_cop),
#                 "recaudo": float(recaudo),
#                 "asesor": c.asesor,
#                 "correo": c.correo,
#             })

#         return templates.TemplateResponse(
#             "index.html",
#             {
#                 "request": request,
#                 "view": "flat",
#                 "filas": filas,
#                 "flash_msg": flash_msg,
#                 "flash_type": flash_type,
#                 "current_qs": current_qs,  # <- usa esto en los href
#             }
#         )

#     agrupados = {}
#     for c in clientes:
#         if c.nit_cliente not in agrupados:
#             agrupados[c.nit_cliente] = {
#                 "nit_cliente": c.nit_cliente,
#                 "razon_social": c.razon_social,
#                 "telefono": c.telefono,
#                 "celular": c.celular,
#                 "correo": c.correo,
#                 "asesor": c.asesor,
#                 "facturas": []
#             }

#         valor_docto = Decimal(str(c.valor_docto or 0))
#         total_cop = Decimal(str(c.total_cop or 0))
#         recaudo = valor_docto - total_cop

#         agrupados[c.nit_cliente]["facturas"].append({
#             "id": c.id,
#             "nro_docto_cruce": c.nro_docto_cruce,
#             "dias_vencidos": c.dias_vencidos,
#             "fecha_docto": c.fecha_docto,
#             "fecha_vcto": c.fecha_vcto,
#             "valor_docto": float(valor_docto),
#             "total_cop": float(total_cop),
#             "recaudo": float(recaudo),
#             "fecha_gestion": c.fecha_gestion,
#             "tipo": c.tipo,
#             "asesor": c.asesor,
#             "observaciones": [obs.texto for obs in c.observaciones]
#         })

#     for cl in agrupados.values():
#         dias = [f["dias_vencidos"] for f in cl["facturas"] if f["dias_vencidos"] is not None]
#         cl["max_dias"] = max(dias) if dias else None

#     return templates.TemplateResponse(
#         "index.html",
#         {
#             "request": request,
#             "clientes": list(agrupados.values()),
#             "flash_msg": flash_msg,
#             "flash_type": flash_type,
#             "current_qs": current_qs,  # <- usa esto en los href
#         }
#     )


# ------------------- Imports y configuración inicial -------------------

# from fastapi.responses import JSONResponse, RedirectResponse
# from fastapi import UploadFile, File, Form, Query, status
# from fastapi.responses import StreamingResponse
# from database import SessionLocal, engine, Base
# from fastapi.templating import Jinja2Templates
# from fastapi import FastAPI, Request, Depends
# from fastapi.staticfiles import StaticFiles
# from fastapi import HTTPException
# from sqlalchemy.orm import Session
# from sqlalchemy import text  # <-- usar SQL crudo MSSQL-safe
# from datetime import datetime
# from typing import Optional
# from decimal import Decimal
# from io import BytesIO
# import pandas as pd
# import models
# import math
# import crud
# import re


# # ------------------- Inicialización -------------------


# app = FastAPI()
# templates = Jinja2Templates(directory="templates")

# app.mount("/static", StaticFiles(directory="static"), name="static")
# app.mount("/assets", StaticFiles(directory="assets"), name="assets")

# # ------------------- Helpers -------------------
# def fmt_money(value):
#     """Filtro para mostrar números como moneda"""
#     if value is None or value == "":
#         return "-"
#     try:
#         v = Decimal(str(value))
#         return f"{v:,.2f}"
#     except Exception:
#         return str(value)

# templates.env.filters["fmt_money"] = fmt_money


# def get_db():
#     db = SessionLocal()
#     try:
#         yield db
#     finally:
#         db.close()


# def to_float(val):
#     try:
#         if pd.isna(val):
#             return None
#         val_str = str(val).replace(",", ".")
#         val_str = re.sub(r"[^0-9.]", "", val_str)
#         return float(val_str) if val_str else None
#     except (ValueError, TypeError):
#         return None
    
# def clean_phone(val):
#     """Normaliza teléfonos: quita .0, separadores y deja solo dígitos."""
#     if val is None or (isinstance(val, float) and pd.isna(val)):
#         return None
#     s = str(val).strip()

#     # 3046304674.0 -> 3046304674
#     if s.endswith(".0") and s[:-2].isdigit():
#         s = s[:-2]

#     # 3.046304674E+09 -> 3046304674
#     try:
#         if re.fullmatch(r"\d+(\.\d+)?[eE][+-]?\d+", s):
#             s = str(int(float(s)))
#     except Exception:
#         pass

#     # 1234.00 -> 1234 si es entero exacto
#     if re.fullmatch(r"\d+\.\d+", s):
#         try:
#             f = float(s)
#             if f.is_integer():
#                 s = str(int(f))
#         except Exception:
#             pass

#     # Dejar solo dígitos (quita espacios, guiones, paréntesis, etc.)
#     s = re.sub(r"[^\d]", "", s)
#     return s or None


# # ---------- Helpers nuevos para borrado MSSQL-safe ----------
# def _chunk(iterable, size):
#     buf = []
#     for x in iterable:
#         buf.append(x)
#         if len(buf) >= size:
#             yield buf
#             buf = []
#     if buf:
#         yield buf

# def delete_by_pairs_mssql(db: Session, table_name: str, key_pairs, batch_size: int = 900):
#     """
#     Elimina filas por (nit_cliente, nro_docto_cruce) usando
#     DELETE .. FROM .. JOIN (VALUES ...)  [compatible SQL Server].
#     batch_size controla cuántos pares por lote (900 -> 1800 params).
#     """
#     pairs = list(key_pairs)
#     if not pairs:
#         return
#     for batch in _chunk(pairs, batch_size):
#         values_clause = []
#         params = {}
#         for i, (nit, doc) in enumerate(batch):
#             values_clause.append(f"(:nit{i}, :doc{i})")
#             params[f"nit{i}"] = str(nit) if nit is not None else None
#             params[f"doc{i}"] = str(doc) if doc is not None else None

#         sql = f"""
#         DELETE c
#         FROM {table_name} AS c
#         JOIN (VALUES {", ".join(values_clause)}) AS v(nit_cliente, nro_docto_cruce)
#           ON v.nit_cliente = c.nit_cliente
#          AND v.nro_docto_cruce = c.nro_docto_cruce;
#         """
#         db.execute(text(sql), params)


# # ------------------- Exportar cartera a Excel -------------------
# @app.get("/exportar_cartera.xlsx", name="exportar_cartera_xlsx")
# def exportar_cartera_xlsx(
#     db: Session = Depends(get_db),
#     min_dias: Optional[str] = Query(None),
#     max_dias: Optional[str] = Query(None),
#     sort: Optional[str] = Query("dias_desc"),
# ):
#     # --- Helpers ---
#     def to_int_or_none(val: Optional[str]):
#         try:
#             return int(val) if val not in (None, "") else None
#         except (ValueError, TypeError):
#             return None

#     min_val = to_int_or_none(min_dias)
#     max_val = to_int_or_none(max_dias)

#     # --- Query + filtros + orden ---
#     query = db.query(models.Cliente)
#     if min_val is not None:
#         query = query.filter(models.Cliente.dias_vencidos >= min_val)
#     if max_val is not None:
#         query = query.filter(models.Cliente.dias_vencidos <= max_val)

#     order_map = {
#         "dias_desc": models.Cliente.dias_vencidos.desc(),
#         "dias_asc": models.Cliente.dias_vencidos.asc(),
#     }
#     query = query.order_by(order_map.get(sort, order_map["dias_desc"]))
#     clientes = query.all()

#     # --- Filas planas ---
#     rows = []
#     for c in clientes:
#         valor_docto = Decimal(str(c.valor_docto or 0))
#         total_cop = Decimal(str(c.total_cop or 0))
#         recaudo = Decimal(str(c.recaudo)) if c.recaudo is not None else (valor_docto - total_cop)

#         obs_txt = ""
#         if getattr(c, "observaciones", None):
#             obs_txt = "\n".join(
#                 f"{(o.fecha_creacion.strftime('%Y-%m-%d %H:%M') if o.fecha_creacion else '')} - {o.texto}"
#                 for o in c.observaciones
#             )

#         rows.append({
#             "ID": c.id,
#             "Razón social": c.razon_social,
#             "NIT": c.nit_cliente,
#             "Docto cruce": c.nro_docto_cruce,
#             "Días vencidos": c.dias_vencidos,
#             "Fecha docto": c.fecha_docto,
#             "Fecha vcto": c.fecha_vcto,
#             "Valor docto": float(valor_docto),
#             "Total COP (saldo)": float(total_cop),
#             "Recaudo": float(recaudo),
#             "Teléfono": c.telefono,
#             "Celular": c.celular,
#             "Asesor": c.asesor,
#             "Fecha gestión": c.fecha_gestion,
#             "Tipo": c.tipo,
#             "Observaciones": obs_txt,
#         })

#     df = pd.DataFrame(rows)

#     # Normaliza fechas a datetime (Excel-friendly)
#     for col in ["Fecha docto", "Fecha vcto", "Fecha gestión"]:
#         if col in df.columns:
#             df[col] = pd.to_datetime(df[col], errors="coerce")

#     # --- Resumen por cliente ---
#     if not df.empty:
#         resumen = (
#             df.groupby(["NIT", "Razón social"], dropna=False)
#               .agg({
#                   "Valor docto": "sum",
#                   "Total COP (saldo)": "sum",
#                   "Recaudo": "sum",
#                   "Días vencidos": "max",
#               })
#               .rename(columns={
#                   "Valor docto": "Total Valor Docto",
#                   "Total COP (saldo)": "Saldo Total",
#                   "Recaudo": "Recaudo Total",
#                   "Días vencidos": "Max Días Vencidos",
#               })
#               .reset_index()
#         )
#         ref = (
#             df.sort_values(["NIT"]).groupby(["NIT", "Razón social"], dropna=False)
#               .agg({"Teléfono": "first", "Celular": "first", "Asesor": "first"})
#               .reset_index()
#         )
#         resumen = resumen.merge(ref, on=["NIT", "Razón social"], how="left")

#         facturas = (
#             df.groupby(["NIT", "Razón social"], dropna=False)
#               .size()
#               .reset_index(name="# Facturas")
#         )
#         resumen = resumen.merge(facturas, on=["NIT", "Razón social"], how="left")
#         resumen["# Facturas"] = resumen["# Facturas"].fillna(0).astype(int)
#     else:
#         resumen = pd.DataFrame(columns=[
#             "NIT","Razón social","Total Valor Docto","Saldo Total","Recaudo Total",
#             "Max Días Vencidos","Teléfono","Celular","Asesor","# Facturas"
#         ])

#     # --- Excel en memoria ---
#     output = BytesIO()
#     with pd.ExcelWriter(output, engine="xlsxwriter",
#                         datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as writer:
#         # Escribir hojas
#         df.to_excel(writer, index=False, sheet_name="cartera")
#         resumen.to_excel(writer, index=False, sheet_name="resumen")

#         wb = writer.book
#         ws1 = writer.sheets["cartera"]
#         ws2 = writer.sheets["resumen"]

#         # Congelar encabezado
#         for ws in (ws1, ws2):
#             ws.freeze_panes(1, 0)

#         # --- Formatos ---
#         fmt_header = wb.add_format({
#             'bold': True, 'bg_color': '#D9E1F2', 'border': 1,
#             'align': 'center', 'valign': 'vcenter'
#         })
#         fmt_money0 = wb.add_format({'num_format': '#,##0', 'border': 1})
#         fmt_int = wb.add_format({'num_format': '#,##0', 'border': 1})
#         fmt_date = wb.add_format({'num_format': 'yyyy-mm-dd', 'border': 1})
#         fmt_text = wb.add_format({'text_wrap': True, 'border': 1, 'valign': 'top'})
#         fmt_default = wb.add_format({'border': 1})

#         def table_with_formats(ws, df_sheet, money_cols, date_cols, wrap_cols):
#             """Crea tabla con formatos por columna y calcula anchos."""
#             rows, cols = df_sheet.shape
#             headers = list(df_sheet.columns)

#             # Mapear formato por columna
#             per_col_format = {}
#             for col in headers:
#                 if col in money_cols:
#                     per_col_format[col] = fmt_money0
#                 elif col in date_cols:
#                     per_col_format[col] = fmt_date
#                 elif col in wrap_cols:
#                     per_col_format[col] = fmt_text
#                 elif df_sheet[col].dtype.kind in ("i", "u"):  # enteros
#                     per_col_format[col] = fmt_int
#                 else:
#                     per_col_format[col] = fmt_default

#             # Definir columnas de la tabla con formato (esto sí aplica dentro de la tabla)
#             columns_def = [{'header': h, 'format': per_col_format[h]} for h in headers]

#             # Crear tabla (si hay datos). Si no, solo pintamos encabezados.
#             if rows > 0:
#                 ws.add_table(0, 0, rows, cols - 1, {
#                     'style': 'Table Style Medium 9',
#                     'banded_rows': True,
#                     'columns': columns_def
#                 })
#             else:
#                 for c_idx, h in enumerate(headers):
#                     ws.write(0, c_idx, h, fmt_header)

#             # Auto–ancho aproximado por columna
#             col_widths = {}
#             for c_idx, col in enumerate(headers):
#                 max_len = len(str(col)) + 2
#                 serie = df_sheet[col]

#                 if col in money_cols:
#                     # medir como miles sin decimales
#                     for v in serie.dropna():
#                         try:
#                             s = f"{int(round(float(v))):,}"
#                             max_len = max(max_len, len(s))
#                         except Exception:
#                             pass
#                 elif col in date_cols:
#                     max_len = max(max_len, 12)
#                 else:
#                     # considerar saltos de línea (tomar la línea más larga)
#                     for v in serie.dropna():
#                         s = str(v).replace("\r", "")
#                         max_len = max(max_len, max((len(seg) for seg in s.split("\n")), default=0))
#                     # límite práctico
#                     if col in wrap_cols:
#                         max_len = min(max(max_len, 40), 80)
#                     else:
#                         max_len = min(max_len, 40)

#                 width = max_len + 1
#                 col_widths[col] = width
#                 ws.set_column(c_idx, c_idx, width, per_col_format[col])

#             return col_widths, per_col_format

#         def autofit_row_heights(ws, df_sheet, wrap_cols, col_widths):
#             """Ajusta alto de fila según # líneas estimadas en columnas wrap."""
#             if df_sheet.empty or not wrap_cols:
#                 return
#             base_height = 15  # pts aprox por línea
#             max_height = 300

#             for r in range(1, len(df_sheet) + 1):  # +1 por header
#                 max_lines = 1
#                 for col in wrap_cols:
#                     if col not in df_sheet.columns:
#                         continue
#                     val = df_sheet.iloc[r - 1][col]
#                     if pd.isna(val) or val is None:
#                         continue
#                     text = str(val).replace("\r", "")
#                     if text == "":
#                         continue

#                     # Estimar caracteres por línea a partir del ancho de columna
#                     col_w = int(col_widths.get(col, 40))
#                     chars_per_line = max(col_w - 2, 10)

#                     total_lines = 0
#                     for seg in text.split("\n"):
#                         seg = seg.strip()
#                         if seg == "":
#                             total_lines += 1
#                         else:
#                             total_lines += math.ceil(len(seg) / chars_per_line)

#                     max_lines = max(max_lines, total_lines)

#                 ws.set_row(r, min(base_height * max_lines + 4, max_height))

#         # Columnas por tipo (cartera)
#         money_cols_cartera = ["Valor docto", "Total COP (saldo)", "Recaudo"]
#         date_cols_cartera = ["Fecha docto", "Fecha vcto", "Fecha gestión"]
#         wrap_cols_cartera = ["Razón social", "Observaciones"]  # ⇦ ajustar ancho/alto por texto

#         # Columnas por tipo (resumen)
#         money_cols_resumen = ["Total Valor Docto", "Saldo Total", "Recaudo Total"]
#         date_cols_resumen = []
#         wrap_cols_resumen = ["Razón social"]

#         # Aplicar tabla + formatos + anchos
#         widths_cartera, _ = table_with_formats(ws1, df, money_cols_cartera, date_cols_cartera, wrap_cols_cartera)
#         widths_resumen, _ = table_with_formats(ws2, resumen, money_cols_resumen, date_cols_resumen, wrap_cols_resumen)

#         # Ajustar alturas (solo donde hay wrap)
#         autofit_row_heights(ws1, df, wrap_cols_cartera, widths_cartera)
#         autofit_row_heights(ws2, resumen, wrap_cols_resumen, widths_resumen)

#     output.seek(0)
#     filename = f"cartera_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
#     return StreamingResponse(
#         output,
#         media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
#         headers={"Content-Disposition": f'attachment; filename="{filename}"'}
#     )


# # ------------------- Importar cartera -------------------
# @app.post("/importar_excel")
# async def importar_excel(file: UploadFile = File(...), db: Session = Depends(get_db)):
#     try:
#         contents = await file.read()
#         df = pd.read_excel(
#             BytesIO(contents),
#             dtype={
#                 "Nit cliente despacho": str,
#                 "Nro. docto. cruce": str,
#                 "Celular": str,
#                 "Teléfono": str,
#             }
#         )

#         rename_map = {
#             "Razón social": "razon_social",
#             "Nit cliente despacho": "nit_cliente",
#             "Nro. docto. cruce": "nro_docto_cruce",
#             "Fecha docto.": "fecha_docto",
#             "Dias vencidos": "dias_vencidos",
#             "Valor docto": "valor_docto",
#             "Total COP": "total_cop",
#             "Fecha vcto.": "fecha_vcto",
#             "Celular": "celular",
#             "Teléfono": "telefono",
#             "Razón social vend. cliente": "asesor"
#         }
#         df.rename(columns=rename_map, inplace=True)
#         for col in ("telefono", "celular"):
#             if col in df.columns:
#                 df[col] = df[col].apply(clean_phone)

#         if df.empty:
#             return RedirectResponse(url="/?msg=Archivo%20vac%C3%ADo%20o%20sin%20filas%20v%C3%A1lidas&msg_type=info", status_code=303)

#         # Claves en Excel
#         excel_claves = set(
#             (str(row["nit_cliente"]).strip(), str(row["nro_docto_cruce"]).strip())
#             for _, row in df.iterrows()
#             if row.get("nit_cliente") and row.get("nro_docto_cruce")
#         )

#         # Claves en BD
#         bd_clientes = db.query(models.Cliente).all()
#         bd_claves = set((str(c.nit_cliente), str(c.nro_docto_cruce)) for c in bd_clientes)

#         # Eliminar clientes que no están en Excel — MSSQL SAFE (sin abrir begin extra)
#         claves_a_eliminar = bd_claves - excel_claves
#         if claves_a_eliminar:
#             delete_by_pairs_mssql(db, "cartera", claves_a_eliminar, batch_size=900)

#         # Insertar o actualizar
#         for _, row in df.iterrows():
#             nit = str(row.get("nit_cliente")).strip() if row.get("nit_cliente") else None
#             docto = str(row.get("nro_docto_cruce")).strip() if row.get("nro_docto_cruce") else None
#             if not nit or not docto:
#                 continue

#             valor_docto = to_float(row.get("valor_docto")) or 0.0
#             total_excel = to_float(row.get("total_cop")) if row.get("total_cop") is not None else valor_docto

#             # calcular recaudo y normalizar a 2 decimales (no negativo)
#             recaudo_calc = round(max((valor_docto or 0.0) - (total_excel or 0.0), 0.0), 2)

#             cliente = db.query(models.Cliente).filter_by(
#                 nit_cliente=nit,
#                 nro_docto_cruce=docto
#             ).first()

#             if cliente:
#                 # Actualizar existente
#                 cliente.razon_social = row.get("razon_social")
#                 cliente.dias_vencidos = row.get("dias_vencidos")
#                 cliente.fecha_docto = row.get("fecha_docto")
#                 cliente.fecha_vcto = row.get("fecha_vcto")
#                 cliente.valor_docto = valor_docto
#                 cliente.total_cop = total_excel
#                 cliente.recaudo = recaudo_calc
#                 cliente.telefono = clean_phone(row.get("telefono"))
#                 cliente.celular  = clean_phone(row.get("celular"))
#                 cliente.asesor = row.get("asesor")
#             else:
#                 # Insertar nuevo
#                 nuevo = models.Cliente(
#                     razon_social=row.get("razon_social"),
#                     nit_cliente=nit,
#                     nro_docto_cruce=docto,
#                     dias_vencidos=row.get("dias_vencidos"),
#                     fecha_docto=row.get("fecha_docto"),
#                     fecha_vcto=row.get("fecha_vcto"),
#                     valor_docto=valor_docto,
#                     total_cop=total_excel,
#                     recaudo=recaudo_calc,
#                     telefono=clean_phone(row.get("telefono")),
#                     celular=clean_phone(row.get("celular")),
#                     asesor=row.get("asesor"),
#                 )
#                 db.add(nuevo)

#         db.commit()
#         return RedirectResponse(url="/?msg=Archivo%20subido%20correctamente&msg_type=success", status_code=303)

#     except Exception as e:
#         print("❌ Error importar_excel:", e)
#         return RedirectResponse(url="/?msg=Error%20al%20subir%20el%20archivo&msg_type=error", status_code=303)


# # ------------------- Observaciones -------------------
# @app.post("/cliente/{cliente_id}/observacion")
# def agregar_observacion(cliente_id: int, texto: str = Form(...), db: Session = Depends(get_db)):
#     crud.add_observacion(db, cliente_id, texto)
#     return RedirectResponse(
#         url=f"/cliente/{cliente_id}?msg=Observaci%C3%B3n%20guardada&msg_type=success",
#         status_code=303
#     )

# # ------------------- Actualizar cliente (editar recaudo) -------------------
# @app.post("/cliente/{cliente_id}/update")
# async def update_cliente(
#     request: Request,
#     cliente_id: int,
#     db: Session = Depends(get_db)
# ):
#     form_data = await request.form()

#     # Buscar cliente
#     cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
#     if not cliente:
#         return RedirectResponse(url="/?msg=Cliente%20no%20encontrado&msg_type=error", status_code=status.HTTP_303_SEE_OTHER)

#     # ---- Campos texto: actualizar solo si llegan con valor (no sobrescribir con vacío) ----
#     for campo in ["telefono", "celular", "fecha_gestion", "tipo", "asesor"]:
#         val = form_data.get(campo)
#         if val is not None and val != "":
#             setattr(cliente, campo, val)

#     # ---- NUMÉRICOS coherentes: identidad recaudo = valor_docto - total_cop ----
#     valor_docto = float(cliente.valor_docto or 0.0)

#     raw_recaudo = form_data.get("recaudo")
#     raw_total_cop = form_data.get("total_cop")

#     has_recaudo = raw_recaudo is not None and raw_recaudo.strip() != ""
#     has_total_cop = raw_total_cop is not None and raw_total_cop.strip() != ""

#     if has_recaudo:
#         # Interpreta recaudo como ABONO (delta)
#         delta = to_float(raw_recaudo)
#         if delta is not None:
#             old_total = float(cliente.total_cop or 0.0)
#             new_total = max(round(old_total - delta, 2), 0.0)
#             cliente.total_cop = new_total
#             cliente.recaudo = round(max(valor_docto - new_total, 0.0), 2)

#     elif has_total_cop:
#         # Ajuste explícito de saldo total_cop => recalcular recaudo
#         nuevo_total = to_float(raw_total_cop)
#         if nuevo_total is not None:
#             nuevo_total = max(nuevo_total, 0.0)
#             # No dejar que el saldo supere el valor del documento
#             if nuevo_total > valor_docto:
#                 nuevo_total = valor_docto
#             cliente.total_cop = round(nuevo_total, 2)
#             cliente.recaudo = round(max(valor_docto - nuevo_total, 0.0), 2)

#     # ---- Nueva observación ----
#     nueva_obs = form_data.get("observaciones")
#     if nueva_obs and nueva_obs.strip():
#         obs = models.Observacion(
#             cliente_id=cliente.id,
#             texto=nueva_obs.strip(),
#         )
#         db.add(obs)

#     db.commit()
#     return RedirectResponse(
#         url=f"/cliente/{cliente_id}?msg=Cambios%20guardados&msg_type=success",
#         status_code=status.HTTP_303_SEE_OTHER
#     )

# # ------------------- historial cliente -------------------
# @app.get("/cliente/{cliente_id}/historial")
# def historial_cliente(cliente_id: int, db: Session = Depends(get_db)):
#     cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
#     if not cliente:
#         return JSONResponse(content=[], status_code=200)

#     historial = [
#         {
#             "texto": obs.texto,
#             "fecha": obs.fecha_creacion.strftime("%d/%m/%Y %H:%M")
#             if obs.fecha_creacion else None
#         }
#         for obs in cliente.observaciones
#     ]

#     return JSONResponse(content=historial, status_code=200)

# # ------------------- Vista cliente -------------------
# @app.get("/cliente/{cliente_id}")
# def ver_cliente(cliente_id: int, request: Request, db: Session = Depends(get_db)):
#     cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
#     if not cliente:
#         raise HTTPException(status_code=404, detail="Cliente no encontrado")

#     # Cliente anterior
#     prev_cliente = (
#         db.query(models.Cliente)
#         .filter(models.Cliente.id < cliente_id)
#         .order_by(models.Cliente.id.desc())
#         .first()
#     )
#     prev_id = prev_cliente.id if prev_cliente else None

#     # Cliente siguiente
#     next_cliente = (
#         db.query(models.Cliente)
#         .filter(models.Cliente.id > cliente_id)
#         .order_by(models.Cliente.id.asc())
#         .first()
#     )
#     next_id = next_cliente.id if next_cliente else None

#     return templates.TemplateResponse(
#         "cliente.html",
#         {
#             "request": request,
#             "cliente": cliente,
#             "prev_id": prev_id,
#             "next_id": next_id,
#             # flash
#             "flash_msg": request.query_params.get("msg"),
#             "flash_type": request.query_params.get("msg_type"),
#         }
#     )


# # ------------------- Index agrupado -------------------
# @app.get("/")
# def index(
#     request: Request,
#     db: Session = Depends(get_db),
#     view: Optional[str] = Query(None),
#     min_dias: Optional[str] = Query(None),
#     max_dias: Optional[str] = Query(None),
#     sort: Optional[str] = Query("dias_desc"),
# ):
#     # ---- Helpers locales ----
#     def to_int_or_none(val: Optional[str]):
#         try:
#             return int(val) if val not in (None, "") else None
#         except (ValueError, TypeError):
#             return None

#     min_val = to_int_or_none(min_dias)
#     max_val = to_int_or_none(max_dias)

#     # ---- Base query ----
#     query = db.query(models.Cliente)

#     # ---- Filtros de días vencidos (solo si llegan válidos) ----
#     if min_val is not None:
#         query = query.filter(models.Cliente.dias_vencidos >= min_val)
#     if max_val is not None:
#         query = query.filter(models.Cliente.dias_vencidos <= max_val)

#     # ---- Ordenamiento seguro por mapa ----
#     order_map = {
#         "dias_desc": models.Cliente.dias_vencidos.desc(),
#         "dias_asc": models.Cliente.dias_vencidos.asc(),
#     }
#     query = query.order_by(order_map.get(sort, order_map["dias_desc"]))

#     clientes = query.all()

#     # ---- Flash messages desde query y automáticos por filtros ----
#     flash_msg = request.query_params.get("msg")
#     flash_type = request.query_params.get("msg_type")
#     if (min_val is not None) or (max_val is not None):
#         if not flash_msg:
#             flash_msg = "Filtros aplicados"
#         if not flash_type:
#             flash_type = "info"

#     # ====================================
#     # VISTA PLANA
#     # ====================================
#     if view == "flat":
#         filas = []
#         for c in clientes:
#             valor_docto = Decimal(str(c.valor_docto or 0))
#             total_cop = Decimal(str(c.total_cop or 0))
#             recaudo = valor_docto - total_cop

#             filas.append({
#                 "id": c.id,
#                 "razon_social": c.razon_social,
#                 "nit_cliente": c.nit_cliente,
#                 "nro_docto_cruce": c.nro_docto_cruce,
#                 "dias_vencidos": c.dias_vencidos,
#                 "fecha_docto": c.fecha_docto,
#                 "fecha_vcto": c.fecha_vcto,
#                 "valor_docto": float(valor_docto),
#                 "total_cop": float(total_cop),
#                 "recaudo": float(recaudo),
#                 "asesor": c.asesor,
#             })

#         return templates.TemplateResponse(
#             "index.html",
#             {
#                 "request": request,
#                 "view": "flat",
#                 "filas": filas,
#                 # flash
#                 "flash_msg": flash_msg,
#                 "flash_type": flash_type,
#             }
#         )

#     # ====================================
#     # VISTA AGRUPADA
#     # ====================================
#     agrupados = {}
#     for c in clientes:
#         if c.nit_cliente not in agrupados:
#             agrupados[c.nit_cliente] = {
#                 "nit_cliente": c.nit_cliente,
#                 "razon_social": c.razon_social,
#                 "telefono": c.telefono,
#                 "celular": c.celular,
#                 "asesor": c.asesor,
#                 "facturas": []
#             }

#         valor_docto = Decimal(str(c.valor_docto or 0))
#         total_cop = Decimal(str(c.total_cop or 0))
#         recaudo = valor_docto - total_cop

#         agrupados[c.nit_cliente]["facturas"].append({
#             "id": c.id,
#             "nro_docto_cruce": c.nro_docto_cruce,
#             "dias_vencidos": c.dias_vencidos,
#             "fecha_docto": c.fecha_docto,
#             "fecha_vcto": c.fecha_vcto,
#             "valor_docto": float(valor_docto),
#             "total_cop": float(total_cop),
#             "recaudo": float(recaudo),
#             "fecha_gestion": c.fecha_gestion,
#             "tipo": c.tipo,
#             "asesor": c.asesor,
#             "observaciones": [obs.texto for obs in c.observaciones]
#         })

#     # 👉 calcular max_dias por cliente
#     for cl in agrupados.values():
#         dias = [f["dias_vencidos"] for f in cl["facturas"] if f["dias_vencidos"] is not None]
#         cl["max_dias"] = max(dias) if dias else None

#     return templates.TemplateResponse(
#         "index.html",
#         {
#             "request": request,
#             "clientes": list(agrupados.values()),
#             # flash
#             "flash_msg": flash_msg,
#             "flash_type": flash_type,
#         }
#     )







# ------------------- Endpoints Cartas de Cobro -------------------
@app.get('/generar_cartas_masivas', name='generar_cartas_masivas')
def cartas_masivas(db: Session = Depends(get_db)):
    generador_cobros.generar(db_session=db)
    return RedirectResponse(
        url='/?msg=Generaci%C3%B3n%20de%20cartas%20completa.%20Archivos%20abiertos%20o%20enviados.&msg_type=success',
        status_code=303
    )

@app.get('/cliente/{cliente_id}/generar_carta', name='generar_carta')
def carta_cliente(cliente_id: int, request: Request, db: Session = Depends(get_db)):
    cliente = db.query(models.Cliente).filter(models.Cliente.id == cliente_id).first()
    if not cliente:
        return RedirectResponse(url='/?msg=Cliente%20no%20encontrado&msg_type=error', status_code=303)
    generador_cobros.generar(db_session=db, nit_especifico=cliente.nit_cliente)
    return RedirectResponse(
        url=f'/cliente/{cliente_id}?msg=Carta%20generada%20correctamente.&msg_type=success',
        status_code=303
    )

