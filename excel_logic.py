import pandas as pd
from datetime import date, timedelta, datetime
from typing import List, Dict, Tuple, Optional
import os
import io
import openpyxl
from openpyxl.styles import PatternFill


# ---------------------------------------
# Utilidades internas
# ---------------------------------------
def _is_blank_series(s: pd.Series) -> bool:
    """True si toda la serie es NaN o strings vacíos."""
    if s is None:
        return True
    if s.isna().all():
        return True
    s_str = s.astype(str).str.strip().str.lower()
    return (s_str.eq('') | s_str.eq('nan') | s_str.eq('none') | s_str.eq('null')).all()


def _prefix_for_file(plan_staff_file: str) -> str:
    """Prefijo de badge basado en el archivo."""
    base = os.path.basename(plan_staff_file).lower()
    if 'newmont' in base:
        return 'NM'
    return 'ID'


def _normalize_status(v: object) -> Tuple[Optional[str], Optional[str]]:
    """
    Normaliza celdas a ('ON'|'ON NS'|'OFF'|None, 'Day Shift'|'Night Shift'|None).
    - Valores numéricos o 'OK' se consideran ON (día).
    - 'ON' -> ON (día)
    - 'ON NS' o 'NIGHT' -> ON NS (noche)
    - 'OFF', 'Break', 'KO' -> OFF
    - Cualquier otro valor -> None (ignorar)
    """
    if v is None:
        return None, None
    s = str(v).strip().upper()
    if s in ("", "NONE", "NAN", "NULL"):
        return None, None
    if s in ("OFF", "BREAK", "KO", "LEAVE"):
        return "OFF", None
    if s == "ON":
        return "ON", "Day Shift"
    if "ON NS" in s or "NIGHT" in s:
        return "ON NS", "Night Shift"
    # dígitos o 'OK'
    if s.isdigit() or s == "OK":
        return "ON", "Day Shift"
    return None, None


def _is_date_header(col) -> bool:
    """Detecta si el encabezado es una fecha (datetime o pandas.Timestamp)."""
    if isinstance(col, datetime):
        return True
    # pandas Timestamp
    return getattr(col, '__class__', None).__name__ == 'Timestamp'


def _to_pydate(col) -> Optional[date]:
    """Convierte encabezado a date."""
    if isinstance(col, datetime):
        return col.date()
    try:
        return col.to_pydatetime().date()  # pandas.Timestamp
    except Exception:
        return None


# ---------------------------------------
# Lecturas auxiliares / previews
# ---------------------------------------
def get_schedule_preview(plan_staff_file: str) -> pd.DataFrame:
    """
    Carga un DataFrame con columnas: ROLE, NAME, BADGE y columnas fecha.
    Si el archivo no existe o falla, retorna df vacío.
    """
    if not os.path.exists(plan_staff_file):
        return pd.DataFrame()

    try:
        df = pd.read_excel(plan_staff_file, engine='openpyxl')
        # Detectar columnas base
        cols = list(df.columns)
        base_cols = []
        for c in ["ROLE", "NAME", "BADGE", "TEAM", "Discipline", "Company ID"]:
            if c in cols:
                base_cols.append(c)
        # Normalizar a ROLE/NAME/BADGE si vienen en formato Newmont
        if "Discipline" in base_cols and "Company ID" in base_cols:
            df["ROLE"] = df["Discipline"]
            base_cols.append("ROLE")
        # Fechas
        date_cols = [c for c in cols if _is_date_header(c)]
        keep = []
        # Mantener ROLE/NAME/BADGE si existen
        for key in ["ROLE", "NAME", "BADGE"]:
            if key in df.columns:
                keep.append(key)
        # si no había BADGE pero hay Company ID/Discipline, también
        if "Company ID" in df.columns and "BADGE" not in keep:
            df["BADGE"] = df["Company ID"].astype(str)
            keep.append("BADGE")
        # armar df final
        keep += date_cols
        if not keep:
            return pd.DataFrame()
        return df[keep]
    except Exception:
        return pd.DataFrame()


def get_roles_from_excel(plan_staff_file: str) -> list:
    """Lista de roles únicos (si está disponible)."""
    if not os.path.exists(plan_staff_file):
        return [f"Role not available ({os.path.basename(plan_staff_file)} not found)"]
    try:
        wb = openpyxl.load_workbook(plan_staff_file, read_only=True, data_only=True)
        ws = wb.active
        header_map = {cell.value: cell.column for cell in ws[1]}
        role_header = "ROLE" if "ROLE" in header_map else ("Discipline" if "Discipline" in header_map else None)
        if not role_header:
            return ["ROLE/Discipline column not found"]
        col_idx = header_map[role_header]
        roles = {
            ws.cell(row=i, column=col_idx).value
            for i in range(2, ws.max_row + 1)
            if ws.cell(row=i, column=col_idx).value
        }
        return sorted(list(roles))
    except Exception:
        return ["Error reading Excel"]


def get_users_from_excel(plan_staff_file: str) -> list:
    """
    Extrae usuarios (name, role, badge) de la planilla.
    Soporta:
      - RGM: NAME, ROLE, BADGE
      - Newmont: Last Name, First Name, Discipline, Company ID
    Si no hay badge, genera uno estable (prefijo NM o ID + secuencia).
    """
    if not os.path.exists(plan_staff_file):
        return []
    try:
        df = pd.read_excel(plan_staff_file, engine='openpyxl')

        rgm_cols = ['NAME', 'ROLE', 'BADGE']
        newmont_cols = ['Last Name', 'First Name', 'Discipline', 'Company ID']

        users_df = None

        if all(col in df.columns for col in rgm_cols):
            users_df = df[rgm_cols].copy()
            # Badges faltantes
            if _is_blank_series(users_df['BADGE']):
                prefix = _prefix_for_file(plan_staff_file)
                users_df['BADGE'] = [f"{prefix}{i+1:05d}" for i in range(len(users_df))]
            else:
                prefix = _prefix_for_file(plan_staff_file)
                badge_series = users_df['BADGE'].astype(str)
                is_missing = users_df['BADGE'].isna() | badge_series.str.strip().eq('') | badge_series.str.lower().isin(['nan', 'none', 'null'])
                seq = (f"{prefix}{i+1:05d}" for i in range(is_missing.sum()))
                users_df.loc[is_missing, 'BADGE'] = [next(seq) for _ in range(is_missing.sum())]

        elif all(col in df.columns for col in newmont_cols):
            df_copy = df[newmont_cols].copy()
            df_copy['NAME'] = df_copy['Last Name'].astype(str).strip() + ', ' + df_copy['First Name'].astype(str).strip()
            df_copy.rename(columns={'Discipline': 'ROLE', 'Company ID': 'BADGE'}, inplace=True)
            users_df = df_copy[['NAME', 'ROLE', 'BADGE']]
            if _is_blank_series(users_df['BADGE']):
                prefix = _prefix_for_file(plan_staff_file)
                users_df['BADGE'] = [f"{prefix}{i+1:05d}" for i in range(len(users_df))]
        else:
            # Fallback si vienen NAME/ROLE solamente
            if all(col in df.columns for col in ['NAME', 'ROLE']):
                prefix = _prefix_for_file(plan_staff_file)
                users_df = df[['NAME', 'ROLE']].copy()
                users_df['BADGE'] = [f"{prefix}{i+1:05d}" for i in range(len(users_df))]
            else:
                return []

        users_df['NAME'] = users_df['NAME'].astype(str).str.strip()
        users_df['ROLE'] = users_df['ROLE'].astype(str).str.strip()
        users_df['BADGE'] = users_df['BADGE'].astype(str).str.strip()

        users_df = users_df[(users_df['NAME'] != '') & (users_df['BADGE'] != '')]
        users_df.drop_duplicates(subset=['BADGE'], keep='first', inplace=True)
        users_df.rename(columns={'NAME': 'name', 'ROLE': 'role', 'BADGE': 'badge'}, inplace=True)
        return users_df.to_dict('records')
    except Exception:
        return []


# ---------------------------------------
# Escritura / actualización de plan staff (FR-01 y soporte general)
# ---------------------------------------
def update_plan_staff_excel(plan_staff_file: str, username: str, role: str, badge: str,
                            schedule_status: Optional[str], shift_type: Optional[str],
                            schedule_start: date, schedule_end: date) -> Tuple[bool, str]:
    """
    Actualiza (o crea si no existe) la fila del empleado en el Excel:
    - Busca por BADGE y, si no, por NAME.
    - Escribe ON/ON NS/OFF con colores.
    - Si schedule_status es None, limpia el rango.
    """
    try:
        # Abrir o crear
        if os.path.exists(plan_staff_file):
            wb = openpyxl.load_workbook(plan_staff_file)
            ws = wb.active
        else:
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Operations_best_opt"
            headers = ["TEAM", "ROLE", "NAME", "BADGE"]
            for col_idx, h in enumerate(headers, start=1):
                ws.cell(row=1, column=col_idx, value=h)

        # Colores
        green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # ON día
        red   = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # OFF
        yel   = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")  # ON NS noche

        # Mapas de cabecera
        header_map = {cell.value: cell.column for cell in ws[1] if isinstance(cell.value, str)}
        date_map: Dict[date, int] = {}
        for cell in ws[1]:
            v = cell.value
            if isinstance(v, datetime):
                date_map[v.date()] = cell.column

        # localizar fila por BADGE o NAME
        row_idx = None
        if "BADGE" in header_map:
            for i in range(2, ws.max_row + 1):
                v = ws.cell(row=i, column=header_map["BADGE"]).value
                if v and str(v) == str(badge):
                    row_idx = i
                    break
        if not row_idx and "NAME" in header_map:
            for i in range(2, ws.max_row + 1):
                v = ws.cell(row=i, column=header_map["NAME"]).value
                if v and str(v).strip().lower() == str(username).strip().lower():
                    row_idx = i
                    break
        if not row_idx:
            row_idx = ws.max_row + 1

        # Datos fijos
        if "NAME" in header_map:
            ws.cell(row=row_idx, column=header_map["NAME"], value=username)
        if "ROLE" in header_map:
            ws.cell(row=row_idx, column=header_map["ROLE"], value=role)
        if "BADGE" in header_map:
            ws.cell(row=row_idx, column=header_map["BADGE"], value=badge)

        # texto/estilo por estado
        text = None
        fill = None
        if schedule_status == "ON":
            text = "ON"
            fill = green
            if shift_type == "Night Shift":
                text = "ON NS"
                fill = yel
        elif schedule_status == "OFF":
            text = "OFF"
            fill = red
        elif schedule_status is None:
            text = None
            fill = None

        # Escribir/limpiar rango
        d = schedule_start
        while d <= schedule_end:
            # Crear columna de fecha si no existe en el template
            if d not in date_map:
                new_col = ws.max_column + 1
                ws.cell(row=1, column=new_col, value=datetime(d.year, d.month, d.day))
                date_map[d] = new_col
            col_idx = date_map[d]
            cell = ws.cell(row=row_idx, column=col_idx)
            if text is None:
                cell.value = None
                cell.fill = PatternFill(fill_type=None)
            else:
                cell.value = text
                cell.fill = fill
            d += timedelta(days=1)

        wb.save(plan_staff_file)
        return True, f"Plan staff updated for {username}."
    except Exception as e:
        return False, f"Error updating plan staff: {e}"


# ---------------------------------------
# RF-01: Detección de conflictos (sobrescritura)
# ---------------------------------------
def find_conflicts(plan_staff_file: str, username: str, badge: str,
                   schedule_start: date, schedule_end: date) -> List[Dict]:
    """
    Devuelve [{'date': date, 'existing': 'ON/ON NS/OFF/...'}] si hay valores ya escritos en el rango.
    Busca fila por BADGE y luego por NAME, igual que update_plan_staff_excel.
    """
    try:
        if not os.path.exists(plan_staff_file):
            return []
        wb = openpyxl.load_workbook(plan_staff_file, data_only=True)
        ws = wb.active

        header_map = {cell.value: cell.column for cell in ws[1] if isinstance(cell.value, str)}
        date_map = {cell.value.date(): cell.column for cell in ws[1] if isinstance(cell.value, datetime)}

        # localizar fila por BADGE y luego por NAME
        row_idx = None
        if "BADGE" in header_map:
            for i in range(2, ws.max_row + 1):
                v = ws.cell(row=i, column=header_map["BADGE"]).value
                if v and str(v) == str(badge):
                    row_idx = i
                    break
        if not row_idx and "NAME" in header_map:
            for i in range(2, ws.max_row + 1):
                v = ws.cell(row=i, column=header_map["NAME"]).value
                if v and str(v).strip().lower() == str(username).strip().lower():
                    row_idx = i
                    break
        if not row_idx:
            return []

        conflicts: List[Dict] = []
        d = schedule_start
        while d <= schedule_end:
            if d in date_map:
                col = date_map[d]
                val = ws.cell(row=row_idx, column=col).value
                if val not in (None, '', ' '):
                    conflicts.append({"date": d, "existing": str(val)})
            d += timedelta(days=1)
        return conflicts
    except Exception:
        return []


# ---------------------------------------
# RF-02: Importar Excel -> DB (usuarios + schedules)
# ---------------------------------------
def import_excel_to_db(plan_staff_file: str, source: str) -> Tuple[int, int, int]:
    """
    Procesa .xlsx y almacena en la BD:
      - Usuarios (name, role, badge)
      - Schedules día-a-día (ON/ON NS/OFF)
    Devuelve: (nuevos_usuarios, usuarios_omitidos, upserts_schedule)
    """
    from database_logic import add_users_bulk, get_all_users, upsert_schedule_day  # import diferido

    users_in_file = get_users_from_excel(plan_staff_file)
    if not users_in_file:
        return (0, 0, 0)

    # Insertar usuarios (evitando duplicados por badge)
    before = get_all_users(source)
    before_badges = {u['badge'] for u in before}
    inserted = add_users_bulk(users_in_file, source)
    after = get_all_users(source)
    after_badges = {u['badge'] for u in after}
    skipped = len(before_badges & after_badges)  # aproximado para el mensaje

    # Importar schedules
    upserts = 0
    try:
        df = pd.read_excel(plan_staff_file, engine='openpyxl')
        # detectar identificadores
        name_field = 'NAME' if 'NAME' in df.columns else None
        role_field = 'ROLE' if 'ROLE' in df.columns else ('Discipline' if 'Discipline' in df.columns else None)
        badge_field = 'BADGE' if 'BADGE' in df.columns else ('Company ID' if 'Company ID' in df.columns else None)

        if not badge_field:
            return inserted, skipped, 0

        date_cols = [c for c in df.columns if _is_date_header(c)]
        for _, row in df.iterrows():
            name = str(row[name_field]).strip() if name_field else ""
            role = str(row[role_field]).strip() if role_field else ""
            badge = str(row[badge_field]).strip()
            if not badge:
                continue
            for dcol in date_cols:
                d_py = _to_pydate(dcol)
                if not d_py:
                    continue
                status, shift = _normalize_status(row[dcol])
                if status:
                    upsert_schedule_day(badge, d_py, status, shift, source)
                    upserts += 1
    except Exception:
        pass

    return (inserted, skipped, upserts)


# ---------------------------------------
# RF-03: Exportar Excel desde la BD (manteniendo formato base)
# ---------------------------------------
def export_plan_from_db(template_path: str, users: List[Dict], schedules: List[Dict], output_path: str) -> Tuple[bool, str]:
    """
    users: [{'name','role','badge'}]
    schedules: [{'badge','date':'YYYY-MM-DD','status':'ON|ON NS|OFF'}]
    """
    try:
        if not os.path.exists(template_path):
            return False, f"Template '{os.path.basename(template_path)}' not found."

        wb = openpyxl.load_workbook(template_path)
        ws = wb.active

        header_map = {cell.value: cell.column for cell in ws[1] if isinstance(cell.value, str)}
        for h in ["NAME", "ROLE", "BADGE"]:
            if h not in header_map:
                return False, f"Required header '{h}' not found in template."

        # Mapas de fechas existentes
        date_map: Dict[date, int] = {}
        for c in ws[1]:
            if isinstance(c.value, datetime):
                date_map[c.value.date()] = c.column

        # Colores
        green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        red   = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        yel   = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")

        # Escribir usuarios en orden
        start_row = 2
        for idx, u in enumerate(users):
            r = start_row + idx
            ws.cell(row=r, column=header_map["NAME"], value=u["name"])
            ws.cell(row=r, column=header_map["ROLE"], value=u["role"])
            ws.cell(row=r, column=header_map["BADGE"], value=u["badge"])

        # Poner schedules
        # Asegurar columnas para todas las fechas presentes
        used_dates = sorted({datetime.strptime(s['date'], "%Y-%m-%d").date() for s in schedules})
        for d in used_dates:
            if d not in date_map:
                new_col = ws.max_column + 1
                ws.cell(row=1, column=new_col, value=datetime(d.year, d.month, d.day))
                date_map[d] = new_col

        # Mapa rápido badge->row
        badge_row = {}
        for idx, u in enumerate(users):
            r = start_row + idx
            badge_row[u['badge']] = r

        for s in schedules:
            b = s['badge']
            r = badge_row.get(b)
            if not r:
                continue
            d = datetime.strptime(s['date'], "%Y-%m-%d").date()
            col = date_map.get(d)
            if not col:
                continue
            cell = ws.cell(row=r, column=col)
            st = s.get('status')
            if st == "OFF":
                cell.value = "OFF"; cell.fill = red
            elif st == "ON NS":
                cell.value = "ON NS"; cell.fill = yel
            elif st == "ON":
                cell.value = "ON"; cell.fill = green

        wb.save(output_path)
        return True, f"Plan successfully exported to '{os.path.basename(output_path)}'."
    except Exception as e:
        return False, f"Export error: {e}"


# ---------------------------------------
# Generar reporte de transporte (opcional, para botón "Generate Report")
# ---------------------------------------
def generate_transport_report(plan_staff_file: str, start_date: date, end_date: date) -> Tuple[bytes, str]:
    """
    Genera un archivo Excel simple (bytes) con dos bloques "IN/OUT" en la hoja 'travel list'.
    Este método es robusto y no depende del formato exacto del template.
    """
    # Leer usuarios
    users = get_users_from_excel(plan_staff_file)
    # Armar DataFrame básico
    df_in = pd.DataFrame([{"DEPT": u['role'], "NAME": u['name'], "DATE": start_date.strftime("%Y-%m-%d")} for u in users])
    df_out = pd.DataFrame([{"DEPT": u['role'], "NAME": u['name'], "DATE": end_date.strftime("%Y-%m-%d")} for u in users])

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        # hoja base
        df_in.to_excel(writer, sheet_name='travel list', startrow=6, index=False)
        ws = writer.book["travel list"]
        ws.cell(row=2, column=1, value="MERIAN TRANSPORTATION REQUEST")
        ws.cell(row=5, column=1, value="IN")
        ws.cell(row=5, column=5, value="OUT")
        # más etiquetas
        ws.cell(row=6, column=1, value="DEPT")
        ws.cell(row=6, column=2, value="NAME")
        ws.cell(row=6, column=3, value="DATE")
    return output.getvalue(), f"Report successfully generated from {os.path.basename(plan_staff_file)}."
