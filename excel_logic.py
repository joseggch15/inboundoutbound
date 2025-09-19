# excel_logic.py
# ============================================================
# Utilidades de Excel para PlanStaff con SSoT (SQLite),
# validación de estructura, importación validada,
# exportación/regeneración desde BD, comparación Excel↔BD,
# y utilidades de reporte.
#
# Esta versión corrige el error de Pylance:
#   "variant no está definido"
# garantizando que la variable 'variant' se inicializa y se
# propaga correctamente en validate_excel_structure y en
# cualquier flujo que la use (meta['variant']).
# ============================================================

from __future__ import annotations

import os
import io
from datetime import date, timedelta, datetime
from typing import List, Dict, Tuple, Optional, Set

import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill
from openpyxl.comments import Comment

# ============================================================
# Helpers / Normalización
# ============================================================

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
    - 'OFF', 'Break', 'KO', 'Leave' -> OFF
    - Cualquier otro valor -> None (ignorar)
    Nota: Tipos personalizados (códigos) no se normalizan aquí (se tratan fuera).
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
    return getattr(col, '__class__', None).__name__ == 'Timestamp'


def _to_pydate(col) -> Optional[date]:
    """Convierte encabezado a date."""
    if isinstance(col, datetime):
        return col.date()
    try:
        return col.to_pydatetime().date()  # pandas.Timestamp
    except Exception:
        return None


def _fill_for_base_status(status: Optional[str]) -> Optional[PatternFill]:
    """Devuelve PatternFill para estados base ('ON', 'OFF', 'ON NS')."""
    if status is None:
        return None
    s = str(status).strip().upper()
    green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # ON día
    red   = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # OFF
    yel   = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")  # ON NS
    if s == "ON":
        return green
    if s == "OFF":
        return red
    if s == "ON NS":
        return yel
    return None


# ============================================================
# Lecturas auxiliares / previews
# ============================================================

def get_schedule_preview(plan_staff_file: str) -> pd.DataFrame:
    """
    Carga un DataFrame con columnas: ROLE, NAME, BADGE y columnas fecha.
    Aplica FR-07: limpia NaN/None en celdas de fechas para no mostrar 'nan'.
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
            if "BADGE" not in df.columns and "Company ID" in df.columns:
                df["BADGE"] = df["Company ID"]
        # Fechas
        date_cols = [c for c in cols if _is_date_header(c)]
        keep: List[str] = []
        # Mantener ROLE/NAME/BADGE si existen
        for key in ["ROLE", "NAME", "BADGE"]:
            if key in df.columns:
                keep.append(key)
        # si no había BADGE pero hay Company ID, también
        if "Company ID" in df.columns and "BADGE" not in keep:
            df["BADGE"] = df["Company ID"].astype(str)
            keep.append("BADGE")
        # armar df final
        keep += date_cols
        if not keep:
            return pd.DataFrame()

        df_out = df[keep].copy()

        # FR-07: limpiar NaN/None -> '' SOLO para visualización en la UI (no altera importación)
        for c in date_cols:
            if c in df_out.columns:
                df_out[c] = df_out[c].apply(
                    lambda v: "" if (v is None or str(v).strip().lower() in ("nan", "none", "null")) else v
                )

        return df_out
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
            df_copy['NAME'] = df_copy['Last Name'].astype(str).str.strip() + ', ' + df_copy['First Name'].astype(str).str.strip()
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


# ============================================================
# Escritura / actualización del plan staff (Excel)
# ============================================================

def update_plan_staff_excel(
    plan_staff_file: str,
    username: str,
    role: str,
    badge: str,
    schedule_status: Optional[str],
    shift_type: Optional[str],
    schedule_start: date,
    schedule_end: date,
    source: str,
    in_time: Optional[str] = None,
    out_time: Optional[str] = None
) -> Tuple[bool, str]:
    """
    Actualiza (o crea si no existe) la fila del empleado en el Excel:
    - Busca por BADGE y, si no, por NAME.
    - Escribe:
        * Estados base -> 'ON' / 'ON NS' / 'OFF' con colores legacy.
        * Tipos personalizados -> código (p.ej. 'SOP') y color del tipo.
      Además añade un comentario con 'IN-OUT' (HH:MM-HH:MM) si viene in_time/out_time.
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

        # Colores base
        green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # ON día
        red   = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # OFF
        yel   = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")  # ON NS noche

        # Mapa de tipos personalizados (para colores)
        try:
            from database_logic import get_shift_type_map  # import diferido para evitar ciclos
            custom_map = get_shift_type_map(source)
        except Exception:
            custom_map = {}

        def _fill_for_status(status: Optional[str]) -> Optional[PatternFill]:
            if status is None:
                return None
            s = str(status).strip().upper()
            if s == "ON":
                return green
            if s == "ON NS":
                return yel
            if s == "OFF":
                return red
            # código personalizado
            info = custom_map.get(s)
            if info and info.get('color_hex'):
                hex6 = info['color_hex'].lstrip('#').upper()
                return PatternFill(start_color=hex6, end_color=hex6, fill_type="solid")
            return None

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
        if schedule_status in ("ON", "OFF", "ON NS"):
            # base
            text = "ON" if schedule_status == "ON" else ("ON NS" if schedule_status == "ON NS" else "OFF")
            fill = _fill_for_status(text)
        elif schedule_status is None:
            text = None
            fill = None
        else:
            # personalizado -> escribir código
            text = str(schedule_status).strip().upper()
            fill = _fill_for_status(text)

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
                cell.comment = None
                cell.fill = PatternFill(fill_type=None)
            else:
                cell.value = text
                cell.fill = fill if fill else PatternFill(fill_type=None)
                # Comentario con horarios para personalizados
                if text not in ("ON", "ON NS", "OFF") and in_time and out_time:
                    cell.comment = Comment(f"{in_time}-{out_time}", "ShiftType")
                else:
                    cell.comment = None
            d += timedelta(days=1)

        wb.save(plan_staff_file)
        return True, f"Plan staff updated for {username}."
    except Exception as e:
        return False, f"Error updating plan staff: {e}"


# ============================================================
# FR-01: Detección de conflictos (sobrescritura)
# ============================================================

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


# ============================================================
# FR-02: Importar Excel -> DB (usuarios + schedules) con validación
# ============================================================

def import_excel_to_db(plan_staff_file: str, source: str) -> Tuple[int, int, int]:
    """
    Procesa .xlsx y almacena en la BD:
      - Usuarios (name, role, badge)
      - Schedules día-a-día (ON/ON NS/OFF)
    Devuelve: (nuevos_usuarios, usuarios_omitidos, upserts_schedule)

    Si el archivo NO es válido (estructura), levanta ValueError con detalle.
    """
    # Validación previa estricta
    ok, errors, _meta = validate_excel_structure(plan_staff_file)
    if not ok:
        raise ValueError("Invalid Plan Staff structure:\n" + "\n".join(f"- {e}" for e in errors))

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

    # Importar schedules (solo estados base reconocidos)
    upserts = 0
    try:
        df = pd.read_excel(plan_staff_file, engine='openpyxl')
        # detectar identificadores
        role_field = 'ROLE' if 'ROLE' in df.columns else ('Discipline' if 'Discipline' in df.columns else None)
        badge_field = 'BADGE' if 'BADGE' in df.columns else ('Company ID' if 'Company ID' in df.columns else None)

        if not badge_field:
            return inserted, skipped, 0

        date_cols = [c for c in df.columns if _is_date_header(c)]
        for _, row in df.iterrows():
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


# ============================================================
# FR-03: Exportar Excel desde la BD (manteniendo plantilla)
# ============================================================

def export_plan_from_db(
    template_path: str,
    users: List[Dict],
    schedules: List[Dict],
    output_path: str,
    source: str
) -> Tuple[bool, str]:
    """
    users: [{'name','role','badge'}]
    schedules: [{'badge','date':'YYYY-MM-DD','status','shift_type','in_time','out_time'}]
    Soporta plantillas RGM (NAME/ROLE/BADGE) y Newmont (Last/First/Discipline/Company ID).
    """
    try:
        if not os.path.exists(template_path):
            return False, f"Template '{os.path.basename(template_path)}' not found."

        try:
            from database_logic import get_shift_type_map  # import diferido
            custom_map = get_shift_type_map(source)
        except Exception:
            custom_map = {}

        wb = openpyxl.load_workbook(template_path)
        ws = wb.active

        header_map = {cell.value: cell.column for cell in ws[1] if isinstance(cell.value, str)}

        # Detectar variante de plantilla
        variant = None
        if all(h in header_map for h in ("NAME", "ROLE", "BADGE")):
            variant = "RGM"
        elif all(h in header_map for h in ("Last Name", "First Name", "Discipline", "Company ID")):
            variant = "Newmont"
        else:
            return False, "Unsupported template: expected RGM (NAME/ROLE/BADGE) or Newmont (Last/First/Discipline/Company ID)."

        # Mapas de fechas existentes
        date_map: Dict[date, int] = {}
        for c in ws[1]:
            if isinstance(c.value, datetime):
                date_map[c.value.date()] = c.column

        # Colores base
        green = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        red   = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        yel   = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")

        def _fill_for(status: Optional[str]) -> Optional[PatternFill]:
            if status is None:
                return None
            s = str(status).strip().upper()
            if s == "OFF":
                return red
            if s == "ON NS":
                return yel
            if s == "ON":
                return green
            info = custom_map.get(s)
            if info and info.get('color_hex'):
                hex6 = info['color_hex'].lstrip('#').upper()
                return PatternFill(start_color=hex6, end_color=hex6, fill_type="solid")
            return None

        # Schedules -> mapa por badge y fecha
        sched_by_badge: Dict[str, Dict[str, Dict[str, Optional[str]]]] = {}
        for s in schedules:
            b = str(s.get('badge', '')).strip()
            d = str(s.get('date', '')).strip()  # 'YYYY-MM-DD'
            if not b or not d:
                continue
            sched_by_badge.setdefault(b, {})[d] = {
                'status': s.get('status'),
                'shift_type': s.get('shift_type'),
                'in_time': s.get('in_time'),
                'out_time': s.get('out_time')
            }

        # Helper para nombre en Newmont
        def split_name(full: str) -> Tuple[str, str]:
            if not full:
                return "", ""
            s = str(full).strip()
            if "," in s:
                last, first = s.split(",", 1)
                return last.strip(), first.strip()
            parts = s.split()
            if len(parts) >= 2:
                return parts[-1], " ".join(parts[:-1])
            return s, ""

        # Escribir/actualizar usuarios
        # Índice rápido por badge ya existente
        badge_col = header_map["BADGE"] if variant == "RGM" else header_map["Company ID"]
        existing_rows_by_badge: Dict[str, int] = {}
        for i in range(2, ws.max_row + 1):
            v = ws.cell(row=i, column=badge_col).value
            if v:
                existing_rows_by_badge[str(v).strip()] = i

        for u in users:
            name = u.get('name', '').strip()
            role = u.get('role', '').strip()
            badge = str(u.get('badge', '')).strip()
            if not badge:
                continue

            row_idx = existing_rows_by_badge.get(badge)
            if not row_idx:
                row_idx = ws.max_row + 1
                existing_rows_by_badge[badge] = row_idx

            if variant == "RGM":
                ws.cell(row=row_idx, column=header_map["NAME"], value=name)
                ws.cell(row=row_idx, column=header_map["ROLE"], value=role)
                ws.cell(row=row_idx, column=header_map["BADGE"], value=badge)
            else:
                last, first = split_name(name)
                ws.cell(row=row_idx, column=header_map["Last Name"], value=last)
                ws.cell(row=row_idx, column=header_map["First Name"], value=first)
                ws.cell(row=row_idx, column=header_map["Discipline"], value=role)
                ws.cell(row=row_idx, column=header_map["Company ID"], value=badge)

            # Rellenar días
            # Fechas ordenadas por las que ya existen en plantilla
            dates_sorted = sorted(date_map.keys())
            for d in dates_sorted:
                col_idx = date_map[d]
                cell = ws.cell(row=row_idx, column=col_idx)
                info = sched_by_badge.get(badge, {}).get(d.isoformat(), None)
                if info:
                    st = (info.get('status') or '').strip().upper() if info.get('status') else None
                    cell.value = st
                    cell.fill = _fill_for(st) or PatternFill(fill_type=None)
                    # comentario para personalizados si tenemos HH:MM
                    if st and st not in ("ON", "ON NS", "OFF"):
                        it = (info.get('in_time') or '').strip()
                        ot = (info.get('out_time') or '').strip()
                        if it and ot:
                            cell.comment = Comment(f"{it}-{ot}", "ShiftType")
                        else:
                            cell.comment = None
                else:
                    cell.value = None
                    cell.fill = PatternFill(fill_type=None)
                    cell.comment = None

        # Expandir columnas si hay fechas en BD que no existían en plantilla
        # (opcional; se puede agregar al final)
        all_sched_dates: Set[date] = set()
        for b, per_day in sched_by_badge.items():
            for d_str in per_day.keys():
                try:
                    y, m, dy = d_str.split("-")
                    all_sched_dates.add(date(int(y), int(m), int(dy)))
                except Exception:
                    pass
        missing_dates = sorted([d for d in all_sched_dates if d not in date_map])
        for d in missing_dates:
            new_col = ws.max_column + 1
            ws.cell(row=1, column=new_col, value=datetime(d.year, d.month, d.day))
            date_map[d] = new_col
            # escribir valores para cada usuario
            for u in users:
                badge = str(u.get('badge', '')).strip()
                if not badge:
                    continue
                row_idx = existing_rows_by_badge.get(badge)
                if not row_idx:
                    continue
                info = sched_by_badge.get(badge, {}).get(d.isoformat(), None)
                cell = ws.cell(row=row_idx, column=new_col)
                if info:
                    st = (info.get('status') or '').strip().upper() if info.get('status') else None
                    cell.value = st
                    cell.fill = _fill_for(st) or PatternFill(fill_type=None)
                    if st and st not in ("ON", "ON NS", "OFF"):
                        it = info.get('in_time')
                        ot = info.get('out_time')
                        if it and ot:
                            cell.comment = Comment(f"{it}-{ot}", "ShiftType")
                        else:
                            cell.comment = None
                else:
                    cell.value = None
                    cell.fill = PatternFill(fill_type=None)
                    cell.comment = None

        wb.save(output_path)
        return True, f"Plan successfully exported to '{os.path.basename(output_path)}'."
    except Exception as e:
        return False, f"Export error: {e}"


# ============================================================
# Reporte de transporte (IN/OUT) — sin cambios en lógica base
# ============================================================

def generate_transport_report(plan_staff_file: str, start_date: date, end_date: date) -> Tuple[bytes, str]:
    """
    Genera el Excel de transporte considerando:
      - Estados base: 'ON' (IN=06:00:00, OUT=12:00:00) y 'ON NS' (IN=12:00:00, OUT=06:00:00).
      - Códigos personalizados: usa las horas configuradas en shift_types (BD). Si no hay, toma
        el comentario de la celda 'HH:MM-HH:MM' como fallback.

    Detección de eventos dentro del rango [start_date, end_date]:
      - ENTRADA (IN / TRAVEL TO SITE): día d 'trabaja' y status(d-1) != status(d)
      - SALIDA  (OUT / TRAVEL FROM SITE): día d 'trabaja' y status(d+1) != status(d)
      (Se consideran cambios ON↔ON NS y también cambios entre códigos personalizados.)

    ✅ RF-TR-01 / RF-TR-02 (ajuste):
      Además de los eventos dentro del rango, **siempre** se incluyen los primeros
      eventos **posteriores** a end_date:
        • la primera ENTRADA (> end_date)
        • la primera SALIDA  (> end_date)
      Esto se hace aunque ya existan entradas/salidas dentro del rango para el empleado,
      y se evita la duplicidad por fecha.

    """
    # ---- 1) Inferir source por nombre de archivo ----
    fname = os.path.basename(plan_staff_file).lower()
    source = "Newmont" if "newmont" in fname else "RGM"

    # ---- 2) Cargar mapping de tipos personalizados desde BD ----
    try:
        from database_logic import get_shift_type_map  # import diferido para evitar ciclos
        custom_map: Dict[str, Dict] = get_shift_type_map(source)  # code -> {in_time, out_time, .}
        custom_map = {k.strip().upper(): v for k, v in custom_map.items()}
    except Exception:
        custom_map = {}

    ## NUEVO CAMBIO ## - Importar helper de ubicación
    try:
        from database_logic import get_user_location_for_date
    except Exception:
        def get_user_location_for_date(b, d):
            return (None, None)
    ## FIN NUEVO CAMBIO ##

    # ---- 3) Abrir plan staff ----
    try:
        wb_src = openpyxl.load_workbook(plan_staff_file, data_only=True)
        ws_src = wb_src.active
    except Exception:
        # Retornar archivo vacío con la hoja 'travel list'
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "travel list"
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out.read(), "Unsupported Plan Staff format."

    # ---- 4) Resolver esquema de columnas (RGM vs Newmont) ----
    header_map: Dict[str, int] = {}
    date_cols: Dict[int, date] = {}  # col_index -> date
    for c in ws_src[1]:
        v = c.value
        if isinstance(v, str):
            header_map[v] = c.column
        elif isinstance(v, datetime):
            date_cols[c.column] = v.date()

    is_rgm = all(h in header_map for h in ("NAME", "ROLE", "BADGE"))
    is_new = all(h in header_map for h in ("Last Name", "First Name", "Discipline", "Company ID"))
    if not (is_rgm or is_new):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "travel list"
        out = io.BytesIO()
        wb.save(out)
        out.seek(0)
        return out.read(), "Unsupported Plan Staff format."

    # Orden de fechas
    dates_sorted: List[Tuple[int, date]] = sorted(date_cols.items(), key=lambda x: x[1])

    # ---- 5) Helpers ----
    OFF_LIKE = {"OFF", "BREAK", "KO", "LEAVE"}

    def _norm_status(cell_val: Optional[str]) -> Optional[str]:
        if cell_val is None:
            return None
        s = str(cell_val).strip()
        if not s:
            return None
        su = s.upper()
        if su in OFF_LIKE:
            return "OFF"
        if su in ("ON", "ON NS"):
            return su
        if su.isdigit() or su == "OK" or "DAY" in su:
            return "ON"
        return su  # personalizado

    def _is_working(status: Optional[str]) -> bool:
        if not status:
            return False
        if status in ("ON", "ON NS"):
            return True
        if status == "OFF":
            return False
        return True  # personalizado -> cuenta como trabajando

    def _hhmm_to_hhmmss(t: str) -> str:
        t = (t or "").strip()
        if len(t) == 5:
            return t + ":00"
        return t

    def _times_for(status: str, io_kind: str, cell_comment: Optional[str]) -> str:
        su = (status or "").strip().upper()
        if su == "ON":
            return "06:00:00" if io_kind == "IN" else "12:00:00"
        if su == "ON NS":
            return "12:00:00" if io_kind == "IN" else "06:00:00"

        info = custom_map.get(su)
        if info:
            t = info.get("in_time") if io_kind == "IN" else info.get("out_time")
            if t:
                return _hhmm_to_hhmmss(t)

        if cell_comment:
            txt = str(cell_comment).strip()
            if "-" in txt:
                left, right = txt.split("-", 1)
                t = left if io_kind == "IN" else right
                return _hhmm_to_hhmmss(t.strip())
        return "06:00:00" if io_kind == "IN" else "12:00:00"

    # ---- 6) Workbook de salida ----
    company_default = "PLGims"

    wb_out = openpyxl.Workbook()
    ws = wb_out.active
    ws.title = "travel list"

    ws.cell(row=2, column=1, value="MERIAN TRANSPORTATION REQUEST")
    ws.cell(row=5, column=1, value="IN")
    ws.cell(row=5, column=2, value="TRAVEL TO SITE")
    ws.cell(row=5, column=11, value="OUT")
    ws.cell(row=5, column=12, value="TRAVEL FROM SITE")

    headers_in = ["#", "NAME", "FIRST NAME", "GID", "COMPANY", "DEPT", "FROM", "DATE", "TIME"]
    for i, h in enumerate(headers_in, start=1):
        ws.cell(row=6, column=i, value=h)
    headers_out = ["#", "NAME", "FIRST NAME", "GID", "COMPANY", "DEPT", "TO", "DATE", "TIME"]
    for i, h in enumerate(headers_out, start=11):
        ws.cell(row=6, column=i, value=h)

    # ---- 7) Recorrer filas (personas) y detectar IN/OUT ----
    r_in = 7
    r_out = 7
    idx_in = 1
    idx_out = 1

    # Campos de identificación
    if is_rgm:
        name_col = header_map["NAME"]
        role_col = header_map["ROLE"]
        badge_col = header_map["BADGE"]
        def get_name(row):  # Last, First (heurística)
            nm = ws_src.cell(row=row, column=name_col).value
            nm = str(nm).strip() if nm else ""
            if "," in nm:
                last, first = nm.split(",", 1)
                return last.strip(), first.strip()
            parts = nm.split()
            if len(parts) >= 2:
                return parts[-1], " ".join(parts[:-1])
            return nm, ""
    else:
        ln_col = header_map["Last Name"]
        fn_col = header_map["First Name"]
        role_col = header_map["Discipline"]
        badge_col = header_map["Company ID"]
        def get_name(row):
            ln = ws_src.cell(row=row, column=ln_col).value
            fn = ws_src.cell(row=row, column=fn_col).value
            return (str(ln).strip() if ln else ""), (str(fn).strip() if fn else "")

    for r in range(2, ws_src.max_row + 1):
        badge = ws_src.cell(row=r, column=badge_col).value
        if not badge:
            continue
        badge = str(badge).strip()
        role = ws_src.cell(row=r, column=role_col).value
        role = str(role).strip() if role else ""
        last, first = get_name(r)

        # Construir secuencia de estados por fecha
        per_day: Dict[date, Tuple[Optional[str], Optional[str]]] = {}  # date -> (status, comment_text)
        for c_idx, d in dates_sorted:
            v = ws_src.cell(row=r, column=c_idx).value
            st = _norm_status(v)
            if st:
                cm = ws_src.cell(row=r, column=c_idx).comment
                per_day[d] = (st, cm.text if cm else None)

        if not per_day:
            continue

        # Conjuntos para evitar filas duplicadas por fecha
        added_in_dates: Set[date] = set()
        added_out_dates: Set[date] = set()

        next_entry_after_range = None
        next_exit_after_range = None

        # Fechas ordenadas completas
        all_dates = [d for _, d in dates_sorted]
        n = len(all_dates)
        for i, d in enumerate(all_dates):
            if d < start_date:
                continue

            st_d = per_day.get(d, (None, None))[0]
            if not _is_working(st_d):
                continue

            prev_d = all_dates[i - 1] if i - 1 >= 0 else None
            next_d = all_dates[i + 1] if i + 1 < n else None
            st_prev = per_day.get(prev_d, (None, None))[0] if prev_d else None
            st_next = per_day.get(next_d, (None, None))[0] if next_d else None

            # Entrada si cambia respecto al anterior
            if st_prev != st_d:
                cmt = per_day.get(d, (None, None))[1]
                if start_date <= d <= end_date and d not in added_in_dates:
                    ## NUEVO CAMBIO ##
                    pu, _do = get_user_location_for_date(badge, d)
                    ws.cell(row=r_in, column=1, value=idx_in)
                    ws.cell(row=r_in, column=2, value=last)
                    ws.cell(row=r_in, column=3, value=first)
                    ws.cell(row=r_in, column=4, value=badge)
                    ws.cell(row=r_in, column=5, value=company_default)
                    ws.cell(row=r_in, column=6, value=role)
                    ws.cell(row=r_in, column=7, value=pu or "")   # FROM = Pick Up Location
                    ws.cell(row=r_in, column=8, value=d.strftime("%Y-%m-%d"))
                    ws.cell(row=r_in, column=9, value=_times_for(st_d, "IN", cmt))
                    ## FIN NUEVO CAMBIO ##
                    added_in_dates.add(d)
                    r_in += 1
                    idx_in += 1
                elif d > end_date and next_entry_after_range is None:
                    next_entry_after_range = (d, st_d, cmt)

            # Salida si cambia respecto al siguiente
            if st_next != st_d:
                cmt = per_day.get(d, (None, None))[1]
                if start_date <= d <= end_date and d not in added_out_dates:
                    ## NUEVO CAMBIO ##
                    _pu, do = get_user_location_for_date(badge, d)
                    ws.cell(row=r_out, column=11, value=idx_out)
                    ws.cell(row=r_out, column=12, value=last)
                    ws.cell(row=r_out, column=13, value=first)
                    ws.cell(row=r_out, column=14, value=badge)
                    ws.cell(row=r_out, column=15, value=company_default)
                    ws.cell(row=r_out, column=16, value=role)
                    ws.cell(row=r_out, column=17, value=do or "")  # TO = Drop off location
                    ws.cell(row=r_out, column=18, value=d.strftime("%Y-%m-%d"))
                    ws.cell(row=r_out, column=19, value=_times_for(st_d, "OUT", cmt))
                    ## FIN NUEVO CAMBIO ##
                    added_out_dates.add(d)
                    r_out += 1
                    idx_out += 1
                elif d > end_date and next_exit_after_range is None:
                    next_exit_after_range = (d, st_d, cmt)

        # ✅ SIEMPRE agregar el primer IN/OUT posterior al rango, si existen
        if next_entry_after_range:
            d, st, cmt = next_entry_after_range
            if d not in added_in_dates:
                ## NUEVO CAMBIO ##
                pu, _do = get_user_location_for_date(badge, d)
                ws.cell(row=r_in, column=1, value=idx_in)
                ws.cell(row=r_in, column=2, value=last)
                ws.cell(row=r_in, column=3, value=first)
                ws.cell(row=r_in, column=4, value=badge)
                ws.cell(row=r_in, column=5, value=company_default)
                ws.cell(row=r_in, column=6, value=role)
                ws.cell(row=r_in, column=7, value=pu or "")
                ws.cell(row=r_in, column=8, value=d.strftime("%Y-%m-%d"))
                ws.cell(row=r_in, column=9, value=_times_for(st, "IN", cmt))
                ## FIN NUEVO CAMBIO ##
                r_in += 1
                idx_in += 1

        if next_exit_after_range:
            d, st, cmt = next_exit_after_range
            if d not in added_out_dates:
                ## NUEVO CAMBIO ##
                _pu, do = get_user_location_for_date(badge, d)
                ws.cell(row=r_out, column=11, value=idx_out)
                ws.cell(row=r_out, column=12, value=last)
                ws.cell(row=r_out, column=13, value=first)
                ws.cell(row=r_out, column=14, value=badge)
                ws.cell(row=r_out, column=15, value=company_default)
                ws.cell(row=r_out, column=16, value=role)
                ws.cell(row=r_out, column=17, value=do or "")
                ws.cell(row=r_out, column=18, value=d.strftime("%Y-%m-%d"))
                ws.cell(row=r_out, column=19, value=_times_for(st, "OUT", cmt))
                ## FIN NUEVO CAMBIO ##
                r_out += 1
                idx_out += 1

    out = io.BytesIO()
    wb_out.save(out)
    out.seek(0)
    return out.read(), "Transportation report generated."



# ============================================================
# Utilidad: Propagar cambios de código/color a Excel (inmediato)
# ============================================================

def apply_shift_type_update_to_excel(plan_staff_file: str, source: str, old_code: str, new_code: str, color_hex: str) -> Tuple[bool, str]:
    """
    Reemplaza en TODO el archivo Excel el código viejo por el nuevo y aplica el color indicado.
    No altera comentarios ni otros contenidos.
    """
    try:
        if not os.path.exists(plan_staff_file):
            return False, f"Plan staff '{os.path.basename(plan_staff_file)}' not found."

        wb = openpyxl.load_workbook(plan_staff_file)
        ws = wb.active

        hex6 = color_hex.lstrip('#').upper()
        fill = PatternFill(start_color=hex6, end_color=hex6, fill_type="solid")

        # Hallar columnas de fecha para no tocar cabeceras no relacionadas
        date_cols = set()
        for c in ws[1]:
            if isinstance(c.value, datetime):
                date_cols.add(c.column)

        # Buscar celdas con old_code y reemplazar (solo en columnas de fecha)
        for r in range(2, ws.max_row + 1):
            for c in date_cols:
                cell = ws.cell(row=r, column=c)
                if str(cell.value).strip().upper() == str(old_code).strip().upper():
                    cell.value = new_code
                    cell.fill = fill

        wb.save(plan_staff_file)
        return True, "Excel updated with new shift code/color."
    except Exception as e:
        return False, f"Excel update error: {e}"


# ============================================================
# VALIDACIÓN de estructura y salud del archivo (SSoT guardrails)
# ============================================================

def validate_excel_structure(plan_staff_file: str) -> Tuple[bool, List[str], Dict]:
    """
    Valida que el Excel sea una planilla soportada (RGM o Newmont) y que posea columnas de fecha.
    Devuelve: (ok, errors, meta) con meta['variant'] = 'RGM' | 'Newmont' | None y meta['date_columns'].
    """
    errors: List[str] = []
    meta: Dict = {'variant': None, 'date_columns': 0, 'headers': []}

    if not os.path.exists(plan_staff_file):
        errors.append(f"File not found: {plan_staff_file}")
        return False, errors, meta

    try:
        wb = openpyxl.load_workbook(plan_staff_file, read_only=True, data_only=True)
        ws = wb.active
    except Exception as e:
        errors.append(f"Cannot open workbook: {e}")
        return False, errors, meta

    header_map: Dict[str, int] = {}
    date_count = 0
    for c in ws[1]:
        v = c.value
        if isinstance(v, str):
            header_map[v] = c.column
        elif isinstance(v, datetime):
            date_count += 1

    meta['headers'] = list(header_map.keys())
    meta['date_columns'] = int(date_count)

    # Detectar variante
    variant: Optional[str] = None
    if all(h in header_map for h in ("NAME", "ROLE", "BADGE")):
        variant = "RGM"
    elif all(h in header_map for h in ("Last Name", "First Name", "Discipline", "Company ID")):
        variant = "Newmont"
    else:
        errors.append("Unsupported format. Expected headers for RGM (NAME/ROLE/BADGE) or Newmont (Last Name/First Name/Discipline/Company ID).")

    meta['variant'] = variant  # <- SIEMPRE definido en meta (corrige Pylance)

    if variant is None:
        # Sin variante, no seguimos validando otras reglas
        return False, errors, meta

    if meta['date_columns'] == 0:
        errors.append("No date columns detected in first row (datetime cells).")

    # Confirmar headers mínimos por variante
    if variant == "RGM":
        for h in ("NAME", "ROLE", "BADGE"):
            if h not in header_map:
                errors.append(f"Required header '{h}' missing.")
    else:
        for h in ("Last Name", "First Name", "Discipline", "Company ID"):
            if h not in header_map:
                errors.append(f"Required header '{h}' missing.")

    return (len(errors) == 0), errors, meta


# ============================================================
# Comparación Excel ↔ BD (coherencia con SSoT)
# ============================================================

def check_db_sync_with_excel(plan_staff_file: str, source: str) -> Dict:
    """
    Compara usuarios y schedules entre Excel y BD.
    Retorna:
      {
        'users_in_excel': int,
        'users_in_db': int,
        'missing_badges_in_db': [badge,...],     # En Excel pero NO en BD
        'extra_badges_in_db': [badge,...],       # En BD pero NO en Excel
        'schedule_mismatches': [
          {'badge':..., 'date':'YYYY-MM-DD', 'excel': 'ON', 'db':'OFF'}, ...
        ]
      }
    """
    report = {
        'users_in_excel': 0,
        'users_in_db': 0,
        'missing_badges_in_db': [],
        'extra_badges_in_db': [],
        'schedule_mismatches': []
    }

    ok, _errors, _meta = validate_excel_structure(plan_staff_file)
    if not ok:
        return report

    # --- BD
    try:
        from database_logic import get_all_users, get_schedules_for_source
        users_db = get_all_users(source)
        sched_db = get_schedules_for_source(source)
    except Exception:
        users_db = []
        sched_db = []

    db_badges = {str(u.get('badge', '')).strip() for u in users_db if u.get('badge')}
    report['users_in_db'] = len(db_badges)

    sched_db_map: Dict[str, Dict[str, str]] = {}
    for s in sched_db:
        b = str(s.get('badge', '')).strip()
        d = str(s.get('date', '')).strip()
        st = (s.get('status') or '').strip().upper() if s.get('status') else None
        if not b or not d:
            continue
        sched_db_map.setdefault(b, {})[d] = st

    # --- Excel
    try:
        wb = openpyxl.load_workbook(plan_staff_file, data_only=True)
        ws = wb.active
    except Exception:
        return report

    header_map: Dict[str, int] = {}
    date_cols: Dict[int, date] = {}
    for c in ws[1]:
        v = c.value
        if isinstance(v, str):
            header_map[v] = c.column
        elif isinstance(v, datetime):
            date_cols[c.column] = v.date()

    # Determine variant & badge column
    if all(h in header_map for h in ("NAME", "ROLE", "BADGE")):
        badge_col = header_map["BADGE"]
    elif all(h in header_map for h in ("Last Name", "First Name", "Discipline", "Company ID")):
        badge_col = header_map["Company ID"]
    else:
        # No soportado
        return report

    excel_badges: Set[str] = set()
    sched_excel_map: Dict[str, Dict[str, str]] = {}

    def _norm_for_compare(val: object) -> Optional[str]:
        if val is None:
            return None
        s = str(val).strip()
        if s == "":
            return None
        u = s.upper()
        if u in ("OFF", "BREAK", "KO", "LEAVE"):
            return "OFF"
        if u in ("ON", "ON NS"):
            return u
        if u.isdigit() or u == "OK" or "DAY" in u:
            return "ON"
        return u  # personalizado

    for r in range(2, ws.max_row + 1):
        b = ws.cell(row=r, column=badge_col).value
        if not b:
            continue
        b = str(b).strip()
        if not b:
            continue
        excel_badges.add(b)

        # Por fecha
        for c_idx, d in date_cols.items():
            val = ws.cell(row=r, column=c_idx).value
            st = _norm_for_compare(val)
            if st is not None:
                sched_excel_map.setdefault(b, {})[d.isoformat()] = st

    report['users_in_excel'] = len(excel_badges)

    # Diferencias en usuarios
    report['missing_badges_in_db'] = sorted(list(excel_badges - db_badges))
    report['extra_badges_in_db'] = sorted(list(db_badges - excel_badges))

    # Mismatches de schedule (solo comparar cuando ambos tienen valor)
    mismatches: List[Dict] = []
    for b, per_day in sched_excel_map.items():
        for d_str, st_excel in per_day.items():
            st_db = sched_db_map.get(b, {}).get(d_str)
            if st_db is None:
                continue
            if (st_excel or '').upper() != (st_db or '').upper():
                mismatches.append({
                    'badge': b,
                    'date': d_str,
                    'excel': st_excel,
                    'db': st_db
                })
    report['schedule_mismatches'] = mismatches

    return report


# ============================================================
# REGENERACIÓN desde BD (SSoT) — independiente del archivo
# ============================================================

def regenerate_plan_from_db(plan_staff_file: str, source: str) -> Tuple[bool, str]:
    """
    Regenera el archivo PlanStaff (en la ruta indicada) a partir de la BD (SSoT).
    - Si el archivo NO existe o su estructura es inválida, crea una plantilla mínima RGM (NAME/ROLE/BADGE).
    - Luego exporta el estado actual (usuarios+schedules) a dicho archivo.
    """
    try:
        from database_logic import get_all_users, get_schedules_for_source
        users = get_all_users(source)
        schedules = get_schedules_for_source(source)
    except Exception as e:
        return False, f"DB error: {e}"

    # Validar si la plantilla actual es utilizable
    ok, _errors, _meta = validate_excel_structure(plan_staff_file)

    if not ok:
        # Crear plantilla mínima (RGM-like compatible con export_plan_from_db)
        try:
            os.makedirs(os.path.dirname(plan_staff_file), exist_ok=True) if os.path.dirname(plan_staff_file) else None
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "Operations_best_opt"
            headers = ["TEAM", "ROLE", "NAME", "BADGE"]
            for col_idx, h in enumerate(headers, start=1):
                ws.cell(row=1, column=col_idx, value=h)
            wb.save(plan_staff_file)
        except Exception as e:
            return False, f"Cannot create template: {e}"

    # Exportar desde BD usando la plantilla (soporta RGM y Newmont si ya existe)
    ok2, msg = export_plan_from_db(plan_staff_file, users, schedules, plan_staff_file, source)
    if ok2:
        return True, f"PlanStaff file regenerated from DB (SSoT): {os.path.basename(plan_staff_file)}"
    else:
        return False, msg