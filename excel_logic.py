import pandas as pd
from datetime import date, timedelta, datetime
import io
import openpyxl
from openpyxl.styles import PatternFill
import os


# --------------------------
# Utilidades internas
# --------------------------
def _is_blank_series(s: pd.Series) -> bool:
    """Return True if the entire series is NaN or empty strings."""
    if s is None:
        return True
    if s.isna().all():
        return True
    s_str = s.astype(str).str.strip().str.lower()
    return (s_str.eq('') | s_str.eq('nan') | s_str.eq('none') | s_str.eq('null')).all()

def _prefix_for_file(plan_staff_file: str) -> str:
    """Badge prefix based on file (NM for Newmont by default)."""
    base = os.path.basename(plan_staff_file).lower()
    if 'newmont' in base:
        return 'NM'
    return 'ID'


# --------------------------
# Lecturas auxiliares
# --------------------------
def get_roles_from_excel(plan_staff_file: str) -> list:
    """Return the list of unique roles from the given Excel file."""
    if not os.path.exists(plan_staff_file):
        return [f"Role not available ({os.path.basename(plan_staff_file)} not found)"]
    try:
        workbook = openpyxl.load_workbook(plan_staff_file, read_only=True, data_only=True)
        sheet = workbook.active

        header_map = {cell.value: cell.column for cell in sheet[1]}
        role_header = "ROLE" if "ROLE" in header_map else ("Discipline" if "Discipline" in header_map else None)
        if not role_header:
            return ["ROLE/Discipline column not found"]

        role_col_idx = header_map[role_header]
        roles = {
            sheet.cell(row=i, column=role_col_idx).value
            for i in range(2, sheet.max_row + 1)
            if sheet.cell(row=i, column=role_col_idx).value
        }
        return sorted(list(roles))
    except Exception as e:
        print(f"Error reading roles from Excel: {e}")
        return ["Error reading Excel"]

def get_users_from_excel(plan_staff_file: str) -> list:
    """
    Extract a list of users (name, role, badge) from the Excel file.
    - Accepts RGM format (NAME, ROLE, BADGE).
    - Accepts Newmont format (Last Name, First Name, Discipline, Company ID).
    - If BADGE does not exist or is blank, a stable one is GENERATED.
    """
    if not os.path.exists(plan_staff_file):
        print(f"Import file not found: {plan_staff_file}")
        return []
    try:
        df = pd.read_excel(plan_staff_file, engine='openpyxl')

        # Expected layouts
        rgm_cols = ['NAME', 'ROLE', 'BADGE']
        newmont_cols = ['Last Name', 'First Name', 'Discipline', 'Company ID']

        users_df = None

        if all(col in df.columns for col in rgm_cols):
            users_df = df[rgm_cols].copy()

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
            if all(col in df.columns for col in ['NAME', 'ROLE']):
                users_df = df[['NAME', 'ROLE']].copy()
                prefix = _prefix_for_file(plan_staff_file)
                users_df['BADGE'] = [f"{prefix}{i+1:05d}" for i in range(len(users_df))]
            else:
                print(f"Error: Required columns not found in {os.path.basename(plan_staff_file)}.")
                print(f"Expected RGM format ({rgm_cols}) or Newmont ({newmont_cols}).")
                return []

        users_df['NAME'] = users_df['NAME'].astype(str).str.strip()
        users_df['ROLE'] = users_df['ROLE'].astype(str).str.strip()
        users_df['BADGE'] = users_df['BADGE'].astype(str).str.strip()

        users_df = users_df[(users_df['NAME'] != '') & (users_df['BADGE'] != '')]
        users_df.drop_duplicates(subset=['BADGE'], keep='first', inplace=True)

        users_df.rename(columns={'NAME': 'name', 'ROLE': 'role', 'BADGE': 'badge'}, inplace=True)
        return users_df.to_dict('records')

    except Exception as e:
        print(f"Error processing the Excel file: {e}")
        return []

# --------------------------
# Escritura/lectura de plan staff
# --------------------------
def update_plan_staff_excel(plan_staff_file: str, username: str, role: str, badge: str,
                            schedule_status: str, shift_type: str,
                            schedule_start: date, schedule_end: date):
    """
    Actualiza (o crea si no existe) la fila del empleado en el Excel de plan staff.
    - Busca primero por BADGE y, si no lo encuentra, por NAME.
    - Escribe ON/ON NS/OFF con sus colores.
    - Si schedule_status es None (opción "Do Not Mark Days"), limpia el rango: deja celdas vacías y SIN color.
    """
    try:
        # Abrir o crear libro
        if os.path.exists(plan_staff_file):
            workbook = openpyxl.load_workbook(plan_staff_file)
            sheet = workbook.active
        else:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Operations_best_opt"
            headers = ["TEAM", "ROLE", "NAME", "BADGE"]
            for col_idx, h in enumerate(headers, start=1):
                sheet.cell(row=1, column=col_idx).value = h

        # Colores usados en la vista
        green_fill  = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")  # ON (día)
        red_fill    = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")  # OFF
        yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid")  # ON NS (noche)

        # Mapas de cabeceras (texto) y fechas (encabezados datetime)
        header_map = {cell.value: cell.column for cell in sheet[1] if isinstance(cell.value, str)}
        date_map = {}
        for cell in sheet[1]:
            v = cell.value
            if isinstance(v, datetime):
                date_map[v.date()] = cell.column  # solo fechas válidas

        # Localizar fila del empleado
        employee_row_idx = None

        # 1) Buscar por BADGE
        badge_col_idx = header_map.get("BADGE")
        if badge_col_idx:
            for i in range(2, sheet.max_row + 1):
                cell_value = sheet.cell(row=i, column=badge_col_idx).value
                if cell_value and str(cell_value) == str(badge):
                    employee_row_idx = i
                    break

        # 2) Si no, por NAME
        if not employee_row_idx and "NAME" in header_map:
            name_col_idx = header_map["NAME"]
            for i in range(2, sheet.max_row + 1):
                cell_value = sheet.cell(row=i, column=name_col_idx).value
                if cell_value and str(cell_value).strip().lower() == str(username).strip().lower():
                    employee_row_idx = i
                    break

        # 3) Si no existe, crear nueva fila
        if not employee_row_idx:
            employee_row_idx = sheet.max_row + 1

        # Escribir datos básicos (no cambia colores)
        if "NAME" in header_map:
            sheet.cell(row=employee_row_idx, column=header_map["NAME"]).value = username
        if "ROLE" in header_map:
            sheet.cell(row=employee_row_idx, column=header_map["ROLE"]).value = role
        if "BADGE" in header_map:
            sheet.cell(row=employee_row_idx, column=header_map["BADGE"]).value = badge

        # Determinar texto/estilo según estado
        cell_text = ""
        fill_color = None
        if schedule_status == "ON":
            if shift_type == "Night Shift":
                cell_text = "ON NS"
                fill_color = yellow_fill
            else:
                cell_text = "ON"
                fill_color = green_fill
        elif schedule_status == "OFF":
            cell_text = "OFF"
            fill_color = red_fill
        elif schedule_status is None:
            # "Do Not Mark Days": limpiar texto y color en el rango
            cell_text = None
            fill_color = None

        # Marcar (o limpiar) el rango indicado
        if schedule_start and schedule_end:
            delta = schedule_end - schedule_start
            for i in range(delta.days + 1):
                day_to_mark = schedule_start + timedelta(days=i)
                if day_to_mark in date_map:
                    col_idx = date_map[day_to_mark]
                    cell_to_update = sheet.cell(row=employee_row_idx, column=col_idx)
                    if schedule_status is None:
                        # LIMPIAR: sin texto y sin relleno
                        cell_to_update.value = None
                        cell_to_update.fill = PatternFill(fill_type=None)
                    else:
                        # ESCRIBIR: ON / ON NS / OFF
                        cell_to_update.value = cell_text
                        if fill_color:
                            cell_to_update.fill = fill_color

        workbook.save(plan_staff_file)
        return True, f"{os.path.basename(plan_staff_file)} updated successfully."
    except Exception as e:
        return False, f"Error saving to Excel: {e}"

def get_schedule_preview(plan_staff_file: str) -> pd.DataFrame:
    """Read the given Excel file and return it as a pandas DataFrame."""
    if not os.path.exists(plan_staff_file):
        return pd.DataFrame({"Message": [f"File {os.path.basename(plan_staff_file)} not found."]})
    try:
        df = pd.read_excel(plan_staff_file, engine='openpyxl')
        return df.fillna('')
    except Exception as e:
        return pd.DataFrame({"Error": f"Could not read the Excel file: {e}"}, index=[0])

def generate_transport_excel_from_planstaff(plan_staff_file: str, report_start: date, report_end: date) -> tuple:
    """Generate a transportation report by analyzing the given Excel file."""
    try:
        if not os.path.exists(plan_staff_file):
            return None, f"File {os.path.basename(plan_staff_file)} not found."
        df = pd.read_excel(plan_staff_file, engine='openpyxl').fillna('OFF')
    except Exception as e:
        return None, f"Could not read the Excel file: {e}"

    info_cols = ['NAME', 'ROLE']
    for col in info_cols:
        if col not in df.columns:
            return None, f"Required Excel column is missing: '{col}'"

    all_date_cols = sorted([col for col in df.columns if isinstance(col, datetime)])
    if not all_date_cols:
        return None, f"No date columns found in {os.path.basename(plan_staff_file)}."

    date_col_map = {col.date(): col for col in all_date_cols}
    travel_in_records, travel_out_records = [], []

    for _, row in df.iterrows():
        username = row.get('NAME')
        if not isinstance(username, str) or not username.strip():
            continue
        role = row.get('ROLE', 'N/A')

        for d in pd.date_range(start=report_start, end=report_end):
            current_day = d.date()
            prev_day = current_day - timedelta(days=1)
            next_day = current_day + timedelta(days=1)

            if prev_day in date_col_map and current_day in date_col_map:
                prev_status = str(row.get(date_col_map[prev_day], 'OFF')).upper()
                curr_status = str(row.get(date_col_map[current_day], 'OFF')).upper()
                if 'ON' not in prev_status and 'ON' in curr_status:
                    time_in = "12:00:00" if curr_status == 'ON NS' else "06:00:00"
                    travel_in_records.append({"NAME": username, "DEPT": role, "DATE": current_day.strftime('%Y-%m-%d'), "TIME": time_in})

            if next_day in date_col_map and current_day in date_col_map:
                curr_status = str(row.get(date_col_map[current_day], 'OFF')).upper()
                next_status = str(row.get(date_col_map[next_day], 'OFF')).upper()
                if 'ON' in curr_status and 'ON' not in next_status:
                    time_out = "06:00:00" if curr_status == 'ON NS' else "12:00:00"
                    travel_out_records.append({"NAME": username, "DEPT": role, "DATE": current_day.strftime('%Y-%m-%d'), "TIME": time_out})

    if not travel_in_records and not travel_out_records:
        return None, "No staff check-ins or check-outs found in the date range."

    def parse_name(username_str):
        if ',' in username_str:
            parts = username_str.split(',', 1)
            return pd.Series([parts[0].strip(), parts[1].strip()])
        else:
            parts = username_str.split()
            last_name = parts[-1] if parts else ''
            first_name = " ".join(parts[:-1]) if len(parts) > 1 else ''
            return pd.Series([last_name, first_name])

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        pd.DataFrame().to_excel(writer, sheet_name='travel list', index=False)
        ws = writer.sheets['travel list']
        cols_in = ['#', 'NAME', 'FIRST NAME', 'GID', 'COMPANY', 'DEPT', 'FROM', 'DATE', 'TIME']
        cols_out = ['#', 'NAME', 'FIRST NAME', 'GID', 'COMPANY', 'DEPT', 'TO', 'DATE', 'TIME']

        df_in_processed = pd.DataFrame()
        if travel_in_records:
            df_in = pd.DataFrame(travel_in_records)
            df_in[['NAME', 'FIRST NAME']] = df_in['NAME'].apply(parse_name)
            df_in = df_in.drop_duplicates().sort_values(by=["DEPT", "NAME", "DATE"]).reset_index(drop=True)
            df_in["GID"], df_in["COMPANY"], df_in["FROM"] = "", "PLGims", "PBO"
            df_in.insert(0, '#', range(1, 1 + len(df_in)))
            df_in_processed = df_in.reindex(columns=cols_in)
            df_in_processed.to_excel(writer, sheet_name='travel list', startrow=5, index=False, header=True)

        df_out_processed = pd.DataFrame()
        if travel_out_records:
            df_out = pd.DataFrame(travel_out_records)
            df_out[['NAME', 'FIRST NAME']] = df_out['NAME'].apply(parse_name)
            df_out = df_out.drop_duplicates().sort_values(by=["DEPT", "NAME", "DATE"]).reset_index(drop=True)
            df_out["GID"], df_out["COMPANY"], df_out["TO"] = "", "PLGims", "PBO"
            df_out.insert(0, '#', range(1, 1 + len(df_out)))
            start_col_out = len(cols_in) + 1 if not df_in_processed.empty else 0
            df_out_processed = df_out.reindex(columns=cols_out)
            df_out_processed.to_excel(writer, sheet_name='travel list', startrow=5, startcol=start_col_out, index=False, header=True)

        ws.cell(row=2, column=1, value="MERIAN TRANSPORTATION REQUEST")
        start_col_out_label = len(cols_in) + 2 if not df_in_processed.empty else 1
        if not df_in_processed.empty:
            ws.cell(row=5, column=1, value="IN")
            ws.cell(row=5, column=2, value="TRAVEL TO SITE")
        if not df_out_processed.empty:
            ws.cell(row=5, column=start_col_out_label, value="OUT")
            ws.cell(row=5, column=start_col_out_label + 1, value="TRAVEL FROM SITE")

    return output.getvalue(), f"Report successfully generated from {os.path.basename(plan_staff_file)}."
