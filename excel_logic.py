import pandas as pd
from datetime import date, timedelta, datetime
import io
import openpyxl
from openpyxl.styles import PatternFill
import os

def get_roles_from_excel(plan_staff_file: str) -> list:
    """Obtiene la lista de roles únicos desde el archivo Excel especificado."""
    if not os.path.exists(plan_staff_file):
        return [f"Rol no disponible ({os.path.basename(plan_staff_file)} no encontrado)"]
    try:
        workbook = openpyxl.load_workbook(plan_staff_file, read_only=True, data_only=True)
        sheet = workbook.active
        header_map = {cell.value: cell.column for cell in sheet[1]}
        if "ROLE" not in header_map: return ["Columna ROLE no encontrada"]
        role_col_idx = header_map["ROLE"]
        roles = {sheet.cell(row=i, column=role_col_idx).value for i in range(2, sheet.max_row + 1) if sheet.cell(row=i, column=role_col_idx).value}
        return sorted(list(roles))
    except Exception as e:
        print(f"Error al leer roles de Excel: {e}")
        return ["Error al leer Excel"]

def get_users_from_excel(plan_staff_file: str) -> list:
    """
    Extrae una lista de usuarios (Nombre, Rol, Badge) del archivo Excel.
    Se asegura de que no haya duplicados basados en el 'BADGE'.
    """
    if not os.path.exists(plan_staff_file):
        print(f"Archivo para importación no encontrado: {plan_staff_file}")
        return []
    try:
        df = pd.read_excel(plan_staff_file, engine='openpyxl')
        
        required_cols = ['NAME', 'ROLE', 'BADGE']
        if not all(col in df.columns for col in required_cols):
            print(f"Faltan columnas requeridas para importar usuarios. Se necesitan: {required_cols}")
            return []

        # Seleccionar, limpiar y eliminar duplicados
        users_df = df[required_cols].copy()
        users_df.dropna(subset=['NAME', 'BADGE'], inplace=True)
        users_df = users_df[users_df['NAME'].str.strip() != '']
        users_df.drop_duplicates(subset=['BADGE'], keep='first', inplace=True)
        
        # Asegurarse de que los datos tienen el tipo correcto
        users_df['BADGE'] = users_df['BADGE'].astype(str)
        users_df['ROLE'] = users_df['ROLE'].fillna('N/A')

        # Renombrar columnas para que coincidan con los parámetros de la base de datos
        users_df.rename(columns={'NAME': 'name', 'ROLE': 'role', 'BADGE': 'badge'}, inplace=True)

        return users_df.to_dict('records')

    except Exception as e:
        print(f"Error al leer usuarios de Excel: {e}")
        return []

def update_plan_staff_excel(plan_staff_file: str, username: str, role: str, badge: str,
                              schedule_status: str, shift_type: str, 
                              schedule_start: date, schedule_end: date):
    """Actualiza o crea una entrada para un empleado en el archivo Excel especificado."""
    try:
        if os.path.exists(plan_staff_file):
            workbook = openpyxl.load_workbook(plan_staff_file)
            sheet = workbook.active
        else:
            workbook = openpyxl.Workbook()
            sheet = workbook.active
            sheet.title = "Operations_best_opt"
            headers = ["TEAM", "ROLE", "NAME", "BADGE"]
            sheet.append(headers)

        green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
        red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
        yellow_fill = PatternFill(start_color="FFFF99", end_color="FFFF99", fill_type="solid") 
        
        header_map = {cell.value: cell.column for cell in sheet[1]}
        date_map = {cell.value.date() if isinstance(cell.value, datetime) else None: cell.column for cell in sheet[1]}
        
        employee_row_idx = None
        badge_col_idx = header_map.get("BADGE")
        if badge_col_idx:
            for i in range(2, sheet.max_row + 1):
                cell_value = sheet.cell(row=i, column=badge_col_idx).value
                if cell_value and str(cell_value) == str(badge):
                    employee_row_idx = i
                    break
        
        if not employee_row_idx:
            employee_row_idx = sheet.max_row + 1

        if "NAME" in header_map:
            sheet.cell(row=employee_row_idx, column=header_map["NAME"]).value = username
        if "ROLE" in header_map:
            sheet.cell(row=employee_row_idx, column=header_map["ROLE"]).value = role
        if "BADGE" in header_map:
            sheet.cell(row=employee_row_idx, column=header_map["BADGE"]).value = badge
        
        cell_text = ""
        fill_color = None

        if schedule_status == "ON":
            if shift_type == "Night Shift":
                cell_text = "ON NS" 
                fill_color = yellow_fill 
            else: # Day Shift
                cell_text = "ON"
                fill_color = green_fill
        elif schedule_status == "OFF":
            cell_text = "OFF"
            fill_color = red_fill
        
        if cell_text and schedule_start and schedule_end:
            delta = schedule_end - schedule_start
            for i in range(delta.days + 1):
                day_to_mark = schedule_start + timedelta(days=i)
                if day_to_mark in date_map:
                    col_idx = date_map[day_to_mark]
                    cell_to_update = sheet.cell(row=employee_row_idx, column=col_idx)
                    cell_to_update.value = cell_text
                    if fill_color:
                        cell_to_update.fill = fill_color
        
        workbook.save(plan_staff_file)
        return True, f"{os.path.basename(plan_staff_file)} actualizado exitosamente."
    except Exception as e:
        return False, f"Error al guardar en Excel: {e}"

def get_schedule_preview(plan_staff_file: str) -> pd.DataFrame:
    """Lee el archivo Excel especificado y lo devuelve como un DataFrame de pandas."""
    if not os.path.exists(plan_staff_file):
        return pd.DataFrame({"Mensaje": [f"No se encontró el archivo {os.path.basename(plan_staff_file)}."]})
    try:
        df = pd.read_excel(plan_staff_file, engine='openpyxl')
        return df.fillna('')
    except Exception as e:
        return pd.DataFrame({"Error": [f"No se pudo leer el archivo Excel: {e}"]})

def generate_transport_excel_from_planstaff(plan_staff_file: str, report_start: date, report_end: date) -> tuple:
    """Genera un reporte de transporte analizando el archivo Excel especificado."""
    try:
        if not os.path.exists(plan_staff_file):
            return None, f"No se encontró el archivo {os.path.basename(plan_staff_file)}."
        df = pd.read_excel(plan_staff_file, engine='openpyxl').fillna('OFF')
    except Exception as e:
        return None, f"No se pudo leer el archivo Excel: {e}"

    info_cols = ['NAME', 'ROLE']
    for col in info_cols:
        if col not in df.columns:
            return None, f"Falta la columna requerida en Excel: '{col}'"

    all_date_cols = sorted([col for col in df.columns if isinstance(col, datetime)])
    if not all_date_cols:
        return None, f"No se encontraron columnas de fecha en {os.path.basename(plan_staff_file)}."

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
        return None, "No se encontraron entradas o salidas de personal en el rango de fechas."

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
            
    return output.getvalue(), f"Reporte generado exitosamente desde {os.path.basename(plan_staff_file)}."