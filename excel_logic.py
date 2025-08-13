import pandas as pd
from datetime import date, timedelta, datetime
import io
import openpyxl
from openpyxl.styles import PatternFill
import os

PLAN_STAFF_FILE = "PlanStaff.xlsx"

def get_roles_from_excel() -> list:
    if not os.path.exists(PLAN_STAFF_FILE):
        return ["Rol no disponible (PlanStaff.xlsx no encontrado)"]
    try:
        workbook = openpyxl.load_workbook(PLAN_STAFF_FILE, read_only=True, data_only=True)
        sheet = workbook.active
        header_map = {cell.value: cell.column for cell in sheet[1]}
        if "ROLE" not in header_map: return ["Columna ROLE no encontrada"]
        role_col_idx = header_map["ROLE"]
        roles = {sheet.cell(row=i, column=role_col_idx).value for i in range(2, sheet.max_row + 1) if sheet.cell(row=i, column=role_col_idx).value}
        return sorted(list(roles))
    except Exception as e:
        print(f"Error al leer roles de Excel: {e}")
        return ["Error al leer Excel"]

def update_plan_staff_excel(username: str, role: str, badge: str,
                            schedule_status: str, shift_type: str, 
                            schedule_start: date, schedule_end: date):
    try:
        if os.path.exists(PLAN_STAFF_FILE):
            workbook = openpyxl.load_workbook(PLAN_STAFF_FILE)
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
                # <<< CAMBIO: Se establece el texto como "ON NS" para el turno de noche.
                cell_text = "ON NS" 
                fill_color = yellow_fill 
            else: # Day Shift
                # Para el turno de día, podemos usar "ON DS" o simplemente "ON". Usemos "ON".
                cell_text = "ON"
                fill_color = green_fill
        elif schedule_status == "OFF":
            cell_text = "OFF"
            fill_color = red_fill
        
        if cell_text and schedule_start and schedule_end:
            today = date.today()
            delta = schedule_end - schedule_start
            for i in range(delta.days + 1):
                day_to_mark = schedule_start + timedelta(days=i)
                
                if day_to_mark >= today and day_to_mark in date_map:
                    col_idx = date_map[day_to_mark]
                    cell_to_update = sheet.cell(row=employee_row_idx, column=col_idx)
                    
                    cell_to_update.value = cell_text
                    if fill_color:
                        cell_to_update.fill = fill_color
        
        workbook.save(PLAN_STAFF_FILE)
        return True, "PlanStaff.xlsx actualizado exitosamente."
    except Exception as e:
        return False, f"Error al guardar en Excel: {e}"

def get_schedule_preview() -> pd.DataFrame:
    if not os.path.exists(PLAN_STAFF_FILE):
        return pd.DataFrame({"Mensaje": [f"No se encontró el archivo {PLAN_STAFF_FILE}."]})
    try:
        df = pd.read_excel(PLAN_STAFF_FILE, engine='openpyxl')
        return df.fillna('')
    except Exception as e:
        return pd.DataFrame({"Error": [f"No se pudo leer el archivo Excel: {e}"]})

def generate_transport_excel_from_db(records: list) -> bytes:
    travel_in_records, travel_out_records = [], []
    time_in, time_out = "06:00:00", "18:00:00"
    for record in records:
        username, role = record['username'], record['role']
        start_date_str, end_date_str = record['start_date'], record['end_date']
        start_date = datetime.strptime(start_date_str, '%Y-%m-%d').date()
        end_date = datetime.strptime(end_date_str, '%Y-%m-%d').date()
        
        if ',' in username:
            parts = username.split(',', 1); last_name, first_name = parts[0].strip(), parts[1].strip()
        else:
            parts = username.split(); last_name = parts[-1] if parts else ''; first_name = " ".join(parts[:-1]) if len(parts) > 1 else ''
        
        travel_in_records.append({"NAME": last_name, "FIRST NAME": first_name, "DEPT": role, "DATE": start_date.strftime('%Y-%m-%d'), "TIME": time_in})
        travel_out_records.append({"NAME": last_name, "FIRST NAME": first_name, "DEPT": role, "DATE": end_date.strftime('%Y-%m-%d'), "TIME": time_out})
    
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        cols_in = ['#', 'NAME', 'FIRST NAME', 'GID', 'COMPANY', 'DEPT', 'FROM', 'DATE', 'TIME']
        cols_out = ['#', 'NAME', 'FIRST NAME', 'GID', 'COMPANY', 'DEPT', 'TO', 'DATE', 'TIME']
        df_in, df_out = pd.DataFrame(travel_in_records), pd.DataFrame(travel_out_records)
        
        pd.DataFrame().to_excel(writer, sheet_name='travel list', index=False)
        ws = writer.sheets['travel list']
        
        if not df_in.empty:
            df_in = df_in.drop_duplicates().sort_values(by=["DEPT", "NAME", "DATE"]).reset_index(drop=True)
            df_in["GID"], df_in["COMPANY"], df_in["FROM"] = "", "PLGims", "PBO"
            df_in.insert(0, '#', range(1, 1 + len(df_in)))
            df_in.reindex(columns=cols_in).to_excel(writer, sheet_name='travel list', startrow=5, index=False, header=True)
        
        if not df_out.empty:
            df_out = df_out.drop_duplicates().sort_values(by=["DEPT", "NAME", "DATE"]).reset_index(drop=True)
            df_out["GID"], df_out["COMPANY"], df_out["TO"] = "", "PLGims", "PBO"
            df_out.insert(0, '#', range(1, 1 + len(df_out)))
            df_out.reindex(columns=cols_out).to_excel(writer, sheet_name='travel list', startrow=5, startcol=len(cols_in) + 1, index=False, header=True)
        
        ws.cell(row=2, column=1, value="MERIAN TRANSPORTATION REQUEST")
        ws.cell(row=5, column=1, value="IN")
        ws.cell(row=5, column=2, value="TRAVEL TO SITE")
        ws.cell(row=5, column=len(cols_in) + 2, value="OUT")
        ws.cell(row=5, column=len(cols_in) + 3, value="TRAVEL FROM SITE")
        
    return output.getvalue()