import json
from PyQt6.QtCore import QMimeData
from PyQt6.QtGui import QClipboard
from PyQt6.QtWidgets import QApplication, QTableWidgetItem

class ScheduleClipboardService:
    MIME_TYPE = "application/x-transport-shifts"

    @staticmethod
    def copy_to_clipboard(table_widget, row_identities, date_col_dates, custom_shift_map):
        selection = table_widget.selectedRanges()
        if not selection:
            return

        # Get bounding box of selection
        top = min(r.topRow() for r in selection)
        bottom = max(r.bottomRow() for r in selection)
        left = min(r.leftColumn() for r in selection)
        right = max(r.rightColumn() for r in selection)

        row_count = bottom - top + 1
        col_count = right - left + 1

        # Matrix to hold data
        grid_data = [[None for _ in range(col_count)] for _ in range(row_count)]
        text_buffer = []

        for r in range(row_count):
            row_text = []
            for c in range(col_count):
                abs_row = top + r
                abs_col = left + c
                
                item = table_widget.item(abs_row, abs_col)
                text = item.text() if item else ""
                
                # Construct Rich Data Object
                # In a real scenario, you might want to fetch from DB to get the 'remark' 
                # if it's not stored in the item's UserRole. 
                # For now, we infer basic info from text + custom_map.
                
                shift_info = {
                    "text": text,
                    "clean_code": text.strip().upper(),
                    # We could store remarks here if we stored them in ItemData
                }
                
                grid_data[r][c] = shift_info
                row_text.append(text)
            text_buffer.append("\t".join(row_text))

        plain_text = "\n".join(text_buffer)
        json_data = json.dumps({
            "rows": row_count,
            "cols": col_count,
            "grid": grid_data
        })

        mime = QMimeData()
        mime.setText(plain_text)
        mime.setData(ScheduleClipboardService.MIME_TYPE, json_data.encode('utf-8'))
        
        QApplication.clipboard().setMimeData(mime)

    @staticmethod
    def parse_clipboard():
        clipboard = QApplication.clipboard()
        mime = clipboard.mimeData()

        if mime.hasFormat(ScheduleClipboardService.MIME_TYPE):
            try:
                data = mime.data(ScheduleClipboardService.MIME_TYPE).data()
                return json.loads(data.decode('utf-8')), "JSON"
            except Exception:
                pass # Fallback to text
        
        if mime.hasText():
            text = mime.text()
            # Convert TSV to grid
            rows = text.split('\n')
            # Remove empty trailing row if common in copy
            if rows and not rows[-1].strip():
                rows.pop()
                
            grid = []
            for row in rows:
                cols = row.split('\t')
                grid_row = [{"text": c.strip(), "clean_code": c.strip().upper()} for c in cols]
                grid.append(grid_row)
            
            return {
                "rows": len(grid),
                "cols": len(grid[0]) if grid else 0,
                "grid": grid
            }, "TEXT"
            
        return None, None