# undo_commands.py
# Extracted from main_window.py — Undo/Redo command classes for schedule editing.

import sqlite3
import logging
from datetime import timedelta

from PyQt6.QtGui import QUndoCommand

import database_logic as db
import excel_logic as excel

logger = logging.getLogger(__name__)


# =============================================================================
# UNDO / REDO COMMAND  (Command Pattern sobre lógica existente — SSoT = SQLite)
# =============================================================================
class UndoCellChangeCommand(QUndoCommand):
    """
    Encapsula un cambio de celda/rango para Ctrl+Z / Ctrl+Y.

    PRINCIPIOS:
    • NO duplica reglas de negocio: delega a métodos existentes de PlanStaffWidget.
    • SSoT siempre es SQLite: undo/redo escriben en DB primero, luego actualizan
      Excel y la UI.
    • El primer redo() no re-ejecuta la escritura (ya la hizo _on_schedule_cell_changed).
    • _patch_cells_in_table() actualiza solo las celdas afectadas sin recargar Excel.
    """

    def __init__(
        self,
        widget,                  # referencia a PlanStaffWidget
        badge: str,
        role: str,
        username: str,
        start_date,              # datetime.date
        end_date,                # datetime.date
        old_schedule_map: dict,  # {date_str: {status, shift_type, in_time, out_time, remark, force_new_entry}}
        old_pickup,
        old_dropoff,
        new_status: str,
        new_shift_type,
        new_in_time,
        new_out_time,
        new_remark,
        new_pickup,
        new_dropoff,
        description: str = "Edit schedule",
    ):
        super().__init__(description)
        self._widget       = widget
        self._badge        = badge
        self._role         = role
        self._username     = username
        self._start_date   = start_date
        self._end_date     = end_date
        self._old_map      = old_schedule_map   # snapshot antes del cambio
        self._old_pickup   = old_pickup
        self._old_dropoff  = old_dropoff
        self._new_status       = new_status
        self._new_shift_type   = new_shift_type
        self._new_in_time      = new_in_time
        self._new_out_time     = new_out_time
        self._new_remark       = new_remark
        self._new_pickup       = new_pickup
        self._new_dropoff      = new_dropoff
        self._first_redo   = True   # skip primer redo: la escritura ya ocurrió

    # ------------------------------------------------------------------ redo --
    def redo(self):
        if self._first_redo:
            self._first_redo = False
            return   # ya aplicado por _on_schedule_cell_changed

        widget = self._widget
        try:
            db.upsert_schedule_range(
                self._badge, self._start_date, self._end_date,
                self._new_status, self._new_shift_type, widget.source,
                self._new_in_time, self._new_out_time, self._new_remark,
            )
            widget._consolidate_and_record_logistics(
                self._badge, self._role, self._username,
                self._start_date, self._end_date, self._new_status,
            )
            if self._new_pickup or self._new_dropoff:
                db.assign_user_location_range(
                    self._badge, self._start_date, self._end_date,
                    self._new_pickup, self._new_dropoff,
                )
            excel.update_plan_staff_excel(
                widget.excel_file,
                self._username, self._role, self._badge,
                self._new_status, self._new_shift_type,
                self._start_date, self._end_date,
                widget.source, self._new_in_time, self._new_out_time,
            )
            widget._patch_cells_in_table(
                self._badge, self._start_date, self._end_date, self._new_status
            )
            db.log_event(
                widget.logged_username, widget.source,
                "SHIFT_REDO",
                f"{self._badge} {self._start_date}..{self._end_date} → {self._new_status}",
            )
        except Exception as e:
            print(f"[UndoCmd.redo] Error: {e}")

    # ------------------------------------------------------------------ undo --
    def undo(self):
        widget = self._widget
        try:
            # 1. Restaurar cada día con sus datos originales (o borrar si no existía)
            self._restore_old_schedule()

            # 2. Recalcular operaciones logísticas con el status antiguo
            first_day_str = self._start_date.isoformat()
            old_day = self._old_map.get(first_day_str) or {}
            old_status      = old_day.get("status") or ""
            old_shift_type  = old_day.get("shift_type")
            old_in_time     = old_day.get("in_time")
            old_out_time    = old_day.get("out_time")

            widget._consolidate_and_record_logistics(
                self._badge, self._role, self._username,
                self._start_date, self._end_date, old_status,
            )

            # 3. Restaurar ubicaciones si cambiaaron
            if self._old_pickup is not None or self._old_dropoff is not None:
                db.assign_user_location_range(
                    self._badge, self._start_date, self._end_date,
                    self._old_pickup, self._old_dropoff,
                )

            # 4. Actualizar Excel con estado anterior
            excel.update_plan_staff_excel(
                widget.excel_file,
                self._username, self._role, self._badge,
                old_status, old_shift_type,
                self._start_date, self._end_date,
                widget.source, old_in_time, old_out_time,
            )

            # 5. Actualizar sólo las celdas afectadas en la tabla (sin recargar Excel)
            widget._patch_cells_in_table(
                self._badge, self._start_date, self._end_date, old_status
            )
            db.log_event(
                widget.logged_username, widget.source,
                "SHIFT_UNDO",
                f"{self._badge} {self._start_date}..{self._end_date} ← {old_status}",
            )
        except Exception as e:
            print(f"[UndoCmd.undo] Error: {e}")

    # --------------------------------------------------------- helper privado --
    def _restore_old_schedule(self):
        """Restaura registros en DB día a día desde el snapshot old_map."""
        widget = self._widget
        d = self._start_date
        while d <= self._end_date:
            day_str = d.isoformat()
            old_rec = self._old_map.get(day_str)
            if not old_rec or old_rec.get("status") is None:
                # El día no existía antes → borrar el registro nuevo
                db.clear_schedule_range(self._badge, d, d, widget.source)
            else:
                db.upsert_schedule_day(
                    self._badge, d,
                    old_rec.get("status", ""),
                    old_rec.get("shift_type"),
                    widget.source,
                    old_rec.get("in_time"),
                    old_rec.get("out_time"),
                    old_rec.get("remark"),
                    old_rec.get("force_new_entry"),
                )
            d += timedelta(days=1)


# =============================================================================
# UNDO / REDO (SNAPSHOT) — para operaciones multi-celda como drag-fill
# =============================================================================
class UndoScheduleSnapshotCommand(QUndoCommand):
    """
    Undo/Redo basado en *snapshots* para drag-fill (y futuras ops masivas).

    PRINCIPIOS:
    • Guarda un mapa completo (old y new) de la ventana afectada, incluidos
      los bordes adyacentes al rango (para respetar force_new_entry).
    • undo()/redo() restauran EXACTAMENTE ese estado en DB (SSoT).
    • Luego reconcilia operaciones logísticas, sincroniza Excel y parchea UI.
    • El primer redo() es no-op: el cambio original ya ocurrió en _apply_fill_from_anchor.
    """

    def __init__(
        self,
        widget,                     # PlanStaffWidget
        badge: str,
        role: str,
        username: str,
        snapshot_start,             # datetime.date — inicio de la ventana capturada
        snapshot_end,               # datetime.date — fin de la ventana capturada
        op_start,                   # datetime.date — para consolidación logística
        op_end,                     # datetime.date
        old_schedule_map: dict,     # {date_iso: {...}}  — estado ANTES del drag
        new_schedule_map: dict,     # {date_iso: {...}}  — estado DESPUÉS del drag
        description: str = "Drag fill",
    ):
        super().__init__(description)
        self._widget         = widget
        self._badge          = badge
        self._role           = role
        self._username       = username
        self._snapshot_start = snapshot_start
        self._snapshot_end   = snapshot_end
        self._op_start       = op_start
        self._op_end         = op_end
        self._old_map        = old_schedule_map or {}
        self._new_map        = new_schedule_map or {}
        self._first_redo     = True   # el drag original ya ocurrió

    # ------------------------------------------------------------------ redo --
    def redo(self):
        if self._first_redo:
            self._first_redo = False
            return  # ya aplicado por _apply_fill_from_anchor
        self._apply_snapshot(self._new_map, action="SHIFT_REDO_DRAGFILL")

    # ------------------------------------------------------------------ undo --
    def undo(self):
        self._apply_snapshot(self._old_map, action="SHIFT_UNDO_DRAGFILL")

    # --------------------------------------------------------------- helpers --
    def _apply_snapshot(self, target_map: dict, action: str):
        """Restaura target_map en DB → reconsolida → sincroniza Excel → parchea UI."""
        import sqlite3 as _sqlite3
        widget = self._widget
        try:
            # 1) DB (SSoT) — transacción atómica día a día
            conn = _sqlite3.connect(db.DB_FILE)
            cur  = conn.cursor()
            cur.execute("BEGIN TRANSACTION")
            try:
                d = self._snapshot_start
                while d <= self._snapshot_end:
                    day_str = d.isoformat()
                    rec = target_map.get(day_str)
                    if rec and rec.get("status") is not None:
                        db.upsert_schedule_day(
                            self._badge, d,
                            rec.get("status") or "",
                            rec.get("shift_type"),
                            widget.source,
                            rec.get("in_time"),
                            rec.get("out_time"),
                            rec.get("remark"),
                            rec.get("force_new_entry"),
                            cursor=cur,
                        )
                    else:
                        cur.execute(
                            "DELETE FROM schedules WHERE badge = ? AND source = ? AND date = ?",
                            (self._badge, widget.source, day_str),
                        )
                    d += timedelta(days=1)
                conn.commit()
            except Exception:
                conn.rollback()
                raise
            finally:
                conn.close()

            # 2) Operaciones logísticas — reconsolidar con el status resultante
            status_for_consol = ""
            d = self._op_start
            while d <= self._op_end:
                rec = target_map.get(d.isoformat())
                st  = (rec.get("status") or "").strip().upper() if rec else ""
                if st and db.is_working_status(st, widget.source):
                    status_for_consol = st
                    break
                d += timedelta(days=1)

            widget._consolidate_and_record_logistics(
                self._badge, self._role, self._username,
                self._op_start, self._op_end, status_for_consol,
            )

            # 3) Excel — sincronizar en runs contiguos (minimiza escrituras)
            self._sync_excel_from_map(target_map)

            # 4) UI — parchear sólo celdas afectadas desde el mapa (soporta statuses mixtos)
            widget._patch_cells_in_table_from_map(
                self._badge, self._snapshot_start, self._snapshot_end, target_map
            )

            db.log_event(
                widget.logged_username, widget.source,
                action,
                f"{self._badge} {self._snapshot_start}..{self._snapshot_end}",
            )
        except Exception as e:
            print(f"[UndoScheduleSnapshotCommand] Error ({action}): {e}")

    def _sync_excel_from_map(self, target_map: dict):
        """Sincroniza Excel agrupando días consecutivos con el mismo status en runs."""
        widget = self._widget

        def norm_tuple(rec):
            if not rec:
                return (None, None, None, None)
            st = rec.get("status") or None
            return (st, rec.get("shift_type"), rec.get("in_time"), rec.get("out_time"))

        # Agrupar en runs de mismo (status, shift_type, in_time, out_time)
        runs      = []
        curr_t    = None
        run_start = None
        d = self._snapshot_start

        while d <= self._snapshot_end:
            t = norm_tuple(target_map.get(d.isoformat()))
            if curr_t is None:
                curr_t    = t
                run_start = d
            elif t != curr_t:
                runs.append((run_start, d - timedelta(days=1), curr_t))
                curr_t    = t
                run_start = d
            d += timedelta(days=1)

        if curr_t is not None and run_start is not None:
            runs.append((run_start, self._snapshot_end, curr_t))

        for rs, re_, (st, shift_type, in_time, out_time) in runs:
            excel.update_plan_staff_excel(
                widget.excel_file,
                self._username, self._role, self._badge,
                st, shift_type,
                rs, re_,
                widget.source,
                in_time, out_time,
            )
