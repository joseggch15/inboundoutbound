# main.py
import sys
import time
import pandas as pd # Necesario para el tipado
from PyQt6.QtWidgets import QApplication, QDialog, QWidget, QVBoxLayout, QPushButton, QLabel, QMessageBox
from PyQt6.QtCore import Qt, QTimer, QThread, pyqtSignal, QObject

from main_window import MainWindow, AdminMainWindow
from ui_login import LoginWindow, LoadingWindow
from ui.theme import apply_app_theme
import database_logic as db
import excel_logic as excel

# ---------------------------------------------------------
# CLASE WORKER PARA CARGA EN SEGUNDO PLANO
# ---------------------------------------------------------
class StartupWorker(QThread):
    """
    Realiza las tareas pesadas de inicialización (BD, Excel I/O)
    fuera del hilo principal para que la UI no se congele.
    """
    progress_updated = pyqtSignal(int, str)
    finished_success = pyqtSignal(object) # Retorna payload de datos
    finished_error = pyqtSignal(str)

    def __init__(self, excel_file, source):
        super().__init__()
        self.excel_file = excel_file
        self.source = source

    def run(self):
        try:
            # Pasa 1: Setup de Base de Datos (10%)
            self.progress_updated.emit(10, "Initializing Database (SSoT)...")
            db.setup_database()
            time.sleep(0.3) # Pequeña pausa estética para que el usuario lea

            # Paso 2: Validación de estructura Excel (30%)
            if self.excel_file:
                self.progress_updated.emit(30, f"Validating structure: {self.excel_file}...")
                exists = excel.os.path.exists(self.excel_file)
                if not exists:
                    # No fallamos aquí, dejamos que la UI maneje el archivo faltante, 
                    # pero reportamos progreso.
                    pass 
                else:
                    # Validación real (IO costoso)
                    ok, errs, meta = excel.validate_excel_structure(self.excel_file)
                    if not ok:
                        # Opcional: Podríamos abortar, pero mejor dejamos que la app abra y muestre errores
                        pass

            # Paso 3: Pre-carga de datos pesados (Schedule Preview) (70%)
            # Esto es lo que normalmente congela la UI al iniciar MainWindow
            self.progress_updated.emit(50, "Reading heavy schedule data...")
            preloaded_data = None
            if self.excel_file and excel.os.path.exists(self.excel_file):
                # Leemos el DataFrame AQUÍ en el hilo secundario
                df_preview = excel.get_schedule_preview(self.excel_file)
                preloaded_data = df_preview
            
            self.progress_updated.emit(90, "Finalizing UI components...")
            time.sleep(0.2)
            
            # Paso 4: Finalización (100%)
            self.progress_updated.emit(100, "Ready.")
            
            # Retornamos los datos pre-cargados
            payload = {
                "schedule_df": preloaded_data
            }
            self.finished_success.emit(payload)

        except Exception as e:
            self.finished_error.emit(str(e))

# ---------------------------------------------------------
# LAUNCHER MODIFICADO
# ---------------------------------------------------------
class LauncherWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Inbound - Outbound PLG")
        # ... (Tu código de UI existente se mantiene igual) ...
        self.setMinimumSize(420, 220)
        self.main_app_window = None
        self._login_payload = None
        self._loading = None
        
        # UI Setup (Idéntico a tu código original)
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(12)
        layout.setContentsMargins(24, 24, 24, 24)
        title = QLabel("Transport & Operations Manager")
        font = title.font(); font.setPointSize(20); font.setBold(True)
        title.setFont(font)
        login_button = QPushButton("Sign In")
        login_button.setFixedSize(220, 48)
        login_button.setProperty("variant", "primary")
        layout.addWidget(title, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addSpacing(20)
        layout.addWidget(login_button, alignment=Qt.AlignmentFlag.AlignCenter)
        self.setLayout(layout)
        login_button.clicked.connect(self.start_login_process)

    def start_login_process(self):
        login_dialog = LoginWindow(self)
        if login_dialog.exec() == QDialog.DialogCode.Accepted:
            self._login_payload = {
                "user_role": login_dialog.user_role,
                "excel_file": login_dialog.excel_file,
                "logged_username": login_dialog.username,
                "can_manage_shift_types": bool(getattr(login_dialog, "can_manage_shift_types", False)),
            }
            self.hide()
            
            # INICIO DEL PROCESO REAL DE CARGA
            self._start_loading_sequence()

    def _start_loading_sequence(self):
        role = self._login_payload["user_role"]
        excel_file = self._login_payload["excel_file"]

        # 1. Mostrar pantalla de carga
        self._loading = LoadingWindow(role=role)
        self._loading.show()

        # 2. Iniciar Worker Thread
        self.worker = StartupWorker(excel_file, role)
        self.worker.progress_updated.connect(self._loading.update_progress)
        self.worker.finished_success.connect(self._on_loading_finished)
        self.worker.finished_error.connect(self._on_loading_error)
        self.worker.start()

    def _on_loading_finished(self, payload):
        # El hilo terminó, los datos pesados están en 'payload'
        # Cerramos pantalla de carga
        if self._loading:
            self._loading.close()
            self._loading = None
        
        # Abrimos la ventana principal inyectando los datos (si es posible)
        self._open_main_window(payload)

    def _on_loading_error(self, error_msg):
        if self._loading:
            self._loading.close()
        QMessageBox.critical(self, "Startup Error", f"Failed to initialize:\n{error_msg}")
        self.show() # Volver al launcher

    def _open_main_window(self, preloaded_payload):
        p = self._login_payload or {}
        role = p.get("user_role")
        
        # Extraemos el DataFrame pre-cargado
        schedule_df = preloaded_payload.get("schedule_df")

        if role == "Administrator":
            self.main_app_window = AdminMainWindow(
                logged_username=p.get("logged_username") or "",
                rgm_excel="PlanStaffRGM.xlsx",
                newmont_excel="PlanStaffNewmont.xlsx",
            )
        else:
            # Aquí pasamos el DF pre-cargado al MainWindow
            self.main_app_window = MainWindow(
                user_role=role,
                excel_file=p.get("excel_file") or "",
                logged_username=p.get("logged_username") or "",
                can_manage_shift_types=p.get("can_manage_shift_types", False),
                preloaded_data=schedule_df  # <--- INYECCIÓN DE DEPENDENCIA
            )

        self.main_app_window.logout_signal.connect(self.handle_logout)
        self.main_app_window.show()

    def handle_logout(self):
        self.main_app_window = None
        self.show()

if __name__ == '__main__':
    # ... (Resto del código igual) ...
    try:
        QApplication.setHighDpiScaleFactorRoundingPolicy(
            Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
        )
    except Exception:
        pass
    app = QApplication(sys.argv)
    app.setApplicationName("Inbound - Outbound PLG")
    apply_app_theme(app)
    launcher = LauncherWindow()
    launcher.show()
    sys.exit(app.exec())