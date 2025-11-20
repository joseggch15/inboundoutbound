# main.py
import sys
import os
import time
from PyQt6.QtWidgets import QApplication, QDialog, QWidget, QVBoxLayout, QPushButton, QLabel, QMessageBox
from PyQt6.QtCore import Qt, QTimer, QThread, pyqtSignal

from main_window import MainWindow, AdminMainWindow
from ui_login import LoginWindow, LoadingWindow
from ui.theme import apply_app_theme

# Importamos la lógica para ejecutarla en el hilo de carga
import database_logic as db
import excel_logic as excel

class AppInitializer(QThread):
    """
    Hilo de trabajo (Backend Thread).
    Realiza la carga de datos pesados (BD, Validación Excel) sin congelar la UI.
    Se detiene al 80% para permitir que el Hilo Principal construya la interfaz gráfica.
    """
    progress_updated = pyqtSignal(int, str)
    finished_success = pyqtSignal()
    finished_error = pyqtSignal(str)

    def __init__(self, excel_file, source):
        super().__init__()
        self.excel_file = excel_file
        self.source = source

    def run(self):
        try:
            # PASO 1: Base de Datos (10-20%)
            self.progress_updated.emit(10, "Connecting to secure database...")
            time.sleep(0.1) # Pequeña pausa técnica
            db.setup_database()
            
            # PASO 2: Integridad de Datos (30%)
            self.progress_updated.emit(30, "Verifying SSoT integrity...")
            # Forzamos una lectura a la BD para "calentar" la conexión y caché
            db.get_all_users(self.source)
            
            # PASO 3: Validación de Excel (50-70%)
            self.progress_updated.emit(50, f"Analyzing file: {os.path.basename(self.excel_file)}...")
            
            if self.excel_file and os.path.exists(self.excel_file):
                # Esta función puede tardar si el archivo es grande
                ok, _, _ = excel.validate_excel_structure(self.excel_file)
                if not ok:
                    self.progress_updated.emit(60, "Warning: Structure check flagged issues.")
                else:
                    self.progress_updated.emit(70, "Excel structure verified.")
            else:
                 self.progress_updated.emit(50, "File not found. Skipping validation.")

            # PASO 4: Preparación final del Backend (80%)
            self.progress_updated.emit(80, "Preparing User Interface...")
            time.sleep(0.2)
            
            # Aquí terminamos el trabajo de fondo. 
            # El 20% restante es construcción de GUI en el Main Thread.
            self.finished_success.emit()

        except Exception as e:
            self.finished_error.emit(str(e))


class LauncherWindow(QWidget):
    """
    Initial window with a single Log In button.
    """
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Inbound - Outbound PLG")
        self.setMinimumSize(420, 220)
        self.main_app_window = None
        self._login_payload = None
        self._loading_window = None
        self._initializer_thread = None

        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.setSpacing(12)
        layout.setContentsMargins(24, 24, 24, 24)

        title = QLabel("Transport & Operations Manager")
        font = title.font(); font.setPointSize(20); font.setBold(True)
        title.setFont(font)

        login_button = QPushButton("Log In")
        login_button.setFixedSize(220, 48)
        login_button.setProperty("variant", "primary")

        layout.addWidget(title, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addSpacing(20)
        layout.addWidget(login_button, alignment=Qt.AlignmentFlag.AlignCenter)

        self.setLayout(layout)
        login_button.clicked.connect(self.start_login_process)

    def start_login_process(self):
        """
        Handles the sign-in flow and starts the real loading process.
        """
        login_dialog = LoginWindow(self)
        if login_dialog.exec() == QDialog.DialogCode.Accepted:
            # 1. Guardar credenciales
            self._login_payload = {
                "user_role": login_dialog.user_role,
                "excel_file": login_dialog.excel_file,
                "logged_username": login_dialog.username,
                "can_manage_shift_types": bool(getattr(login_dialog, "can_manage_shift_types", False)),
            }

            # 2. Ocultar Launcher
            self.hide()

            # 3. Mostrar Splash Screen (LoadingWindow optimizada)
            self._loading_window = LoadingWindow(role=self._login_payload["user_role"])
            self._loading_window.show()

            # 4. Iniciar carga real en segundo plano
            self.start_app_initialization()

    def start_app_initialization(self):
        """Configura e inicia el hilo de inicialización."""
        role = self._login_payload["user_role"]
        excel_file = self._login_payload["excel_file"]
        # Si es admin, usamos RGM por defecto para validaciones iniciales
        source = role if role in ["RGM", "Newmont"] else "RGM"

        self._initializer_thread = AppInitializer(excel_file, source)
        
        # Conectar señales
        self._initializer_thread.progress_updated.connect(self._loading_window.update_progress)
        self._initializer_thread.finished_success.connect(self._on_backend_ready)
        self._initializer_thread.finished_error.connect(self._on_initialization_error)
        
        # Arrancar hilo
        self._initializer_thread.start()

    def _on_backend_ready(self):
        """
        El Backend terminó (80%). Ahora construimos la UI pesada (20%)
        MANTENIENDO la ventana de carga visible para evitar 'congelamientos'.
        """
        try:
            p = self._login_payload
            role = p.get("user_role")

            # --- FASE 1: Instanciación Pesada (80% -> 90%) ---
            self._loading_window.update_progress(85, "Building Dashboard components...")
            # Forzamos el redibujado para que el usuario vea el cambio de texto
            QApplication.processEvents() 
            
            if role == "Administrator":
                self.main_app_window = AdminMainWindow(
                    logged_username=p.get("logged_username") or "",
                    rgm_excel="PlanStaffRGM.xlsx",
                    newmont_excel="PlanStaffNewmont.xlsx",
                )
            else:
                self.main_app_window = MainWindow(
                    user_role=role,
                    excel_file=p.get("excel_file") or "",
                    logged_username=p.get("logged_username") or "",
                    can_manage_shift_types=p.get("can_manage_shift_types", False),
                )
            
            # Conectar logout
            self.main_app_window.logout_signal.connect(self.handle_logout)

            # --- FASE 2: Finalización (90% -> 100%) ---
            self._loading_window.update_progress(100, "Starting application...")
            
            # Pequeño retraso para que el usuario vea el 100% antes del switch
            QTimer.singleShot(200, self._finish_loading_sequence)

        except Exception as e:
            self._on_initialization_error(f"UI Build Failed: {e}")

    def _finish_loading_sequence(self):
        """Intercambio limpio de ventanas."""
        if self.main_app_window:
            self.main_app_window.show() # Mostrar ventana principal
        
        if self._loading_window:
            self._loading_window.close() # Cerrar splash solo ahora
            self._loading_window = None

    def _on_initialization_error(self, error_msg):
        """Manejo de errores fatales durante la carga."""
        if self._loading_window:
            self._loading_window.close()
        
        QMessageBox.critical(self, "Initialization Error", f"Critical error during startup:\n{error_msg}")
        self.show() # Volver al launcher

    def handle_logout(self):
        """Shows the launcher again after signing out."""
        self.main_app_window = None
        self.show()


if __name__ == '__main__':
    # Qt6 ya maneja HiDPI por defecto; solo ajustamos la política de redondeo si está disponible
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