# widgets/report_settings_widget.py
# Extracted from main_window.py — ReportSettingsWidget

import json
import logging
from PyQt6.QtWidgets import (
    QWidget, QGridLayout, QHBoxLayout, QLabel, QLineEdit,
    QPushButton, QTableWidget, QTableWidgetItem, QHeaderView,
    QMessageBox, QColorDialog, QAbstractItemView, QFontComboBox,
)
from PyQt6.QtGui import QColor, QFont

import database_logic as db
from constants import RGM_REPORT_HEADERS, NEWMONT_REPORT_HEADERS, Source

logger = logging.getLogger(__name__)


class ReportSettingsWidget(QWidget):
    def __init__(self, username: str, source: str):
        super().__init__()
        self.username = username
        self.source = source
        self._color_buttons = {}

        layout = QGridLayout(self)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setHorizontalSpacing(15)
        layout.setVerticalSpacing(10)

        # Font Name
        self.font_name_combo = QFontComboBox()
        layout.addWidget(QLabel("Font Name:"), 0, 0)
        layout.addWidget(self.font_name_combo, 0, 1)

        # Font Color
        self.font_color_btn = self._create_color_button("font_color")
        layout.addWidget(QLabel("Body Font Color:"), 1, 0)
        layout.addWidget(self.font_color_btn, 1, 1)

        # Date Format
        self.date_format_input = QLineEdit()
        self.date_format_input.setPlaceholderText("e.g., dd/mm/yyyy, yyyy-mm-dd, etc.")
        layout.addWidget(QLabel("Report Date Format:"), 2, 0)
        layout.addWidget(self.date_format_input, 2, 1)

        # Header BG Color
        self.header_bg_btn = self._create_color_button("header_bg_color")
        layout.addWidget(QLabel("Default Header Background:"), 3, 0)
        layout.addWidget(self.header_bg_btn, 3, 1)

        # Header Font Color
        self.header_font_btn = self._create_color_button("header_font_color")
        layout.addWidget(QLabel("Header Font Color:"), 4, 0)
        layout.addWidget(self.header_font_btn, 4, 1)

        # Column specific colors
        self.col_table = QTableWidget(0, 2)
        self.col_table.setHorizontalHeaderLabels(["Column Name", "Header Color"])
        self.col_table.horizontalHeader().setSectionResizeMode(
            0, QHeaderView.ResizeMode.Stretch
        )
        self.col_table.horizontalHeader().setSectionResizeMode(
            1, QHeaderView.ResizeMode.ResizeToContents
        )
        self.col_table.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.col_table.cellDoubleClicked.connect(self._pick_column_color)
        layout.addWidget(
            QLabel("Column-specific Header Colors (double-click to change):"),
            5, 0, 1, 2,
        )
        layout.addWidget(self.col_table, 6, 0, 1, 2)

        # Action Buttons
        self.save_btn = QPushButton("Save Report Settings")
        self.save_btn.setProperty("variant", "primary")
        self.save_btn.clicked.connect(self.save_settings)

        self.reset_btn = QPushButton("Reset to Default")
        self.reset_btn.setProperty("variant", "secondary")
        self.reset_btn.clicked.connect(self._reset_settings)

        button_layout = QHBoxLayout()
        button_layout.addStretch()
        button_layout.addWidget(self.reset_btn)
        button_layout.addWidget(self.save_btn)
        button_layout.addStretch()

        layout.addLayout(button_layout, 7, 0, 1, 2)

        self.load_settings()

    def _create_color_button(self, key):
        btn = QPushButton("Pick Color...")
        btn.setProperty("key", key)
        btn.clicked.connect(self._pick_color)
        self._color_buttons[key] = btn
        return btn

    def _update_button_color(self, btn, color_hex):
        btn.setText(color_hex.upper())
        btn.setStyleSheet(
            f"background-color: {color_hex}; color: {'#FFFFFF' if QColor(color_hex).lightness() < 128 else '#000000'};"
        )

    def _pick_color(self):
        sender_btn = self.sender()
        key = sender_btn.property("key")
        current_color = sender_btn.text()
        color = QColorDialog.getColor(QColor(current_color), self, "Pick a Color")
        if color.isValid():
            self._update_button_color(sender_btn, color.name())

    def _pick_column_color(self, row, column):
        if column != 1:
            return
        header_item = self.col_table.item(row, 0)
        color_item = self.col_table.item(row, 1)
        if not header_item or not color_item:
            return

        current_color_hex = color_item.text()
        color = QColorDialog.getColor(
            QColor(current_color_hex), self, f"Color for {header_item.text()}"
        )
        if color.isValid():
            color_item.setText(color.name())
            color_item.setBackground(color)
            color_item.setForeground(
                QColor("#FFFFFF" if color.lightness() < 128 else "#000000")
            )

    def load_settings(self):
        settings = db.get_report_settings(self.username, self.source)
        self.font_name_combo.setCurrentFont(QFont(settings["font_name"]))
        self._update_button_color(self.font_color_btn, settings["font_color"])
        self.date_format_input.setText(settings.get("date_format", "dd/mm/yyyy"))
        self._update_button_color(self.header_bg_btn, settings["header_bg_color"])
        self._update_button_color(self.header_font_btn, settings["header_font_color"])

        headers = RGM_REPORT_HEADERS if self.source == Source.RGM.value else NEWMONT_REPORT_HEADERS
        self.col_table.setRowCount(0)
        col_colors = settings.get("column_colors", {})
        for header in headers:
            row_pos = self.col_table.rowCount()
            self.col_table.insertRow(row_pos)
            self.col_table.setItem(row_pos, 0, QTableWidgetItem(header))
            color_hex = col_colors.get(header, settings["header_bg_color"])
            color_item = QTableWidgetItem(color_hex)
            color_item.setBackground(QColor(color_hex))
            color_item.setForeground(
                QColor("#FFFFFF" if QColor(color_hex).lightness() < 128 else "#000000")
            )
            self.col_table.setItem(row_pos, 1, color_item)

    def save_settings(self):
        col_colors = {}
        for row in range(self.col_table.rowCount()):
            header = self.col_table.item(row, 0).text()
            color = self.col_table.item(row, 1).text()
            col_colors[header] = color

        settings = {
            "font_name": self.font_name_combo.currentFont().family(),
            "font_color": self.font_color_btn.text(),
            "date_format": self.date_format_input.text().strip(),
            "header_bg_color": self.header_bg_btn.text(),
            "header_font_color": self.header_font_btn.text(),
            "column_colors": col_colors,
        }
        db.save_report_settings(self.username, self.source, settings)
        db.log_event(
            self.username,
            self.source,
            "SETTINGS_UPDATE",
            f"Report layout changed: {json.dumps(settings)}",
        )
        QMessageBox.information(
            self,
            "Settings Saved",
            "Report layout settings have been saved successfully.",
        )

    def _reset_settings(self):
        reply = QMessageBox.question(
            self,
            "Confirm Reset",
            "Are you sure you want to reset the report settings for this profile to their defaults?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )

        if reply == QMessageBox.StandardButton.Yes:
            db.delete_report_settings(self.username, self.source)
            self.load_settings()
            QMessageBox.information(
                self,
                "Settings Reset",
                "Report layout settings have been reset to default.",
            )
