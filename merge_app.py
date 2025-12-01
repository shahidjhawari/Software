# QA LAB Merge Application — Professional UI
# Developer: Shahid Iqbal — © All Rights Reserved

import sys
import os
import pandas as pd
from PyQt5.QtGui import QIcon, QFont
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QFileDialog,
    QListWidget, QVBoxLayout, QHBoxLayout, QMessageBox, QFrame
)
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QStyleFactory


class MergeApp(QWidget):
    def __init__(self):
        super().__init__()

        # ---------------------
        # APP ICON
        # ---------------------
        base_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(base_dir, "usgroup.ico")
        self.setWindowIcon(QIcon(icon_path))

        self.report_files = []
        self.dataentry_file = None

        # ---------------------
        # WINDOW SETTINGS
        # ---------------------
        self.setWindowTitle("QA LAB — Auto Data Enter")
        self.setGeometry(200, 200, 1100, 680)
        QApplication.setStyle(QStyleFactory.create("Fusion"))

        self.setStyleSheet("""
            QWidget {
                background-color: #1e1e1e;
                color: #e8e8e8;
                font-family: Segoe UI;
                font-size: 11pt;
            }
            QLabel {
                font-size: 13pt;
                font-weight: bold;
                padding: 4px;
            }
            QListWidget {
                background: #2b2b2b;
                border: 1px solid #444;
                padding: 8px;
                border-radius: 6px;
            }
            QPushButton {
                background-color: #0066cc;
                color: white;
                padding: 10px 15px;
                border: none;
                border-radius: 6px;
                font-size: 11.5pt;
            }
            QPushButton:hover {
                background-color: #1a75ff;
            }
            QPushButton:pressed {
                background-color: #004c99;
            }
            QFrame {
                background-color: #2a2a2a;
                border: 1px solid #444;
                border-radius: 10px;
                padding: 12px;
            }
        """)

        # Layouts
        main_layout = QHBoxLayout()
        left_box = QFrame()
        right_box = QFrame()

        left_layout = QVBoxLayout()
        right_layout = QVBoxLayout()

        # ---------------------
        # LEFT PANEL
        # ---------------------
        report_title = QLabel("Report Files")
        self.report_list = QListWidget()

        self.btn_add_reports = QPushButton("Add Report Files")
        self.btn_add_reports.clicked.connect(self.load_reports)

        self.btn_clear_reports = QPushButton("Clear List")
        self.btn_clear_reports.clicked.connect(self.clear_reports)

        left_layout.addWidget(report_title)
        left_layout.addWidget(self.report_list)
        left_layout.addWidget(self.btn_add_reports)
        left_layout.addWidget(self.btn_clear_reports)

        left_box.setLayout(left_layout)

        # ---------------------
        # RIGHT PANEL
        # ---------------------
        data_title = QLabel("Development Data Entry File")
        self.data_display = QLabel("No File Selected")
        self.btn_dataentry = QPushButton("Select Data Entry File")
        self.btn_dataentry.clicked.connect(self.load_dataentry)

        right_layout.addWidget(data_title)
        right_layout.addWidget(self.data_display)
        right_layout.addWidget(self.btn_dataentry)

        right_box.setLayout(right_layout)

        # ---------------------
        # MERGE BUTTON
        # ---------------------
        self.btn_merge = QPushButton("Merge & Save Output")
        self.btn_merge.setFixedHeight(50)
        self.btn_merge.setStyleSheet("""
            QPushButton {
                font-size: 13pt;
                background-color: #008000;
                border-radius: 6px;
            }
            QPushButton:hover {
                background-color: #00a300;
            }
            QPushButton:pressed {
                background-color: #006600;
            }
        """)
        self.btn_merge.clicked.connect(self.merge_files)

        # ---------------------
        # FOOTER
        # ---------------------
        footer = QLabel("Developed by Shahid Iqbal — © All Rights Reserved")
        footer.setAlignment(Qt.AlignCenter)
        footer.setStyleSheet("font-size: 10pt; padding: 10px; color: #bbbbbb;")

        # ---------------------
        # ORGANIZE LAYOUT
        # ---------------------
        main_layout.addWidget(left_box, 2)
        main_layout.addWidget(right_box, 1)

        wrapper = QVBoxLayout()
        wrapper.addLayout(main_layout)
        wrapper.addWidget(self.btn_merge)
        wrapper.addWidget(footer)

        self.setLayout(wrapper)

    # -----------------------------
    # Load Reports
    # -----------------------------
    def load_reports(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Report Files", "", "Excel Files (*.xlsx *.xls)")
        if files:
            for f in files:
                if f not in self.report_files:
                    self.report_files.append(f)
                    self.report_list.addItem(f)

    def clear_reports(self):
        self.report_files = []
        self.report_list.clear()

    # -----------------------------
    # Load DataEntry Base File
    # -----------------------------
    def load_dataentry(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select Data Entry File", "", "Excel Files (*.xlsx *.xls)")
        if file:
            self.dataentry_file = file
            self.data_display.setText(file)

    # -----------------------------
    # Extract LAST value after matched name
    # -----------------------------
    def extract_value(self, df, colname):
        colname = str(colname).strip().lower()

        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                cell = str(df.iat[i, j]).strip().lower()

                if cell == colname:
                    values = []
                    for k in range(j + 1, df.shape[1]):
                        v = df.iat[k, j] if pd.notna(df.iat[k, j]) else None
                        if pd.notna(df.iat[i, k]) and str(df.iat[i, k]).strip() != "":
                            values.append(df.iat[i, k])

                    return values[-1] if values else "-"

        return "-"

    # -----------------------------
    # Merge Logic
    # -----------------------------
    def merge_files(self):
        if not self.dataentry_file:
            QMessageBox.warning(self, "Error", "Please select DataEntry file first.")
            return
        
        if not self.report_files:
            QMessageBox.warning(self, "Error", "Please upload at least one Report file.")
            return

        df_base = pd.read_excel(self.dataentry_file)
        base_columns = list(df_base.columns)

        appended_rows = []

        for fpath in self.report_files:
            df_r = pd.read_excel(fpath, header=None)

            row = {}
            for col in base_columns:
                row[col] = self.extract_value(df_r, col)

            appended_rows.append(row)

        df_append = pd.DataFrame(appended_rows, columns=base_columns)
        df_final = pd.concat([df_base, df_append], ignore_index=True)

        save_path, _ = QFileDialog.getSaveFileName(self, "Save Merged File", "", "Excel Files (*.xlsx)")
        if not save_path:
            return
        if not save_path.lower().endswith(".xlsx"):
            save_path += ".xlsx"

        df_final.to_excel(save_path, index=False)
        QMessageBox.information(self, "Success", "Merged file saved successfully!")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MergeApp()
    window.show()
    sys.exit(app.exec_())
