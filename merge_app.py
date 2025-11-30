# QA LAB Merge Application — Flexible Report Reader (Final Updated Logic)
# Developer: Shahid Iqbal — © All Rights Reserved

import sys
import os
import pandas as pd
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QFileDialog,
    QListWidget, QVBoxLayout, QHBoxLayout, QMessageBox
)
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QStyleFactory


class MergeApp(QWidget):
    def __init__(self):
        super().__init__()

        # Permanent Software Icon
        self.app_icon_path = "/mnt/data/30a1bb27-cbe8-4041-9ee1-1efe58fdf9a7.png"
        self.setWindowIcon(QIcon(self.app_icon_path))

        self.report_files = []
        self.dataentry_file = None

        self.setWindowTitle("QA LAB Data Merge Software — Developed by Shahid Iqbal")
        self.setGeometry(200, 200, 1050, 650)

        QApplication.setStyle(QStyleFactory.create("Fusion"))

        main_layout = QHBoxLayout()
        left_layout = QVBoxLayout()
        right_layout = QVBoxLayout()

        # Left side
        self.report_label = QLabel("Upload Report Files:")
        self.report_list = QListWidget()
        self.btn_add_reports = QPushButton("Add Reports")
        self.btn_add_reports.clicked.connect(self.load_reports)
        self.btn_clear_reports = QPushButton("Clear Reports")
        self.btn_clear_reports.clicked.connect(self.clear_reports)

        left_layout.addWidget(self.report_label)
        left_layout.addWidget(self.report_list)
        left_layout.addWidget(self.btn_add_reports)
        left_layout.addWidget(self.btn_clear_reports)

        # Right side
        self.data_label = QLabel("Select Base DataEntry File:")
        self.data_display = QLabel("No File Selected")
        self.btn_dataentry = QPushButton("Select DataEntry File")
        self.btn_dataentry.clicked.connect(self.load_dataentry)

        right_layout.addWidget(self.data_label)
        right_layout.addWidget(self.data_display)
        right_layout.addWidget(self.btn_dataentry)

        # Merge button
        self.btn_merge = QPushButton("Merge & Save Output")
        self.btn_merge.setFixedHeight(45)
        self.btn_merge.clicked.connect(self.merge_files)

        main_layout.addLayout(left_layout, 2)
        main_layout.addLayout(right_layout, 1)

        wrapper = QVBoxLayout()
        wrapper.addLayout(main_layout)
        wrapper.addWidget(self.btn_merge)

        dev_label = QLabel("Developed by Shahid Iqbal — © All Rights Reserved")
        dev_label.setAlignment(Qt.AlignCenter)
        wrapper.addWidget(dev_label)

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
    # Extract LAST value after the matched name
    # -----------------------------
    def extract_value(self, df, colname):
        colname = str(colname).strip().lower()

        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                cell = str(df.iat[i, j]).strip().lower()

                # Column name matched
                if cell == colname:

                    # Collect ALL values in right-side columns
                    values = []
                    for k in range(j + 1, df.shape[1]):
                        val = df.iat[i, k]
                        if pd.notna(val) and str(val).strip() != "":
                            values.append(val)

                    # If multiple values found → return LAST value
                    if values:
                        return values[-1]

                    return "-"

        return "-"  # Not found at all

    # -----------------------------
    # Merge Logic (FINAL)
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

            df_r = pd.read_excel(fpath, header=None)  # Flexible format

            row = {}
            for col in base_columns:
                val = self.extract_value(df_r, col)
                row[col] = val

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
