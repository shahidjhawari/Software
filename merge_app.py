import sys
import os
import pandas as pd
from openpyxl import load_workbook
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QFileDialog,
    QListWidget, QVBoxLayout, QHBoxLayout, QMessageBox, QFrame
)
from PyQt5.QtWidgets import QStyleFactory


class MergeApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("QA LAB — Auto Data Entry Software")
        self.setGeometry(200, 200, 1100, 680)
        QApplication.setStyle(QStyleFactory.create("Fusion"))

        self.report_files = []
        self.dataentry_file = None

        layout = QVBoxLayout(self)

        self.report_list = QListWidget()
        btn_reports = QPushButton("Add Report Files")
        btn_reports.clicked.connect(self.load_reports)

        btn_data = QPushButton("Select Data Entry File")
        btn_data.clicked.connect(self.load_dataentry)

        self.lbl_data = QLabel("No Data Entry File Selected")

        btn_merge = QPushButton("MERGE & SAVE")
        btn_merge.setFixedHeight(45)
        btn_merge.clicked.connect(self.merge_files)

        layout.addWidget(self.report_list)
        layout.addWidget(btn_reports)
        layout.addWidget(self.lbl_data)
        layout.addWidget(btn_data)
        layout.addWidget(btn_merge)

    # ------------------------------------------------
    def load_reports(self):
        files, _ = QFileDialog.getOpenFileNames(self, "", "", "Excel Files (*.xlsx *.xls)")
        for f in files:
            if f not in self.report_files:
                self.report_files.append(f)
                self.report_list.addItem(f)

    def load_dataentry(self):
        file, _ = QFileDialog.getOpenFileName(self, "", "", "Excel Files (*.xlsx *.xls)")
        if file:
            self.dataentry_file = file
            self.lbl_data.setText(os.path.basename(file))

    # ------------------------------------------------
    def last_value(self, row):
        for cell in reversed(row):
            val = str(cell).strip()
            if val and val.lower() != "nan":
                return val
        return ""

    # ------------------------------------------------
    def extract_dynamic(self, df, columns):
        extracted = {}

        for i in range(len(df)):
            row = df.iloc[i].fillna("").astype(str)
            text = " ".join(row).lower()

            for col in columns:
                key = col.lower()

                if key in text and col not in extracted:
                    value = self.last_value(row)

                    # اگر اسی row میں value نہ ہو تو اگلی row دیکھو
                    if value.lower() == key and i + 1 < len(df):
                        value = self.last_value(
                            df.iloc[i + 1].fillna("").astype(str)
                        )

                    extracted[col] = value

        return extracted

    # ------------------------------------------------
    def merge_files(self):
        if not self.dataentry_file or not self.report_files:
            QMessageBox.warning(self, "Error", "Please select files first")
            return

        base_df = pd.read_excel(self.dataentry_file)
        headers = list(base_df.columns)

        wb = load_workbook(self.dataentry_file)
        ws = wb.active

        for rpt in self.report_files:
            df = pd.read_excel(rpt, header=None)
            data = self.extract_dynamic(df, headers)

            new_row = []
            for col in headers:
                new_row.append(data.get(col, ""))

            ws.append(new_row)

        save_path, _ = QFileDialog.getSaveFileName(self, "", "", "Excel Files (*.xlsx)")
        if save_path:
            wb.save(save_path)
            QMessageBox.information(self, "Success", "All values auto-extracted successfully!")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MergeApp()
    win.show()
    sys.exit(app.exec_())
