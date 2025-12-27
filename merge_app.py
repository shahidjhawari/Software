import sys
import os
import pandas as pd
from openpyxl import load_workbook
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel,
    QFileDialog, QListWidget, QVBoxLayout, QMessageBox
)
from PyQt5.QtWidgets import QStyleFactory


# ---------------- HELPER ----------------
def is_number(x):
    try:
        float(x)
        return True
    except:
        return False


def last_numeric(row):
    for v in reversed(row):
        if is_number(v):
            return v
    return ""


# ---------------- APP ----------------
class MergeApp(QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("QA LAB — Accurate Auto Entry")
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

    # ---------------- FILE LOAD ----------------
    def load_reports(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "", "", "Excel Files (*.xlsx *.xls)"
        )
        for f in files:
            if f not in self.report_files:
                self.report_files.append(f)
                self.report_list.addItem(f)

    def load_dataentry(self):
        file, _ = QFileDialog.getOpenFileName(
            self, "", "", "Excel Files (*.xlsx *.xls)"
        )
        if file:
            self.dataentry_file = file
            self.lbl_data.setText(os.path.basename(file))

    # ---------------- CORE EXTRACTION ----------------
    def extract_report(self, df):
        df = df.fillna("").astype(str)
        result = {}
        current_test = ""

        for i in range(len(df)):
            row = df.iloc[i]
            text = " ".join(row).lower()

            # Detect test block
            if "tear strength" in text:
                current_test = "Tear"
            elif "tensile strength" in text:
                current_test = "Tensile"
            elif "color fastness to rubbing" in text:
                current_test = "Rubbing"
            elif "shade change" in text:
                current_test = "Shade Change"
            elif "staining" in text:
                current_test = "Staining"
            elif "ph value" in text:
                current_test = "pH"
            elif text.strip() == "temp":
                current_test = "Temp"

            # Warp / Weft rows
            if "warp" in text:
                val = last_numeric(row)
                if current_test:
                    result[f"{current_test} Warp"] = val

            if "weft" in text:
                val = last_numeric(row)
                if current_test:
                    result[f"{current_test} Weft"] = val

            # Single value tests
            if current_test in ["Shade Change", "Staining", "pH", "Temp"]:
                val = last_numeric(row)
                if val:
                    result[current_test] = val

        return result

    # ---------------- MERGE ----------------
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
            extracted = self.extract_report(df)

            new_row = []
            for h in headers:
                new_row.append(extracted.get(h, ""))

            ws.append(new_row)

        save_path, _ = QFileDialog.getSaveFileName(
            self, "", "", "Excel Files (*.xlsx)"
        )
        if save_path:
            wb.save(save_path)
            QMessageBox.information(
                self, "Success",
                "QA Report extracted EXACTLY as per structure ✔"
            )


# ---------------- RUN ----------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MergeApp()
    win.show()
    sys.exit(app.exec_())
