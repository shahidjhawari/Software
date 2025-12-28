import sys
import os
import pandas as pd
from openpyxl import load_workbook
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel,
    QFileDialog, QListWidget, QVBoxLayout, QMessageBox
)
from PyQt5.QtWidgets import QStyleFactory


# ---------------- HELPERS ----------------
def is_number(x):
    try:
        float(x)
        return True
    except:
        return False


def clean(val):
    return str(val).strip()


def next_value(row, start_idx):
    for j in range(start_idx + 1, len(row)):
        if str(row[j]).strip():
            return clean(row[j])
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
        rubbing_found = {"Dry": False, "Wet": False}

        for i in range(len(df)):
            row = df.iloc[i]
            row_lower = [c.lower().strip() for c in row]
            text = " ".join(row_lower)

            # -------- HEADER --------
            for idx, cell in enumerate(row_lower):

                if cell == "date":
                    result["Date"] = next_value(row, idx)

                elif cell == "customer":
                    result["Customer"] = next_value(row, idx)

                elif "order#" in cell or cell == "order":
                    result["Order#"] = next_value(row, idx)

                elif "fabric code" in cell:
                    result["Fabric Code"] = next_value(row, idx)

                elif "sample status" in cell:
                    result["Sample Status"] = next_value(row, idx)

                elif cell == "article":
                    result["Article"] = next_value(row, idx)

                elif "wash ref" in cell:
                    result["Wash ref"] = next_value(row, idx)

                elif cell == "reference":
                    result["Reference"] = next_value(row, idx)

                elif cell == "remarks":
                    result["Remarks"] = next_value(row, idx)

            # -------- WEIGHT (Column 9 FIXED) --------
            if "weight" in row_lower:
                if len(row) > 9 and is_number(row[9]):
                    result["Weight"] = row[9]

            # -------- TEST BLOCK --------
            if "tear strength" in text:
                current_test = "Tear"

            elif "tensile strength" in text:
                current_test = "Tensile"

            elif "color fastness to rubbing" in text:
                current_test = "Rubbing"
                rubbing_found = {"Dry": False, "Wet": False}

            elif "color fastness to home laundering" in text:
                current_test = "Home Laundering"

            # -------- Tear / Tensile --------
            if "warp" in text and current_test in ["Tear", "Tensile"]:
                if is_number(row[9]):
                    result[f"{current_test} Warp"] = row[9]

            if "weft" in text and current_test in ["Tear", "Tensile"]:
                if is_number(row[9]):
                    result[f"{current_test} Weft"] = row[9]

            # -------- Rubbing --------
            if current_test == "Rubbing":
                if "dry" in text and is_number(row[9]):
                    result["Rubbing Dry"] = row[9]
                    rubbing_found["Dry"] = True

                if "wet" in text and is_number(row[9]):
                    result["Rubbing Wet"] = row[9]
                    rubbing_found["Wet"] = True

            # -------- Home Laundering --------
            if current_test == "Home Laundering":
                if "shade change" in text and is_number(row[9]):
                    result["Shade Change"] = row[9]

                if "staining" in text and is_number(row[9]):
                    result["Staining"] = row[9]

            # -------- pH (Column 9 FIXED) --------
            if "ph value" in row_lower:
                if len(row) > 9 and is_number(row[9]):
                    result["pH"] = row[9]

            # -------- TEMP --------
            if "temp" in row_lower:
                if len(row) > 9 and is_number(row[9]):
                    result["Temp"] = row[9]

        if "Rubbing Dry" in result and "Rubbing Wet" not in result:
            result["Rubbing Wet"] = "-"

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
            ws.append([extracted.get(h, "") for h in headers])

        save_path, _ = QFileDialog.getSaveFileName(
            self, "", "", "Excel Files (*.xlsx)"
        )
        if save_path:
            wb.save(save_path)
            QMessageBox.information(
                self, "Success",
                "Weight + pH + Temp values extracted correctly ✔"
            )


# ---------------- RUN ----------------
if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MergeApp()
    win.show()
    sys.exit(app.exec_())
