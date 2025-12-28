import sys
import os
import pandas as pd
from openpyxl import load_workbook
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel,
    QFileDialog, QListWidget, QVBoxLayout, QHBoxLayout,
    QMessageBox, QFrame
)
from PyQt5.QtWidgets import QStyleFactory
from PyQt5.QtCore import Qt


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

        self.setStyleSheet("""
            QWidget {
                background-color: #0f172a;
                color: #e5e7eb;
                font-family: Segoe UI;
                font-size: 14px;
            }
            QLabel#title {
                font-size: 26px;
                font-weight: 600;
                color: #38bdf8;
                padding: 12px;
            }
            QPushButton {
                background-color: #1e293b;
                border: 1px solid #334155;
                border-radius: 8px;
                padding: 10px;
                font-size: 14px;
            }
            QPushButton:hover {
                background-color: #334155;
            }
            QPushButton#merge {
                background-color: #0284c7;
                font-size: 16px;
                font-weight: bold;
                height: 48px;
            }
            QPushButton#merge:hover {
                background-color: #0369a1;
            }
            QListWidget {
                background-color: #020617;
                border: 1px solid #334155;
                border-radius: 8px;
                padding: 6px;
            }
            QLabel#footer {
                font-size: 12px;
                color: #94a3b8;
                padding: 10px;
            }
        """)

        main_layout = QVBoxLayout(self)
        main_layout.setSpacing(14)

        title = QLabel("QA Laboratory Report Merger")
        title.setObjectName("title")
        title.setAlignment(Qt.AlignCenter)

        card = QFrame()
        card.setStyleSheet("""
            QFrame {
                background-color: #020617;
                border-radius: 12px;
                padding: 16px;
            }
        """)

        card_layout = QVBoxLayout(card)
        card_layout.setSpacing(12)

        self.report_list = QListWidget()

        btn_reports = QPushButton("➕ Add Report Files")
        btn_reports.clicked.connect(self.load_reports)

        btn_data = QPushButton("📄 Select Data Entry File")
        btn_data.clicked.connect(self.load_dataentry)

        self.lbl_data = QLabel("No Data Entry File Selected")
        self.lbl_data.setStyleSheet("color:#38bdf8; padding:6px;")

        btn_merge = QPushButton("MERGE & SAVE")
        btn_merge.setObjectName("merge")
        btn_merge.clicked.connect(self.merge_files)

        card_layout.addWidget(QLabel("Selected Report Files"))
        card_layout.addWidget(self.report_list)
        card_layout.addWidget(btn_reports)
        card_layout.addWidget(self.lbl_data)
        card_layout.addWidget(btn_data)
        card_layout.addSpacing(6)
        card_layout.addWidget(btn_merge)

        footer = QLabel("© 2025 QA LAB Software — Developed by Shahid Iqbal")
        footer.setObjectName("footer")
        footer.setAlignment(Qt.AlignCenter)

        main_layout.addWidget(title)
        main_layout.addWidget(card)
        main_layout.addStretch()
        main_layout.addWidget(footer)

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

            if "weight" in row_lower and len(row) > 9 and is_number(row[9]):
                result["Weight"] = row[9]

            if "tear strength" in text:
                current_test = "Tear"
            elif "tensile strength" in text:
                current_test = "Tensile"
            elif "color fastness to rubbing" in text:
                current_test = "Rubbing"
                rubbing_found = {"Dry": False, "Wet": False}
            elif "color fastness to home laundering" in text:
                current_test = "Home Laundering"

            if "warp" in text and current_test in ["Tear", "Tensile"] and is_number(row[9]):
                result[f"{current_test} Warp"] = row[9]
            if "weft" in text and current_test in ["Tear", "Tensile"] and is_number(row[9]):
                result[f"{current_test} Weft"] = row[9]

            if current_test == "Rubbing":
                if "dry" in text and is_number(row[9]):
                    result["Rubbing Dry"] = row[9]
                if "wet" in text and is_number(row[9]):
                    result["Rubbing Wet"] = row[9]

            if current_test == "Home Laundering":
                if "shade change" in text and is_number(row[9]):
                    result["Shade Change"] = row[9]
                if "staining" in text and is_number(row[9]):
                    result["Staining"] = row[9]

            if "ph value" in row_lower and len(row) > 9 and is_number(row[9]):
                result["pH"] = row[9]

            if "temp" in row_lower and len(row) > 9 and is_number(row[9]):
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
