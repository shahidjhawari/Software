{"id":"51592","variant":"standard","title":"Updated Merge Software Code (Preserves Excel Formatting)"}
import sys
import os
import pandas as pd
from openpyxl import load_workbook
from PyQt5.QtGui import QIcon
from PyQt5.QtWidgets import (
    QApplication, QWidget, QPushButton, QLabel, QFileDialog,
    QListWidget, QVBoxLayout, QHBoxLayout, QMessageBox, QFrame
)
from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QStyleFactory


class ReportList(QListWidget):
    def __init__(self):
        super().__init__()
        self.setDragDropMode(QListWidget.InternalMove)
        self.setAcceptDrops(True)
        self.setDefaultDropAction(Qt.MoveAction)


class MergeApp(QWidget):
    def __init__(self):
        super().__init__()

        base_dir = os.path.dirname(os.path.abspath(__file__))
        icon_path = os.path.join(base_dir, "usgroup.ico")

        self.setWindowIcon(QIcon(icon_path))
        self.setWindowTitle("QA LAB — Auto Data Entry Software")
        self.setGeometry(200, 200, 1100, 680)
        QApplication.setStyle(QStyleFactory.create("Fusion"))

        self.report_files = []
        self.dataentry_file = None
        self.setAcceptDrops(True)

        # UI Styles
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
                padding: 10px;
                border: none;
                border-radius: 6px;
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

        # LAYOUTS
        main_layout = QHBoxLayout()
        left_box = QFrame()
        right_box = QFrame()

        left_layout = QVBoxLayout()
        right_layout = QVBoxLayout()

        # LEFT PANEL
        report_title = QLabel("Report Files")
        self.report_list = ReportList()

        self.btn_add_reports = QPushButton("Add Report Files")
        self.btn_add_reports.clicked.connect(self.load_reports)

        self.btn_clear_reports = QPushButton("Clear List")
        self.btn_clear_reports.clicked.connect(self.clear_reports)

        left_layout.addWidget(report_title)
        left_layout.addWidget(self.report_list)
        left_layout.addWidget(self.btn_add_reports)
        left_layout.addWidget(self.btn_clear_reports)
        left_box.setLayout(left_layout)

        # RIGHT PANEL
        data_title = QLabel("Data Entry Main File")
        self.data_display = QLabel("No File Selected")

        self.btn_dataentry = QPushButton("Select Data Entry File")
        self.btn_dataentry.clicked.connect(self.load_dataentry)

        right_layout.addWidget(data_title)
        right_layout.addWidget(self.data_display)
        right_layout.addWidget(self.btn_dataentry)
        right_box.setLayout(right_layout)

        # MERGE BUTTON
        self.btn_merge = QPushButton("Merge & Save Output")
        self.btn_merge.setFixedHeight(50)
        self.btn_merge.clicked.connect(self.merge_files)

        # FOOTER
        footer = QLabel("Developed by Shahid Iqbal — © All Rights Reserved")
        footer.setAlignment(Qt.AlignCenter)

        # MAIN
        main_layout.addWidget(left_box, 2)
        main_layout.addWidget(right_box, 1)

        wrapper = QVBoxLayout()
        wrapper.addLayout(main_layout)
        wrapper.addWidget(self.btn_merge)
        wrapper.addWidget(footer)
        self.setLayout(wrapper)

    # DRAG DROP
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            file_path = url.toLocalFile().lower()
            if file_path.endswith((".xlsx", ".xls")):
                if file_path not in self.report_files:
                    self.report_files.append(file_path)
                    self.report_list.addItem(file_path)

    # LOAD REPORTS
    def load_reports(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Select Reports", "", "Excel Files (*.xlsx *.xls)")
        if files:
            for f in files:
                if f not in self.report_files:
                    self.report_files.append(f)
                    self.report_list.addItem(f)

    def clear_reports(self):
        self.report_files = []
        self.report_list.clear()

    # LOAD BASE FILE
    def load_dataentry(self):
        file, _ = QFileDialog.getOpenFileName(self, "Select Data Entry File", "", "Excel Files (*.xlsx *.xls)")
        if file:
            self.dataentry_file = file
            self.data_display.setText(file)

    # EXTARCT VALUE
    def extract_value(self, df, colname):
        colname = str(colname).strip().lower()

        for i in range(df.shape[0]):
            for j in range(df.shape[1]):
                if str(df.iat[i, j]).strip().lower() == colname:

                    values = []
                    for k in range(j + 1, df.shape[1]):
                        v = df.iat[i, k]
                        if pd.notna(v) and str(v).strip() != "":
                            values.append(v)

                    return values[-1] if values else "-"

        return "-"

    # MERGE WITHOUT CHANGING FORMAT
    def merge_files(self):
        if not self.dataentry_file:
            QMessageBox.warning(self, "Error", "Please select Data Entry file.")
            return

        if self.report_list.count() == 0:
            QMessageBox.warning(self, "Error", "Please add at least one report.")
            return

        # Load Base File (Keep Formatting)
        wb = load_workbook(self.dataentry_file)
        ws = wb.active

        df_base = pd.read_excel(self.dataentry_file)
        base_columns = list(df_base.columns)

        ordered_reports = [self.report_list.item(i).text() for i in range(self.report_list.count())]

        for file in ordered_reports:
            df_r = pd.read_excel(file, header=None)

            row_values = [self.extract_value(df_r, col) for col in base_columns]

            ws.append(row_values)   # 💥 Formatting preserved fully

        save_path, _ = QFileDialog.getSaveFileName(self, "Save Output File", "", "Excel Files (*.xlsx)")
        if not save_path:
            return
        if not save_path.lower().endswith(".xlsx"):
            save_path += ".xlsx"

        wb.save(save_path)

        QMessageBox.information(self, "Success", "Merged file saved with original formatting intact!")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MergeApp()
    window.show()
    sys.exit(app.exec_())
