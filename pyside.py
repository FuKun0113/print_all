import sys
import os
import win32print
import win32ui
from PySide6.QtCore import Qt
from PySide6.QtGui import QIcon
from PySide6.QtWidgets import QApplication, QMainWindow, QWidget, QVBoxLayout, QPushButton, QListWidget, QListWidgetItem, QComboBox, QCheckBox, QLabel, QSpinBox,QFileDialog

class BatchPrintApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Batch Print")
        self.setGeometry(100, 100, 600, 400)
        self.setWindowIcon(QIcon("printer_icon.png"))  # Replace with your icon file

        self.central_widget = QWidget(self)
        self.setCentralWidget(self.central_widget)

        self.printer_status_label = QLabel(self)
        self.central_layout = QVBoxLayout(self.central_widget)

        self.printer_list = QComboBox(self)
        self.populate_printer_list()
        self.central_layout.addWidget(self.printer_list)

        self.printer_status_label.setAlignment(Qt.AlignCenter)
        self.central_layout.addWidget(self.printer_status_label)

        self.print_list = QListWidget(self)
        self.central_layout.addWidget(self.print_list)

        self.add_files_button = QPushButton("Add Files", self)
        self.add_files_button.clicked.connect(self.add_files)
        self.central_layout.addWidget(self.add_files_button)

        self.print_button = QPushButton("Print", self)
        self.print_button.clicked.connect(self.print_files)
        self.central_layout.addWidget(self.print_button)

    def populate_printer_list(self):
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
        for printer_info in printers:
            _, _, printer_name, _= printer_info
            status = win32print.GetPrinter(win32print.OpenPrinter(printer_name), 2)['Status']
            self.printer_list.addItem(f"{printer_name} {'在线' if status == 0 else '离线'}")

    def add_files(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFiles)
        if file_dialog.exec_():
            selected_files = file_dialog.selectedFiles()
            for file_path in selected_files:
                item = QListWidgetItem(file_path)
                item.setFlags(item.flags() | Qt.ItemIsEditable)
                self.print_list.addItem(item)

    def print_files(self):
        selected_printer = self.printer_list.currentText()
        printer_name = selected_printer.split()[0]
        printer_handle = win32print.OpenPrinter(printer_name)
        for row in range(self.print_list.count()):
            item = self.print_list.item(row)
            file_path = item.text()
            if item.checkState() == Qt.Checked:
                self.print_file(printer_handle, file_path)

        win32print.ClosePrinter(printer_handle)

    def print_file(self, printer_handle, file_path):
        if not os.path.exists(file_path):
            return

        hprinter = win32ui.CreateDC()
        hprinter.CreatePrinterDC(printer_handle)

        with open(file_path, "rb") as file:
            data = file.read()

        hprinter.StartDoc(file_path)
        hprinter.StartPage()
        hprinter.WritePrinter(data)
        hprinter.EndPage()
        hprinter.EndDoc()

        hprinter.DeleteDC()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = BatchPrintApp()
    window.show()
    sys.exit(app.exec_())