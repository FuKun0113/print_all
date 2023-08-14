import os
import sys
import ctypes
import win32print
import win32api
import win32con
from PySide6.QtWidgets import (QApplication, QWidget, QVBoxLayout, QPushButton, QLabel, QComboBox,
                               QFileDialog, QTableWidget, QTableWidgetItem, QCheckBox, QHeaderView, QMessageBox,
                               QSpinBox, QLineEdit, QHBoxLayout, QAbstractItemView)
from PySide6.QtCore import Qt

# Check for admin privileges
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

# If not admin, request for admin privileges
if not is_admin():
    ctypes.windll.shell32.ShellExecuteW(None, "runas", sys.executable, " ".join(sys.argv), None, 1)
    sys.exit(0)
def get_prints():
    prints_list = []
    printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS)
    for printer_info in printers:
        _, _, printer_name, _ = printer_info
        status_code = win32print.GetPrinter(win32print.OpenPrinter(printer_name), 2)['Status']

        if status_code == 0:
            status = "就绪"
        elif status_code == 128:
            status = "离线"
        elif status_code == 131072:
            status = "墨粉/墨水不足"
        else:
            status = "未知"

        prints_list.append(f"{printer_name} ({status})")
    return prints_list


class PrintApp(QMainWindow):

    def __init__(self):
        super().__init__()

        self.setWindowTitle("批量打印软件")
        main_layout = QVBoxLayout()

        # Printer selection group
        printer_group = QGroupBox("打印机选择")
        printer_layout = QHBoxLayout()
        self.printer_label = QLabel("选择打印机:")
        self.printer_combo = QComboBox()
        self.folder_input = QLineEdit("请选择文件夹")
        self.folder_select_btn = QPushButton("选择文件夹")

        printer_layout.addWidget(self.printer_label)
        printer_layout.addWidget(self.printer_combo)
        printer_layout.addWidget(self.folder_input)
        printer_layout.addWidget(self.folder_select_btn)
        printer_group.setLayout(printer_layout)
        main_layout.addWidget(printer_group)

        # File selection group
        file_group = QGroupBox("文件选择")
        file_layout = QVBoxLayout()
        self.file_list = QTableWidget()
        self.file_list.setColumnCount(6)
        self.file_list.setHorizontalHeaderLabels(["文件名", "是否打印", "单/双面", "颜色", "份数", "删除"])
        self.file_list.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.file_list.setDragDropOverwriteMode(False)
        self.file_list.setDragDropMode(QAbstractItemView.InternalMove)
        file_layout.addWidget(self.file_list)
        file_group.setLayout(file_layout)
        main_layout.addWidget(file_group)

        # Buttons
        button_layout = QHBoxLayout()
        self.add_files_btn = QPushButton("添加文件")
        self.clear_list_btn = QPushButton("清空列表")
        button_layout.addWidget(self.add_files_btn)
        button_layout.addWidget(self.clear_list_btn)
        main_layout.addLayout(button_layout)

        # Print button
        self.print_btn = QPushButton("开始打印")
        main_layout.addWidget(self.print_btn)

        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

        # Allow dropping files
        self.setAcceptDrops(True)

        # Connect signals
        self.add_files_btn.clicked.connect(self.add_files)
        self.clear_list_btn.clicked.connect(self.clear_files)
        self.print_btn.clicked.connect(self.print_files)
        self.folder_select_btn.clicked.connect(self.select_folder)

        self.init_ui()

    def init_ui(self):
        printers = get_prints()
        self.printer_combo.addItems(printers)

    def add_files(self):
        file_paths, _ = QFileDialog.getOpenFileNames(self, "选择文件", "",
                                                     "所有文件 (*.*);;文本文件 (*.txt);;PDF文件 (*.pdf);;Word文件 (*.docx);;Excel文件 (*.xlsx);;PowerPoint文件 (*.pptx)")
        for file_path in file_paths:
            self.add_file_to_list(file_path)

    def clear_files(self):
        self.file_list.setRowCount(0)

    def select_folder(self):
        folder_path = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder_path:
            self.folder_input.setText(folder_path)
            for root, _, files in os.walk(folder_path):
                for file_name in files:
                    self.add_file_to_list(os.path.join(root, file_name))

    def add_file_to_list(self, file_path):
        row_position = self.file_list.rowCount()
        self.file_list.insertRow(row_position)

        # File name
        self.file_list.setItem(row_position, 0, QTableWidgetItem(os.path.basename(file_path)))

        # Checkbox for printing
        print_checkbox = QCheckBox()
        print_checkbox.setChecked(True)
        self.file_list.setCellWidget(row_position, 1, print_checkbox)

        # Single/double sided printing
        duplex_combobox = QComboBox()
        duplex_combobox.addItems(['单面', '双面'])
        self.file_list.setCellWidget(row_position, 2, duplex_combobox)

        # Color/Monochrome
        color_combobox = QComboBox()
        color_combobox.addItems(['彩色', '黑白'])
        self.file_list.setCellWidget(row_position, 3, color_combobox)

        # Number of copies
        copies_spinbox = QSpinBox()
        copies_spinbox.setValue(1)
        self.file_list.setCellWidget(row_position, 4, copies_spinbox)

        # Delete button
        delete_button = QPushButton("删除")
        delete_button.clicked.connect(lambda _, row=row_position: self.file_list.removeRow(row))
        self.file_list.setCellWidget(row_position, 5, delete_button)

    def print_files(self):
        # Get selected printer
        printer_name = self.printer_combo.currentText().split(" (")[0]

        # Iterate over all rows and print if checkbox is checked
        for row in range(self.file_list.rowCount()):
            if self.file_list.cellWidget(row, 1).isChecked():
                file_name = self.file_list.item(row, 0).text()
                folder_path = self.folder_input.text()
                file_path = os.path.join(folder_path, file_name)

                duplex = 1 if self.file_list.cellWidget(row, 2).currentText() == "单面" else 2
                color = 1 if self.file_list.cellWidget(row, 3).currentText() == "彩色" else 2
                copies = self.file_list.cellWidget(row, 4).value()

                # Set printer settings
                original_settings = self.set_printer_settings(printer_name, duplex, color)
                for _ in range(copies):
                    win32api.ShellExecute(0, "printto", file_path, '"%s"' % printer_name, ".", 0)
                self.restore_printer_settings(printer_name, original_settings)

        QMessageBox.information(self, "打印", "打印中...")
        QMessageBox.information(self, "打印", "打印完成!")

    def set_printer_settings(self, printer_name, duplex, color):
        printer_handle = win32print.OpenPrinter(printer_name)
        printer_info = win32print.GetPrinter(printer_handle, 2)
        pdc = printer_info["pDevMode"]
        pdc.Fields = pdc.Fields | win32con.DM_DUPLEX
        pdc.Fields = pdc.Fields | win32con.DM_COLOR
        pdc.Duplex = duplex
        pdc.Color = color
        printer_info["pDevMode"] = pdc
        printer_info["pSecurityDescriptor"] = None
        win32print.SetPrinter(printer_handle, 2, printer_info, 0)
        win32print.ClosePrinter(printer_handle)
        return printer_info

    def restore_printer_settings(self, printer_name, printer_info):
        printer_handle = win32print.OpenPrinter(printer_name)
        printer_info["pSecurityDescriptor"] = None
        win32print.SetPrinter(printer_handle, 2, printer_info, 0)
        win32print.ClosePrinter(printer_handle)

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            self.add_file_to_list(url.toLocalFile())


if __name__ == "__main__":
    app = QApplication([])
    window = PrintApp()
    window.resize(800, 600)
    window.show()
    app.exec_()
