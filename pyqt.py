import sys
import os
import win32print
from PyQt6.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QPushButton, QListWidget, QListWidgetItem, \
    QComboBox, QCheckBox, QLabel, QSpinBox, QColorDialog, QFileDialog, QWidget, QHBoxLayout
from PyQt6.QtGui import QColor


class BatchPrintApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("批量打印程序")
        self.setGeometry(100, 100, 800, 600)

        self.printer_list = self.get_printer_list()

        self.print_list_widget = QListWidget()
        self.add_file_button = QPushButton("添加文件")
        self.printer_combo = QComboBox()
        self.printer_combo.addItems(self.printer_list)
        self.printer_status_label = QLabel()

        self.print_button = QPushButton("开始打印")

        layout = QVBoxLayout()
        layout.addWidget(self.add_file_button)
        layout.addWidget(self.print_list_widget)
        layout.addWidget(QLabel("选择打印机:"))
        layout.addWidget(self.printer_combo)
        layout.addWidget(self.printer_status_label)
        layout.addWidget(self.print_button)

        central_widget = QWidget()
        central_widget.setLayout(layout)
        self.setCentralWidget(central_widget)

        self.add_file_button.clicked.connect(self.add_files)
        self.print_button.clicked.connect(self.start_printing)
        self.printer_combo.currentIndexChanged.connect(self.update_printer_status)

        self.update_printer_status()

        self.print_settings = {}
        self.default_print_settings = {
            'print': True,
            'double_sided': True,
            'color': True,
            'copies': 1
        }

    def add_files(self):
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.FileMode.ExistingFiles)
        file_paths, _ = file_dialog.getOpenFileNames(self, "选择文件", "", "所有文件 (*)")

        for file_path in file_paths:
            item_widget = self.create_settings_widget()  # 创建包含横向布局的 widget
            filename_label = item_widget.layout().itemAt(0).widget()  # 获取 QLabel
            filename_label.setText(os.path.basename(file_path))  # 设置文件名

            item = item_widget.item  # 获取存储的 item
            self.print_list_widget.addItem(item)
            self.print_list_widget.setItemWidget(item, item_widget)
            self.print_settings[file_path] = self.default_print_settings

    def create_settings_widget(self):
        settings_widget = QWidget()
        settings_layout = QHBoxLayout()  # 使用 QHBoxLayout

        item = QListWidgetItem()
        settings_widget.item = item  # 将 item 存储在 widget 上
        settings_widget.setLayout(settings_layout)

        # 创建一个 QLabel 用于显示文件名
        filename_label = QLabel()

        print_checkbox = QCheckBox("打印")
        double_sided_checkbox = QCheckBox("双面打印")
        color_checkbox = QCheckBox("彩色打印")
        copies_spinbox = QSpinBox()
        copies_spinbox.setRange(1, 100)

        color_button = QPushButton("选择颜色")
        color_button.clicked.connect(self.choose_color)

        settings_layout.addWidget(filename_label)  # 将 QLabel 添加到布局
        settings_layout.addWidget(print_checkbox)
        settings_layout.addWidget(double_sided_checkbox)
        settings_layout.addWidget(color_checkbox)
        settings_layout.addWidget(QLabel("打印份数:"))
        settings_layout.addWidget(copies_spinbox)
        settings_layout.addWidget(color_button)

        item.setSizeHint(settings_widget.sizeHint())  # 设置 item 的大小

        return settings_widget

    def choose_color(self):
        color_dialog = QColorDialog()
        color = color_dialog.getColor()
        if color.isValid():
            sender = self.sender()
            item = self.print_list_widget.itemAt(sender.pos())
            file_path = self.print_list_widget.itemWidget(item).item.text()
            self.print_settings[file_path]['color'] = color

    def start_printing(self):
        selected_printer = self.printer_combo.currentText()
        for i in range(self.print_list_widget.count()):
            item = self.print_list_widget.item(i)
            file_path = self.print_list_widget.itemWidget(item).item.text()
            settings = self.print_settings[file_path]

            if settings['print']:
                printer_info = win32print.GetPrinter(selected_printer)
                if printer_info['Status'] == win32print.PRINTER_STATUS_OFFLINE:
                    print("打印机离线:", selected_printer)
                else:
                    print("开始打印文件:", file_path)
                    # 实际打印操作，使用 pywin32 的打印功能

    def update_printer_status(self):
        selected_printer = win32print.OpenPrinter(self.printer_combo.currentText())


        printer_info = win32print.GetPrinter(selected_printer, 2)
        if printer_info['Status'] == win32print.PRINTER_STATUS_OFFLINE:
            self.printer_status_label.setText("离线")
            self.printer_status_label.setStyleSheet("color: red;")
        else:
            self.printer_status_label.setText("在线")
            self.printer_status_label.setStyleSheet("color: green;")

    def get_printer_list(self):
        printer_list = []
        printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL, None, 1)
        for printer in printers:
            printer_list.append(printer[2])
        return printer_list


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = BatchPrintApp()
    window.show()
    sys.exit(app.exec())
