import sys
import os
import win32print
import win32api
import win32con
import math
import time
import json
from datetime import datetime
from enum import Enum
from dataclasses import dataclass
from typing import Dict, List, Optional
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                            QHBoxLayout, QLabel, QComboBox, QLineEdit, 
                            QPushButton, QScrollArea, QFrame, QSpinBox,
                            QFileDialog, QMessageBox, QStyledItemDelegate, QStyle, 
                            QCheckBox, QProgressBar, QTabWidget, QTableWidget,
                            QTableWidgetItem, QMenu)
from PyQt6.QtCore import Qt, QPropertyAnimation, QRect, QEasingCurve, QSize, QTimer, QUrl, pyqtSignal
from PyQt6.QtGui import QFont, QIcon, QPainter, QColor, QPixmap, QImage
from PyQt6.QtPdf import QPdfDocument
from PyQt6.QtPdfWidgets import QPdfView
from docx import Document
from openpyxl import load_workbook
from PIL import Image
import io

class CustomComboBox(QComboBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setItemDelegate(CustomItemDelegate(self))
        self.view().window().setWindowFlags(Qt.WindowType.Popup | Qt.WindowType.FramelessWindowHint | Qt.WindowType.NoDropShadowWindowHint)
        self.view().window().setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        # 设置动画
        self.animation = QPropertyAnimation(self.view().window(), b"geometry")
        self.animation.setDuration(200)
        self.animation.setEasingCurve(QEasingCurve.Type.OutQuad)
        
        # 添加呼吸动画定时器
        self._opacity = 1.0
        self._scale = 1.0
        self.animation_step = 0
        self.breath_timer = QTimer(self)
        self.breath_timer.timeout.connect(self._update_opacity)
        self.breath_timer.start(16)  # 约60fps的更新频率
        
        # 动画参数
        self.animation_speed = 0.03  # 控制动画速度
        self.base_scale = 1.0  # 基础缩放
        self.scale_range = 0.05  # 缩放范围 ±5%
        self.opacity_min = 0.6  # 最小透明度
        self.opacity_max = 1.0  # 最大透明度

        # 设置样式
        self.setStyleSheet("""
            QComboBox {
                border: 1px solid #dcdde1;
                border-radius: 4px;
                padding: 8px 30px 8px 12px;
                background: white;
                font-size: 13px;
                color: #2c3e50;
                min-height: 20px;
            }
            QComboBox:hover {
                border-color: #3498db;
            }
            QComboBox:focus {
                border-color: #3498db;
            }
            QComboBox::drop-down {
                border: none;
                width: 30px;
                background: transparent;
                padding-right: 5px;
            }
            QComboBox::down-arrow {
                image: url(down-arrow.svg);
                width: 12px;
                height: 12px;
            }
            QComboBox QAbstractItemView {
                border: 1px solid #dcdde1;
                border-radius: 4px;
                background: white;
                outline: none;
                selection-background-color: transparent;
            }
            QComboBox QAbstractItemView::item {
                height: 35px;
                padding: 0 12px;
                border: none;
                color: #2c3e50;
            }
            QComboBox QAbstractItemView::item:hover {
                background-color: #f5f6fa;
            }
            QComboBox QAbstractItemView::item:selected {
                background-color: #f0f0f0;
            }
        """)

    def showPopup(self):
        # 计算弹出框位置
        popup = self.view().window()
        
        # 获取ComboBox在屏幕上的位置和大小
        pos = self.mapToGlobal(QRect(0, 0, self.width(), self.height()).bottomLeft())
        
        # 设置弹出框位置和大小
        popup_width = self.width()  # 与ComboBox同宽
        popup_height = min(200, self.view().sizeHintForRow(0) * self.count() + 4)  # 限制最大高度
        
        # 调整Y坐标，确保紧贴ComboBox底部
        pos.setY(pos.y())  # 直接使用底部位置
        
        # 设置弹出框样式
        popup.setStyleSheet("""
            QListView {
                border: 1px solid #dcdde1;
                border-radius: 4px;
                background-color: white;
                outline: none;
                margin: 0px;
            }
        """)
        
        # 创建阴影效果
        geo = QRect(pos.x(), pos.y(), popup_width, 0)
        popup.setGeometry(geo)
        
        # 显示弹出框
        super().showPopup()
        
        # 开始动画
        self.animation.setStartValue(geo)
        geo.setHeight(popup_height)
        self.animation.setEndValue(geo)
        self.animation.start()

    def _update_opacity(self):
        self.animation_step += self.animation_speed
        
        # 使用多个正弦函数组合创建更自然的动画效果
        wave1 = math.sin(self.animation_step)
        wave2 = math.sin(self.animation_step * 0.5) * 0.3  # 添加一个较慢的波动
        wave = (wave1 + wave2) / 1.3  # 归一化到 -1 到 1 的范围
        
        # 将波形转换到0-1范围
        normalized_wave = (wave + 1) / 2
        
        # 计算透明度，使用缓动函数使变化更平滑
        opacity_range = self.opacity_max - self.opacity_min
        self._opacity = self.opacity_min + opacity_range * self._ease_in_out(normalized_wave)
        
        # 计算缩放，缩放效果要比透明度更微妙
        scale_factor = normalized_wave * 2 - 1  # 转回 -1 到 1
        self._scale = self.base_scale + self.scale_range * self._ease_in_out(scale_factor)
        
        # 强制更新视图
        self.view().viewport().update()
        self.update()

    def _ease_in_out(self, t):
        # 使用三次方缓动函数使动画更平滑
        if t < 0.5:
            return 4 * t * t * t
        else:
            return 1 - pow(-2 * t + 2, 3) / 2

    def get_opacity(self):
        return self._opacity

    def get_scale(self):
        return self._scale

class CustomItemDelegate(QStyledItemDelegate):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.combo_box = parent

    def paint(self, painter, option, index):
        painter.save()
        
        # 绘制背景
        if option.state & QStyle.StateFlag.State_MouseOver:
            painter.fillRect(option.rect, QColor("#f5f6fa"))
        elif option.state & QStyle.StateFlag.State_Selected:
            painter.fillRect(option.rect, QColor("#f0f0f0"))
        else:
            painter.fillRect(option.rect, QColor("white"))
            
        # 获取打印机状态
        printer_data = index.data(Qt.ItemDataRole.UserRole)
        if printer_data:
            name, status = printer_data
        else:
            name = index.data(Qt.ItemDataRole.DisplayRole)
            status = None
            
        # 绘制状态指示点
        if status is not None:
            # 计算正方形区域，保持宽高相等
            base_dot_size = min(option.rect.height() - 10, 12)  # 基础大小
            
            # 如果是就绪状态，应用缩放效果
            if status == 0 and self.combo_box:
                dot_size = int(base_dot_size * self.combo_box.get_scale())
            else:
                dot_size = base_dot_size
                
            # 确保圆点始终居中
            x_offset = int((base_dot_size - dot_size) / 2)
            y_offset = int((base_dot_size - dot_size) / 2)
            
            status_rect = QRect(
                int(option.rect.left() + 10 + x_offset),
                int(option.rect.top() + (option.rect.height() - base_dot_size) // 2 + y_offset),
                dot_size,
                dot_size
            )
            
            # 定义打印机状态常量
            PRINTER_STATUS_PAUSED = 1
            PRINTER_STATUS_ERROR = 2
            PRINTER_STATUS_PAPER_JAM = 8
            PRINTER_STATUS_PAPER_OUT = 16
            PRINTER_STATUS_PAPER_PROBLEM = 64
            PRINTER_STATUS_OFFLINE = 128
            PRINTER_STATUS_OUTPUT_BIN_FULL = 2048
            PRINTER_STATUS_NO_TONER = 262144
            PRINTER_STATUS_DOOR_OPEN = 4194304
            
            # 根据状态设置颜色
            if status == 0:  # 就绪状态
                status_color = QColor("#2ecc71")  # 绿色 - 就绪
                # 应用呼吸效果的透明度
                if self.combo_box:
                    status_color.setAlphaF(self.combo_box.get_opacity())
            elif status & (PRINTER_STATUS_ERROR | PRINTER_STATUS_OFFLINE | PRINTER_STATUS_NO_TONER | PRINTER_STATUS_DOOR_OPEN):
                status_color = QColor("#e74c3c")  # 红色 - 错误/离线
            elif status & (PRINTER_STATUS_PAPER_JAM | PRINTER_STATUS_PAPER_OUT | 
                         PRINTER_STATUS_PAPER_PROBLEM | PRINTER_STATUS_OUTPUT_BIN_FULL | PRINTER_STATUS_PAUSED):
                status_color = QColor("#f1c40f")  # 黄色 - 警告
            else:
                status_color = QColor("#95a5a6")  # 灰色 - 其他状态
                
            painter.setBrush(status_color)
            painter.setPen(Qt.PenStyle.NoPen)
            painter.drawEllipse(status_rect)
        
        # 设置文本颜色和字体
        painter.setPen(QColor("#2c3e50"))
        font = painter.font()
        font.setPointSize(10)  # 设置字体大小
        painter.setFont(font)
        
        # 计算文本绘制区域
        text_left_margin = 30 if status is not None else 10
        text_rect = option.rect.adjusted(text_left_margin, 0, -10, 0)
        
        # 绘制文本，确保垂直居中
        text = name if name else index.data(Qt.ItemDataRole.DisplayRole)
        painter.drawText(text_rect, Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, text)
        
        painter.restore()

    def sizeHint(self, option, index):
        # 设置每个项的高度为35像素
        return QSize(option.rect.width(), 35)

class SidesItemDelegate(QStyledItemDelegate):
    def paint(self, painter, option, index):
        painter.save()
        
        # 绘制背景
        if option.state & QStyle.StateFlag.State_MouseOver:
            painter.fillRect(option.rect, QColor("#f5f6fa"))
        elif option.state & QStyle.StateFlag.State_Selected:
            painter.fillRect(option.rect, QColor("#f0f0f0"))
        else:
            painter.fillRect(option.rect, QColor("white"))
        
        # 设置文本颜色和字体
        painter.setPen(QColor("#2c3e50"))
        font = painter.font()
        font.setPointSize(10)  # 设置字体大小
        painter.setFont(font)
        
        # 绘制文本，减小左右边距
        text_rect = option.rect.adjusted(8, 0, -8, 0)  # 左右各留8像素的边距
        text = index.data(Qt.ItemDataRole.DisplayRole)
        painter.drawText(text_rect, Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft, text)
        
        # 如果是选中项，绘制左边框高亮
        if option.state & QStyle.StateFlag.State_Selected:
            highlight_rect = QRect(option.rect.left(), option.rect.top(), 3, option.rect.height())
            painter.fillRect(highlight_rect, QColor("#3498db"))
        
        painter.restore()

    def sizeHint(self, option, index):
        return QSize(option.rect.width(), 45)  # 减小高度到30像素

class SidesComboBox(QComboBox):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setItemDelegate(SidesItemDelegate(self))
        self.view().window().setWindowFlags(Qt.WindowType.Popup | Qt.WindowType.FramelessWindowHint | Qt.WindowType.NoDropShadowWindowHint)
        self.view().window().setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        
        # 设置动画
        self.animation = QPropertyAnimation(self.view().window(), b"geometry")
        self.animation.setDuration(200)
        self.animation.setEasingCurve(QEasingCurve.Type.OutQuad)
        
        # 设置样式
        self.setStyleSheet("""
            QComboBox {
                padding: 8px 8px 8px 8px;  /* 减小内边距 */
                border: 1px solid #dcdde1;
                border-radius: 4px;
                background: white;
                font-size: 13px;
                color: #2c3e50;
                min-height: 20px;
            }
            QComboBox:hover {
                border-color: #3498db;
            }
            QComboBox:focus {
                border-color: #3498db;
            }
            QComboBox::drop-down {
                border: none;
                width: 20px;  /* 减小下拉箭头区域的宽度 */
            }
            QComboBox::down-arrow {
                image: url(down-arrow.svg);
                width: 12px;
                height: 12px;
            }
            QComboBox QAbstractItemView {
                border: none;
                background: white;
                outline: none;
                selection-background-color: transparent;
                margin: 0px;
                padding: 0px;
            }
            QComboBox QAbstractItemView::item {
                height: 10px;  /* 减小高度到30像素 */
                padding: 0;
                border: none;
                color: #2c3e50;
                background: white;
            }
            QComboBox QAbstractItemView::item:hover {
                background-color: transparent;
            }
            QComboBox QAbstractItemView::item:selected {
                background-color: transparent;
            }
        """)

    def showPopup(self):
        # 计算弹出框位置
        popup = self.view().window()
        
        # 获取ComboBox在屏幕上的位置和大小
        pos = self.mapToGlobal(QRect(0, 0, self.width(), self.height()).bottomLeft())
        
        # 设置弹出框位置和大小
        popup_width = self.width()  # 与ComboBox同宽
        popup_height = self.view().sizeHintForRow(0) * self.count() + 4  # 根据项目数量计算高度
        
        # 调整Y坐标，确保紧贴ComboBox底部
        pos.setY(pos.y() + 1)  # 向下偏移1像素以避免边框重叠
        
        # 设置弹出框样式
        popup.setStyleSheet("""
            QListView {
                border: 1px solid #dcdde1;
                border-radius: 4px;
                background-color: white;
                outline: none;
                margin: 0px;
            }
        """)
        
        # 创建阴影效果
        geo = QRect(pos.x(), pos.y(), popup_width, 0)
        popup.setGeometry(geo)
        
        # 显示弹出框
        super().showPopup()
        
        # 开始动画
        self.animation.setStartValue(geo)
        geo.setHeight(popup_height)
        self.animation.setEndValue(geo)
        self.animation.start()

class PreviewWindow(QMainWindow):
    def __init__(self, file_path, parent=None):
        super().__init__(parent)
        self.file_path = file_path
        self.setWindowTitle("文件预览")
        self.setMinimumSize(800, 600)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # 创建工具栏
        toolbar = QFrame()
        toolbar.setStyleSheet("""
            QFrame {
                background: white;
                border-bottom: 1px solid #dcdde1;
                padding: 5px;
            }
        """)
        toolbar_layout = QHBoxLayout(toolbar)
        
        # 添加缩放控件
        zoom_label = QLabel("缩放:")
        self.zoom_combo = QComboBox()
        self.zoom_combo.addItems(["50%", "75%", "100%", "125%", "150%", "200%"])
        self.zoom_combo.setCurrentText("100%")
        self.zoom_combo.currentTextChanged.connect(self.change_zoom)
        
        # 添加适应窗口按钮
        self.fit_btn = QPushButton("适应窗口")
        self.fit_btn.clicked.connect(self.fit_to_window)
        self.fit_btn.setFixedWidth(80)
        
        toolbar_layout.addStretch()
        toolbar_layout.addWidget(zoom_label)
        toolbar_layout.addWidget(self.zoom_combo)
        toolbar_layout.addWidget(self.fit_btn)
        
        layout.addWidget(toolbar)
        
        # 创建预览区域
        preview_widget = QWidget()
        preview_layout = QVBoxLayout(preview_widget)
        
        # 创建 PDF 预览区域
        self.pdf_view = QPdfView(preview_widget)
        self.pdf_view.setPageMode(QPdfView.PageMode.SinglePage)
        self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitInView)
        
        # 创建 PDF 文档对象
        self.pdf_document = QPdfDocument(self)
        self.pdf_view.setDocument(self.pdf_document)
        
        preview_layout.addWidget(self.pdf_view)
        layout.addWidget(preview_widget)
        
        # 加载文件
        self.load_file()
    
    def load_file(self):
        ext = os.path.splitext(self.file_path)[1].lower()
        
        try:
            if ext == '.pdf':
                self.load_pdf()
            elif ext in ['.jpg', '.jpeg', '.png']:
                self.load_image()
            elif ext in ['.docx', '.doc']:
                self.convert_to_pdf_and_load()
            elif ext in ['.xlsx', '.xls']:
                self.convert_to_pdf_and_load()
            else:
                self.pdf_view.hide()
                label = QLabel("暂不支持此文件格式的预览")
                label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                self.centralWidget().layout().addWidget(label)
        except Exception as e:
            self.pdf_view.hide()
            label = QLabel(f"预览失败: {str(e)}")
            label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.centralWidget().layout().addWidget(label)
    
    def load_pdf(self):
        self.pdf_document.load(self.file_path)
        self.pdf_view.show()
    
    def load_image(self):
        # 将图片转换为 PDF 后显示
        try:
            img = Image.open(self.file_path)
            pdf_path = os.path.join(os.path.dirname(self.file_path), "_temp.pdf")
            img.save(pdf_path, "PDF", resolution=100.0)
            self.pdf_document.load(pdf_path)
            self.pdf_view.show()
            # 删除临时文件
            os.remove(pdf_path)
        except Exception as e:
            self.pdf_view.hide()
            label = QLabel(f"图片预览失败: {str(e)}")
            label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            self.centralWidget().layout().addWidget(label)
    
    def convert_to_pdf_and_load(self):
        # TODO: 实现文档转换为 PDF 的功能
        self.pdf_view.hide()
        label = QLabel("文档预览功能即将推出")
        label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.centralWidget().layout().addWidget(label)
    
    def change_zoom(self, zoom_text):
        zoom_value = float(zoom_text.strip('%')) / 100
        self.pdf_view.setZoomFactor(zoom_value)
    
    def fit_to_window(self):
        self.pdf_view.setZoomMode(QPdfView.ZoomMode.FitInView)
    
    def closeEvent(self, event):
        self.pdf_document.close()
        super().closeEvent(event)

class PrintStatus(Enum):
    WAITING = "等待打印"
    PRINTING = "正在打印"
    PAUSED = "已暂停"
    COMPLETED = "已完成"
    FAILED = "打印失败"
    CANCELLED = "已取消"

@dataclass
class PrintSettings:
    paper_size: str
    orientation: str
    page_range: str
    color_mode: str
    sides_option: str
    copies: int

class PrintTask:
    def __init__(self, file_path: str, settings: PrintSettings):
        self.file_path = file_path
        self.file_name = os.path.basename(file_path)
        self.settings = settings
        self.status = PrintStatus.WAITING
        self.progress = 0
        self.start_time: Optional[float] = None
        self.end_time: Optional[float] = None
        self.error_message: Optional[str] = None
        self.total_pages = 1
        self.current_page = 0
    
    def start(self):
        self.status = PrintStatus.PRINTING
        self.start_time = time.time()
    
    def complete(self):
        self.status = PrintStatus.COMPLETED
        self.progress = 100
        self.end_time = time.time()
    
    def fail(self, error_message: str):
        self.status = PrintStatus.FAILED
        self.error_message = error_message
        self.end_time = time.time()
    
    def cancel(self):
        self.status = PrintStatus.CANCELLED
        self.end_time = time.time()
    
    def pause(self):
        if self.status == PrintStatus.PRINTING:
            self.status = PrintStatus.PAUSED
    
    def resume(self):
        if self.status == PrintStatus.PAUSED:
            self.status = PrintStatus.PRINTING
    
    def update_progress(self, current_page: int, total_pages: int):
        self.current_page = current_page
        self.total_pages = total_pages
        self.progress = int((current_page / total_pages) * 100)

class PrintQueue:
    def __init__(self):
        self.waiting_tasks: List[PrintTask] = []
        self.current_task: Optional[PrintTask] = None
        self.completed_tasks: List[PrintTask] = []
        self.history_file = "print_history.json"
        self.load_history()
    
    def add_task(self, task: PrintTask):
        self.waiting_tasks.append(task)
    
    def start_next_task(self) -> Optional[PrintTask]:
        if not self.current_task and self.waiting_tasks:
            self.current_task = self.waiting_tasks.pop(0)
            self.current_task.start()
            return self.current_task
        return None
    
    def complete_current_task(self):
        if self.current_task:
            self.current_task.complete()
            self.completed_tasks.append(self.current_task)
            self.current_task = None
            self.save_history()
    
    def fail_current_task(self, error_message: str):
        if self.current_task:
            self.current_task.fail(error_message)
            self.completed_tasks.append(self.current_task)
            self.current_task = None
            self.save_history()
    
    def cancel_task(self, task: PrintTask):
        if task == self.current_task:
            task.cancel()
            self.completed_tasks.append(task)
            self.current_task = None
        else:
            self.waiting_tasks.remove(task)
            task.cancel()
            self.completed_tasks.append(task)
        self.save_history()
    
    def pause_current_task(self):
        if self.current_task:
            self.current_task.pause()
    
    def resume_current_task(self):
        if self.current_task:
            self.current_task.resume()
    
    def save_history(self):
        history_data = []
        for task in self.completed_tasks:
            history_data.append({
                'file_name': task.file_name,
                'file_path': task.file_path,
                'status': task.status.value,
                'start_time': task.start_time,
                'end_time': task.end_time,
                'error_message': task.error_message,
                'settings': {
                    'paper_size': task.settings.paper_size,
                    'orientation': task.settings.orientation,
                    'page_range': task.settings.page_range,
                    'color_mode': task.settings.color_mode,
                    'sides_option': task.settings.sides_option,
                    'copies': task.settings.copies
                }
            })
        
        with open(self.history_file, 'w', encoding='utf-8') as f:
            json.dump(history_data, f, ensure_ascii=False, indent=2)
    
    def load_history(self):
        if not os.path.exists(self.history_file):
            return
        
        try:
            with open(self.history_file, 'r', encoding='utf-8') as f:
                history_data = json.load(f)
            
            for task_data in history_data:
                settings = PrintSettings(
                    paper_size=task_data['settings']['paper_size'],
                    orientation=task_data['settings']['orientation'],
                    page_range=task_data['settings']['page_range'],
                    color_mode=task_data['settings']['color_mode'],
                    sides_option=task_data['settings']['sides_option'],
                    copies=task_data['settings']['copies']
                )
                
                task = PrintTask(task_data['file_path'], settings)
                task.file_name = task_data['file_name']
                task.status = PrintStatus(task_data['status'])
                task.start_time = task_data['start_time']
                task.end_time = task_data['end_time']
                task.error_message = task_data['error_message']
                
                self.completed_tasks.append(task)
        except Exception as e:
            print(f"Error loading print history: {str(e)}")

class PrintQueueWidget(QWidget):
    def __init__(self, print_queue: PrintQueue, parent=None):
        super().__init__(parent)
        self.print_queue = print_queue
        self.setup_ui()
        
        # 更新定时器
        self.update_timer = QTimer()
        self.update_timer.timeout.connect(self.update_display)
        self.update_timer.start(1000)  # 每秒更新一次
    
    def setup_ui(self):
        layout = QVBoxLayout(self)
        
        # 创建选项卡
        tabs = QTabWidget()
        
        # 当前任务选项卡
        current_tab = QWidget()
        current_layout = QVBoxLayout(current_tab)
        
        # 当前任务信息
        self.current_task_info = QLabel("当前没有打印任务")
        current_layout.addWidget(self.current_task_info)
        
        # 进度条
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        current_layout.addWidget(self.progress_bar)
        
        # 控制按钮
        control_layout = QHBoxLayout()
        self.pause_btn = QPushButton("暂停")
        self.pause_btn.clicked.connect(self.toggle_pause)
        self.cancel_btn = QPushButton("取消")
        self.cancel_btn.clicked.connect(self.cancel_current)
        
        control_layout.addWidget(self.pause_btn)
        control_layout.addWidget(self.cancel_btn)
        current_layout.addLayout(control_layout)
        
        # 等待任务列表
        self.waiting_table = QTableWidget()
        self.waiting_table.setColumnCount(5)
        self.waiting_table.setHorizontalHeaderLabels(["文件名", "页数", "打印设置", "状态", "操作"])
        current_layout.addWidget(self.waiting_table)
        
        # 历史记录选项卡
        history_tab = QWidget()
        history_layout = QVBoxLayout(history_tab)
        
        self.history_table = QTableWidget()
        self.history_table.setColumnCount(6)
        self.history_table.setHorizontalHeaderLabels(["文件名", "打印时间", "打印设置", "状态", "耗时", "备注"])
        history_layout.addWidget(self.history_table)
        
        # 添加选项卡
        tabs.addTab(current_tab, "当前任务")
        tabs.addTab(history_tab, "历史记录")
        
        layout.addWidget(tabs)
    
    def update_display(self):
        # 更新当前任务信息
        if self.print_queue.current_task:
            task = self.print_queue.current_task
            self.current_task_info.setText(
                f"正在打印: {task.file_name}\n"
                f"页数: {task.current_page}/{task.total_pages}"
            )
            self.progress_bar.setValue(task.progress)
            
            # 更新按钮状态
            self.pause_btn.setEnabled(True)
            self.cancel_btn.setEnabled(True)
            self.pause_btn.setText("继续" if task.status == PrintStatus.PAUSED else "暂停")
        else:
            self.current_task_info.setText("当前没有打印任务")
            self.progress_bar.setValue(0)
            self.pause_btn.setEnabled(False)
            self.cancel_btn.setEnabled(False)
        
        # 更新等待任务列表
        self.waiting_table.setRowCount(len(self.print_queue.waiting_tasks))
        for i, task in enumerate(self.print_queue.waiting_tasks):
            self.waiting_table.setItem(i, 0, QTableWidgetItem(task.file_name))
            self.waiting_table.setItem(i, 1, QTableWidgetItem(str(task.total_pages)))
            self.waiting_table.setItem(i, 2, QTableWidgetItem(
                f"{task.settings.paper_size}, "
                f"{task.settings.orientation}, "
                f"{task.settings.color_mode}"
            ))
            self.waiting_table.setItem(i, 3, QTableWidgetItem(task.status.value))
            
            # 添加取消按钮
            cancel_btn = QPushButton("取消")
            cancel_btn.clicked.connect(lambda checked, t=task: self.cancel_task(t))
            self.waiting_table.setCellWidget(i, 4, cancel_btn)
        
        # 更新历史记录
        self.history_table.setRowCount(len(self.print_queue.completed_tasks))
        for i, task in enumerate(reversed(self.print_queue.completed_tasks)):
            self.history_table.setItem(i, 0, QTableWidgetItem(task.file_name))
            
            # 格式化打印时间
            if task.start_time:
                start_time = datetime.fromtimestamp(task.start_time).strftime("%Y-%m-%d %H:%M:%S")
                self.history_table.setItem(i, 1, QTableWidgetItem(start_time))
            
            self.history_table.setItem(i, 2, QTableWidgetItem(
                f"{task.settings.paper_size}, "
                f"{task.settings.orientation}, "
                f"{task.settings.color_mode}"
            ))
            
            status_item = QTableWidgetItem(task.status.value)
            if task.status == PrintStatus.COMPLETED:
                status_item.setBackground(QColor("#2ecc71"))  # 绿色
            elif task.status == PrintStatus.FAILED:
                status_item.setBackground(QColor("#e74c3c"))  # 红色
            self.history_table.setItem(i, 3, status_item)
            
            # 计算耗时
            if task.start_time and task.end_time:
                duration = task.end_time - task.start_time
                duration_str = f"{duration:.1f}秒"
                self.history_table.setItem(i, 4, QTableWidgetItem(duration_str))
            
            # 添加错误信息或备注
            if task.error_message:
                self.history_table.setItem(i, 5, QTableWidgetItem(task.error_message))
    
    def toggle_pause(self):
        if self.print_queue.current_task:
            if self.print_queue.current_task.status == PrintStatus.PAUSED:
                self.print_queue.resume_current_task()
            else:
                self.print_queue.pause_current_task()
    
    def cancel_current(self):
        if self.print_queue.current_task:
            self.print_queue.cancel_task(self.print_queue.current_task)
    
    def cancel_task(self, task: PrintTask):
        self.print_queue.cancel_task(task)

class PrintQueueWindow(QMainWindow):
    def __init__(self, print_queue: PrintQueue, parent=None):
        super().__init__(parent)
        self.setWindowTitle("打印任务管理")
        self.setMinimumSize(800, 600)
        
        # 创建中央部件
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        
        # 添加打印队列管理器
        self.queue_widget = PrintQueueWidget(print_queue)
        layout.addWidget(self.queue_widget)
        
        # 设置窗口样式
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f6fa;
            }
            QTabWidget::pane {
                border: 1px solid #dcdde1;
                border-radius: 4px;
                background: white;
                padding: 10px;
            }
            QTabWidget::tab-bar {
                left: 5px;
            }
            QTabBar::tab {
                background: #f8f9fa;
                border: 1px solid #dcdde1;
                border-bottom: none;
                border-top-left-radius: 4px;
                border-top-right-radius: 4px;
                padding: 8px 15px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background: white;
                border-bottom: none;
                margin-bottom: -1px;
            }
            QTableWidget {
                border: none;
                background: white;
                gridline-color: #f1f2f6;
            }
            QTableWidget::item {
                padding: 8px;
            }
            QTableWidget::item:selected {
                background: #f5f6fa;
                color: #2c3e50;
            }
            QHeaderView::section {
                background: #f8f9fa;
                padding: 8px;
                border: none;
                border-right: 1px solid #dcdde1;
                border-bottom: 1px solid #dcdde1;
            }
            QPushButton {
                padding: 8px 15px;
                border: none;
                border-radius: 4px;
                background-color: #3498db;
                color: white;
                font-size: 13px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QPushButton:disabled {
                background-color: #bdc3c7;
            }
            QProgressBar {
                border: 1px solid #dcdde1;
                border-radius: 4px;
                text-align: center;
                background: white;
            }
            QProgressBar::chunk {
                background-color: #3498db;
                border-radius: 3px;
            }
        """)

class PrinterApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("云雀批量打印工具")
        self.setMinimumSize(1000, 800)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #f5f6fa;
            }
            QLabel {
                font-size: 14px;
                color: #2f3542;
            }
            QLineEdit {
                padding: 8px 15px;
                border: 1px solid #dcdde1;
                border-radius: 4px;
                background-color: white;
                font-size: 13px;
            }
            QLineEdit:focus {
                border-color: #3498db;
            }
            QPushButton {
                padding: 8px 15px;
                border: none;
                border-radius: 4px;
                background-color: #3498db;
                color: white;
                font-size: 13px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
            QSpinBox {
                padding: 8px 15px;
                border: 1px solid #dcdde1;
                border-radius: 4px;
                background-color: white;
                font-size: 13px;
            }
            QSpinBox::up-button, QSpinBox::down-button {
                width: 20px;
                border: none;
                background: #f8f9fa;
            }
            QSpinBox::up-button:hover, QSpinBox::down-button:hover {
                background: #f1f2f6;
            }
            QScrollArea {
                border: 1px solid #dcdde1;
                border-radius: 4px;
                background-color: white;
            }
            QFrame#toolbar {
                background-color: white;
                border: 1px solid #dcdde1;
                border-radius: 4px;
                margin: 10px;
                padding: 15px;
            }
            QLabel#header_label {
                font-weight: bold;
                color: #2c3e50;
                padding: 5px;
                min-width: 80px;
            }
            QWidget#list_item {
                background-color: #f8f9fa;
                border-radius: 4px;
                margin: 2px 5px;
            }
            QWidget#list_item:hover {
                background-color: #f1f2f6;
            }
        """)
        
        # 创建主窗口部件
        main_widget = QWidget()
        self.setCentralWidget(main_widget)
        
        # 创建主布局
        layout = QVBoxLayout(main_widget)
        layout.setSpacing(15)
        layout.setContentsMargins(20, 20, 20, 20)
        
        # 添加说明文字
        description = QLabel("本工具支持excel、pdf、word、文本、图片等文件的批量打印")
        description.setStyleSheet("font-size: 16px; color: #2c3e50; padding: 10px;")
        description.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(description)
        
        # 工具栏
        toolbar = QFrame()
        toolbar.setObjectName("toolbar")
        toolbar_layout = QVBoxLayout(toolbar)
        toolbar_layout.setSpacing(15)
        
        # 创建工具栏
        toolbar_row = QHBoxLayout()
        toolbar_row.setSpacing(20)
        
        # 打印机选择
        printer_label = QLabel("选择打印机:")
        printer_label.setObjectName("header_label")
        toolbar_row.addWidget(printer_label)
        
        self.printer_combo = CustomComboBox()
        self.update_printer_list()
        toolbar_row.addWidget(self.printer_combo)
        
        # 文件夹路径输入框
        path_label = QLabel("文件夹路径:")
        path_label.setObjectName("header_label")
        toolbar_row.addWidget(path_label)
        
        self.path_input = QLineEdit()
        self.path_input.setPlaceholderText("请选择文件夹")
        self.path_input.setMinimumWidth(300)
        toolbar_row.addWidget(self.path_input)
        
        # 选择文件夹按钮
        self.select_folder_btn = QPushButton("选择文件夹")
        self.select_folder_btn.clicked.connect(self.select_folder)
        toolbar_row.addWidget(self.select_folder_btn)
        
        toolbar_layout.addLayout(toolbar_row)
        layout.addWidget(toolbar)
        
        # 文件列表区域
        list_container = QFrame()
        list_container.setStyleSheet("""
            QFrame {
                background-color: white;
                border: 1px solid #dcdde1;
                border-radius: 4px;
                padding: 10px;
            }
        """)
        list_layout = QVBoxLayout(list_container)
        list_layout.setSpacing(10)
        
        # 添加表头
        header = QWidget()
        header_layout = QHBoxLayout(header)
        header_layout.setContentsMargins(15, 5, 15, 5)
        header_layout.setSpacing(15)
        
        # 设置表头标签
        header_widths = [250, 80, 100, 100, 120, 100, 100, 100]  # 调整列宽
        headers = ["文件名", "格式", "纸张", "方向", "页面范围", "颜色", "单双面", "份数"]
        
        for text, width in zip(headers, header_widths):
            label = QLabel(text)
            label.setObjectName("header_label")
            label.setFixedWidth(width)
            label.setStyleSheet("""
                QLabel#header_label {
                    font-weight: bold;
                    color: #2c3e50;
                    font-size: 14px;
                    padding: 8px;
                    background-color: transparent;
                }
            """)
            header_layout.addWidget(label)
        
        header_layout.addStretch()
        list_layout.addWidget(header)
        
        # 滚动区域
        self.list_area = QScrollArea()
        self.list_widget = QWidget()
        self.list_layout = QVBoxLayout(self.list_widget)
        self.list_layout.setSpacing(5)
        self.list_layout.setContentsMargins(15, 5, 15, 5)
        
        self.list_area.setWidget(self.list_widget)
        self.list_area.setWidgetResizable(True)
        list_layout.addWidget(self.list_area)
        
        layout.addWidget(list_container)
        
        # 打印按钮
        self.print_btn = QPushButton("开始打印")
        self.print_btn.setStyleSheet("""
            QPushButton {
                background-color: #2ecc71;
                color: white;
                padding: 12px;
                font-size: 15px;
                font-weight: bold;
                min-width: 200px;
            }
            QPushButton:hover {
                background-color: #27ae60;
            }
        """)
        self.print_btn.clicked.connect(self.print_files)
        
        # 将打印按钮居中显示
        btn_container = QWidget()
        btn_layout = QHBoxLayout(btn_container)
        btn_layout.addStretch()
        btn_layout.addWidget(self.print_btn)
        btn_layout.addStretch()
        layout.addWidget(btn_container)
        
        # 存储文件项的列表
        self.file_items = []
        
        # 在工具栏中添加搜索和排序控件
        toolbar_row2 = QHBoxLayout()
        toolbar_row2.setSpacing(20)
        
        # 添加搜索框
        search_label = QLabel("搜索文件:")
        search_label.setObjectName("header_label")
        toolbar_row2.addWidget(search_label)
        
        self.search_input = QLineEdit()
        self.search_input.setPlaceholderText("输入文件名进行搜索")
        self.search_input.textChanged.connect(self.filter_files)
        toolbar_row2.addWidget(self.search_input)
        
        # 添加排序选项
        sort_label = QLabel("排序方式:")
        sort_label.setObjectName("header_label")
        toolbar_row2.addWidget(sort_label)
        
        self.sort_combo = CustomComboBox()
        self.sort_combo.addItems(["名称升序", "名称降序", "类型升序", "类型降序", "大小升序", "大小降序"])
        self.sort_combo.currentTextChanged.connect(self.sort_files)
        toolbar_row2.addWidget(self.sort_combo)
        
        # 添加批量选择按钮
        self.select_all_btn = QPushButton("全选")
        self.select_all_btn.clicked.connect(self.toggle_select_all)
        self.select_all_btn.setFixedWidth(80)
        toolbar_row2.addWidget(self.select_all_btn)
        
        toolbar_row2.addStretch()
        toolbar_layout.addLayout(toolbar_row2)
        
        # 添加文件选择状态存储
        self.file_selections = {}
        
        # 添加定时器以定期更新打印机状态
        self.status_timer = QTimer()
        self.status_timer.timeout.connect(self.update_printer_status)
        self.status_timer.start(5000)  # 每5秒更新一次状态
        
        # 启用拖放功能
        self.setAcceptDrops(True)
        
        # 创建打印队列
        self.print_queue = PrintQueue()
        
        # 添加打印任务管理按钮
        self.queue_btn = QPushButton("打印任务管理")
        self.queue_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                padding: 8px 15px;
                border: none;
                border-radius: 4px;
                font-size: 13px;
                min-width: 120px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        self.queue_btn.clicked.connect(self.show_queue_window)
        
        # 将打印任务管理按钮添加到按钮容器中
        btn_layout.addWidget(self.queue_btn)
        btn_layout.addWidget(self.print_btn)
        btn_layout.addStretch()
        
        # 创建打印任务管理窗口（但不显示）
        self.queue_window = None
    
    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择文件夹")
        if folder:
            self.path_input.setText(folder)
            self.load_file_list()  # 自动加载文件列表
    
    def load_file_list(self):
        path = self.path_input.text()
        if not path:
            QMessageBox.warning(self, "警告", "请先选择文件夹！")
            return
            
        # 清除现有列表
        for item in self.file_items:
            item.deleteLater()
        self.file_items.clear()
        self.file_selections.clear()
        
        # 获取文件列表
        files = []
        supported_extensions = ['.pdf', '.doc', '.docx', '.xls', '.xlsx', '.txt', '.jpg', '.jpeg', '.png']
        
        for file in os.listdir(path):
            ext = os.path.splitext(file)[1].lower()
            if ext not in supported_extensions:
                continue
            
            file_path = os.path.join(path, file)
            size = os.path.getsize(file_path)
            files.append({
                'name': file,
                'ext': ext,
                'size': size,
                'path': file_path
            })
        
        # 应用当前排序
        self.sort_files(self.sort_combo.currentText(), files)
        
        # 创建文件列表项
        for file_info in files:
            item = self.create_file_item(file_info)
            self.list_layout.addWidget(item)
            self.file_items.append(item)
            self.file_selections[file_info['name']] = False
        
        if not self.file_items:
            QMessageBox.information(self, "提示", "未找到支持的文件格式！")
    
    def create_file_item(self, file_info):
        item = QWidget()
        item.setObjectName("list_item")
        item_layout = QHBoxLayout(item)
        item_layout.setContentsMargins(15, 5, 15, 5)
        item_layout.setSpacing(15)
        
        # 添加复选框
        checkbox = QCheckBox()
        checkbox.setStyleSheet("""
            QCheckBox {
                padding: 5px;
            }
            QCheckBox::indicator {
                width: 20px;
                height: 20px;
            }
            QCheckBox::indicator:unchecked {
                border: 2px solid #dcdde1;
                background: white;
                border-radius: 4px;
            }
            QCheckBox::indicator:checked {
                border: 2px solid #3498db;
                background: #3498db;
                border-radius: 4px;
            }
        """)
        checkbox.stateChanged.connect(lambda state, name=file_info['name']: self.on_file_selection_changed(name, state))
        item_layout.addWidget(checkbox)
        
        # 文件名
        name_label = QLabel(file_info['name'])
        name_label.setFixedWidth(250)
        name_label.setStyleSheet("""
            padding: 8px;
            background: white;
            border: 1px solid #dcdde1;
            border-radius: 4px;
        """)
        name_label.setToolTip(file_info['name'])
        item_layout.addWidget(name_label)
        
        # 文件格式
        ext_label = QLabel(file_info['ext'][1:].upper())
        ext_label.setFixedWidth(80)
        ext_label.setStyleSheet("""
            padding: 8px;
            background: white;
            border: 1px solid #dcdde1;
            border-radius: 4px;
        """)
        item_layout.addWidget(ext_label)
        
        # 纸张大小选择
        paper_combo = CustomComboBox()
        paper_combo.addItems(["A4", "A3", "B5", "Letter", "Legal"])
        paper_combo.setFixedWidth(100)
        item_layout.addWidget(paper_combo)
        
        # 打印方向选择
        orientation_combo = CustomComboBox()
        orientation_combo.addItems(["纵向", "横向"])
        orientation_combo.setFixedWidth(100)
        item_layout.addWidget(orientation_combo)
        
        # 页面范围输入框
        page_range_input = QLineEdit()
        page_range_input.setPlaceholderText("全部")
        page_range_input.setToolTip("输入格式：1-3,5,7-9")
        page_range_input.setFixedWidth(120)
        page_range_input.setStyleSheet("""
            QLineEdit {
                padding: 8px;
                border: 1px solid #dcdde1;
                border-radius: 4px;
                background: white;
            }
            QLineEdit:focus {
                border-color: #3498db;
            }
        """)
        item_layout.addWidget(page_range_input)
        
        # 颜色模式选择
        color_combo = CustomComboBox()
        color_combo.addItems(["彩色", "黑白"])
        color_combo.setFixedWidth(100)
        item_layout.addWidget(color_combo)
        
        # 单双面选择
        sides_combo = SidesComboBox()
        sides_combo.addItems(["单面", "双面长边", "双面短边"])
        sides_combo.setFixedWidth(100)
        item_layout.addWidget(sides_combo)
        
        # 打印份数
        copies_spin = QSpinBox()
        copies_spin.setMinimum(1)
        copies_spin.setMaximum(99)
        copies_spin.setFixedWidth(100)
        copies_spin.setStyleSheet("""
            QSpinBox {
                padding: 8px;
                border: 1px solid #dcdde1;
                border-radius: 4px;
                background: white;
            }
            QSpinBox::up-button, QSpinBox::down-button {
                width: 20px;
                border: none;
                background: #f8f9fa;
            }
            QSpinBox::up-button:hover, QSpinBox::down-button:hover {
                background: #f1f2f6;
            }
        """)
        item_layout.addWidget(copies_spin)
        
        # 添加预览按钮
        preview_btn = QPushButton("预览")
        preview_btn.setFixedWidth(60)
        preview_btn.setStyleSheet("""
            QPushButton {
                background-color: #3498db;
                color: white;
                border: none;
                border-radius: 4px;
                padding: 8px;
                font-size: 12px;
                min-width: 60px;
            }
            QPushButton:hover {
                background-color: #2980b9;
            }
        """)
        preview_btn.clicked.connect(lambda checked, f=file_info['name']: self.preview_file(f))
        item_layout.addWidget(preview_btn)
        
        item_layout.addStretch()
        return item
    
    def preview_file(self, file_name):
        file_path = os.path.join(self.path_input.text(), file_name)
        preview_window = PreviewWindow(file_path, self)
        preview_window.show()
    
    def print_files(self):
        if not self.file_items:
            QMessageBox.warning(self, "警告", "请先选择要打印的文件！")
            return
        
        printer_name = self.printer_combo.currentText()
        folder_path = self.path_input.text()
        
        # 添加选中的文件到打印队列
        for item in self.file_items:
            layout = item.layout()
            checkbox = layout.itemAt(0).widget()
            
            if not checkbox.isChecked():
                continue
            
            file_name = layout.itemAt(1).widget().text()
            settings = PrintSettings(
                paper_size=layout.itemAt(3).widget().currentText(),
                orientation=layout.itemAt(4).widget().currentText(),
                page_range=layout.itemAt(5).widget().text(),
                color_mode=layout.itemAt(6).widget().currentText(),
                sides_option=layout.itemAt(7).widget().currentText(),
                copies=layout.itemAt(8).widget().value()
            )
            
            task = PrintTask(os.path.join(folder_path, file_name), settings)
            self.print_queue.add_task(task)
        
        # 开始打印队列中的第一个任务
        self.process_print_queue()
    
    def process_print_queue(self):
        if self.print_queue.current_task or not self.print_queue.waiting_tasks:
            return
        
        task = self.print_queue.start_next_task()
        if not task:
            return
        
        try:
            printer_name = self.printer_combo.currentText()
            
            # 设置默认打印机
            win32print.SetDefaultPrinter(printer_name)
            
            # 打开打印机
            handle = win32print.OpenPrinter(printer_name)
            try:
                # 获取打印机默认设置
                properties = win32print.GetPrinter(handle, 2)
                devmode = properties['pDevMode']
                
                # 设置打印参数
                PAPER_SIZES = {
                    "A4": win32con.DMPAPER_A4,
                    "A3": win32con.DMPAPER_A3,
                    "B5": win32con.DMPAPER_B5,
                    "Letter": win32con.DMPAPER_LETTER,
                    "Legal": win32con.DMPAPER_LEGAL
                }
                
                if task.settings.paper_size in PAPER_SIZES:
                    devmode.PaperSize = PAPER_SIZES[task.settings.paper_size]
                
                if task.settings.orientation == "横向":
                    devmode.Orientation = win32con.DMORIENT_LANDSCAPE
                else:
                    devmode.Orientation = win32con.DMORIENT_PORTRAIT
                
                if task.settings.sides_option == "单面":
                    devmode.Duplex = win32con.DMDUP_SIMPLEX
                elif task.settings.sides_option == "双面长边":
                    devmode.Duplex = win32con.DMDUP_VERTICAL
                elif task.settings.sides_option == "双面短边":
                    devmode.Duplex = win32con.DMDUP_HORIZONTAL
                
                if task.settings.color_mode == "彩色":
                    devmode.Color = 1
                else:
                    devmode.Color = 2
                
                try:
                    win32print.SetPrinter(handle, 2, properties, 0)
                except:
                    pass
                
                # 打印文件
                for _ in range(task.settings.copies):
                    if task.status == PrintStatus.CANCELLED:
                        break
                    
                    while task.status == PrintStatus.PAUSED:
                        time.sleep(0.1)
                    
                    try:
                        with open(task.file_path, 'rb') as f:
                            data = f.read()
                        
                        job = win32print.StartDocPrinter(handle, 1, (task.file_name, None, "RAW"))
                        try:
                            win32print.StartPagePrinter(handle)
                            win32print.WritePrinter(handle, data)
                            win32print.EndPagePrinter(handle)
                            task.update_progress(1, 1)  # 临时方案，后续添加实际页数统计
                        finally:
                            win32print.EndDocPrinter(handle)
                    
                    except Exception as e:
                        raise Exception(f"打印失败: {str(e)}")
                
                if task.status != PrintStatus.CANCELLED:
                    self.print_queue.complete_current_task()
                
            finally:
                win32print.ClosePrinter(handle)
        
        except Exception as e:
            self.print_queue.fail_current_task(str(e))
        
        # 继续处理下一个任务
        QTimer.singleShot(100, self.process_print_queue)
    
    def get_printer_status(self, printer_name):
        try:
            handle = win32print.OpenPrinter(printer_name)
            try:
                info = win32print.GetPrinter(handle, 2)
                return info['Status']
            finally:
                win32print.ClosePrinter(handle)
        except:
            return 0
            
    def update_printer_list(self):
        self.printer_combo.clear()
        self.printers = []
        
        # 获取打印机列表和状态
        for printer in win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS):
            name = printer[2]
            status = self.get_printer_status(name)
            self.printers.append((name, status))
            
            # 设置ComboBox项，包含名称和状态
            self.printer_combo.addItem(name)
            index = self.printer_combo.count() - 1
            self.printer_combo.setItemData(index, (name, status), Qt.ItemDataRole.UserRole)
            
        # 设置默认打印机
        self.default_printer = win32print.GetDefaultPrinter()
        default_index = next((i for i, p in enumerate(self.printers) if p[0] == self.default_printer), 0)
        self.printer_combo.setCurrentIndex(default_index)
        
    def update_printer_status(self):
        for i in range(self.printer_combo.count()):
            name = self.printer_combo.itemText(i)
            status = self.get_printer_status(name)
            self.printer_combo.setItemData(i, (name, status), Qt.ItemDataRole.UserRole)
        self.printer_combo.update()
    
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
    
    def dropEvent(self, event):
        files = []
        for url in event.mimeData().urls():
            file_path = url.toLocalFile()
            if os.path.isfile(file_path):
                files.append(file_path)
            elif os.path.isdir(file_path):
                # 如果是文件夹，设置为当前文件夹并加载文件
                self.path_input.setText(file_path)
                self.load_file_list()
                return
        
        if files:
            # 如果拖入的是文件，将它们复制到当前文件夹
            current_path = self.path_input.text()
            if not current_path:
                QMessageBox.warning(self, "警告", "请先选择目标文件夹！")
                return
            
            for file_path in files:
                file_name = os.path.basename(file_path)
                target_path = os.path.join(current_path, file_name)
                try:
                    import shutil
                    shutil.copy2(file_path, target_path)
                except Exception as e:
                    QMessageBox.warning(self, "错误", f"复制文件失败: {str(e)}")
            
            # 重新加载文件列表
            self.load_file_list()
    
    def on_page_range_changed(self, text):
        if text == "指定页面":
            self.page_range_input.show()
        else:
            self.page_range_input.hide()
    
    def on_file_selection_changed(self, file_name, state):
        self.file_selections[file_name] = (state == Qt.CheckState.Checked.value)
        
        # 更新全选按钮状态
        all_selected = all(self.file_selections.values())
        any_selected = any(self.file_selections.values())
        
        if all_selected:
            self.select_all_btn.setText("取消全选")
        elif any_selected:
            self.select_all_btn.setText("全选")
        else:
            self.select_all_btn.setText("全选")
    
    def toggle_select_all(self):
        # 检查当前是否全部选中
        all_selected = all(self.file_selections.values())
        
        # 切换所有文件的选中状态
        for item in self.file_items:
            checkbox = item.layout().itemAt(0).widget()
            checkbox.setChecked(not all_selected)
        
        # 更新按钮文本
        self.select_all_btn.setText("取消全选" if not all_selected else "全选")
    
    def filter_files(self):
        search_text = self.search_input.text().lower()
        
        for item in self.file_items:
            file_name = item.layout().itemAt(1).widget().text().lower()
            item.setVisible(search_text in file_name)
    
    def sort_files(self, sort_option, files=None):
        if files is None:
            # 获取当前显示的文件列表
            files = []
            for item in self.file_items:
                layout = item.layout()
                name = layout.itemAt(1).widget().text()
                ext = layout.itemAt(2).widget().text().lower()
                path = os.path.join(self.path_input.text(), name)
                size = os.path.getsize(path)
                files.append({
                    'name': name,
                    'ext': f".{ext}",
                    'size': size,
                    'path': path
                })
        
        # 根据选项排序
        if sort_option == "名称升序":
            files.sort(key=lambda x: x['name'].lower())
        elif sort_option == "名称降序":
            files.sort(key=lambda x: x['name'].lower(), reverse=True)
        elif sort_option == "类型升序":
            files.sort(key=lambda x: x['ext'].lower())
        elif sort_option == "类型降序":
            files.sort(key=lambda x: x['ext'].lower(), reverse=True)
        elif sort_option == "大小升序":
            files.sort(key=lambda x: x['size'])
        elif sort_option == "大小降序":
            files.sort(key=lambda x: x['size'], reverse=True)
        
        if files is not None:
            return
            
        # 重新排列文件列表
        for item in self.file_items:
            self.list_layout.removeWidget(item)
            item.hide()
        
        # 按新顺序添加文件
        for file_info in files:
            for item in self.file_items:
                if item.layout().itemAt(1).widget().text() == file_info['name']:
                    self.list_layout.addWidget(item)
                    item.show()
                    break
    
    def show_queue_window(self):
        if not self.queue_window:
            self.queue_window = PrintQueueWindow(self.print_queue, self)
        self.queue_window.show()
        self.queue_window.activateWindow()  # 将窗口提升到最前

if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")  # 使用Fusion风格
    window = PrinterApp()
    window.show()
    sys.exit(app.exec()) 