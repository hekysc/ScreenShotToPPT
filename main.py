import sys, os, time
import getpass
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout,
    QFileDialog, QComboBox, QTextEdit, QSpinBox, QListView
)
from PyQt5.QtCore import Qt, QTimer, QPoint, QSettings
from PyQt5.QtGui import QIcon
from capture import (
    get_window_list, capture_window, 
    is_different,create_placeholder_image,is_image_black
)
from ppt_generator import generate_ppt
from img_convertor import (
    hoverview_img_generator,
    preview_img_generator,
    iconview_img_generator,
    pil_image_to_qpixmap
)
from ui_style import (
    create_styled_label,style_input_widget,
    create_styled_info,style_btn,
    style_preview_widget
)
from PIL import Image
import ctypes
import re

class MainApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("窗口截图与PPT生成器")
        self.setWindowIcon(QIcon("icons/app_icon.png"))
        self.setGeometry(100, 100, 400, 600)

        self.last_image = None
        self.capture_folder = sys.path[0]
        self.capture_count = 0
        self.start_time = None
        self.timer = QTimer()
        self.timer.timeout.connect(self.capture_loop)

        # 新增：应用设置对象（组织名、应用名可自定义）
        username = re.sub(r'[^A-Za-z0-9_-]', '_', getpass.getuser())
        self.settings = QSettings(username, "ScreenShotToPPT")

        self.init_ui()

        # 新增：加载保存过的参数
        self.load_settings()

    def init_ui(self):
        layout = QVBoxLayout()

        # 截图窗口选择
        window_widget=QWidget()
        window_layout=QHBoxLayout(window_widget)
        window_layout.setContentsMargins(8, 8, 8, 8)

        window_label=create_styled_label("截图窗口")
        window_layout.addWidget(window_label)

        self.combo = QComboBox()
        self.img_manager = {}
        self.get_combo_list()

        style_input_widget(self.combo)
        window_layout.addWidget(self.combo)

        # 刷新窗口按钮
        self.widow_refresh_btn = QPushButton("刷新")
        style_btn(self.widow_refresh_btn)
        self.widow_refresh_btn.setToolTip("选择保存文件夹")
        self.widow_refresh_btn.setFixedWidth(28)
        self.widow_refresh_btn.clicked.connect(self.refresh_window_list)
        window_layout.addWidget(self.widow_refresh_btn)

        # 建立预览窗口（空白）
        self.hover_preview = HoverPreview()
        style_preview_widget(self.hover_preview)

        # 将icon和title填入列表窗口
        hover_view = HoverListView(self.img_manager, self.hover_preview)
        self.combo.setView(hover_view)

        # 截图间隔-创建容器
        interval_widget = QWidget()
        interval_layout = QHBoxLayout(interval_widget)
        interval_layout.setContentsMargins(8, 8, 8, 8)

        # 截图间隔-标签        
        interval_label = create_styled_label("截图间隔(s)")
        interval_layout.addWidget(interval_label)
        
        # 截图间隔-数值输入框 
        self.interval_spin = QSpinBox()
        self.interval_spin.setRange(1, 60)
        self.interval_spin.setValue(5)
        self.interval_spin.valueChanged.connect(lambda _: self.save_settings())
        style_input_widget(self.interval_spin)
        interval_layout.addWidget(self.interval_spin)

        # 文件夹选择容器
        folder_widget=QWidget()
        folder_layout = QHBoxLayout(folder_widget)
        
        # 标签
        folder_label=create_styled_label("截图保存文件夹")
        folder_layout.addWidget(folder_label)

        self.folder_path_info=create_styled_info(self.capture_folder)
        folder_layout.addWidget(self.folder_path_info)

        self.folder_btn = QPushButton("...")
        style_btn(self.folder_btn)
        self.folder_btn.setToolTip("选择保存文件夹")
        self.folder_btn.setFixedWidth(28)
        self.folder_btn.clicked.connect(self.select_folder)
        folder_layout.addWidget(self.folder_btn)

        # 控制按钮
        scr_shot_widget=QWidget()
        scr_shot_layout = QHBoxLayout(scr_shot_widget)

        self.start_btn = QPushButton("开始截图")
        style_btn(self.start_btn)
        self.start_btn.clicked.connect(self.start_capture)
        scr_shot_layout.addWidget(self.start_btn,alignment=Qt.AlignLeft)
        scr_shot_layout.setSpacing(20)  # 标签和按钮之间 8px 间距

        self.stop_btn = QPushButton("停止截图")
        style_btn(self.stop_btn,bg_color="#EC4630")
        self.stop_btn.clicked.connect(self.stop_capture)
        scr_shot_layout.addWidget(self.stop_btn,alignment=Qt.AlignLeft)

        # ppt按钮
        ppt_widget=QWidget()
        ppt_layout = QHBoxLayout(ppt_widget)        
        
        self.ppt_btn = QPushButton("生成PPT")
        style_btn(self.ppt_btn)
        self.ppt_btn.clicked.connect(self.generate_ppt)
        ppt_layout.addWidget(self.ppt_btn,alignment=Qt.AlignLeft)

        # 建立预览窗口
        self.preview = QLabel("截图预览")
        self.preview.setAlignment(Qt.AlignCenter)
        self.preview.setFixedHeight(300)
        self.preview.setStyleSheet("""
            border: 1px solid #ccc;
            background-color: #f9f9f9;
            font-size: 12pt;
            color: #888;
        """)

        # 日志输出
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.append(f"共发现窗口数: {len(self.img_manager)}")

        layout.addWidget(window_widget)
        layout.addWidget(interval_widget)
        layout.addWidget(folder_widget)
        layout.addWidget(scr_shot_widget)
        layout.addWidget(ppt_widget)
        layout.addWidget(self.preview)
        layout.addWidget(QLabel("执行日志:"))
        layout.addWidget(self.log)

        self.stop_btn.setVisible(False)

        self.setLayout(layout)
            
    def get_combo_list(self):
        self.combo.clear()
        self.img_manager.clear()

        for title, hwnd in get_window_list():
            img = capture_window(hwnd)
            # 情况 1: 截图正常
            if img and img.width > 30 and img.height > 30 and not is_image_black(img):
                img=img
                status_flag="succeed"
            
            # 情况 2: 截图为 None 或尺寸过小
            elif img is None or img.width <= 30 or img.height <= 30:
                placeholder = create_placeholder_image(text="Fail", color=(200, 0, 0))  # 背景
                img=placeholder
                status_flag="fail"

            # 情况 3: 全黑图
            elif is_image_black(img):
                placeholder = create_placeholder_image(text="Fail, Full BLACK", color=(150, 0, 0))  # 背景
                img=placeholder
                status_flag="fail"

            hoverview_img=hoverview_img_generator(img)
            preview_img=preview_img_generator(img)
            iconview_img=iconview_img_generator(img)

            self.img_manager[title] = [
                title, hwnd, img, preview_img, hoverview_img,  
                iconview_img, status_flag
                ]
            self.combo.addItem(iconview_img, title,status_flag)

        # 按 flag 排序（成功在前，失败在后）
        self.refresh_combo(sort_index=3, reverse=False)

    def log_message(self, msg):
        timestamp = time.strftime("%H:%M:%S")
        self.log.append(f"[{timestamp}] {msg}")

    def refresh_window_list(self):
        self.get_combo_list()
        hover_view = HoverListView(self.img_manager, self.hover_preview)
        self.combo.setView(hover_view)
        self.log_message(f"已刷新窗口列表")

    def select_folder(self):
        dialog = QFileDialog(self)
        dialog.setFileMode(QFileDialog.Directory)
        dialog.setOption(QFileDialog.ShowDirsOnly, True)
        dialog.setWindowTitle("选择截图保存文件夹")
        if dialog.exec_():
            folder = dialog.selectedFiles()[0]
            self.capture_folder = folder
            self.folder_path_info.setText(folder)
            self.log_message(f"截图将保存到: {folder}")
        # 新增：立刻保存
        self.save_settings()

    def start_capture(self):
        if not self.capture_folder:
            self.log_message("请先选择截图保存文件夹")
            return
        self.capture_count = 0
        self.start_time = time.time()
        self.last_image = None
        self.timer.start(self.interval_spin.value() * 1000)
        self.log_message("开始截图...")
        self.start_btn.setVisible(False)
        self.stop_btn.setVisible(True)

    def stop_capture(self):
        self.timer.stop()
        elapsed = time.time() - self.start_time if self.start_time else 0
        self.log_message(f"截图结束，耗时 {elapsed:.1f} 秒，共保存 {self.capture_count} 张截图")
        self.stop_btn.setVisible(False)
        self.start_btn.setVisible(True)

    def capture_loop(self):
        title = self.combo.currentText()
        hwnd = self.img_manager[title][1]
        img = capture_window(hwnd)

        if img is None or img.size[0] < 30 or img.size[1] < 30:
            img = Image.new("RGB", (50, 50), (255, 0, 0))
            self.log_message("截图失败，使用红色占位图")
        else:
            if self.last_image is None or is_different(self.last_image, img):
                self.capture_count += 1
                title = title.strip().replace(" ", "_").replace(":", "_")
                safe_title = re.sub(r'[\\/*?:"<>|]', "_", title)
                timestamp = time.strftime("%Y%m%d_%H%M%S")
                filename = f"{safe_title}_{timestamp}.png"
                path = os.path.join(self.capture_folder, filename)
                img.save(path)
                self.log_message(f"保存截图: {filename}")
                self.last_image = img
            else:
                self.log_message("截图内容相似，未保存")

        pixmap = pil_image_to_qpixmap(img)
        self.preview.setPixmap(pixmap.scaled(400, 300, Qt.KeepAspectRatio,Qt.SmoothTransformation))

    def generate_ppt(self):
        folder = self.capture_folder

        # 如果未指定截图文件夹，则让用户选择
        if not folder or not os.path.exists(folder):
            folder = QFileDialog.getExistingDirectory(self, "请选择截图文件夹")
            if not folder:
                self.log_message("未指定截图文件夹，操作取消")
                return

        # 获取截图文件列表
        files = sorted([
            f for f in os.listdir(folder)
            if f.lower().endswith(".png")
        ])
        if not files:
            self.log_message("截图文件夹中没有可用的图片")
            return

        # 选择保存路径
        path, _ = QFileDialog.getSaveFileName(self, "保存PPT", "", "PowerPoint (*.pptx)")
        if not path:
            self.log_message("未指定PPT保存路径，操作取消")
            return

        try:
            generate_ppt(folder, files, path)
            self.log_message(f"PPT已保存到: {path}")
        except Exception as e:
            self.log_message(f"PPT生成失败: {e}")

    def refresh_combo(self, sort_index=3, reverse=False):
        """
        按 QComboBox 每项数据元组的某个字段排序并刷新显示
        :param sort_index: 用于排序的字段索引，例如：
                        0=title, 1=hwnd, 2=image, 3=flag
        :param reverse: 是否倒序
        """
        # 1. 取出所有条目的 (text, userData) 二元组
        items_data = []
        for i in range(self.combo.count()):
            icon=self.combo.itemIcon(i)
            text = self.combo.itemText(i)
            data = self.combo.itemData(i)  # (title, hwnd, image多个, flag)
            items_data.append((icon,text, data))

        # 2. 按指定字段排序
        items_data.sort(key=lambda item: item[2][sort_index], reverse=reverse)

        # 3. 重建 QComboBox
        self.combo.clear()
        for icon,text, data in items_data:
            self.combo.addItem(icon,text, data)

    def save_settings(self):
        """保存上次参数：截图路径、间隔秒数"""
        try:
            self.settings.setValue("capture_folder", self.capture_folder)
            self.settings.setValue("interval_seconds", int(self.interval_spin.value()))
            self.settings.sync()  # 立即落盘
        except Exception as e:
            self.log_message(f"保存设置失败: {e}")

    def load_settings(self):
        """加载上次参数（若不存在则保持默认值）"""
        try:
            folder = self.settings.value("capture_folder", type=str)
            if folder and os.path.isdir(folder):
                self.capture_folder = folder
                self.folder_path_info.setText(folder)

            interval = self.settings.value("interval_seconds", type=int)
            if interval and 1 <= interval <= 60:
                self.interval_spin.setValue(interval)

            # 同步 UI（如果 load 比 init_ui 晚执行）
            # 已经在上面 setText/setValue 处理，无需额外操作
        except Exception as e:
            self.log_message(f"加载设置失败: {e}")

    def closeEvent(self, event):
        # 退出前保存一次
        self.save_settings()
        super().closeEvent(event)

# 悬浮显示鼠标响应类
class HoverListView(QListView):
    # def __init__(self, img_manager, preview_label):
    def __init__(self, img_manager, preview_widget):
        super().__init__()
        self.img_manager = img_manager
        # self.preview_label = preview_label
        self.preview_widget=preview_widget
        self.setMouseTracking(True)

    def mouseMoveEvent(self, event):
        try:
            index = self.indexAt(event.pos())
            if index.isValid():
                title = index.data()
                pixmap = self.img_manager.get(title)[4]
                if pixmap:
                    self.preview_widget.update_content(title, pixmap)
                    # Ensure widget size matches content before moving
                    self.preview_widget.adjustSize()
                    global_pos = self.viewport().mapToGlobal(event.pos())
                    self.preview_widget.move(global_pos + QPoint(20, 20))
                    self.preview_widget.show()
                else:
                    self.preview_widget.hide()
            else:
                self.preview_widget.hide()
        except Exception as e:
            print("悬浮预览失败:", e)
            self.preview_widget.hide()

        super().mouseMoveEvent(event)

    def leaveEvent(self, event):
        self.preview_widget.hide()
        super().leaveEvent(event)

# 空白悬浮窗口类，允许更新显示内容
class HoverPreview(QWidget):
    def __init__(self, width=200, height=150):
        super().__init__()
        self.setWindowFlags(Qt.ToolTip)

        self.title_label = QLabel()
        self.title_label.setWordWrap(True)
        self.title_label.setAlignment(Qt.AlignCenter)
        # self.title_label.setStyleSheet("""
        #     QLabel {
        #         font-size: 9pt;
        #         color: #333;
        #         background-color: #fff;
        #         padding: 2px;
        #     }
        # """)

        self.image_label = QLabel()
        self.image_label.setAlignment(Qt.AlignCenter)
        # self.image_label.setStyleSheet("""
        #     QLabel {
        #         border: 1px solid #888;
        #         background-color: #333;
        #         border-radius: 6px;
        #     }
        # """)

        layout = QVBoxLayout()
        layout.setContentsMargins(0,0,0,0)
        layout.setSpacing(0)
        layout.addWidget(self.title_label)
        layout.addWidget(self.image_label)
        self.setLayout(layout)

        self.hide()

    def update_content(self, title: str, pixmap):
        # Update labels then adjust to content so geometry matches size hints
        self.title_label.setText(title)
        self.image_label.setPixmap(pixmap)
        self.title_label.adjustSize()
        self.image_label.adjustSize()
        self.adjustSize()

ctypes.windll.user32.SetProcessDPIAware()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    sys.exit(app.exec_())
