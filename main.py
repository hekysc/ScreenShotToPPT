import sys, os, time
from PyQt5.QtWidgets import (
    QApplication, QWidget, QLabel, QPushButton, QVBoxLayout, QHBoxLayout,
    QFileDialog, QComboBox, QTextEdit, QSpinBox, QListView,QScrollArea,QGridLayout
)
from PyQt5.QtCore import Qt, QTimer, QPoint
from PyQt5.QtGui import QPixmap, QImage,QIcon
from capture import (
    get_window_list, capture_window, 
    is_different,create_placeholder_image,get_window_list_new,is_image_black
)
from ppt_generator import generate_ppt
from ui_style import (
    create_styled_label,style_input_widget,
    create_styled_info,style_btn,
    style_preview_widget
)
from PIL import Image, ImageQt
import ctypes
import re
from io import BytesIO

class MainApp(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("窗口截图与PPT生成器")
        self.setWindowIcon(QIcon("icons/app_icon.png"))
        self.setGeometry(100, 100, 400, 600)

        self.last_image = None
        self.capture_folder = "..\\test"
        self.capture_count = 0
        self.start_time = None
        self.timer = QTimer()
        self.timer.timeout.connect(self.capture_loop)

        self.init_ui()

    def init_ui(self):
        layout = QVBoxLayout()

        # 截图窗口选择
        window_widget=QWidget()
        window_layout=QHBoxLayout(window_widget)
        window_layout.setContentsMargins(8, 8, 8, 8)

        window_label=create_styled_label("截图窗口")
        window_layout.addWidget(window_label)

        self.combo = QComboBox()
        self.window_map = {}

        self.thumb_manager = SafeThumbnailManager()
        # self.succeed_thumb=SafeThumbnailManager()
        # self.failed_thumb=SafeThumbnailManager()

        for title, hwnd in get_window_list():
            # print(title)
            img = capture_window(hwnd)
            if img and img.width > 30 and img.height > 30 and not is_image_black(img):
                img=img
                status_flag="succeed"
                # self.succeed_thumb.add(title,img)
                # self.thumb_manager.add(title, img)
                # self.add_window_item(title, hwnd, img)
                # self.window_map[title] = hwnd
            # 情况 2: 截图为 None 或尺寸过小
            elif img is None or img.width <= 30 or img.height <= 30:
                # print(f"[无法截图] {title}")
                placeholder = create_placeholder_image(text="Fail", color=(200, 0, 0))  # 背景
                img=placeholder
                status_flag="fail"
                # self.failed_thumb.add(title,img)
                # self.thumb_manager.add(title, placeholder)
                # self.add_window_item(title, hwnd, placeholder)
                # self.window_map[title] = hwnd
                # 情况 3: 全黑图
            elif is_image_black(img):
                placeholder = create_placeholder_image(text="Fail, Full BLACK", color=(150, 0, 0))  # 背景
                img=placeholder
                status_flag="fail"
                # self.failed_thumb.add(title,img)
                # self.thumb_manager.add(title, placeholder)
                # self.add_window_item(title, hwnd, placeholder)
                # self.window_map[title] = hwnd
            self.thumb_manager.add(title, img)
            # self.thumb_manager=self.succeed_thumb+self.failed_thumb
            self.add_window_item(title, hwnd, img,status_flag)
            self.window_map[title] = hwnd

        # 按 flag 排序（成功在前，失败在后）
        self.refresh_combo(sort_index=3, reverse=False)

        style_input_widget(self.combo)
        window_layout.addWidget(self.combo)

        # 预览窗口（空白）
        self.hover_preview = HoverPreview()
        style_preview_widget(self.hover_preview)

        # 列表窗口填充
        hover_view = HoverListView(self.thumb_manager, self.hover_preview)
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
        style_input_widget(self.interval_spin)
        interval_layout.addWidget(self.interval_spin)

        # 文件夹选择
        folder_widget=QWidget()
        folder_layout = QHBoxLayout(folder_widget)
        
        folder_label=create_styled_label("截图保存文件夹")
        folder_layout.addWidget(folder_label)
        
        self.folder_btn = QPushButton("...")
        style_btn(self.folder_btn)
        self.folder_btn.setToolTip("选择保存文件夹")
        self.folder_btn.setFixedWidth(28)
        self.folder_btn.clicked.connect(self.select_folder)
        folder_layout.addWidget(self.folder_btn)

        self.folder_path_info=create_styled_info(self.capture_folder)
        # self.folder_path_label.setStyleSheet("color: gray; font-size: 10pt;")
        folder_layout.addWidget(self.folder_path_info)

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

        # 日志输出
        self.log = QTextEdit()
        self.log.setReadOnly(True)
        self.log.append(f"共发现窗口数: {len(self.window_map)}")
        # print(self.window_map)

        # 图片预览
        self.preview = QLabel("截图预览")
        self.preview.setAlignment(Qt.AlignCenter)
        self.preview.setFixedHeight(300)
        self.preview.setStyleSheet("""
            border: 1px solid #ccc;
            background-color: #f9f9f9;
            font-size: 12pt;
            color: #888;
        """)

        layout.addWidget(window_widget)
        layout.addWidget(interval_widget)
        layout.addWidget(folder_widget)
        # layout.addWidget(self.folder_btn)
        layout.addWidget(scr_shot_widget)
        # layout.addWidget(self.stop_btn)
        layout.addWidget(ppt_widget)
        layout.addWidget(self.preview)
        layout.addWidget(QLabel("执行日志:"))
        layout.addWidget(self.log)

        self.stop_btn.setVisible(False)

        self.setLayout(layout)
            
    def add_window_item(self,title: str, hwnd, image, flag):

        # qt_img = ImageQt.ImageQt(image.resize((64, 48)))  # 缩略图尺寸
        # pixmap = QPixmap.fromImage(QImage(qt_img))
        pixmap=pil_image_to_qpixmap(image.resize((64, 48)))
        icon = QIcon(pixmap)
        self.combo.addItem(icon, title,flag)

    def log_message(self, msg):
        timestamp = time.strftime("%H:%M:%S")
        self.log.append(f"[{timestamp}] {msg}")

    def select_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "选择截图保存文件夹")
        if folder:
            self.capture_folder = folder
            self.folder_path_info.setText(folder)
            self.log_message(f"截图将保存到: {folder}")

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
        hwnd = self.window_map.get(title)
        img = capture_window(hwnd)

        if img is None or img.size[0] < 30 or img.size[1] < 30:
            img = Image.new("RGB", (50, 50), (255, 0, 0))
            self.log_message("截图失败，使用红色占位图")
        else:
            if self.last_image is None or is_different(self.last_image, img):
                self.capture_count += 1
                title = self.combo.currentText().strip().replace(" ", "_").replace(":", "_")
                safe_title = re.sub(r'[\\/*?:"<>|]', "_", title)
                timestamp = time.strftime("%Y%m%d_%H%M%S")
                filename = f"{safe_title}_{timestamp}.png"
                path = os.path.join(self.capture_folder, filename)
                img.save(path)
                self.log_message(f"保存截图: {filename}")
                self.last_image = img
            else:
                self.log_message("截图内容相似，未保存")

        # qt_img = ImageQt.ImageQt(img)
        pixmap = pil_image_to_qpixmap(img)
        self.preview.setPixmap(pixmap.scaled(400, 300, Qt.KeepAspectRatio,Qt.SmoothTransformation))
        # print("caputure:",type(img),"___pixmap：",type(pixmap))

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
            text = self.combo.itemText(i)
            data = self.combo.itemData(i)  # (title, hwnd, image, flag)
            items_data.append((text, data))

        # 2. 按指定字段排序
        items_data.sort(key=lambda item: item[1][sort_index], reverse=reverse)

        # 3. 重建 QComboBox
        self.combo.clear()
        for text, data in items_data:
            self.combo.addItem(text, data)

# 悬浮显示
class HoverListView(QListView):
    # def __init__(self, thumb_manager, preview_label):
    def __init__(self, thumb_manager, preview_widget):
        super().__init__()
        self.thumb_manager = thumb_manager
        # self.preview_label = preview_label
        self.preview_widget=preview_widget
        self.setMouseTracking(True)

    def mouseMoveEvent(self, event):
        try:
            index = self.indexAt(event.pos())
            if index.isValid():
                title = index.data()
                # pixmap = self.thumb_map.get(title)
                pixmap = self.thumb_manager.get(title)
                # safe_title = re.sub(r'[\\/*?:"<>|]', "_", title)
                # pixmap.save(f"{safe_title}.png")
                # print("thumb_mamager.get:",type(pixmap))
                if pixmap:
                    self.preview_widget.update_content(title, pixmap)
                    # safe_title = re.sub(r'[\\/*?:"<>|]', "_", title)
                    # pixmap.save(f"{safe_title}.png")
                    global_pos = self.viewport().mapToGlobal(event.pos())
                    self.preview_widget.move(global_pos + QPoint(20, 20))
                    self.preview_widget.show()
                else:
                    # self.preview_label.hide()
                    self.preview_widget.hide()
            else:
                # self.preview_label.hide()
                self.preview_widget.hide()
        except Exception as e:
            print("悬浮预览失败:", e)
            # self.preview_label.hide()
            self.preview_widget.hide()

        super().mouseMoveEvent(event)

    def leaveEvent(self, event):
        # self.preview_label.hide()
        self.preview_widget.hide()
        super().leaveEvent(event)

class SafeThumbnailManager:
    def __init__(self, size=(200, 150)):
        self.size = size
        self.pixmap_cache = {}  # title -> QPixmap

    def add(self, title: str, pil_image: Image.Image):
        try:
            if pil_image and pil_image.size[0] > 30 and pil_image.size[1] > 30:
                # print(title)
                # img = pil_image.convert("RGB")  # 避免透明通道问题
                img = pil_image
                thumbnail = img.resize(self.size)
                pixmap = pil_image_to_qpixmap(thumbnail)
                # qt_img = ImageQt.ImageQt(thumbnail)
                # pixmap = QPixmap.fromImage(QImage(qt_img)) # ⚠️ 可能不稳定
                # safe_title = re.sub(r'[\\/*?:"<>|]', "_", title)
                # pixmap.save(f"{safe_title}.png")
                # self.pixmap_cache[title] = pixmap
                if pixmap:
                    self.pixmap_cache[title] = pixmap
                else:
                    print(f"[缩略图转换失败] {title}")
                # safe_title = re.sub(r'[\\/*?:"<>|]', "_", title)
                # self.pixmap_cache[title].save(f"{safe_title}.png")
        except Exception as e:
            print(f"[缩略图处理失败] {title}: {e}")

    def get(self, title: str):
        # print("get:",type(self.pixmap_cache))
        # safe_title = re.sub(r'[\\/*?:"<>|]', "_", title)
        # self.pixmap_cache.get(title).save(f"{safe_title}.png")
        return self.pixmap_cache[title]
        # return self.pixmap_cache[title]

    def has(self, title: str):
        return title in self.pixmap_cache
    
    def get_status(self, title: str):
        return title in self.pixmap_cache

class HoverPreview(QWidget):
    def __init__(self, width=200, height=150):
        super().__init__()
        self.setWindowFlags(Qt.ToolTip)
        # self.setAttribute(Qt.WA_TranslucentBackground)

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
        # self.image_label.setFixedSize(width, height)
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
        self.title_label.setText(title)
        # qt_img = ImageQt.ImageQt(pixmap)
        # pixmap = QPixmap.fromImage(QImage(qt_img))
        self.image_label.setPixmap(pixmap)
        # print("update_content:",type(pixmap))

def pil_image_to_qpixmap(pil_img: Image.Image) -> QPixmap:
    try:
        pil_img = pil_img.convert("RGB")  # 强制 RGB，避免透明通道
        buffer = BytesIO()
        pil_img.save(buffer, format="PNG")
        qt_img = QImage.fromData(buffer.getvalue())
        return QPixmap.fromImage(qt_img)
    except Exception as e:
        print("图像转换失败:", e)
        return None

ctypes.windll.user32.SetProcessDPIAware()

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MainApp()
    window.show()
    sys.exit(app.exec_())
