from PyQt5.QtWidgets import QLabel
from PyQt5.QtWidgets import QPushButton
from PyQt5.QtWidgets import QSizePolicy
from PyQt5.QtCore import Qt

def create_styled_label(text: str, width: int = 120) -> QLabel:
    label = QLabel(text)
    label.setFixedWidth(width)
    label.setStyleSheet("""
        QLabel {
            color: #333;
            font-weight: bold;
            font-size: 11pt;
            padding-right: 6px;
        }
    """)
    return label

def create_styled_info(text: str, min_width: int = 150) -> QLabel:
    label = QLabel(text)
    label.setStyleSheet("""
        QLabel {
            color: gray;
            font-weight: bold;
            font-size: 10pt;
            padding-right: 6px;
            /* text-decoration: underline;    添加下划线 */
            border: 1px solid #aaa;

        }
    """)
    # 自动换行 + 自动高度
    label.setWordWrap(True)
    label.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Minimum)
    label.setMinimumHeight(0)
    # 允许选择复制文本（并保留链接点击能力）
    label.setTextInteractionFlags(Qt.TextSelectableByMouse | Qt.LinksAccessibleByMouse)
    if min_width:
        label.setMinimumWidth(min_width)
    return label

def style_preview_label(label: QLabel, width=200, height=150):
    label.setWindowFlags(Qt.ToolTip)
    label.setFixedSize(width, height)
    label.setAlignment(Qt.AlignCenter)
    label.setStyleSheet("""
        QLabel {
            border: 1px solid #888;
            background-color: #ffffff;
            padding: 4px;
            font-size: 10pt;
            color: #333;
            border-radius: 6px;
        }
    """)
    label.hide()

def style_input_widget(widget,min_width=150):

    widget.setStyleSheet(f"""
        /* 基础 QWidget 样式（适用于 QComboBox 本身） */
        QWidget {{
            background-color: #e6f7ff;   /* 浅蓝色背景 */
            border: 1px solid #aaa;      /* 灰色边框 */
            border-radius: 4px;          /* 圆角边框 */
            padding: 4px 1px;            /* 内边距，让文字不贴边 */
            font-size: 10pt;              /* 字体大小 */
            color: #005577;               /* 字体颜色 */
        }}
        
        /* 设置下拉菜单中列表的样式（QAbstractItemView 是下拉菜单的视图） */
        QComboBox QAbstractItemView {{
            outline: 0px;                /* 去掉选中项的虚线框 */
            border: 1px solid #ccc;      /* 设置列表边框 */
            border-radius: 4px;          /* 列表区域圆角 */
        }}
        
        /* 设置垂直滚动条的宽度及背景色 */
        QComboBox QScrollBar:vertical {{
            width: 20px;                 /* 滚动条宽度（关键设置） */
            background-color: #f0f0f0;    /* 滚动条背景色 */
            margin: 0px 0px 0px 0px;     /* 滚动条与内容之间无额外边距 */
        }}
        
        /* 设置垂直滚动条滑块部分的样式 */
        QComboBox QScrollBar::handle:vertical {{
            background: #cfcfcf;          /* 滑块颜色 */
            border-radius: 4px;          /* 滑块圆角，更美观 */
        }}
        
        /* 隐藏垂直滚动条的上下箭头按钮 */
        QComboBox QScrollBar::add-line, QComboBox QScrollBar::sub-line {{
            background: none;
            width: 0px;
            height: 0px;
            sub-line: none;
        }}
        
        /* 隐藏滚动条两侧空白区域的点击区域样式 */
        QComboBox QScrollBar::add-page, QComboBox QScrollBar::sub-page {{
            background: none;
        }}
    """)

    if min_width:
        widget.setMinimumWidth(min_width)
    return widget

def style_preview_widget(widget):
    widget.setStyleSheet("""
        QWidget {
            background-color: #FFFFFF;
            border: 0px solid Black;
            border-radius: 4px;
            padding: 4px 6px;
            font-size: 10pt;
            color: #005577;
        }
    """)
    return widget

def style_btn(btn: QPushButton, 
                 bg_color: str = "#12AD12", 
                 text_color: str = "white",
                 font_size: int = 10,
                 width: int = 200):
    """
    为 QPushButton 设置统一样式
    :param btn: 按钮对象
    :param bg_color: 背景色
    :param text_color: 字体颜色
    :param font_size: 字号
    :param width: 固定宽度（不设则自适应）
    """
    btn.setStyleSheet(f"""
        QPushButton {{
            background-color: {bg_color};
            color: {text_color};
            border: none;
            border-radius: 5px;
            padding: 6px 12px;
            padding-left: 0px;
            padding-right: 0px;
            font-size: {font_size}pt;
            font-weight: bold;
            text-align: center;
            /* 渐变实现立体感 */
            background-image: linear-gradient(to bottom, lighten({bg_color}, 15%), {bg_color});
        }}
        QPushButton:hover {{
            background-image: linear-gradient(to bottom, lighten({bg_color}, 25%), lighten({bg_color}, 5%));
        }}
        QPushButton:pressed {{
            background-image: none;
            background-color: {bg_color};
            padding-top: 7px; /* 按下微微下移模拟凹陷 */
        }}
    """)
    if width:
        btn.setFixedWidth(width)
    # if height:
    #     btn.setFixedHeight(height)
    return btn
