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
    widget.setStyleSheet("""
        QWidget {
            background-color: #e6f7ff;
            border: 1px solid #aaa;
            border-radius: 4px;
            padding: 4px 6px;
            font-size: 10pt;
            color: #005577;
        }
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
