from PyQt5.QtGui import QPixmap, QImage,QIcon
from PIL import Image
from io import BytesIO

def hoverview_img_generator(pil_image,size=(200, 150)):
    img = pil_image
    thumbnail = img.resize(size)
    pixmap = pil_image_to_qpixmap(thumbnail)

    return pixmap

def preview_img_generator(pil_image,size=(200, 150)):
    img = pil_image
    thumbnail = img.resize(size)
    pixmap = pil_image_to_qpixmap(thumbnail)

    return pixmap

def iconview_img_generator(pil_image,size=(200, 150)):
    img=pil_image
    pixmap=pil_image_to_qpixmap(img.resize(size))
    icon = QIcon(pixmap)
    
    return icon

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
