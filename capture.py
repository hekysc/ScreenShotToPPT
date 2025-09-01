import win32gui, win32ui, win32con,win32process
from PIL import Image, ImageChops,ImageDraw,ImageFont,ImageStat

def get_window_list():
    windows = []
    def callback(hwnd, extra):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
            windows.append((win32gui.GetWindowText(hwnd), hwnd))
    win32gui.EnumWindows(callback, None)
    return windows

def capture_window(hwnd):
    try:
        left, top, right, bottom = win32gui.GetWindowRect(hwnd)
        width = right - left
        height = bottom - top

        hwindc = win32gui.GetWindowDC(hwnd)
        srcdc = win32ui.CreateDCFromHandle(hwindc)
        memdc = srcdc.CreateCompatibleDC()

        bmp = win32ui.CreateBitmap()
        bmp.CreateCompatibleBitmap(srcdc, width, height)
        memdc.SelectObject(bmp)
        memdc.BitBlt((0, 0), (width, height), srcdc, (0, 0), win32con.SRCCOPY)

        bmpinfo = bmp.GetInfo()
        bmpstr = bmp.GetBitmapBits(True)

        img = Image.frombuffer(
            'RGB',
            (bmpinfo['bmWidth'], bmpinfo['bmHeight']),
            bmpstr, 'raw', 'BGRX', 0, 1
        )

        win32gui.DeleteObject(bmp.GetHandle())
        memdc.DeleteDC()
        srcdc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwindc)

        return img
    except:
        return None

def is_different(img1, img2, threshold=10):
    diff = ImageChops.difference(img1, img2).convert("L")
    hist = diff.histogram()
    score = sum(i * hist[i] for i in range(256)) / (img1.size[0] * img1.size[1])
    return score > threshold

def create_placeholder_image(size=(200,150),text="Fail",color=(200, 200, 200)):
    img = Image.new("RGB", size, color=color)
    draw = ImageDraw.Draw(img)
    # print(size)

    # 使用默认字体居中绘制文字
    try:
        font = ImageFont.truetype("arial.ttf", 16)
    except:
        font = ImageFont.load_default()

    bbox = draw.textbbox((0, 0), text, font=font)
    text_w, text_h = bbox[2] - bbox[0], bbox[3] - bbox[1]
    position = ((size[0] - text_w) // 2, (size[1] - text_h) // 2)
    draw.text(position, text, fill=(255,255,255), font=font)
    return img

def is_image_black(pil_image):
    """判断图片是否为全黑"""
    if pil_image.mode != "RGB":
        pil_image = pil_image.convert("RGB")
    stat = ImageStat.Stat(pil_image)
    # 如果所有通道平均值都非常低，则视为全黑
    return all(channel < 10 for channel in stat.mean)
