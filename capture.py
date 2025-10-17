import win32gui, win32ui, win32con
from PIL import Image, ImageChops,ImageDraw,ImageFont,ImageStat
import ctypes

def get_window_list():
    windows = []
    exclude_keywords = ["EndpointClientGUI", "Wemail"]

    def callback(hwnd, extra):
        if win32gui.IsWindowVisible(hwnd):
            window_title=win32gui.GetWindowText(hwnd)
            if not window_title:
                return #跳过空标题窗口
            
            for keyword in exclude_keywords:
                if keyword.lower() in window_title.lower():
                    return  # 包含排除关键词，跳过该窗口
                    
            windows.append((window_title, hwnd))
    
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

        # 如果抓到黑图，再尝试 PrintWindow（用ctypes实现）
        if is_image_black(img):
            PW_RENDERFULLCONTENT = 0x00000002
            result = ctypes.windll.user32.PrintWindow(hwnd, memdc.GetSafeHdc(), PW_RENDERFULLCONTENT)
            bmpinfo = bmp.GetInfo()
            bmpstr = bmp.GetBitmapBits(True)
            img = Image.frombuffer('RGB', (bmpinfo['bmWidth'], bmpinfo['bmHeight']), bmpstr, 'raw', 'BGRX', 0, 1)
            # 可以根据 result 做进一步判断

        win32gui.DeleteObject(bmp.GetHandle())
        memdc.DeleteDC()
        srcdc.DeleteDC()
        win32gui.ReleaseDC(hwnd, hwindc)

        # 情况 1: 截图正常
        if img and img.width > 30 and img.height > 30 and not is_image_black(img):
            status_flag="Succeed"
        
        # 情况 2: 截图为 None 或尺寸过小
        elif img is None or img.width <= 30 or img.height <= 30:
            placeholder = create_placeholder_image(text="窗口可能被最小化\n或其它原因，无法截图", color=(200, 0, 0))  # 背景
            img=placeholder
            status_flag="fail"

        # 情况 3: 全黑图
        elif is_image_black(img):
            placeholder = create_placeholder_image(text="窗口可能被最小化\n或其它原因，全黑", color=(150, 0, 0))  # 背景
            img=placeholder
            status_flag="fail"

        return img,status_flag
    except Exception as e:
        # print("Capture error:", e)
        placeholder = create_placeholder_image(text=str(e), color=(200, 0, 0))  # 背景
        img=placeholder
        status_flag="fail"
        return img,status_flag

def get_diff_score(img1, img2, algo=1):
    # 逐象素对比
    if algo==1:
        diff = ImageChops.difference(img1, img2).convert("L")
        hist = diff.histogram()
        score = sum(i * hist[i] for i in range(256)) / (img1.size[0] * img1.size[1])
        return score

def create_placeholder_image(size=(200,150),text="Fail",color=(200, 200, 200)):
    img = Image.new("RGB", size, color=color)
    draw = ImageDraw.Draw(img)
    # print(size)

    try:
        # 指定支持中文的字体路径
        font = ImageFont.truetype("C:/Windows/Fonts/msyh.ttc", 16)
    except Exception as e:
        try:
            font = ImageFont.truetype("C:/Windows/Fonts/SimHei.ttf", 16)
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
