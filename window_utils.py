import win32gui, os

def get_window_list():
    windows = []

    def callback(hwnd, extra):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
            windows.append((win32gui.GetWindowText(hwnd), hwnd))

    win32gui.EnumWindows(callback, None)
    return windows

def image_files_count(folder_path,image_type):
    image_count = 0

    # 检查是否存在这个路径
    if not os.path.isdir(folder_path):
        print("❌ 路径无效或不存在")
        return 0

    # 遍历当前文件夹中的文件（不进入子文件夹）
    for file in os.listdir(folder_path):
        if os.path.isfile(os.path.join(folder_path, file)):
            _, ext = os.path.splitext(file.lower())
            if ext in image_type:
                image_count += 1

    return image_count
