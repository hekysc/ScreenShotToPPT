import win32gui

def get_window_list():
    windows = []

    def callback(hwnd, extra):
        if win32gui.IsWindowVisible(hwnd) and win32gui.GetWindowText(hwnd):
            windows.append((win32gui.GetWindowText(hwnd), hwnd))

    win32gui.EnumWindows(callback, None)
    return windows
