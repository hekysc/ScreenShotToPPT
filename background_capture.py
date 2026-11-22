"""Take a fresh WGC frame for one HWND. Call from a worker thread."""
import asyncio
import ctypes
import uuid
from contextlib import closing

from PIL import Image


def create_device():
    from winrt.windows.graphics.directx.direct3d11.interop import create_direct3d11_device_from_dxgi_device

    pointer = ctypes.c_void_p
    create = ctypes.windll.d3d11.D3D11CreateDevice
    create.argtypes = [pointer, ctypes.c_uint, pointer, ctypes.c_uint, pointer,
                       ctypes.c_uint, ctypes.c_uint, ctypes.POINTER(pointer), pointer, pointer]
    device = pointer()
    # BGRA support, D3D11 SDK version 7. Try hardware, then software (WARP).
    for driver in (1, 5):
        result = create(None, driver, None, 0x20, None, 0, 7, ctypes.byref(device), None, None)
        if result >= 0:
            break
    else:
        raise OSError(f"D3D11CreateDevice: {result:#x}")

    def method(obj, index, returns, *args):
        vtable = ctypes.cast(obj, ctypes.POINTER(ctypes.POINTER(pointer))).contents
        return ctypes.WINFUNCTYPE(returns, pointer, *args)(vtable[index])

    iid = uuid.UUID("54ec77fa-1377-44e6-8c32-88fd5f44c84c").bytes_le  # IDXGIDevice
    dxgi = pointer()
    try:
        query = method(device, 0, ctypes.c_long, ctypes.c_char_p, ctypes.POINTER(pointer))
        result = query(device, iid, ctypes.byref(dxgi))
        if result < 0:
            raise OSError(f"QueryInterface(IDXGIDevice): {result:#x}")
        try:
            return create_direct3d11_device_from_dxgi_device(dxgi.value)
        finally:
            method(dxgi, 2, ctypes.c_ulong)(dxgi)
    finally:
        method(device, 2, ctypes.c_ulong)(device)


async def read_frame(hwnd):
    import win32gui
    from winrt.windows.graphics.capture import Direct3D11CaptureFramePool, GraphicsCaptureSession
    from winrt.windows.graphics.capture.interop import create_for_window
    from winrt.windows.graphics.directx import DirectXPixelFormat
    from winrt.windows.graphics.imaging import SoftwareBitmap
    from winrt.windows.storage.streams import Buffer

    if not win32gui.IsWindow(hwnd):
        raise ValueError("窗口已关闭，请刷新窗口列表")
    if win32gui.IsIconic(hwnd):
        raise ValueError("窗口已最小化，恢复后可被其他窗口遮挡")
    if not GraphicsCaptureSession.is_supported():
        raise RuntimeError("系统不支持 Windows 图形捕获，可选择后台窗口（兼容 GDI）")
    item = create_for_window(hwnd)
    fmt = DirectXPixelFormat.B8_G8_R8_A8_UINT_NORMALIZED
    with closing(create_device()) as device, closing(
        Direct3D11CaptureFramePool.create_free_threaded(device, fmt, 1, item.size)
    ) as pool, closing(pool.create_capture_session(item)) as session:
        session.start_capture()
        while True:
            frame = pool.try_get_next_frame()
            if frame is not None:
                break
            await asyncio.sleep(0.015)
        with closing(frame):
            width, height = frame.content_size.width, frame.content_size.height
            with closing(await SoftwareBitmap.create_copy_from_surface_async(frame.surface)) as bitmap:
                size = bitmap.pixel_width, bitmap.pixel_height
                if not (0 < width <= size[0] and 0 < height <= size[1]):
                    raise ValueError("窗口尺寸正在变化，请稍候重试")
                buffer = Buffer(size[0] * size[1] * 4)
                bitmap.copy_to_buffer(buffer)
                image = Image.frombytes("RGB", size, bytes(buffer), "raw", "BGRX")
                return image.crop((0, 0, width, height)) if size != (width, height) else image


def capture_background(hwnd):
    from winrt.runtime import ApartmentType, init_apartment, uninit_apartment

    init_apartment(ApartmentType.MULTI_THREADED)
    try:
        return asyncio.run(asyncio.wait_for(read_frame(hwnd), timeout=3))
    except asyncio.TimeoutError as error:
        raise RuntimeError("等待窗口新画面超时") from error
    finally:
        uninit_apartment()
