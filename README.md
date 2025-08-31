# ScreenShotToPPT
Automatically copy screen periodically, even the window is not activatied.

# ScreenShotToPPT

一个基于 PyQt5 的 Windows 窗口截图工具，支持：
- 后台窗口截图（尝试使用 Windows PrintWindow）
- 定时自动采集
- 相似度去重（phash）
- 一键导出为 16:9 PowerPoint

> 仅支持 Windows（使用 pywin32），推荐 Python 3.9+。

## 快速开始

```bash
pip install -r requirements.txt
python main.py
