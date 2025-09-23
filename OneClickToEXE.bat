@echo off
chcp 65001 >nul

echo 即将打包main.py文件，是否继续？
choice /C YN /M "按 Y 继续，按 N 取消"
if /I "%errorlevel%"=="1" goto continue
if /I "%errorlevel%"=="2" goto end

:continue
echo 正在执行操作...
pyinstaller main.py -F -i .\icon\icon.ico -w -n ScreenShotToPPT
goto end

:end
echo 操作结束。
