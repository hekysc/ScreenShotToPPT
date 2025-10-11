@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion

:: Step 1: Get the Python script to package
set "script_file="
set /p script_file=请输入待打包的 Python 文件名（如 main.py）: 
if not exist "%script_file%" (
    echo 文件 "%script_file%" 不存在!
    pause
    goto end
)

:: Step 2: Get destination folder
set "dest_folder="
set /p dest_folder=请输入目标文件夹（如 dist）： 
if not exist "%dest_folder%" (
    echo 文件夹 "%dest_folder%" 不存在，将自动创建...
    mkdir "%dest_folder%"
)

:: Step 3: Get icon (optional)
set "icon_path="
set /p icon_path=请输入 exe 图标文件路径（如 .\icon\icon.ico，可留空）: 
if not "!icon_path!"=="" (
    if not exist "!icon_path!" (
        echo 图标文件 "!icon_path!" 不存在，跳过设置图标。
        set "icon_path="
    )
)

:: Step 4: PyInstaller 常用选项选择
echo.
echo 请选择需要的 PyInstaller 选项（可多选，用逗号分隔）:
echo [1] 单文件(-F)
echo [2] 单文件夹(-D)
echo [3] 隐藏命令行窗口(-w)
echo [4] 添加UPX压缩(--upx-dir)
echo [5] 包含全部依赖(--add-data)
echo [6] 生成调试EXE(-d)
echo [7] 不打包Python解释器(--no-embed)
echo [8] 其它（手动输入）
set "options="
set /p options=请输入选项序号（如 1,3,5）: 

set "py_opts="
for %%A in (%options%) do (
    if %%A==1 set py_opts=!py_opts! -F
    if %%A==2 set py_opts=!py_opts! -D
    if %%A==3 set py_opts=!py_opts! -w
    if %%A==4 (
        set /p upx_path=请输入UPX目录路径: 
        if not "!upx_path!"=="" (
            set py_opts=!py_opts! --upx-dir "!upx_path!"
        )
    )
    if %%A==5 (
        set /p add_data=请输入需要包含的数据文件或文件夹（格式：src;dest）: 
        if not "!add_data!"=="" (
            set py_opts=!py_opts! --add-data "!add_data!"
        )
    )
    if %%A==6 set py_opts=!py_opts! -d
    if %%A==7 set py_opts=!py_opts! --no-embed
    if %%A==8 (
        set /p custom_opts=请输入其它 PyInstaller 选项: 
        if not "!custom_opts!"=="" (
            set py_opts=!py_opts! !custom_opts!
        )
    )
)

:: 图标直接加入 py_opts
if not "!icon_path!"=="" (
    set py_opts=!py_opts! -i "!icon_path!"
)

:: Step 5: EXE Name
set "exe_name="
set /p exe_name=请输入生成的 exe 文件名（如 MyApp，可留空，默认为脚本文件名）: 
if "!exe_name!"=="" (
    for %%F in ("!script_file!") do set "exe_name=%%~nF"
)

echo.
echo ========= 打包配置 =========
echo 脚本文件: !script_file!
echo 目标文件夹: !dest_folder!
echo PyInstaller 选项: !py_opts!
echo EXE 名称: !exe_name!
echo ===========================
echo.

echo 是否继续执行此命令？
choice /C YN /M "按 Y 继续，按 N 取消"
if "!errorlevel!"=="2" goto end

:: Step 6: Build pyinstaller command
set "py_cmd=pyinstaller "!script_file!" -n "!exe_name!" --distpath "!dest_folder!" !py_opts!"

echo 正在执行操作...
echo !py_cmd!
!py_cmd!

echo.
echo 操作结束。
pause

:end
endlocal
