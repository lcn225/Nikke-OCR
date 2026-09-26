@echo off
chcp 65001 >nul
setlocal
title NIKKE OCR 采集

rem ---- 切到本脚本所在目录（提权后的子进程也靠这行纠正工作目录）----
cd /d "%~dp0"

rem ============================================================
rem  自我提权
rem  游戏以管理员身份运行时，普通权限的模拟点击会被系统
rem  (UIPI) 静默丢弃，所以必须提权后再跑脚本。
rem  路径内联进 PowerShell 单引号字符串：本机控制台是 936
rem  代码页，中文路径经此转换已验证逐字节正确。
rem ============================================================
net session >nul 2>&1
if errorlevel 1 (
    echo.
    echo   需要管理员权限（游戏是管理员运行的），正在请求提权...
    echo   请在 UAC 弹窗里点「是」。
    echo.
    powershell -NoProfile -ExecutionPolicy Bypass -Command "Start-Process -FilePath '%~f0' -Verb RunAs"
    exit /b
)

where python >nul 2>&1
if errorlevel 1 (
    echo.
    echo   [错误] PATH 里找不到 python。
    echo   请确认 Python 已安装，并且在系统 PATH 中。
    echo.
    pause
    exit /b 1
)

:MENU
cls
color 0B
echo ============================================================
echo    妮姬 OCR 国服采集
echo ============================================================
echo.
echo    目录 : %CD%
echo    权限 : 管理员
echo.
echo    [1] 试跑 6 人       --max 6 --save-shots   约 5 分钟
echo    [2] 全量采集        无参数                 最长约 40 分钟
echo    [3] 离线验证        --offline              用本地截图，不碰屏幕
echo    [4] 检查页面状态    --check-state          只截图判页，不加载 OCR
echo    [5] 校正数据        correct_gui.py         改名字/战力/装备，导出 JSON
echo    [0] 退出
echo.
choice /c 123450 /n /m "请选择: "

if errorlevel 6 goto :END
if errorlevel 5 goto :M5
if errorlevel 4 goto :M4
if errorlevel 3 goto :M3
if errorlevel 2 goto :M2
if errorlevel 1 goto :M1
goto :MENU

:M1
set "ARGS=--max 6 --save-shots"
goto :RUN

:M2
set "ARGS="
goto :RUN

:M3
set "ARGS=--offline"
goto :RUN

:M4
set "ARGS=--check-state"
goto :RUN

rem 校正界面是另一个脚本、不吃 %ARGS%，所以单独一段。
rem 这里的提权对它没有影响（改 JSON 不需要管理员），从别处直接跑也行。
:M5
echo.
echo ------------------------------------------------------------
echo    运行 : python correct_gui.py
echo ------------------------------------------------------------
echo.
python correct_gui.py
set "RC=%errorlevel%"
echo.
echo ------------------------------------------------------------
echo    结束，退出码 %RC%
echo ------------------------------------------------------------
echo    按任意键返回菜单，或直接关掉本窗口
pause >nul
goto :MENU

:RUN
echo.
echo ------------------------------------------------------------
echo    运行 : python collect_cn.py %ARGS%
echo ------------------------------------------------------------
echo.
python collect_cn.py %ARGS%
set "RC=%errorlevel%"
echo.
echo ------------------------------------------------------------
echo    结束，退出码 %RC%
echo ------------------------------------------------------------
echo    按任意键返回菜单，或直接关掉本窗口
pause >nul
goto :MENU

:END
endlocal
