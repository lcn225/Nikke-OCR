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
echo    [2] 采集（参数可选）                        上下勾选 --patch/--max/--save-shots
echo    [3] 离线验证        --offline              用本地截图，不碰屏幕
echo    [4] 检查页面状态    --check-state          只截图判页，不加载 OCR
echo    [5] 增量重扫        --patch --max N        升级了个别角色，只重扫这几个
echo    [6] 校正数据        correct_gui.py         改名字/战力/装备、合并增量补丁
echo    [0] 退出
echo.
choice /c 1234560 /n /m "请选择: "

if errorlevel 7 goto :END
if errorlevel 6 goto :M6
if errorlevel 5 goto :M5
if errorlevel 4 goto :M4
if errorlevel 3 goto :M3
if errorlevel 2 goto :M2
if errorlevel 1 goto :M1
goto :MENU

:M1
set "ARGS=--max 6 --save-shots"
goto :RUN

rem 参数用 pick_params.py 以 ↑↓/空格/回车 勾选，选完它自己就把 collect_cn.py 跑起来 ——
rem 同一个控制台，采集进度和 Ctrl+C 都留在原处。这样换个参数不必再单开一个管理员
rem PowerShell（提权这件事本脚本开头已经做完了）。
rem 为什么不在 bat 里做勾选：choice 只能读固定单键、set /p 必须回车，做方向键循环要靠
rem xcopy 取键之类的歪招；Python 是本项目本来就有的依赖，用它干净得多。
:M2
echo.
echo ------------------------------------------------------------
echo    采集（参数可选）
echo ------------------------------------------------------------
echo.
python pick_params.py
set "RC=%errorlevel%"
echo.
if "%RC%"=="1" echo    （已取消，没有采集）
pause >nul
goto :MENU

:M3
set "ARGS=--offline"
goto :RUN

:M4
set "ARGS=--check-state"
goto :RUN

rem 增量重扫：结果写 output/cn_patch.json，主数据不动（要回校正界面点「合并补丁」才并进去）。
rem 采集端没有跳转能力（只有点「>>」逐格前进一条路），所以"从哪个角色开始"由你在游戏里翻。
:M5
echo.
echo ------------------------------------------------------------
echo    增量重扫（升级了少数角色的装备后，不用全量重来）
echo ------------------------------------------------------------
echo    1) 先在游戏里手动翻到**第一个**要重扫的角色，停在他的信息页
echo    2) 回来填「扫几个」和「都是谁」，名字按扫描顺序对应
echo       名字可留空（那就只用 OCR 读到的名字，可能截断）
echo       逗号分隔，别加空格；全角「，」也认
echo.
set "N="
set /p "N=重扫几个角色（直接回车 = 取消）: "
if not defined N goto :MENU
set "NAMES="
set /p "NAMES=角色名（逗号分隔，回车 = 不填）: "
set "ARGS=--patch --max %N%"
if defined NAMES set ARGS=%ARGS% --as "%NAMES%"
goto :RUN

rem 校正界面是另一个脚本、不吃 %ARGS%，所以单独一段。
rem 这里的提权对它没有影响（改 JSON 不需要管理员），从别处直接跑也行。
:M6
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
