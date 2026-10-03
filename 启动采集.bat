@echo off
chcp 65001 >nul
setlocal
title NIKKE OCR

rem ---- cd to this script's own directory (the elevated child needs it too) ----
cd /d "%~dp0"

rem ============================================================
rem  Self-elevation.
rem  The game runs elevated; clicks from a normal-privilege process
rem  are dropped silently by UIPI. So we must elevate first.
rem  The path is inlined into a PowerShell single-quoted string --
rem  verified byte-exact for this machine's Chinese path (CP936).
rem
rem  NOTE: keep the comments in this file ASCII-only. Chinese text in
rem  rem lines misparses under chcp 65001 on this machine, and a
rem  misparse inside the block below eats the `powershell` token --
rem  which prints the "elevating..." notice, never raises UAC, and
rem  exits. Chinese in echo lines is fine.
rem ============================================================
rem  Admin?  A high-integrity token carries SID S-1-16-12288.
rem  (net session's exit code is not a reliable test under redirection.)
whoami /groups | findstr /C:"S-1-16-12288" >nul 2>&1
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

rem ============================================================
rem  No menu: go straight to the GUI. Collection, incremental-patch
rem  merging and correction are three parts of one job; they used to
rem  be split across a console menu and two scripts, so the user had
rem  to remember "run the CLI, then go merge in the GUI".
rem  All of it now lives in correct_gui.py: the collect button runs a
rem  scan (either scope writes a patch) and hands off to the merge,
rem  which lists EVERY row -- the ones the icons pinned down come
rem  pre-selected, the rest ask.
rem  Diagnostics (--offline / --check-state / --report) stay CLI-only.
rem ============================================================
echo ============================================================
echo    妮姬 OCR 国服采集
echo ============================================================
echo    目录 : %CD%
echo    权限 : 管理员
echo.
echo    正在打开采集/校正界面 ...
echo.
echo    窗口里： [采集…] 跑扫描（增量填个数 / 全量扫到底） · 双击单元格改数据
echo             [保存 JSON] 写盘 · [合并补丁] 手动合并（补丁每条都会列出来）
echo.

python correct_gui.py
set "RC=%errorlevel%"

if not "%RC%"=="0" (
    echo.
    echo ------------------------------------------------------------
    echo    界面退出，返回码 %RC%
    echo ------------------------------------------------------------
    pause
)
endlocal
