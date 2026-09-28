#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""采集参数选择器（控制台 TUI）—— 给 `启动采集.bat` 的 [2] 用。

为的是别再手敲 `--patch --max 31 --save-shots` 这种东西：↑↓ 选行、空格勾选、
←→ 改数字、回车 改数值/开始。选完**就在这个窗口里**把 collect_cn.py 跑起来 ——
采集进度和 Ctrl+C 都留在原处，不会多开一个窗口。

⚠️ 刻意**不 import 任何项目模块**：collect_cn 一进来就会拖上 numpy / PIL / rapidocr，
而菜单必须秒开。所以这里只用标准库。
⚠️ 也刻意不用 tkinter：为了选几个开关弹个窗口太重；而且游戏在前台时，
控制台里按方向键比鼠标去点小方框顺手得多。
"""
import os
import subprocess
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent

# 每一行 = 一个可选参数。on = 勾没勾上；int / text 行多一个 value。
ROWS = [
    {"key": "patch", "kind": "bool", "flag": "--patch", "on": False, "value": None,
     "desc": "增量：结果写 output/cn_patch.json，主数据一个字节都不动"},
    {"key": "max", "kind": "int", "flag": "--max", "on": False, "value": 31,
     "desc": "最多采几个角色（←→ 调，回车 直接输，0 = 不限）"},
    {"key": "shots", "kind": "bool", "flag": "--save-shots", "on": False, "value": None,
     "desc": "每个角色存 5 张截图，约 8MB，很占磁盘"},
    {"key": "as", "kind": "text", "flag": "--as", "on": False, "value": "",
     "desc": "按扫描顺序断言身份，逗号分隔；会自动等于 --patch"},
]
START = len(ROWS)          # 「开始采集」那一行
HELP = "↑↓ 选行 · 空格 勾选 · ←→ 改数字 · 回车 改数值/开始 · Esc 退出"


# ----------------------------------------------------------------------------
# 控制台
# ----------------------------------------------------------------------------
_VT = None


def _enable_vt():
    """让 Windows 控制台认 ANSI 转义（否则清屏/光标定位全是乱码）。成功与否都返回。"""
    global _VT
    if _VT is None:
        _VT = False
        if os.name == "nt":
            try:
                import ctypes
                k = ctypes.windll.kernel32
                h = k.GetStdHandle(-11)          # STD_OUTPUT_HANDLE
                mode = ctypes.c_uint32()
                if k.GetConsoleMode(h, ctypes.byref(mode)):
                    # ENABLE_VIRTUAL_TERMINAL_PROCESSING = 0x0004
                    _VT = bool(k.SetConsoleMode(h, mode.value | 0x0004))
            except Exception:
                _VT = False
        else:
            _VT = True                            # 类 Unix 终端本来就认
    return _VT


def clear():
    """清屏。ANSI 不可用时退回 cls —— 慢一点、会闪，但绝不会画花。"""
    if _enable_vt():
        sys.stdout.write("\033[2J\033[H")
    else:
        os.system("cls" if os.name == "nt" else "clear")
    sys.stdout.flush()


def getkey():
    """读一个键 -> 语义名。**msvcrt 是延迟导入的**：这样本文件在非 Windows 上也能 import，
    菜单逻辑与参数拼装才能脱离 Windows 单独测。"""
    import msvcrt
    ch = msvcrt.getwch()
    if ch in ("\x00", "\xe0"):                   # 方向键等功能键的引导字节
        return {"H": "up", "P": "down", "K": "left", "M": "right",
                "G": "home", "O": "end"}.get(msvcrt.getwch(), "")
    return {"\r": "enter", " ": "space", "\x1b": "esc", "\x08": "back"}.get(ch, ch)


# ----------------------------------------------------------------------------
# 参数拼装（纯函数，可单独测）
# ----------------------------------------------------------------------------
def build_pairs(rows):
    """勾上的行 -> [(flag, value 或 None)]。

    int 行的值为 0 时按「不限」处理，整个参数不发出去 —— 免得 `--max 0` 一个角色都采不到。
    text 行为空同理（勾了个空断言没有意义）。
    """
    pairs = []
    for r in rows:
        if not r["on"]:
            continue
        if r["kind"] == "bool":
            pairs.append((r["flag"], None))
            continue
        v = r["value"]
        if r["kind"] == "int":
            if not v:
                continue
            pairs.append((r["flag"], str(v)))
        else:
            v = (v or "").strip()
            if not v:
                continue
            pairs.append((r["flag"], v))
    return pairs


def display_cmd(pairs):
    """给人看的命令行。只有名字类值加引号（里面可能有中文和逗号）—— 数字加引号很傻。"""
    bits = ["python", "collect_cn.py"]
    for flag, val in pairs:
        if val is None:
            bits.append(flag)
        elif val.isdigit():
            bits.append(f"{flag} {val}")
        else:
            bits.append(f'{flag} "{val}"')
    return " ".join(bits)


def to_argv(pairs):
    argv = []
    for flag, val in pairs:
        argv.append(flag)
        if val is not None:
            argv.append(val)
    return argv


# ----------------------------------------------------------------------------
# 界面
# ----------------------------------------------------------------------------
def draw(cur):
    clear()
    out = ["  采集参数", "  " + HELP, ""]
    for i, r in enumerate(ROWS):
        mark = "▶" if i == cur else " "
        box = "[x]" if r["on"] else "[ ]"
        val = ""
        if r["kind"] == "int":
            val = "" if (r["on"] and not r["value"]) else f" = {r['value']}"
        elif r["kind"] == "text":
            val = f" = {r['value']}" if r["value"] else " = （未填）"
        out.append(f"  {mark} {box} {r['flag']:<13}{val:<14}  {r['desc']}")
    out.append("  " + "─" * 66)
    out.append(f"  {'▶' if cur == START else ' '} [ 开始采集 ]       回车执行")
    out.append("")
    out.append("  将要运行： " + display_cmd(build_pairs(ROWS)))
    if ROWS[0]["on"]:
        out.append("")
        out.append("  提示：勾了 --patch 的话，先在游戏里翻到第一个要扫的角色，再回车。")
    print("\n".join(out))


def edit_value(row):
    """回车进编辑：先清屏、把提示打清楚，再收一行文字（Ctrl+C 取消）。

    这里刻意用 input() 而不是自己收键：中文名字要能退格、能粘贴，自己收键等于重写一遍
    行编辑器，还会把输入法候选搞得一团糟。
    """
    clear()
    if row["kind"] == "int":
        print(f"  {row['flag']} —— 最多采几个角色。0 = 不限。")
        print("  （回车确定，Ctrl+C 取消）\n")
    else:
        print(f"  {row['flag']} —— 按扫描顺序断言身份，逗号分隔，如：")
        print("      红莲：暗影,桃乐丝,阿妮斯")
        print("  （回车确定，Ctrl+C 取消；全角「，」也认）\n")
    try:
        s = input(f"  {row['flag']} = ").strip()
    except (EOFError, KeyboardInterrupt):
        print("\n  （已取消）")
        return
    if row["kind"] == "int":
        if s.isdigit():
            row["value"] = int(s)
            row["on"] = True
        elif s:
            print(f"  「{s}」不是数字，已忽略。")
    else:
        row["value"] = s
        row["on"] = bool(s)


def run():
    pairs = build_pairs(ROWS)
    clear()
    print("  即将运行：\n")
    print("    " + display_cmd(pairs) + "\n")
    print("  5 秒倒计时开始后请切回游戏。\n")
    try:
        rc = subprocess.call([sys.executable, str(HERE / "collect_cn.py")] + to_argv(pairs),
                             cwd=str(HERE))
    except KeyboardInterrupt:
        rc = 130                                  # Ctrl+C：采集端自己也收到了，正常收尾
    print(f"\n  collect_cn.py 退出，返回码 {rc}。")
    return rc


def main():
    try:
        import msvcrt  # noqa: F401
    except ImportError:
        # 非 Windows（比如在 WSL 里手滑跑了它）。说清楚该走哪条路，别甩个 ImportError。
        print("这个选择器要用 Windows 控制台（msvcrt）。直接跑 collect_cn.py 并加参数即可。")
        return 2
    cur = 0
    while True:
        draw(cur)
        k = getkey()
        if k == "esc":
            clear()
            return 1
        if k == "up":
            cur = (cur - 1) % (len(ROWS) + 1)
            continue
        if k == "down":
            cur = (cur + 1) % (len(ROWS) + 1)
            continue
        if cur < len(ROWS):
            row = ROWS[cur]
            if k == "left" and row["kind"] == "int":
                row["value"] = max(0, (row["value"] or 0) - 1)
                row["on"] = True
                continue
            if k == "right" and row["kind"] == "int":
                row["value"] = (row["value"] or 0) + 1
                row["on"] = True
                continue
            if k == "space":
                if row["kind"] == "bool":
                    row["on"] = not row["on"]
                elif row["value"] in (None, ""):
                    edit_value(row)               # 值还空着，勾了没意义 -> 直接去填
                else:
                    row["on"] = not row["on"]
                continue
            if k == "enter":
                if row["kind"] == "bool":
                    row["on"] = not row["on"]
                else:
                    edit_value(row)
                continue
            continue
        if k == "enter":                          # 停在「开始采集」上
            return run()


if __name__ == "__main__":
    sys.exit(main())
