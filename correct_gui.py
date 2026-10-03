#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""国服 NIKKE 采集结果校正界面（tkinter）。

跑法（Windows，项目根目录；`启动采集.bat` 提权后也是打开它）：
    python correct_gui.py                 打开界面，默认读写 output/cn_collect.json
    python correct_gui.py --file X.json   换文件
    python correct_gui.py --report        不开窗，只打印校验基线（无 tkinter 的机器也能跑）
    python correct_gui.py --roundtrip     读入→原样写出→读回，验证明细零丢失、格式零改动
    python correct_gui.py --merge-preview 不开窗，只打印补丁会怎么合并（干跑，不写文件）

**这个界面是唯一的入口**：采集（顶部「采集…」）、合并补丁、校正数据都在这里。
采集跑完会自动接上合并，不必再开另一个程序 —— 以前采集在控制台、合并在界面，
用户得自己记住「跑完命令行还要去点合并」，那一步现在没有了。

采集**只产出补丁**（output/cn_patch.json，带顶层 `_图标`），**主数据一个字节都不动** ——
主数据只由这里的「合并」写入。「增量」和「全量」不再是两条路，只是**扫描范围**不同
（增量填个数、全量扫到最后一个角色），后面的「对比旧表 → 有更新就更新 / 没更新照旧 /
新角色追加」是同一件事。整表替换降级成 collect_cn.py 的 `--replace`，界面不可达。
合并**先分类再全部列出**（见 plan_merge / merge_entries）：补丁每条都带着采集时读到的
图标四维，能唯一确定行号的**预选好、标绿**，但一样显示出来、一样可改 —— 不再有
「全自动就跳过窗口」的旁路，那正是以前窗口瞒着你把几条直接并了的来源。
身份优先看图标（与姓名无关的机器证据）；**姓名永远不足以自动合并** —— 滚动横幅截断后
可能恰好等于另一个真实角色。补丁里为空、而旧行有值的槽默认**保留旧值**（空到底是 T9 真值
还是没读到，光看补丁分辨不了），要覆盖得逐条勾「空槽也覆盖」；预选的那些同样不覆盖空槽。

界面是一张 31 列的大表，一行一个角色：
    姓名 | 战力 | 问题 | 头.等级 头1名 头1值 头2名 头2值 头3名 头3值 | 甲… | 手… | 脚…
双击单元格进入编辑。姓名/词条名是下拉框（可自由输入），等级和数值是只读下拉 —— 数值的
候选就是**该行词条名对应的那 15 档**（`11.11%` 这种写法，内部一律存小数），选就行，不用
记档位也打不错。词条名漏读时没有档位可查，退化成全局并集并提示一句（真正的解法是先定词条名）。

可疑度按「要不要翻回游戏核实」分三级，底色 + 符号直接对应你下一步的动作：
    ▲ 橙底：必须看游戏才能定 —— 姓名图鉴查无此人、姓名多候选、战力缺失、战力违反降序、数值漏读
    ! 黄底：靠已有信息就能修 —— 姓名有唯一候选、数值不在合法档位、重名
    （灰字）：整槽未采集，信息性标记，不是错误
「问题」列写的是**带位置的短标签**（`▲脚3数值漏读`），而且漏读的格子本身会显示成 `?`，
所以不用去翻是哪一格。F3 循环跳到下一个问题行、并自动横向滚到出问题的那一列。
筛选按钮可以只看 ▲ / 只看 ▲+!。

⚠️ 三条刻意的设计边界：
1. **绝不自动改写你的数据**：图鉴里基础角色（`索林`）和「名字：称号」变体（`索林：霜之旅票`）
   是两条独立词条，光看一个错名机器猜不出是哪个，所以候选只提示、由你点选。
   档位校验同理：`align_affix_value` 会把 50% 悄悄对齐成 48.39%，这里只用它提示、不用它覆盖。
2. **本界面的档位校验比采集端弱**：采集时有字色信息能把数值限在配色对应的档位区间，
   这里没有，只能全 15 档查。所以「是合法档位、但落在另一种颜色档位上」这类错读抓不到。
3. **表格用 Canvas 自己画，不是 ttk.Treeview** —— 见 Grid 类的注释，实测差 8~10 倍性能。

本文件刻意**不 import collect_cn**：那会把 rapidocr/onnxruntime/numpy/PIL 一起拖进来
（实测白吃 56MB 内存、启动多 1.1 秒），而界面一行 OCR 代码都用不到。代价是
load_affix_table / save_rows / archive_previous 各留一份 —— 三者都极稳定，且
`--report` 会在 collect_cn 可导入时逐项比对，一旦漂移会报出来。
"""

import argparse
import copy
import difflib
import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
from pathlib import Path

try:
    import tkinter as tk
    from tkinter import ttk, messagebox
except ImportError:  # --report / --roundtrip 在没装 tkinter 的机器上也要能跑
    tk = ttk = messagebox = None

# ----------------------------------------------------------------------------
# 常量
# ----------------------------------------------------------------------------
GUI_SLOTS = ("头", "甲", "手", "脚")
ROSTER_PATH = "docs/chacters.json"
EQUIP_XLSX = "equipment.xlsx"
DEFAULT_JSON = "output/cn_collect.json"
DEFAULT_PATCH = "output/cn_patch.json"   # collect_cn.py 采集的产物（补丁），见「补丁合并」一节
# 「全量」= ALL：不填数量，扫到最后一个角色。collect_cn 遇 LV.1 会自动停，这只是个兜底上限。
FULL_MAX = 500

# 与 collect_cn.py 保持一致（见文件头说明）。空词条占位是游戏里的原文，
# 采集端也拿它当「这块是 T10」的铁证。
EMPTY_AFFIX = "未获得效果"
AFFIX_CANON = ["攻击力增加", "蓄力伤害增加", "防御力增加", "暴击伤害增加", "命中率增加",
               "暴击率增加", "最大装弹数增加", "蓄力速度增加", "优越代码伤害增加"]

# 信息页那四个图标对应的维度，与图鉴的字段名一一对应。与 collect_cn.py 保持一致，
# `--report` 的一致性检查会比对（两处各留一份，防漂移 —— 本文件刻意不 import collect_cn）。
ICON_DIMS = ("属性", "武器", "职业", "企业")
ICON_ALIAS = {"超规格极乐净土": "极乐净土"}   # 图标值 -> 图鉴值，同 collect_cn.ICON_ALIAS

SYM = {"suspect": "▲", "warn": "!"}
ROW_BG = {"suspect": "#ffd6a5", "warn": "#fff2cc", None: "#ffffff"}
F_SUSPECT, F_WARN = "suspect", "warn"
FG_NORMAL, FG_STALE, FG_MARK, FG_SEL = "#000000", "#9a9a9a", "#c00000", "#1a73e8"

CP_MIN, CP_MAX = 1000, 2_000_000  # 战力合理区间
CP_ORDER_RATIO = 0.03             # 战力违反降序的判定阈值，压掉 OCR 抖动

HERE = Path(__file__).resolve().parent      # 起采集子进程时的 cwd，别依赖调用者的 cwd


def is_admin():
    """当前进程是否以管理员身份运行。非 Windows 返回 True（不适用）。

    **为什么采集前必须问这一句**：游戏以管理员身份运行时，普通权限进程发的模拟点击会被
    系统（UIPI）**静默丢弃** —— 不报错、就是点不动。原先这条只由 `启动采集.bat` 的
    `net session` 兜着，而界面是可以被直接 `python correct_gui.py` 起来的，
    没有这道检查就会「点了开始、看着它干等 30 分钟、一行数据都没有」。

    查不出来（异常）时返回 True：宁可放行后失败，也不要因为探测本身出错把人挡在外面。
    """
    if os.name != "nt":
        return True
    try:
        import ctypes
        return bool(ctypes.windll.shell32.IsUserAnAdmin())
    except Exception:
        return True


def load_affix_table(path=EQUIP_XLSX):
    """读 equipment.xlsx「词条表」-> {词条名: [15 档数值]}。与 collect_cn 同逻辑。"""
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["词条表"]
    table = {}
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 1).value
        if not name:
            continue
        name = name.replace("量", "数")   # 「最大装弹量增加」→ 规范名
        vals = [ws.cell(r, c).value for c in range(3, 18)]
        table[name] = [float(v) for v in vals if v is not None]
    return table


def save_rows(rows, out):
    """原子写回（先写 .tmp 再 replace），格式与 collect_cn._save_rows 完全一致。"""
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    tmp = Path(str(out) + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump({"角色": rows}, f, ensure_ascii=False, indent=2)
    tmp.replace(out)


def archive_previous(path):
    """把上一次的结果改名带时间戳留底，返回新路径或 None。与 collect_cn 同逻辑。"""
    p = Path(path)
    if not p.exists():
        return None
    try:
        n = len(json.loads(p.read_text(encoding="utf-8")).get("角色", []))
    except Exception:
        n = 0
    if n == 0:
        return None
    dst = p.with_name(f"{p.stem}.{time.strftime('%Y%m%d-%H%M%S')}{p.suffix}")
    p.replace(dst)
    return dst


def show_scrolled_text(parent, title, body_text, width=760, height=440):
    """可滚动的长文本提示窗，替换 messagebox.showinfo。

    messagebox 会随文本长度一路撑高 —— 合并几十条时它能长到把「确定」顶出屏幕外，
    用户点不到、只能瞎摸。这里固定尺寸 + 滚动条：文本多长都看得完，按钮永远在。
    行为照抄 messagebox.showinfo：模态、只有一个「确定」、Esc 等同确定。
    """
    win = tk.Toplevel(parent)
    win.title(title)
    win.transient(parent)
    win.geometry(f"{width}x{height}")
    win.minsize(420, 240)
    frame = ttk.Frame(win, padding=(10, 10, 10, 6))
    frame.pack(fill="both", expand=True)
    txt = tk.Text(frame, wrap="word", background="#fbfbfb")
    sb = ttk.Scrollbar(frame, orient="vertical", command=txt.yview)
    txt.configure(yscrollcommand=sb.set)
    txt.insert("1.0", body_text)
    txt.configure(state="disabled")      # 只读：文本是内容，不是输入框
    txt.pack(side="left", fill="both", expand=True)
    sb.pack(side="right", fill="y")
    bar = ttk.Frame(win, padding=(10, 0, 10, 10))
    bar.pack(fill="x")
    ttk.Button(bar, text="确定", command=win.destroy).pack(side="right")
    win.protocol("WM_DELETE_WINDOW", win.destroy)
    win.bind("<Escape>", lambda e: win.destroy())
    win.grab_set()
    win.focus_set()
    win.wait_window()


# ----------------------------------------------------------------------------
# 列规格：31 列全部由这张描述符表驱动
# ----------------------------------------------------------------------------
class ColSpec:
    __slots__ = ("col_id", "kind", "slot", "idx", "field", "width", "anchor")

    def __init__(self, col_id, kind, width, anchor, slot=None, idx=None, field=None):
        self.col_id, self.kind, self.width, self.anchor = col_id, kind, width, anchor
        self.slot, self.idx, self.field = slot, idx, field


def build_col_specs():
    specs = [
        ColSpec("姓名", "text", 110, "w", field="姓名"),
        ColSpec("战力", "int", 80, "center", field="战力"),
        ColSpec("问题", "ro", 250, "w"),
    ]
    for slot in GUI_SLOTS:
        specs.append(ColSpec(f"{slot}.等级", "level", 66, "center", slot=slot))
        for i in range(3):
            specs.append(ColSpec(f"{slot}{i + 1}名", "affixname", 116, "w", slot=slot, idx=i))
            specs.append(ColSpec(f"{slot}{i + 1}值", "pct", 68, "center", slot=slot, idx=i))
    return specs


COL_SPECS = build_col_specs()
COL_IDS = [s.col_id for s in COL_SPECS]
COL_WIDTHS = [s.width for s in COL_SPECS]
COL_ANCHORS = [s.anchor for s in COL_SPECS]
COL_INDEX = {s.col_id: i for i, s in enumerate(COL_SPECS)}
SPEC_BY_ID = {s.col_id: s for s in COL_SPECS}


# ----------------------------------------------------------------------------
# 模型层（纯函数，不依赖 tkinter）
# ----------------------------------------------------------------------------
class Prob:
    """一条问题。short 带位置（给人一眼看出是哪一格），col 指到那一列（用来标记/跳转）。"""
    __slots__ = ("level", "code", "short", "detail", "col")

    def __init__(self, level, code, short, detail, col=None):
        self.level, self.code, self.short, self.detail, self.col = level, code, short, detail, col


def empty_slot():
    return {"等级": None, "词条": [{"名称": None, "数值": None} for _ in range(3)]}


def fmt_pct(v):
    """0.0469 -> '4.69%'；None -> ''。用 4 位小数再削尾零，避免浮点毛刺。"""
    if v is None:
        return ""
    return f"{float(v) * 100:.4f}".rstrip("0").rstrip(".") + "%"


def parse_pct_input(text):
    """'11.11%' / '11.11' / '0.1111' -> (小数, None)；空串 -> (None, None)；非法 -> (None, 错误)。"""
    t = (text or "").strip()
    if not t:
        return None, None
    has_pct = t.endswith("%")
    body = t[:-1].strip() if has_pct else t
    try:
        v = float(body)
    except ValueError:
        return None, f"「{text}」不是数字"
    if has_pct or v > 1:
        v = v / 100.0
    if v < 0:
        return None, f"「{text}」是负数"
    return round(v, 4), None


def tier_hit(levels, v):
    """v 是否落在 levels 里（浮点用容差比，避免 0.0469 != 0.0469 这种毛刺）。"""
    return any(abs(float(v) - float(x)) < 1e-9 for x in levels)


def resolve_affix_name(text):
    """用户输入 -> 规范词条名 / EMPTY_AFFIX / None。

    ⚠️ 不能复用 cc.norm_affix：那是给 OCR 噪声用的宽松匹配，实测 norm_affix(['暴击'])
    会按列表顺序瞎猜成「暴击伤害增加」。这里只接受「精确命中」或「唯一包含候选」。
    """
    t = (text or "").strip()
    if not t:
        return None, None
    if t in AFFIX_CANON or t == EMPTY_AFFIX:
        return t, None
    cands = [c for c in AFFIX_CANON if t in c or c in t]
    if len(cands) == 1:
        return cands[0], None
    if len(cands) > 1:
        return None, f"「{t}」命中 {len(cands)} 个词条（{'/'.join(cands)}），请写全"
    return None, f"「{t}」不是规范词条名"


def get_cell(row, spec):
    if spec.kind in ("text", "int"):
        return row.get(spec.field)
    if spec.kind == "ro":
        return None
    slot = row.get(spec.slot) or empty_slot()
    if spec.kind == "level":
        return slot.get("等级")
    aff = (slot.get("词条") or [])[spec.idx]
    return aff.get("名称") if spec.kind == "affixname" else aff.get("数值")


def set_cell(row, spec, raw):
    """把 raw 解析后写回 row。返回 (是否成功, 错误信息)。失败时 row 一个字节都不动。"""
    if spec.kind == "ro":
        return False, "「问题」列是只读的"
    if spec.kind in ("text", "int"):
        val = (raw or "").strip()
        if spec.kind == "int":
            if not val:
                row[spec.field] = None
                return True, None
            try:
                row[spec.field] = int(val.replace(",", ""))
            except ValueError:
                return False, f"战力「{val}」不是整数"
            return True, None
        row[spec.field] = val or None
        return True, None

    slot = row.setdefault(spec.slot, empty_slot())
    if spec.kind == "level":
        val = (raw or "").strip()
        if not val:
            slot["等级"] = None
            return True, None
        if val not in ("0", "1", "2", "3", "4", "5"):
            return False, f"等级「{val}」只能是 0~5 或留空"
        slot["等级"] = int(val)
        return True, None

    affs = slot.setdefault("词条", [{"名称": None, "数值": None} for _ in range(3)])
    aff = affs[spec.idx]
    if spec.kind == "affixname":
        name, err = resolve_affix_name(raw)
        if err:
            return False, err
        aff["名称"] = name
        if name is None:
            aff["数值"] = None
        elif name == EMPTY_AFFIX:
            aff["数值"] = 0
        return True, None

    v, err = parse_pct_input(raw)          # pct
    if err:
        return False, err
    aff["数值"] = v
    return True, None


def render_cell(row, spec):
    """单元格显示文本。永远返回 str，None -> ''（否则会显示字面量 None）。"""
    v = get_cell(row, spec)
    if spec.kind == "pct" and v is not None:
        return fmt_pct(v)
    return "" if v is None else str(v)


def normalize_row(raw):
    """把一行补全到规范形状。返回 (row, [异常说明])，永不抛异常。"""
    msgs = []
    if not isinstance(raw, dict):
        return None, ["不是对象"]
    row = {}
    nm = raw.get("姓名")
    row["姓名"] = nm if isinstance(nm, str) and nm.strip() else None
    if nm is not None and row["姓名"] is None:
        msgs.append(f"姓名 {nm!r} 不是字符串")
    cp = raw.get("战力")
    if cp is None or isinstance(cp, int) and not isinstance(cp, bool):
        row["战力"] = cp
    else:
        try:
            row["战力"] = int(float(cp))
            msgs.append(f"战力 {cp!r} 已转成整数")
        except (TypeError, ValueError):
            row["战力"] = None
            msgs.append(f"战力 {cp!r} 无法识别，已置空")
    for slot in GUI_SLOTS:
        src = raw.get(slot)
        if not isinstance(src, dict):
            row[slot] = empty_slot()
            if src is not None:
                msgs.append(f"{slot} 结构异常，已清空")
            continue
        lv = src.get("等级")
        try:
            lv = None if lv is None or lv == "" else int(lv)
        except (TypeError, ValueError):
            msgs.append(f"{slot} 等级 {lv!r} 无法识别，已置空")
            lv = None
        affs = src.get("词条")
        out = []
        if not isinstance(affs, list):
            affs = []
        for i in range(3):
            a = affs[i] if i < len(affs) else None
            if not isinstance(a, dict):
                out.append({"名称": None, "数值": None})
                continue
            n = a.get("名称")
            n = n if isinstance(n, str) and n.strip() else None
            v = a.get("数值")
            if isinstance(v, str):
                v, _ = parse_pct_input(v)
                msgs.append(f"{slot}{i + 1} 数值是字符串，已转成小数")
            elif v is not None and not isinstance(v, (int, float)):
                v = None
            elif isinstance(v, float):
                v = round(v, 4)
            out.append({"名称": n, "数值": None if v is None else v})
        if len(affs) != 3:
            msgs.append(f"{slot} 词条数 {len(affs)}≠3，已补齐")
        row[slot] = {"等级": lv, "词条": out}
    return row, msgs


def load_rows(path):
    """读 JSON 并逐行归一化。返回 (rows, [说明])，永不抛异常。"""
    p = Path(path)
    if not p.exists():
        return [], [f"文件不存在：{path}"]
    try:
        data = json.loads(p.read_text(encoding="utf-8"))
    except Exception as e:
        return [], [f"JSON 解析失败：{e}（旧数据可能还在 cn_collect.<时间戳>.json 留底里）"]
    if not isinstance(data, dict) or not isinstance(data.get("角色"), list):
        return [], ["顶层结构不是 {'角色': [...]}，无法识别"]
    rows, notes = [], []
    for i, raw in enumerate(data["角色"]):
        r, msgs = normalize_row(raw)
        if r is None:
            notes.append(f"第 {i + 1} 行{msgs[0]}，已跳过")
            continue
        rows.append(r)
        notes += [f"第 {i + 1} 行 {m}" for m in msgs]
    return rows, notes


# ---- 图鉴（docs/chacters.json）----
def norm_roster_name(s):
    """图鉴里半角':'和全角'：'混用，比较前统一。"""
    return re.sub(r"\s+", "", str(s)).replace(":", "：")


def load_roster(path):
    """-> ({"keys": [...], "orig": {归一化名: 原始名}}, 错误说明)。"""
    p = Path(path)
    if not p.exists():
        return None, f"图鉴不存在：{path}"
    try:
        raw = json.loads(p.read_text(encoding="utf-8"))
    except Exception as e:
        return None, f"图鉴解析失败：{e}"
    if not isinstance(raw, dict) or not raw:
        return None, "图鉴不是非空 dict"
    orig = {norm_roster_name(k): k for k in raw}
    return {"keys": sorted(orig), "orig": orig}, None


def load_roster_attrs(path=ROSTER_PATH):
    """图鉴 -> {姓名: {属性, 武器, 职业, 企业}}。读不到返回 None（调用方据此关掉图标定人）。

    与 load_roster 的分工：那个只拿 key 当姓名名单（姓名校验用），这个拿**四维值**给图标
    定人用。四维是**和姓名完全无关**的独立判据 —— 这正是它存在的理由：姓名读的是滚动横幅，
    实测约四成会截断，而截断后可能恰好等于另一个真实角色，那种错在姓名校验里是零告警的。

    同 collect_cn.load_roster_attrs，两处各留一份（本文件不 import collect_cn）。
    """
    try:
        raw = json.loads((Path(__file__).resolve().parent / path).read_text(encoding="utf-8"))
    except Exception:
        return None
    if not isinstance(raw, dict) or not raw:
        return None
    return {k: v for k, v in raw.items()
            if isinstance(v, dict) and all(d in v for d in ICON_DIMS)}


def _is_subseq(short, long_):
    it = iter(long_)
    return all(c in it for c in short)


def name_candidates(bad, roster, limit=3):
    """给一个对不上图鉴的姓名找候选。滚动横幅截帧的典型形态是「连续子串」或「子序列」。"""
    if not bad or not roster:
        return []
    keys = roster["keys"]
    score = {}
    for k in keys:
        if bad in k:
            score[k] = max(score.get(k, 0), 1.0)
        elif len(bad) >= 2 and _is_subseq(bad, k):
            score[k] = max(score.get(k, 0), 0.9)
    for k in difflib.get_close_matches(bad, keys, n=8, cutoff=0.34):
        score[k] = max(score.get(k, 0), difflib.SequenceMatcher(None, bad, k).ratio())
    ranked = sorted(score.items(), key=lambda kv: (-kv[1], len(kv[0])))
    return [roster["orig"][k] for k, _ in ranked[:limit]]


def check_name(name, roster):
    if not roster:
        return []
    if not name:
        return [Prob(F_SUSPECT, "姓名缺失", "姓名缺失", "姓名是空的", "姓名")]
    if norm_roster_name(name) in roster["orig"]:
        return []
    cands = name_candidates(norm_roster_name(name), roster)
    if len(cands) == 1:
        return [Prob(F_WARN, "姓名截断", "姓名截断", f"图鉴里没有「{name}」，疑似截断；候选：{cands[0]}", "姓名")]
    if cands:
        return [Prob(F_SUSPECT, "姓名存疑", "姓名存疑",
                     f"图鉴查无「{name}」，需人工确认；候选：{' / '.join(cands)}", "姓名")]
    return [Prob(F_SUSPECT, "姓名存疑", "姓名存疑",
                 f"图鉴查无「{name}」，且没有相近的名字，需人工确认", "姓名")]


def check_cp_order(rows):
    """战力违反降序的建议。相邻三点局部极值，差值 >=3% 才标；含 None 的三元组跳过。

    成对标记相邻两行是**正确行为**：单看相邻关系无法判断是哪一个读错了，
    两行都标出来让人去比游戏，比猜一个更好。
    """
    out = {}
    cps = [r.get("战力") for r in rows]
    for i in range(1, len(rows) - 1):
        a, b, c = cps[i - 1], cps[i], cps[i + 1]
        if not all(isinstance(x, int) and not isinstance(x, bool) for x in (a, b, c)):
            continue
        lo = min(a, c)
        if not lo:
            continue
        if b < a and b < c and (lo - b) / lo >= CP_ORDER_RATIO:
            out[i] = f"战力疑似读低（相邻 {a} / {b} / {c}）"
        elif b > a and b > c and (b - lo) / lo >= CP_ORDER_RATIO:
            out[i] = f"战力疑似读高（相邻 {a} / {b} / {c}）"
    return out


def slot_is_stale(slot):
    return slot.get("等级") is None and all(
        a.get("名称") is None and a.get("数值") is None for a in slot.get("词条", [])
    )


# ----------------------------------------------------------------------------
# 补丁合并（纯函数，不依赖 tkinter）
# ----------------------------------------------------------------------------
# collect_cn.py 采集写出的 output/cn_patch.json，只经这里并进主数据（主数据的唯一写入者）。
# 身份优先看**图标四维**（纯函数 plan_merge）：能唯一确定的预选好，定不下来的展开提问；
# 两类都会列在合并窗里（merge_entries），没有「静默直接并」这一步。
# 姓名字符串是最后手段 —— 它是 OCR 的产物，约四成会截断，截断后还可能恰好等于另一个真实角色。
def load_patch_icons(path=DEFAULT_PATCH):
    """读补丁文件顶层的图标读数（collect_cn.ICONS_KEY），按下标与「角色」对齐。

    缺失/不是列表 -> []（等于「没有图标证据」，全部退回姓名判定，是保守的一侧）。
    读不到某一条不算错：plan_merge 对越界下标按 None 处理。
    """
    try:
        raw = json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception:
        return []
    icons = raw.get("_图标")
    return icons if isinstance(icons, list) else []


def plan_merge(main_rows, patch_rows, patch_icons, attrs, roster):
    """把补丁分成「机器已唯一确定」与「需要人看」两堆。-> (auto, ambiguous)

      auto      : {补丁下标: 目标旧行下标 或 None}，None = 追加为新角色
      ambiguous : [(补丁下标, 一句话理由, 优先候选旧行下标列表)]
                  最后一项只用来**把下拉框排好序/预选**，不替用户决定。

    **只分类，不决定** —— ambiguous 一律留给界面里的人，绝不猜。

    ⚠️ **auto 只来自图标**（四维是与姓名无关的机器证据）。姓名那一路**永远不 auto**，
    只用来给 ambiguous 排序。理由：姓名是 OCR 的产物，滚动横幅截断后**可能恰好等于另一个
    真实角色**（实测约四成读错，图鉴里 44 条是别条的前缀）—— 「红莲：暗影」截断成「红莲」
    再精确命中主数据里的「红莲」，正是本项目记录在案的那种静默覆盖。所以「姓名精确命中」
    这一条证据**不足以**支撑自动合并，必须有人看一眼。

    图标那一路：
      |C| == 1 -> 图鉴里唯一命中「only」：主数据里恰好一行叫 only 就是它；一行都没有 = 新角色，
                  追加是安全的（错了只是多一行）。但若补丁*姓名*在主数据里另有同名的行，那是
                  「图标说 A、姓名说 B」的冲突，交给人。
      |C| == 0 -> 图鉴里四维全等的人一个都没有：国服特供（婴宁/画皮）或图标读歪了，交给人。
      |C| > 1  -> 图标没能把候选收到一个，交给人。**这一档很常见**：201 条图鉴里有 39 组四维
                  相同的名字对（用户那 52 行里就有 15 行撞），所以它按设计就会弹。
                  但机器仍能帮上忙 —— 「C 里哪个在你这儿」通常就把答案指出来了，所以
                  prefer 把 C∩主数据 排前面预选好，人按一下回车就行。

    ⚠️ |C| > 1 时**绝不**因为「C 里只有一个人在你这儿」就自动并进去：主数据里只有
    「白雪公主」不代表这次扫的不是刚抽到的「白雪公主：纯真年代」—— 那正好会把两个人的
    装备并到一起，而这是整套设计里唯一不可接受的后果。
    """
    auto, ambiguous = {}, []
    attrs_n = {norm_roster_name(k): v for k, v in (attrs or {}).items()}
    for i, pr in enumerate(patch_rows):
        nm = norm_roster_name(pr.get("姓名") or "")
        name_hits = [t for t, r in enumerate(main_rows)
                     if norm_roster_name(r.get("姓名") or "") == nm]
        ic = patch_icons[i] if i < len(patch_icons) else None

        # ---- 图标那一路 ----------------------------------------------------
        if isinstance(ic, dict) and attrs_n and all(ic.get(d) for d in ICON_DIMS):
            want = {d: ICON_ALIAS.get(ic[d], ic[d]) for d in ICON_DIMS}
            C = sorted(n for n, a in attrs_n.items()
                       if all(a.get(d) == want[d] for d in ICON_DIMS))
            shown = "、".join(f"{d}={ic[d]}" for d in ICON_DIMS)
            if len(C) == 1:
                hits = [t for t, r in enumerate(main_rows)
                        if norm_roster_name(r.get("姓名") or "") == C[0]]
                if len(hits) == 1:
                    auto[i] = hits[0]
                    continue
                if not hits:
                    # 图鉴里认得出、主数据里没有 = 新角色。但补丁姓名若有同名旧行，那是
                    # 「图标说这个、姓名说那个」，别自作主张。
                    if name_hits:
                        ambiguous.append(
                            (i, f"图标唯一命中「{C[0]}」，但补丁姓名「{pr.get('姓名')}」"
                                f"在主数据里有 {len(name_hits)} 行同名 —— 冲突，请人工确认",
                             name_hits))
                    else:
                        auto[i] = None
                    continue
                ambiguous.append((i, f"主数据里有 {len(hits)} 行都叫「{C[0]}」，需人工确认", hits))
                continue
            if not C:
                ambiguous.append(
                    (i, f"图标（{shown}）在图鉴里找不到四维全等的人 —— 国服特供（如 婴宁/画皮）"
                        f"属正常，也可能是图标读歪了，请人工定", []))
            else:
                # prefer 只决定下拉的**排序/预选**，不替人决定。两条排序依据：
                # ①姓名与候选**完全相等**的排最前 —— 图标分不开的那几组四维故意相同，
                #   姓名是这里唯一还能用的旁证（但铁律不变：姓名永不足以**自动**合并，
                #   截断后可能恰好等于另一个真实角色）；
                # ②其余按主数据行序（C∩主数据，也就是"这几个候选里哪个在你表里"）。
                prefer = sorted(
                    (t for t, r in enumerate(main_rows)
                     if norm_roster_name(r.get("姓名") or "") in C),
                    key=lambda t: 0 if norm_roster_name(main_rows[t].get("姓名") or "") == nm else 1)
                ambiguous.append(
                    (i, f"图标（{shown}）留下 {len(C)} 个候选（{'、'.join(C)}）—— "
                        f"这些角色的四维故意相同，图标分不开，请人工定", prefer))
            continue

        # ---- 图标缺席：只能靠姓名，而姓名不够格自动并（见 docstring）----------
        cands, why = merge_candidates(pr.get("姓名"), main_rows, roster)
        ambiguous.append((i, why + "（无图标证据）", cands))
    return auto, ambiguous


def merge_entries(patch_rows, auto, ambiguous):
    """plan_merge 的分类结果 -> 按补丁顺序摊平的每一行，供界面渲染 / 干跑打印共用。

    -> [(下标, kind, 目标旧行下标 或 None, 一句话理由, 优先候选列表)]，kind ∈ {"auto","check"}。
      * "auto"  —— 图标唯一确定（target 为 None 表示追加为新角色），预选好但**仍然显示**；
      * "check" —— 定不下来，理由来自 plan_merge，人工定。

    ⚠️ **每条补丁行都必须出现在这里**（plan_merge 保证每条非 auto 即 ambiguous）。
    合并窗现在渲染的就是这个列表的全部 —— 以前只渲染 ambiguous，把 auto 偷偷并了，
    用户既看不到也改不了。摊平后窗口与 --merge-preview 走同一份数据，措辞不会各说各话。
    """
    amb = {i: (why, prefer) for i, why, prefer in ambiguous}
    out = []
    for i in range(len(patch_rows)):
        if i in auto:
            out.append((i, "auto", auto[i], "", []))
        else:
            why, prefer = amb.get(i, ("未分类（plan_merge 漏了这一条）", []))
            out.append((i, "check", None, why, prefer))
    return out


def auto_decisions(auto):
    """plan_merge 的 auto -> apply_patch 吃的 decisions（值是 (目标, 空槽覆盖) 二元组）。

    自动确定的条目**一律不覆盖空槽** —— 空到底是 T9 真值还是没读到，机器分辨不了，
    沿用「默认保留旧值」这条既有保守规则，与有人看着时按下的默认值一致。
    """
    return {i: (t, False) for i, t in auto.items()}


def merge_candidates(patch_name, rows, roster):
    """补丁行的姓名 -> (候选旧行下标列表（最像的在前）, 一句话理由)。

    先精确同名（归一化后），再退回图鉴姓名候选。**只给候选，不替用户决定** ——
    补丁姓名本身可能就是截断或读错，猜错了代价是覆盖掉另一个人的数据。
    """
    nm = norm_roster_name(patch_name or "")
    if not nm:
        return [], "补丁里没有姓名"
    exact = [i for i, r in enumerate(rows) if norm_roster_name(r.get("姓名") or "") == nm]
    if exact:
        return exact, "姓名精确命中"
    if not roster:
        return [], "图鉴未载入，无法给姓名候选"
    wanted = {norm_roster_name(c) for c in name_candidates(nm, roster)}
    hits = [i for i, r in enumerate(rows) if norm_roster_name(r.get("姓名") or "") in wanted]
    return (hits, "姓名相近（需人工确认）") if hits else ([], "没有相近的名字，需人工指定")


def empty_patch_slots(patch_row, old_row):
    """补丁里「空」、而旧行非空的槽。

    **空到底是 T9 的真值，还是这次没读到？光看补丁分辨不了** —— T9 本来就没词条，
    空是对的；而读失败也是空。所以这些槽默认保留旧值，要覆盖得用户显式勾选：
    别把「这次扫到了空」当成「这次更准」。
    """
    return [s for s in GUI_SLOTS
            if slot_is_stale(patch_row.get(s) or {}) and not slot_is_stale(old_row.get(s) or {})]


def merge_row(old_row, patch_row, empty_over=False):
    """把补丁行并进旧行（**原地改 old_row**）。返回 [说明]。

    三条保守规则 —— 姓名、空槽、战力一律「宁可留旧值」：
      * 姓名永远是**旧行的**：合并的前提就是「确认这是同一个人」，改名是另一回事。
      * 空槽默认保留旧值（见 empty_patch_slots），empty_over=True 才用空覆盖。
      * 战力没读到（None）时同样保留旧值。
    """
    notes = []
    for s in GUI_SLOTS:
        if slot_is_stale(patch_row.get(s) or {}) and not empty_over:
            if not slot_is_stale(old_row.get(s) or {}):
                notes.append(f"{s} 补丁为空，保留旧值")
            continue
        old_row[s] = copy.deepcopy(patch_row.get(s) or empty_slot())
    cp = patch_row.get("战力")
    if cp is None:
        if old_row.get("战力") is not None:
            notes.append("战力没读到，保留旧值")
    else:
        old_row["战力"] = cp
    return notes


def sort_by_cp(rows):
    """按战力降序（缺战力的排最后）。游戏名册就是这个序。

    ⚠️ 合并后**必须重排**：覆盖战力会让它在降序里的位置漂，不重排就是凭空造出一处
    「战力违反降序」，而 check_cp_order 是**成对标记相邻两行**的，会连累邻居一起报警。
    重排会让序号漂移 —— 所以确定目标行要在重排**之前**用行身份定好，之后不能再按序号引。
    """
    rows.sort(key=lambda r: (r.get("战力") is None, -(r.get("战力") or 0)))


def apply_patch(rows, patch_rows, decisions):
    """按 decisions 把补丁并进 rows（**原地**，重排由调用方做）。返回 (新增数, 覆盖数, [说明])。

    decisions: {补丁行下标: (目标旧行下标, 空槽是否覆盖)}；**不在里面的按跳过处理**，
    目标下标为 None 表示「作为新角色追加」。
    """
    added = over = 0
    notes = []
    for i, pr in enumerate(patch_rows):
        if i not in decisions:
            continue
        target, empty_over = decisions[i]
        name = pr.get("姓名") or "(空)"
        if target is None:
            rows.append(copy.deepcopy(pr))
            added += 1
            notes.append(f"[{i + 1}] {name}：追加为新角色")
            continue
        if not 0 <= target < len(rows):
            notes.append(f"[{i + 1}] {name}：目标行 {target} 越界，已跳过")
            continue
        ns = merge_row(rows[target], pr, empty_over)
        over += 1
        notes.append(f"[{i + 1}] {name} → 覆盖 #{target + 1} {rows[target].get('姓名') or '(空)'}"
                     + (f"（{'；'.join(ns)}）" if ns else ""))
    return added, over, notes


def merge_preview(main_path, patch_path):
    """干跑：打印每条补丁会怎么合并（**不写任何文件**）。没装 tkinter 的机器也能跑。"""
    if not Path(patch_path).exists():
        print(f"没有补丁文件：{patch_path}")
        return 1
    m = build_model(main_path, verbose=False)
    prows, notes = load_rows(patch_path)
    icons = load_patch_icons(patch_path)
    auto, ambiguous = plan_merge(m.rows, prows, icons, load_roster_attrs(), m.roster)
    entries = merge_entries(prows, auto, ambiguous)
    print(f"主数据 {main_path}：{len(m.rows)} 行")
    print(f"补丁   {patch_path}：{len(prows)} 行" + (f"（结构异常 {len(notes)} 处）" if notes else "")
          + (f"，其中 {len(icons)} 条带图标读数" if icons else "（没有图标读数，"
             f"全部退回姓名判定）"))
    n_auto = sum(1 for _i, k, _t, _w, _p in entries if k == "auto")
    print(f"\n=== 全部 {len(entries)} 条都会出现在合并窗里（{n_auto} 条已由图标预选好，"
          f"{len(entries) - n_auto} 条要你定），逐条如下 ===")
    for i, kind, target, why, prefer in entries:
        pr = prows[i]
        head = (f"  [{i + 1}] {pr.get('姓名') or '(空)'}  战力 {pr.get('战力')}  ")
        if kind == "auto":
            print(head + "【四维唯一】→ "
                  + ("追加为新角色" if target is None
                     else f"#{target + 1} {m.rows[target].get('姓名') or '(空)'}"
                          f"（战力 {m.rows[target].get('战力')}）"))
            continue
        print(head + f"【要你定】{why}")
        shown = prefer or merge_candidates(pr.get("姓名"), m.rows, m.roster)[0]
        for t in shown[:3]:
            amb = empty_patch_slots(pr, m.rows[t])
            print(f"        → #{t + 1} {m.rows[t].get('姓名') or '(空)'}（战力 {m.rows[t].get('战力')}）"
                  + (f"  ⚠ 空槽歧义：{'/'.join(amb)}（默认保留旧值）" if amb else ""))
        if not shown:
            print("        → 无候选：界面里默认「追加为新角色」，也可以手动选一行覆盖")
    print("\n（预选只是排序：窗口里每一条都可以改，改完一起点「应用并写入主数据」。）")
    return 0


def check_row(row, affix_table, name_counts, roster):
    """一行的全部问题 -> [Prob]。short 带位置，col 指到出问题的那一列。"""
    probs = list(check_name(row.get("姓名"), roster))

    cp = row.get("战力")
    if not isinstance(cp, int) or isinstance(cp, bool):
        probs.append(Prob(F_SUSPECT, "战力缺失", "战力缺失", "战力缺失或不是整数", "战力"))
    elif not CP_MIN <= cp <= CP_MAX:
        probs.append(Prob(F_SUSPECT, "战力异常", "战力异常",
                          f"战力 {cp} 超出合理区间 {CP_MIN}~{CP_MAX}", "战力"))

    nm = row.get("姓名")
    if nm and name_counts.get(nm, 0) > 1:
        probs.append(Prob(F_WARN, "重名", "重名", f"「{nm}」在表里出现了 {name_counts[nm]} 次", "姓名"))

    for slot in GUI_SLOTS:
        s = row.get(slot) or {}
        lv = s.get("等级")
        affs = s.get("词条") or []
        real = [a for a in affs if a.get("名称") and a.get("名称") != EMPTY_AFFIX]
        if lv is not None and not 0 <= lv <= 5:
            probs.append(Prob(F_WARN, "等级越界", f"{slot}等级越界", f"{slot} 等级 {lv} 不在 0~5",
                              f"{slot}.等级"))
        if real and lv is None:
            probs.append(Prob(F_SUSPECT, "等级缺失", f"{slot}等级漏读",
                              f"{slot} 有词条但等级没读出来", f"{slot}.等级"))
        for i, a in enumerate(affs):
            n, v = a.get("名称"), a.get("数值")
            tag = f"{slot}{i + 1}"
            if n is None and v is not None:
                probs.append(Prob(F_WARN, "名称错位", f"{tag}名称漏读",
                                  f"{tag} 有数值 {v} 却没有词条名", f"{tag}名"))
            elif n == EMPTY_AFFIX and v not in (0, 0.0, None):
                probs.append(Prob(F_WARN, "空槽非零", f"{tag}空槽非零",
                                  f"{tag} 是未获得效果，数值却是 {v}", f"{tag}值"))
            elif n and n != EMPTY_AFFIX:
                if v is None:
                    probs.append(Prob(F_SUSPECT, "数值漏读", f"{tag}数值漏读",
                                      f"{tag} {n} 的数值没读出来", f"{tag}值"))
                elif affix_table:
                    levels = affix_table.get(n)
                    if levels and not tier_hit(levels, v):
                        near = min(levels, key=lambda x: abs(x - v))
                        probs.append(Prob(F_WARN, "数值越档", f"{tag}数值越档",
                                          f"{tag} {n} = {fmt_pct(v)} 不在合法档位"
                                          f"（最近 {fmt_pct(near)}）", f"{tag}值"))
    return probs


class Model:
    """把行 + 校验结果打包在一起，界面只跟它打交道。"""

    def __init__(self, rows, affix_table, roster, roster_err=None):
        # 一律先归一化：保住「Model.rows 永远是规范形状」这条不变量
        self.rows = []
        for r in rows:
            nr, _ = normalize_row(r)
            self.rows.append(nr if nr is not None
                             else dict({"姓名": None, "战力": None},
                                       **{s: empty_slot() for s in GUI_SLOTS}))
        self.affix_table = affix_table
        self.roster = roster
        self.roster_err = roster_err
        self.notes = []
        self.problems = []
        self.cp_order = {}
        self._counts = None
        self.recompute()

    def recompute(self):
        counts = {}
        for r in self.rows:
            nm = r.get("姓名")
            if nm:
                counts[nm] = counts.get(nm, 0) + 1
        self.cp_order = check_cp_order(self.rows)
        self.problems = []
        for i, r in enumerate(self.rows):
            probs = check_row(r, self.affix_table, counts, self.roster)
            if i in self.cp_order:
                probs.append(Prob(F_SUSPECT, "战力存疑", "战力存疑", self.cp_order[i], "战力"))
            self.problems.append(probs)
        self._counts = None

    def level_of(self, i):
        probs = self.problems[i]
        if any(p.level == F_SUSPECT for p in probs):
            return F_SUSPECT
        return F_WARN if probs else None

    def is_stale(self, i):
        return all(slot_is_stale(self.rows[i].get(s) or {}) for s in GUI_SLOTS)

    def marked_cols(self, i):
        """这一行里「被指到」的列 —— 那些格子会显示成 `?`。"""
        return {p.col for p in self.problems[i] if p.col}

    def problem_text(self, i):
        probs = self.problems[i]
        if not probs:
            return "整槽未采集" if self.is_stale(i) else ""
        return " ".join(f"{SYM[p.level]}{p.short}" for p in probs)

    def detail_text(self, i):
        probs = self.problems[i]
        if not probs:
            return "整槽未采集（信息性，不是错误）" if self.is_stale(i) else "没有问题"
        return " ｜ ".join(p.detail for p in probs)

    def counts(self):
        if self._counts is None:
            suspect = sum(1 for i in range(len(self.rows)) if self.level_of(i) == F_SUSPECT)
            warn = sum(1 for i in range(len(self.rows)) if self.level_of(i) == F_WARN)
            stale = sum(1 for i in range(len(self.rows)) if self.is_stale(i))
            by_code = {}
            for probs in self.problems:
                for p in probs:
                    by_code[p.code] = by_code.get(p.code, 0) + 1
            self._counts = {"suspect": suspect, "warn": warn, "stale": stale, "by_code": by_code}
        return self._counts


# ----------------------------------------------------------------------------
# 报告 / 回归（不需要 tkinter）
# ----------------------------------------------------------------------------
def build_model(path, verbose=True):
    if verbose:
        print(f"读取 {path} …")
    rows, notes = load_rows(path)
    table, terr = None, None
    try:
        table = load_affix_table(EQUIP_XLSX)
    except Exception as e:
        terr = f"词条表没载入（{type(e).__name__}: {e}），档位校验已关闭"
    roster, rerr = load_roster(ROSTER_PATH)
    if verbose:
        print(f"  {len(rows)} 行；结构异常 {len(notes)} 处")
        print(f"  词条表：{'9 词条' if table else terr}")
        print(f"  图鉴：{len(roster['keys']) if roster else 0} 人"
              + ("" if roster else f"（{rerr}，姓名校验已关闭）"))
    m = Model(rows, table, roster, rerr)
    m.notes = notes
    return m


def check_consistency():
    """本文件与 collect_cn.py 的常量/词条表必须一致（两处各留了一份，防漂移）。"""
    try:
        sys.path.insert(0, str(Path(__file__).resolve().parent))
        import collect_cn as cc
    except Exception as e:
        print(f"  跳过（collect_cn 不可导入：{type(e).__name__}）")
        return True
    ok = True
    if list(cc.AFFIX_CANON) != AFFIX_CANON:
        ok = False
        print(f"  ✗ AFFIX_CANON 不一致：{list(cc.AFFIX_CANON)} vs {AFFIX_CANON}")
    if cc.EMPTY_AFFIX != EMPTY_AFFIX:
        ok = False
        print(f"  ✗ EMPTY_AFFIX 不一致：{cc.EMPTY_AFFIX!r} vs {EMPTY_AFFIX!r}")
    if tuple(cc.ICON_DIMS) != ICON_DIMS:
        ok = False
        print(f"  ✗ ICON_DIMS 不一致：{tuple(cc.ICON_DIMS)} vs {ICON_DIMS}")
    if dict(cc.ICON_ALIAS) != ICON_ALIAS:
        ok = False
        print(f"  ✗ ICON_ALIAS 不一致：{dict(cc.ICON_ALIAS)} vs {ICON_ALIAS}")
    if getattr(cc, "ICONS_KEY", None) != "_图标":
        ok = False
        print(f"  ✗ ICONS_KEY 不一致：{getattr(cc, 'ICONS_KEY', None)!r} vs '_图标'")
    ours = load_affix_table(EQUIP_XLSX)
    theirs = cc.load_affix_table(EQUIP_XLSX)
    if ours != theirs:
        diff = [k for k in set(ours) | set(theirs) if ours.get(k) != theirs.get(k)]
        ok = False
        print(f"  ✗ 词条表解析结果不一致：{diff}")
    print("  ✓ 与 collect_cn 一致（AFFIX_CANON / EMPTY_AFFIX / ICON_DIMS / ICON_ALIAS / "
          "ICONS_KEY / 词条表解析）" if ok else "  ↑ 需要修正")
    return ok


def report(path):
    m = build_model(path)
    c = m.counts()
    print("\n=== 可疑分级（按「要不要翻回游戏核实」划）===")
    print(f"  ▲ 必须看游戏 : {c['suspect']} 行")
    print(f"  !  可当场修   : {c['warn']} 行")
    print(f"  ·  整槽未采集 : {c['stale']} 行（信息性，不是错误）")
    print("\n=== 按问题码 ===")
    for code, n in sorted(c["by_code"].items(), key=lambda kv: -kv[1]):
        print(f"  {code:<10} {n}")
    print("\n=== 明细 ===")
    for i, probs in enumerate(m.problems):
        if not probs:
            continue
        print(f"  #{i + 1:>3} {m.rows[i].get('姓名') or '(空)':<12} {SYM[m.level_of(i)]} "
              + " | ".join(p.short for p in probs))
    if Path(DEFAULT_PATCH).exists():
        prows, _ = load_rows(DEFAULT_PATCH)
        auto, amb = plan_merge(m.rows, prows, load_patch_icons(DEFAULT_PATCH),
                               load_roster_attrs(), m.roster)
        print(f"\n=== 待合并的补丁 ===\n  {DEFAULT_PATCH}：{len(prows)} 条等待合并 —— "
              f"{len(auto)} 条图标已唯一确定（窗口里会预选好）、{len(amb)} 条要你定；"
              f"都会列在合并窗里（开界面点「合并补丁」，或先 --merge-preview 干跑看一眼）")
    print("\n=== 与 collect_cn 的一致性 ===")
    check_consistency()
    return m


def roundtrip(path):
    m = build_model(path, verbose=False)
    orig = json.loads(Path(path).read_text(encoding="utf-8"))["角色"]
    tmp = Path("output/_roundtrip_check.json")
    save_rows(m.rows, tmp)
    written = json.loads(tmp.read_text(encoding="utf-8"))["角色"]
    back, _ = load_rows(tmp)
    tmp.unlink()
    same_struct, same_model, same_orig = written == m.rows, back == m.rows, orig == m.rows
    print(f"行数            : {len(m.rows)}")
    print(f"写出结构与模型一致: {same_struct}")
    print(f"读回与模型一致    : {same_model}")
    print(f"与原始文件完全一致: {same_orig}")
    if not same_orig:
        for i, (a, b) in enumerate(zip(orig, m.rows)):
            if a != b:
                print(f"  第一个差异 #{i + 1}: {json.dumps(a, ensure_ascii=False)}")
                print(f"              -> {json.dumps(b, ensure_ascii=False)}")
                break
    return same_struct and same_model and same_orig


# ----------------------------------------------------------------------------
# 表格：Canvas 自己画
# ----------------------------------------------------------------------------
class Grid:
    """Canvas 版大表。

    ⚠️ **为什么不用 ttk.Treeview**：实测本机（1920×1080、31 列 × 52 行）Treeview 每滚
    一步要 **320ms**（横向 321 / 纵向 313），而且成本 = 列数 × 可见行数，去掉 tag 底色、
    换 ttk 主题都不变（vista/clam/alt 全是 300ms+），连只有 3 列的小表都要 59ms ——
    这就是「拖动明显卡顿」的根因，不是数据量的问题。
    同样的内容用 Canvas 画：**横向 39ms / 纵向 27ms**，快 8~10 倍。别改回去。

    代价是 Treeview 白送的东西（表头、选中、单元格命中）都得自己实现 —— 但这个表格是
    规整的行列布局，全部只是加加减减，代价可控。
    """

    ROW_H = 24
    HEAD_H = 26

    def __init__(self, parent, widths, anchors, on_select, on_double, on_cancel):
        self.widths, self.anchors = list(widths), list(anchors)
        self.on_select, self.on_double, self.on_cancel = on_select, on_double, on_cancel
        self.row_ids = []          # 视图第 r 行 -> 模型行号
        self.cells = []            # [r][c] = canvas text item id
        self.rects = []            # [r] = 行背景矩形 item id
        self.sel_row = None

        self.frame = ttk.Frame(parent)
        self.hcv = tk.Canvas(self.frame, height=self.HEAD_H, highlightthickness=0,
                             background="#e9e9e9")
        self.cv = tk.Canvas(self.frame, highlightthickness=0, background="#ffffff",
                            yscrollincrement=self.ROW_H)
        self.xsb = ttk.Scrollbar(self.frame, orient="horizontal", command=self._xview)
        self.ysb = ttk.Scrollbar(self.frame, orient="vertical", command=self._yview)
        self.cv.configure(xscrollcommand=self._xscroll, yscrollcommand=self._yscroll)
        self.hcv.configure(xscrollcommand=lambda *a: None)   # 头部只跟着 body 走

        self.hcv.grid(row=0, column=0, sticky="ew")
        ttk.Label(self.frame, text="", width=2).grid(row=0, column=1, sticky="ns")
        self.cv.grid(row=1, column=0, sticky="nsew")
        self.ysb.grid(row=1, column=1, sticky="ns")
        self.xsb.grid(row=2, column=0, sticky="ew")
        self.frame.rowconfigure(1, weight=1)
        self.frame.columnconfigure(0, weight=1)

        self.cv.bind("<Button-1>", self._click)
        self.cv.bind("<Double-1>", self._double)
        self.cv.bind("<MouseWheel>", self._wheel)
        self.cv.bind("<Configure>", lambda e: (self.on_cancel(), self._draw_header()))
        self.hcv.bind("<Configure>", lambda e: self._draw_header())

    # -- 滚动 -----------------------------------------------------------------
    @staticmethod
    def _scroll_args(a, query):
        """Tk 回调 scrollcommand 时参数个数不固定：通常是 (first, last)，
        但 moveto 路径上只回传一个新起点（实测踩到过 `unknown option "0.0"` 的 TclError，
        异常在回调里会被 Tk 吞掉、只打印 —— 所以这里绝不假设参数个数）。"""
        if len(a) == 2:
            return a
        try:
            return query()               # 查询不会反过来触发回调
        except tk.TclError:
            return None

    def _xscroll(self, *a):
        got = self._scroll_args(a, self.cv.xview)
        if not got:
            return
        try:
            self.xsb.set(got[0], got[1])
            self.hcv.xview_moveto(float(got[0]))   # 表头横向跟 body 同步
        except tk.TclError:
            pass

    def _yscroll(self, *a):
        got = self._scroll_args(a, self.cv.yview)
        if not got:
            return
        try:
            self.ysb.set(got[0], got[1])
        except tk.TclError:
            pass

    def _xview(self, *a):
        self.on_cancel()
        self.cv.xview(*a)

    def _yview(self, *a):
        self.on_cancel()
        self.cv.yview(*a)

    def _wheel(self, ev):
        self.on_cancel()
        self.cv.yview_scroll(-1 * (ev.delta // 120), "units")

    def _draw_header(self):
        self.hcv.delete("all")
        widths = self._fit_widths()
        x = 0
        for col_id, w, anc in zip(COL_IDS, widths, self.anchors):
            tx = x + 6 if anc == "w" else x + w // 2
            self.hcv.create_text(tx, self.HEAD_H // 2, anchor=anc, text=col_id, fill="#333333")
            x += w
            self.hcv.create_line(x - 1, 0, x - 1, self.HEAD_H, fill="#c8c8c8")
        self.hcv.configure(scrollregion=(0, 0, x, self.HEAD_H))

    def _fit_widths(self):
        """列宽：总宽不足窗口时按比例放大填满，避免右边留一大块空白。"""
        total = sum(self.widths)
        avail = max(self.cv.winfo_width(), 100)
        if total >= avail:
            return self.widths
        k = avail / total
        out, used = [], 0
        for w in self.widths[:-1]:
            nw = int(w * k)
            out.append(nw)
            used += nw
        out.append(max(avail - used, self.widths[-1]))
        return out

    # -- 绘制 -----------------------------------------------------------------
    def build(self, row_ids, cell_fn, bg_fn):
        """整表重建。cell_fn(rid, ci) -> (text, fg)；bg_fn(rid) -> 背景色。"""
        self.on_cancel()
        self.cv.delete("all")
        self.row_ids = list(row_ids)
        widths = self._fit_widths()
        self._widths_now = widths
        n, total_w = len(self.row_ids), sum(widths)
        total_h = max(1, n * self.ROW_H)
        self.cv.configure(scrollregion=(0, 0, total_w, total_h))
        self.cells, self.rects = [], []
        for ri, rid in enumerate(self.row_ids):
            y0 = ri * self.ROW_H
            self.rects.append(self.cv.create_rectangle(0, y0, total_w, y0 + self.ROW_H,
                                                       fill=bg_fn(rid), outline=""))
            self.cv.create_line(0, y0 + self.ROW_H, total_w, y0 + self.ROW_H, fill="#e2e2e2")
            row_items, x = [], 0
            for ci, w in enumerate(widths):
                text, fg = cell_fn(rid, ci)
                anc = self.anchors[ci]
                tx = x + 6 if anc == "w" else x + w // 2
                row_items.append(self.cv.create_text(tx, y0 + self.ROW_H // 2, anchor=anc,
                                                     text=text, fill=fg))
                x += w
            self.cells.append(row_items)
        x = 0
        for w in widths[:-1]:                     # 竖分隔线画一次，贯穿全高
            x += w
            self.cv.create_line(x, 0, x, total_h, fill="#ececec")
        self.sel_row = None
        self._draw_header()
        self._draw_sel()

    def update_row(self, view_r, cells, bg):
        for item, (text, fg) in zip(self.cells[view_r], cells):
            self.cv.itemconfigure(item, text=text, fill=fg)
        self.cv.itemconfigure(self.rects[view_r], fill=bg)

    def _draw_sel(self):
        self.cv.delete("sel")
        if self.sel_row is None:
            return
        y0 = self.sel_row * self.ROW_H
        self.cv.create_rectangle(0, y0, sum(self._widths_now), y0 + self.ROW_H,
                                 outline=FG_SEL, width=2, tags="sel")

    def select(self, view_r, scroll=True):
        self.sel_row = view_r
        self._draw_sel()
        if view_r is not None and scroll:
            self.ensure_visible(view_r, None)

    # -- 命中 / 定位 ----------------------------------------------------------
    def cell_at(self, x, y):
        cx, cy = self.cv.canvasx(x), self.cv.canvasy(y)
        n = len(self.row_ids)
        if n == 0 or cy < 0 or cy >= n * self.ROW_H:
            return None, None
        r = int(cy // self.ROW_H)
        acc = 0
        for c, w in enumerate(self._widths_now):
            if cx < acc + w:
                return r, c
            acc += w
        return r, len(self._widths_now) - 1

    def cell_rect(self, view_r, ci):
        """单元格在控件坐标系里的 (x, y, w, h)，供 place() 用。"""
        x = sum(self._widths_now[:ci]) - self.cv.canvasx(0)
        y = view_r * self.ROW_H - self.cv.canvasy(0)
        return x, y, self._widths_now[ci], self.ROW_H

    def ensure_visible(self, view_r, ci=None):
        n, total_w = len(self.row_ids), sum(self._widths_now)
        total_h = max(1, n * self.ROW_H)
        vw, vh = self.cv.winfo_width(), self.cv.winfo_height()
        y0, y1 = view_r * self.ROW_H, (view_r + 1) * self.ROW_H
        vy0 = self.cv.canvasy(0)
        if y0 < vy0:
            self.cv.yview_moveto(y0 / total_h)
        elif y1 > vy0 + vh:
            self.cv.yview_moveto(max(0.0, (y1 - vh)) / total_h)
        if ci is not None:
            x0 = sum(self._widths_now[:ci])
            x1 = x0 + self._widths_now[ci]
            vx0 = self.cv.canvasx(0)
            if x0 < vx0:
                self.cv.xview_moveto(x0 / total_w)
            elif x1 > vx0 + vw:
                self.cv.xview_moveto(max(0.0, (x1 - vw)) / total_w)

    def index_of(self, rid):
        try:
            return self.row_ids.index(rid)
        except ValueError:
            return None

    # -- 事件 -----------------------------------------------------------------
    def _click(self, ev):
        r, _c = self.cell_at(ev.x, ev.y)
        if r is None:
            return
        self.on_cancel()
        self.select(r, scroll=False)
        self.on_select(self.row_ids[r])

    def _double(self, ev):
        r, c = self.cell_at(ev.x, ev.y)
        if r is None:
            return
        self.select(r, scroll=False)
        self.on_double(self.row_ids[r], c)


# ----------------------------------------------------------------------------
# 单元格内联编辑
# ----------------------------------------------------------------------------
class CellEditor:
    """独占的浮层编辑器。生命周期只有 open / finish 两个入口，finish 幂等。"""

    def __init__(self, app):
        self.app = app
        self.w = None
        self.view_r = None
        self.spec = None
        self.orig = None

    def open(self, view_r, ci, display):
        self.finish(commit=True)
        spec = COL_SPECS[ci]
        if spec.kind == "ro":
            self.app.flash(self.app.model.detail_text(self.app.grid.row_ids[view_r]))
            return
        grid = self.app.grid
        grid.ensure_visible(view_r, ci)
        grid.cv.update_idletasks()
        x, y, w, h = grid.cell_rect(view_r, ci)
        w = max(w, 92)
        if spec.kind == "level":
            ed = ttk.Combobox(grid.cv, state="readonly", width=4,
                              values=["", "0", "1", "2", "3", "4", "5"])
            ed.set(display)
        elif spec.kind == "affixname":
            vals = [""] + AFFIX_CANON + [EMPTY_AFFIX]
            ed = ttk.Combobox(grid.cv, state="normal", values=vals)
            ed.set(display)
            self._autofilter(ed, vals)
            w = 140
        elif spec.kind == "text" and self.app.model.roster:
            vals = self.app.name_choices(view_r)
            ed = ttk.Combobox(grid.cv, state="normal", values=vals)
            ed.set(display)
            self._autofilter(ed, vals)
            w = 160
        elif spec.kind == "pct" and self.app.model.affix_table:
            vals, note = self._value_choices(view_r, spec, display)
            ed = ttk.Combobox(grid.cv, state="readonly", values=vals)
            ed.set(display)
            if note:
                self.app.flash(note)
            # 只读下拉对 ↑/↓ 的原生处理各版本 Tk 不一（本地没有 tkinter 可实测），
            # 索性自己接管：原地换一格、不弹列表，按 Enter 才落定。返回 break
            # 就是为了压掉类绑定，行为完全确定。
            ed.bind("<Up>", lambda e, c=ed, v=vals: self._cycle(c, v, -1))
            ed.bind("<Down>", lambda e, c=ed, v=vals: self._cycle(c, v, 1))
            w = 104
        else:
            if spec.kind == "pct":
                # 词条表整个没载入（最常见的原因：没装 openpyxl）-> 没有档位可查。
                # **必须和「这一格真的没有档位」区分开**：两者都表现为「下拉点开是空的」，
                # 长得一模一样，上一次排查正是从那里倒着查了一圈。所以这里退回普通输入框，
                # 并把原因说出来（同 text 列在 roster 缺失时退回 Entry 的做法）。
                self.app.flash("词条表未载入（缺 openpyxl），数值列已退化成手输："
                               "pip install -r requirements.txt", "err")
            ed = ttk.Entry(grid.cv)
            ed.insert(0, display)
            ed.selection_range(0, "end")
        ed.place(x=x, y=y, width=w, height=h)
        ed.focus_set()
        ed.bind("<Return>", lambda e: self._commit_and_move(0, 1))
        ed.bind("<KP_Enter>", lambda e: self._commit_and_move(0, 1))
        ed.bind("<Escape>", lambda e: self.finish(commit=False))
        ed.bind("<Tab>", lambda e: self._commit_and_move(1, 0))
        ed.bind("<Shift-Tab>", lambda e: self._commit_and_move(-1, 0))
        ed.bind("<ISO_Left_Tab>", lambda e: self._commit_and_move(-1, 0))
        # 按**控件类型**判，不按 spec.kind 判：pct 列在词条表没载入时会退化成 Entry，
        # 而「新增一种 kind 就得记得补进这个元组」本身就是个坑（09-27 加 pct 时踩过，
        # 忘了补就会出现「点下拉没反应」）。看 isinstance 就没有第二处要同步。
        if isinstance(ed, ttk.Combobox):
            # ⚠️ 下拉框的弹出列表也会触发 FocusOut。直接提交会把刚弹出来的候选列表
            #    销毁掉 —— 表现就是「点下拉没反应、要点好几遍」（实测踩过）。
            #    所以延后一拍、确认焦点真的离开了自己和 popdown 再提交。
            ed.bind("<FocusOut>", lambda e, widget=ed: self.app.root.after(
                90, lambda: self._focus_left(widget)))
        else:
            ed.bind("<FocusOut>", lambda e: self.finish(commit=True))
        self.w, self.view_r, self.spec, self.orig = ed, view_r, spec, display

    def _value_choices(self, view_r, spec, display):
        """数值列的候选 = **同一个词条的 15 档**（不是全局大杂烩）。

        候选直接是显示串（`11.11%`）而非小数：`equipment.xlsx` 里 15 档按 4 位小数
        格式化后两两不重复（9 个词条都验过），所以拿显示串当选项不会串档，
        也就省掉了「选项 -> 小数」这层映射。写回仍走 `parse_pct_input`。

        返回 (候选, 需要提醒用户的一句话或 None)。

        ⚠️ **不能拿「候选为空」当「没档位可查」**：现值总会被追加进候选（不然 `set()`
        会落空），所以查不到档位时列表至少也有一项。真正的判据是 `model.affix_table`
        有没有载入 —— 由 `open()` 判断，没载入就退化成手输，不给一个点开是空白的下拉。
        """
        rid = self.app.grid.row_ids[view_r]
        row = self.app.model.rows[rid]
        slot = row.get(spec.slot) or {}
        affs = slot.get("词条") or []
        name = affs[spec.idx].get("名称") if spec.idx < len(affs) else None
        table = self.app.model.affix_table or {}
        note = None
        if name == EMPTY_AFFIX:
            levels = [0.0]              # 未获得效果：游戏里这格恒为 0
        elif not table:
            # 词条表整个没载入 —— 和下面「词条名未定」是**两回事**，别混成一句话：
            # 这一种是环境问题（缺 openpyxl），不是这格的数据有问题。这里只负责把话说清；
            # 「退化成手输」由 open() 查 affix_table 决定（候选不会是空列表，见上）。
            levels = []
            note = (f"{spec.col_id} 词条表未载入（缺 openpyxl），没有档位可查："
                    "pip install -r requirements.txt")
        else:
            levels = [float(v) for v in (table.get(name) or [])] if name else []
            if not levels:
                # 词条名漏读（或查不到档位）-> 没有档位可查。给全局并集兜底，免得
                # 双击出来是个空下拉框；但真正的解法是先定词条名，所以说一声。
                levels = sorted({float(v) for lv in table.values() for v in lv})
                note = f"{spec.col_id} 词条名未定，档位取的是全局并集"
        vals = [""] + [fmt_pct(v) for v in levels]
        if display and display not in vals:
            vals.append(display)   # 现值越档时也留在列表里，set() 才不会落空
        return vals, note

    @staticmethod
    def _cycle(combo, values, step):
        """↑/↓ 在候选里挪一格。只改显示不落定，Enter/Tab 才提交（同普通编辑器）。"""
        cur = combo.get()
        try:
            i = values.index(cur)
        except ValueError:
            i = 0 if step > 0 else len(values) - 1
        combo.set(values[(i + step) % len(values)])
        return "break"

    def _focus_left(self, widget):
        """焦点离开 -> 提交。两条「别急着关」的护栏都是实测踩出来的：

        1. 点下拉箭头弹出候选列表**本身就会触发 FocusOut**，直接提交会把列表销毁 ——
           表现是「点一下弹出来又立刻消失」。所以列表只要还开着就绝不提交。
        2. 窗口刚被激活那一下 `focus_get()` 会返回 None（焦点还没落到任何控件上）。
           把「查不到焦点」当成「焦点已离开」的话，**第一次点**就会中招 ——
           用户实测正是「第一次点消失、第二次就正常」（第二次窗口已经有焦点了）。
        """
        if self.w is not widget:
            return
        popdown = None
        try:
            popdown = str(widget.tk.call("ttk::combobox::PopdownWindow", widget))
            if widget.tk.call("winfo", "ismapped", popdown):
                self.app.root.after(120, lambda: self._focus_left(widget))   # 继续盯着
                return
        except Exception:
            pass
        try:
            f = self.app.root.focus_get()
        except Exception:
            f = None
        if f is None:
            return                                    # 焦点去向不明 -> 留着，别误杀
        fs = str(f)
        if fs == str(widget) or fs.startswith(str(widget)):
            return
        if popdown and fs.startswith(popdown):
            return
        self.finish(commit=True)

    @staticmethod
    def _autofilter(combo, all_values):
        """输入即过滤候选。中文名手打太痛苦，这一步最省事。"""
        def on_key(ev):
            if ev.keysym in ("Up", "Down", "Return", "KP_Enter", "Escape",
                             "Tab", "ISO_Left_Tab", "Left", "Right", "Home", "End"):
                return
            txt = combo.get().strip()
            combo["values"] = ([v for v in all_values if txt in v] or all_values) if txt else all_values
        combo.bind("<KeyRelease>", on_key)

    def _commit_and_move(self, dcol, drow):
        self.finish(commit=True)
        self.app.move_selection(drow)
        return "break"

    def finish(self, commit):
        ed, self.w = self.w, None            # ★ 先置空：destroy 会再触发一次 FocusOut
        if ed is None:
            return
        view_r, spec, orig = self.view_r, self.spec, self.orig
        self.view_r = self.spec = self.orig = None
        raw = ed.get()
        try:
            ed.destroy()
        except tk.TclError:
            pass
        if commit and raw != orig:            # finish 内禁止弹窗，反馈一律走状态栏
            self.app.apply_edit(view_r, spec, raw)


# ----------------------------------------------------------------------------
# 补丁合并审阅窗
# ----------------------------------------------------------------------------
MERGE_APPEND = "（追加为新角色）"


class MergeWindow:
    """逐条审阅 output/cn_patch.json **里的每一条**怎么并进主数据。

    ⚠️ **全部显示**（构造参数 auto + ambiguous -> `merge_entries` 摊平）。图标唯一确定的
    那些也列出来、下拉预选好、标「四维唯一」，但**不替用户按下去** —— 以前它们被静默并掉，
    用户既看不到也改不了；放一起看既方便核对，要改也省得去翻 JSON。

    一条一个块：姓名/战力/四槽摘要 + 目标下拉（**优先候选排最前并预选**，然后是其余旧行，
    首项是「追加为新角色」）+「空槽也覆盖」勾选框。四维唯一的压成一行（省高度），
    要人定的才展开带告警。顶部可切「只看待确认」。

    ⚠️ 刻意不复用主界面那张 31 列大表：这里要表达的两件事 ——「并到哪一行」和
    「空槽算不算真值」——都塞不进那 31 列；何况合并必须在**任何写入之前**逐条确认，
    而主表是「编辑即改模型」的。
    """

    def __init__(self, app, patch_path, patch_rows, auto, ambiguous):
        self.app = app
        self.patch_path = patch_path
        self.patch_rows = patch_rows
        self.auto = auto                  # {补丁下标: 目标旧行下标 或 None}
        self.states = []                  # [(补丁下标, 目标 StringVar, 空槽覆盖 BooleanVar)]
        self.blocks = []                  # [(补丁下标, kind, 块 Frame)]，供筛选显示/隐藏

        self.win = tk.Toplevel(app.root)
        self.win.title(f"合并补丁 — {patch_path}")
        self.win.transient(app.root)
        self.win.geometry("980x640")
        self.win.minsize(720, 360)

        entries = merge_entries(patch_rows, auto, ambiguous)
        n_auto = sum(1 for _i, k, _t, _w, _p in entries if k == "auto")
        n_amb = len(entries) - n_auto
        self.n_auto, self.n_amb = n_auto, n_amb

        head = ttk.Frame(self.win, padding=(8, 8, 8, 2))
        head.pack(fill="x")
        ttk.Label(head, text=f"主数据 {app.path}：{len(app.model.rows)} 行　→　"
                             f"补丁 {patch_path}：{len(patch_rows)} 条").pack(anchor="w")
        ttk.Label(head, foreground="#1a7f37",
                  text=f"其中 {n_auto} 条已由图标唯一确定（下拉已预选，标绿）；"
                       f"{n_amb} 条需人工确认。下列条目均可修改。").pack(anchor="w")
        ttk.Label(head, foreground="#666",
                  text="无法确定时选「追加为新角色」，主数据里已有的行不会被改动。"
                  ).pack(anchor="w")

        # 筛选：默认「全部」（用户要的是放一起看），要专挑待确认时再切。
        filt = ttk.Frame(self.win, padding=(8, 2, 8, 0))
        filt.pack(fill="x")
        self.only = tk.StringVar(value="all")
        ttk.Radiobutton(filt, text=f"全部 ({len(entries)})", value="all", variable=self.only,
                        command=self._refilter).pack(side="left")
        ttk.Radiobutton(filt, text=f"只看待确认 ({n_amb})", value="check", variable=self.only,
                        command=self._refilter).pack(side="left", padx=8)

        body = ttk.Frame(self.win)
        body.pack(fill="both", expand=True, padx=8, pady=4)
        cv = tk.Canvas(body, highlightthickness=0)
        sb = ttk.Scrollbar(body, orient="vertical", command=cv.yview)
        inner = ttk.Frame(cv)
        wid = cv.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda e: cv.configure(scrollregion=cv.bbox("all")))
        # 画布窗口默认只有内容宽度，块会挤成一条；跟着画布宽度走才能 fill="x" 铺满
        cv.bind("<Configure>", lambda e: cv.itemconfigure(wid, width=e.width))
        cv.configure(yscrollcommand=sb.set)
        cv.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        cv.bind("<MouseWheel>", lambda e: cv.yview_scroll(-1 * (e.delta // 120), "units"))

        for pi, kind, target, why, prefer in entries:
            self._add_block(inner, pi, patch_rows[pi], kind, target, why, prefer)
        self._refilter()

        bar = ttk.Frame(self.win, padding=(8, 4, 8, 8))
        bar.pack(fill="x")
        ttk.Button(bar, text="应用并写入主数据", command=self.apply).pack(side="left")
        ttk.Button(bar, text="取消", command=self.win.destroy).pack(side="left", padx=8)

    # -- 单条 -----------------------------------------------------------------
    def _row_label(self, idx):
        r = self.app.model.rows[idx]
        cp = r.get("战力")
        # 前缀 #序号 保证标签唯一（重名时也不会有两条一样的选项）
        return f"#{idx + 1} {r.get('姓名') or '(空)'}（战力 {cp if cp is not None else '缺失'}）"

    def _refilter(self):
        """按顶部单选显示/隐藏块。只 pack/forget，不重建 —— 隐藏的块其 StringVar 仍然有效，
        apply() 照旧读得到它，所以切筛选绝不会漏掉哪一条。"""
        only = self.only.get()
        for _i, kind, lf in self.blocks:
            if only == "check" and kind == "auto":
                lf.pack_forget()
            else:
                lf.pack(fill="x", pady=4)   # 按 blocks 顺序重 pack，可见部分仍保持原序

    def _add_block(self, parent, i, pr, kind, target, why, prefer):
        rows = self.app.model.rows
        auto = kind == "auto"
        lf = ttk.LabelFrame(parent, text=f"补丁第 {i + 1} 条", padding=(8, 4))
        lf.pack(fill="x", pady=4)
        self.blocks.append((i, kind, lf))

        # 候选顺序：prefer（图标档给的候选 / 姓名候选）排最前，接其余旧行。选哪一行仍由人按下 ——
        # 这里只决定下拉的**排序与预选**。auto 的 target 一定要在列表里且排最前（它来自图标唯一命中）。
        head_cands = list(prefer or merge_candidates(
            pr.get("姓名"), rows, self.app.model.roster)[0])
        if target is not None and target not in head_cands:
            head_cands = [target] + head_cands
        order = head_cands + [t for t in range(len(rows)) if t not in head_cands]
        labels = [MERGE_APPEND] + [self._row_label(t) for t in order]
        label2idx = {lab: t for lab, t in zip(labels[1:], order)}

        top = ttk.Frame(lf)
        top.pack(fill="x")
        cp = pr.get("战力")
        ttk.Label(top, text=f"{pr.get('姓名') or '(空)'}　战力 {cp if cp is not None else '缺失'}　",
                  font=("", 10, "bold")).pack(side="left")
        if auto:
            # 四维唯一的压成一行：只标来源，不铺四槽摘要（要核对就点开下拉，或去主表看）
            tag = "四维唯一 → 新角色" if target is None else "四维唯一"
            ttk.Label(top, text=f"[{tag}]", foreground="#1a7f37").pack(side="left")
        else:
            ttk.Label(top, text=self._slot_summary(pr), foreground="#444").pack(side="left")

        pick = ttk.Frame(lf)
        pick.pack(fill="x", pady=(2, 0))
        ttk.Label(pick, text="并到：").pack(side="left")
        if auto:
            default = MERGE_APPEND if target is None else self._row_label(target)
        else:
            default = labels[1] if head_cands else MERGE_APPEND
        var = tk.StringVar(value=default)
        cb = ttk.Combobox(pick, state="readonly", values=labels, textvariable=var, width=44)
        cb.pack(side="left")
        over = tk.BooleanVar(value=False)
        ttk.Checkbutton(pick, text="空槽也覆盖", variable=over).pack(side="left", padx=8)
        if why:
            ttk.Label(pick, text=f"（{why}）", foreground="#8a6d00").pack(side="left")

        note = ttk.Label(lf, text="", foreground="#b00020")
        note.pack(anchor="w")
        verbose = not auto   # auto 块默认不占一行废话，只有真有问题（空槽歧义）才冒出来

        def set_note(text, color):
            note.configure(text=text, foreground=color)
            if text or verbose:
                if not note.winfo_manager():
                    note.pack(anchor="w")
            elif note.winfo_manager():
                note.pack_forget()

        def refresh(*_a):
            t = label2idx.get(var.get())
            if t is None:
                set_note("作为新角色追加到末尾 —— 主数据里已有的行一条都不会动。", "#222222")
                return
            amb = empty_patch_slots(pr, rows[t])
            if amb:
                set_note(f"⚠ 空槽歧义：{'、'.join(amb)} 在补丁里是空的，而旧行有值 ——"
                         f"空到底是 T9 真值还是没读到，光看补丁分辨不了。"
                         f"默认保留旧值；确定是真 T9 才勾「空槽也覆盖」。", "#b00020")
            else:
                set_note("该行没有空槽歧义。", "#666666")

        cb.bind("<<ComboboxSelected>>", refresh)
        refresh()
        self.states.append((i, var, over))

    @staticmethod
    def _slot_summary(row):
        bits = []
        for s in GUI_SLOTS:
            slot = row.get(s) or {}
            if slot_is_stale(slot):
                bits.append(f"{s} 空")
                continue
            n = sum(1 for a in (slot.get("词条") or [])
                    if a.get("名称") and a.get("名称") != EMPTY_AFFIX)
            lv = slot.get("等级")
            bits.append(f"{s} Lv{lv if lv is not None else '?'}/{n}条")
        return " | ".join(bits)

    # -- 应用 -----------------------------------------------------------------
    def apply(self):
        # ⚠️ 先用 auto 打底：apply_patch 把「不在 decisions 里的补丁行」当**跳过**处理。
        #    正常情况下每个块都有 states（连隐藏的都有），下面会全部覆盖掉这份打底；
        #    留着它是防「某条没建成块」时被静默丢掉。形状也必须对：值是 (目标, 空槽覆盖)。
        decisions, skip = auto_decisions(self.auto), 0
        for pi, var, over in self.states:
            label = var.get()
            if label == MERGE_APPEND:
                decisions[pi] = (None, False)
                continue
            try:
                idx = int(label.split()[0].lstrip("#")) - 1
            except (ValueError, IndexError):
                skip += 1
                continue
            decisions[pi] = (idx, bool(over.get()))
        added = sum(1 for t, _o in decisions.values() if t is None)
        over_n = len(decisions) - added
        tail = f"、跳过 {skip} 条" if skip else ""
        n_auto = f"其中 {self.n_auto} 条是图标唯一确定的（已按预选执行，除非你改过）；" \
            if self.n_auto else ""
        if not messagebox.askyesno("确认合并", f"将追加 {added} 条、覆盖 {over_n} 条{tail}。\n"
                                              f"{n_auto}主数据 {self.app.path} 会先留底再写入。\n继续？"):
            return
        if self.app.merge_apply(self.patch_path, self.patch_rows, decisions):
            self.win.destroy()


# ----------------------------------------------------------------------------
# 采集对话框
# ----------------------------------------------------------------------------
class CollectDialog:
    """在界面里跑一次采集，跑完**直接接上合并** —— 不必再开另一个程序。

    这是「命令行跑完要去图形界面」那一步消失的地方：采集和校正本来就是一件事的两半，
    中间那次人工搬运（看一眼控制台、记住它让你去开 correct_gui.py、再找到合并按钮）
    没有任何价值。

    ⚠️ **采集跑在子进程里，不是 import**。collect_cn 一进来就拖上
    rapidocr / onnxruntime / numpy / PIL（实测 +56MB 内存、启动 +1.1 秒），而本界面
    一行 OCR 代码都用不到 —— 这条性质必须保住。子进程还顺带继承了 bat 给的提权。
    """

    def __init__(self, app):
        self.app = app
        self.proc = None
        self.q = queue.Queue()
        self.running = False
        self.rc = None           # 采集子进程返回码；跑完 leave_to_merge 要用

        self.win = tk.Toplevel(app.root)
        self.win.title("采集")
        self.win.transient(app.root)
        self.win.geometry("760x520")
        self.win.minsize(560, 360)

        opt = ttk.Frame(self.win, padding=(10, 10, 10, 4))
        opt.pack(fill="x")

        self.mode = tk.StringVar(value="patch")
        ttk.Label(opt, text="模式：").grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(opt, text="增量（重扫几个，填数量）", value="patch",
                        variable=self.mode, command=self._on_mode).grid(row=0, column=1, sticky="w")
        ttk.Radiobutton(opt, text="全量（全部，扫到最后一个角色）", value="full",
                        variable=self.mode, command=self._on_mode).grid(row=0, column=2, sticky="w")

        ttk.Label(opt, text="最多扫几个：").grid(row=1, column=0, sticky="w", pady=(6, 0))
        self.maxv = tk.IntVar(value=3)
        self.maxsp = ttk.Spinbox(opt, from_=1, to=500, width=6, textvariable=self.maxv)
        self.maxsp.grid(row=1, column=1, sticky="w", pady=(6, 0))

        self.shots = tk.BooleanVar(value=False)
        ttk.Checkbutton(opt, text="每个角色存 5 张截图（调试用，约 8MB/角色）",
                        variable=self.shots).grid(row=2, column=1, columnspan=2, sticky="w")

        self.warn = ttk.Label(opt, text="", foreground="#b00020", wraplength=680, justify="left")
        self.warn.grid(row=3, column=0, columnspan=3, sticky="w", pady=(6, 0))

        bar = ttk.Frame(self.win, padding=(10, 0, 10, 4))
        bar.pack(fill="x")
        self.go = ttk.Button(bar, text="开始采集", command=self.start)
        self.go.pack(side="left")
        self.stop = ttk.Button(bar, text="中断", command=self.abort, state="disabled")
        self.stop.pack(side="left", padx=6)
        # 采集跑完才可用：点它 = 收起本窗、进合并窗。跑之前/跑之中置灰，防手滑把没采完的退掉。
        self.exit_btn = ttk.Button(bar, text="保存并退出", command=self.leave_to_merge,
                                   state="disabled")
        self.exit_btn.pack(side="left", padx=6)
        self.status = ttk.Label(bar, text="", foreground="#666")
        self.status.pack(side="left", padx=8)

        body = ttk.Frame(self.win, padding=(10, 0, 10, 10))
        body.pack(fill="both", expand=True)
        self.log = tk.Text(body, wrap="none", height=12, background="#111111",
                           foreground="#dddddd", insertbackground="#dddddd")
        sb = ttk.Scrollbar(body, orient="vertical", command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        self.log.pack(side="left", fill="both", expand=True)
        sb.pack(side="right", fill="y")
        self.log.configure(state="disabled")

        self._on_mode()

    # -- 选项 ----------------------------------------------------------------
    def _on_mode(self):
        full = self.mode.get() == "full"
        # 全量 = ALL：数量由「扫到 LV.1 自动停」决定，输入框置灰（免得误以为它有用）
        self.maxv.set(FULL_MAX if full else 3)
        self.maxsp.configure(state="disabled" if full else "normal")
        self.warn.configure(text=(
            "全量：先把游戏翻到「名单最上面」（第一个角色），从那儿一路扫到最后一个"
            "（遇到 LV.1 自动停）。结果同样是补丁，跑完逐条并入主数据 —— "
            "已扫到的角色就地更新，没扫到的照旧，新角色追加，不会替换整张表。" if full else
            "增量：结果写 output/cn_patch.json，主数据一个字节都不动，跑完自动并进去。\n"
            "开始前请先在游戏里翻到「第一个」要重扫的角色，停在他的信息页。"))

    def _append(self, s):
        self.log.configure(state="normal")
        self.log.insert("end", s)
        self.log.see("end")
        self.log.configure(state="disabled")

    # -- 跑 ------------------------------------------------------------------
    def start(self):
        if self.running:
            return
        if not is_admin():
            messagebox.showerror(
                "需要管理员权限",
                "游戏是以管理员身份运行的，普通权限发出的模拟点击会被系统（UIPI）"
                "静默丢弃 —— 不报错，就是点不动。\n\n"
                "请关掉本窗口，改用 `启动采集.bat` 启动（它会自动请求提权），"
                "或者右键以管理员身份运行。")
            return
        # 两个模式的命令行**完全相同**（都只产出补丁），差别只有数量：增量填的数、全量扫到底。
        # 绝不传 --replace —— 那是界面不可达的逃生口，主数据只能经合并写入。
        full = self.mode.get() == "full"
        n = FULL_MAX if full else max(1, int(self.maxv.get() or 1))
        argv = ["--max", str(n)]
        if self.shots.get():
            argv.append("--save-shots")
        # ⚠️ -u + PYTHONUNBUFFERED：不加的话进度会攒成一坨、最后一次性冒出来，实时性全丢。
        #    显式指定 utf-8：中文 Windows 的子进程 stdout 默认走 936，父进程按 utf-8 解会花屏。
        env = dict(os.environ)
        env["PYTHONIOENCODING"] = "utf-8"
        env["PYTHONUNBUFFERED"] = "1"
        kwargs = {}
        if os.name == "nt":
            kwargs["creationflags"] = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        self._append(f"\n$ python collect_cn.py {' '.join(argv)}\n\n")
        self._append("5 秒倒计时里切回游戏（Alt+Tab），期间别碰鼠标键盘。\n"
                     + ("（全量请确认游戏已停在「名单最上面」那个角色）\n\n" if full else "\n"))
        try:
            self.proc = subprocess.Popen(
                [sys.executable, "-u", "collect_cn.py", *argv], cwd=str(HERE),
                stdout=subprocess.PIPE, stderr=subprocess.STDOUT,
                text=True, encoding="utf-8", errors="replace", env=env, **kwargs)
        except Exception as e:
            messagebox.showerror("起不来", f"{type(e).__name__}: {e}")
            return
        self.running = True
        self.go.configure(state="disabled")
        self.stop.configure(state="normal")
        self.status.configure(text="采集中…（窗口可以放后台，游戏会在前台）")
        threading.Thread(target=self._reader, args=(self.proc,), daemon=True).start()
        self.win.after(100, self._poll)

    def _reader(self, proc):
        """worker 线程：只读管道、只往队列里塞。**绝不碰 Tk**（Tk 不是线程安全的）。"""
        try:
            for line in proc.stdout:
                self.q.put(line)
        except Exception as e:
            self.q.put(f"（读取输出出错：{type(e).__name__}: {e}）\n")
        self.q.put(("__DONE__", proc.wait()))

    def _poll(self):
        """主线程：把队列里的东西搬进文本框。Tk 只在这里被碰。"""
        drained = []
        rc = None
        try:
            while True:
                item = self.q.get_nowait()
                if isinstance(item, tuple) and item[0] == "__DONE__":
                    rc = item[1]
                    break
                drained.append(item)
        except queue.Empty:
            pass
        if drained:
            self._append("".join(drained))
        if rc is None:
            self.win.after(100, self._poll)
            return
        self._append(f"\n--- 采集进程退出（返回码 {rc}）---\n")
        self.running = False
        self.rc = rc
        self.go.configure(state="normal")
        self.stop.configure(state="disabled")
        self.status.configure(text="已结束")
        rows = self.app.collect_finished(self.mode.get(), rc)
        if not rows:
            # 没产生补丁（没扫到 / 中途中断）：不弹「完毕」，留在窗口里看日志、重跑
            return
        self.exit_btn.configure(state="normal")
        self.leave_to_merge()      # 跑完主动问一次，省得用户去摸右上角的 X

    def leave_to_merge(self):
        """收起采集窗、接上合并窗 —— 采集的终点就是逐条确认合并。

        两条路通向这里：跑完自动弹的确认框（建议 ①），和底部的「保存并退出」按钮
        （建议 ②）。取消则留在本窗，之后仍可从主界面点「合并补丁」。
        """
        if self.running:
            return
        if not self.app._patch_rows():
            self.app.flash("没有产生补丁，无从合并 —— 在窗口里看日志、重跑一次", "err")
            return
        if not messagebox.askokcancel(
                "采集完毕", "采集完毕，数据已保存。\n\n"
                "点「确定」关闭本窗口，进入校正窗口逐条确认后并入主数据。\n"
                "点「取消」留在本窗口，之后可从主界面点「合并补丁」。",
                parent=self.win):
            return
        self.win.destroy()
        self.app.after_collect(self.mode.get(), self.rc)

    def abort(self):
        if not self.running or self.proc is None:
            return
        if self.proc.poll() is None:
            self._append("\n（正在中断…）\n")
            self.proc.terminate()


# ----------------------------------------------------------------------------
# 主界面
# ----------------------------------------------------------------------------
class App:
    def __init__(self, root, model, path, patch_path=DEFAULT_PATCH):
        self.root = root
        self.model = model
        self.path = path
        self.patch_path = patch_path
        self.dirty = False
        self.archived = False
        self.backup = None
        self.filter_mode = "all"
        self.editor = CellEditor(self)

        root.title(f"妮姬采集结果校正 — {path}")
        w, h = 1500, 860
        root.geometry(f"{w}x{h}+{max(0, (root.winfo_screenwidth() - w) // 2)}"
                      f"+{max(0, (root.winfo_screenheight() - h) // 3)}")
        root.minsize(900, 500)
        root.protocol("WM_DELETE_WINDOW", self.on_close)
        self._build()
        self.refresh_hint()
        self.rerender()
        self.refresh_status()

    # -- 构建 -----------------------------------------------------------------
    def _build(self):
        bar = ttk.Frame(self.root, padding=(6, 4))
        bar.grid(row=0, column=0, sticky="ew")
        ttk.Button(bar, text="采集…", command=self.open_collect).pack(side="left", padx=(0, 12))
        ttk.Button(bar, text="保存 JSON", command=self.save).pack(side="left")
        ttk.Label(bar, text="(Ctrl+S)").pack(side="left", padx=(2, 10))
        ttk.Button(bar, text="重新载入", command=self.reload).pack(side="left", padx=(0, 12))
        self.merge_btn = ttk.Button(bar, text="合并补丁 …", command=self.open_merge)
        self.merge_btn.pack(side="left", padx=(0, 12))
        self._update_merge_btn()
        ttk.Label(bar, text="显示：").pack(side="left")
        self.filter_var = tk.StringVar(value="all")
        for val, text in (("all", "全部"), ("suspect", "只看 ▲"), ("both", "只看 ▲+!")):
            ttk.Radiobutton(bar, text=text, value=val, variable=self.filter_var,
                            command=self.on_filter).pack(side="left")
        ttk.Button(bar, text="下一个问题行 (F3)",
                   command=self.jump_next_problem).pack(side="left", padx=12)
        self.hint = ttk.Label(bar, text="", foreground="#666")
        self.hint.pack(side="left", padx=8)

        self.grid = Grid(self.root, COL_WIDTHS, COL_ANCHORS,
                         on_select=self.on_row_selected,
                         on_double=self.on_cell_double,
                         on_cancel=lambda: self.editor.finish(True))
        self.grid.frame.grid(row=1, column=0, sticky="nsew")
        self.root.rowconfigure(1, weight=1)
        self.root.columnconfigure(0, weight=1)

        self.stbar = ttk.Label(self.root, anchor="w", relief="sunken", padding=(6, 2))
        self.stbar.grid(row=2, column=0, sticky="ew")
        self.root.bind("<Control-s>", lambda e: self.save())
        self.root.bind("<F3>", lambda e: self.jump_next_problem())

    # -- 渲染 -----------------------------------------------------------------
    def visible(self):
        n = len(self.model.rows)
        if self.filter_mode == "all":
            return list(range(n))
        want = (F_SUSPECT,) if self.filter_mode == "suspect" else (F_SUSPECT, F_WARN)
        return [i for i in range(n) if self.model.level_of(i) in want]

    def cell_text(self, rid, ci):
        """(文本, 颜色)。漏读的格子显示 `?` —— 这样「哪个格子漏了」一眼可见，不用去翻。"""
        spec = COL_SPECS[ci]
        if spec.kind == "ro":
            return self.model.problem_text(rid), FG_MARK if self.model.level_of(rid) else "#666666"
        txt = render_cell(self.model.rows[rid], spec)
        marked = spec.col_id in self.model.marked_cols(rid)
        if not txt and marked:
            return "?", FG_MARK
        if self.model.is_stale(rid):
            return txt, FG_STALE
        return txt, FG_NORMAL

    def row_bg(self, rid):
        return ROW_BG[self.model.level_of(rid)]

    def rerender(self, keep=None):
        self.editor.finish(True)
        if keep is None:
            keep = self.current_index()
        self.grid.build(self.visible(), self.cell_text, self.row_bg)
        view_r = self.grid.index_of(keep) if keep is not None else None
        self.grid.select(view_r, scroll=False)

    def refresh_row(self, rid):
        view_r = self.grid.index_of(rid)
        if view_r is None:
            return
        self.grid.update_row(view_r, [self.cell_text(rid, ci) for ci in range(len(COL_SPECS))],
                             self.row_bg(rid))

    def current_index(self):
        if self.grid.sel_row is None or self.grid.sel_row >= len(self.grid.row_ids):
            return None
        return self.grid.row_ids[self.grid.sel_row]

    # -- 事件 -----------------------------------------------------------------
    def on_row_selected(self, rid):
        self.refresh_status()

    def on_cell_double(self, rid, ci):
        view_r = self.grid.index_of(rid)
        spec = COL_SPECS[ci]
        self.editor.open(view_r, ci, render_cell(self.model.rows[rid], spec)
                         if spec.kind != "ro" else "")

    def apply_edit(self, view_r, spec, raw):
        rid = self.grid.row_ids[view_r]
        ok, err = set_cell(self.model.rows[rid], spec, raw)
        if not ok:
            self.flash(f"未改动：{err}", "err")
            return
        self.model.recompute()
        if self.filter_mode != "all" and rid not in self.visible():
            self.rerender(keep=rid)          # 「只看问题行」下这行已经不满足了 -> 整表重建
        else:
            self.refresh_row(rid)
            self.grid.select(view_r, scroll=False)
        self.dirty = True
        self.root.title(f"* 妮姬采集结果校正 — {self.path}")
        self.refresh_status()

    def move_selection(self, drow):
        if drow == 0 or self.grid.sel_row is None:
            return
        n = len(self.grid.row_ids)
        r = min(n - 1, max(0, self.grid.sel_row + drow))
        self.grid.select(r)
        self.refresh_status()

    # -- 筛选 / 跳转 ----------------------------------------------------------
    def on_filter(self):
        self.filter_mode = self.filter_var.get()
        self.rerender()

    def jump_next_problem(self, ev=None):
        probs = [i for i in self.visible() if self.model.level_of(i)]
        if not probs:
            self.flash("当前筛选下没有问题行", "info")
            return
        cur = self.current_index()
        nxt = next((i for i in probs if cur is not None and i > cur), probs[0])
        view_r = self.grid.index_of(nxt)
        if view_r is None:
            return
        self.grid.select(view_r)
        # 滚到**出问题的那一列**（不是固定的「问题」列）—— 这才叫「找到是哪个漏了」
        cols = self.model.marked_cols(nxt)
        target = COL_INDEX["问题"]
        if cols:
            target = min(COL_INDEX[c] for c in cols if c in COL_INDEX)
        self.grid.ensure_visible(view_r, target)
        self.refresh_status()
        self.flash(f"第 {probs.index(nxt) + 1}/{len(probs)} 个问题行 → "
                   f"#{nxt + 1} {self.model.rows[nxt].get('姓名') or '(空)'}", "warn")

    # -- 保存 / 重载 ----------------------------------------------------------
    def save(self):
        if not self.dirty:
            self.flash("没有修改，未写盘", "info")
            return
        c = self.model.counts()
        if c["suspect"] and not messagebox.askyesno(
                "仍有可疑项", f"还有 {c['suspect']} 行标着 ▲（需人工确认）。仍然写入？"):
            return
        if not self.archived:
            # ★ 只在本次会话首次保存留底：时间戳只到秒，每次保存都留底会被同秒的
            #   下一份静默覆盖 —— 而那正是留底要防的事。
            self.backup = archive_previous(self.path)
            self.archived = True
        try:
            save_rows(self.model.rows, self.path)
        except Exception as e:
            messagebox.showerror("写盘失败", f"{type(e).__name__}: {e}")
            self.flash(f"写盘失败：{e}", "err")
            return
        self.dirty = False
        self.root.title(f"妮姬采集结果校正 — {self.path}")
        self.flash(f"已写入 {self.path}（原文件留底：{self.backup or '无旧数据'}）", "info")

    def reload(self):
        if self.dirty and not messagebox.askyesno("放弃修改", "有未保存的修改，重新载入会丢弃。继续？"):
            return
        self.load_from_disk()
        self.flash(f"已重新载入 {self.path}", "info")

    def load_from_disk(self):
        self.editor.finish(True)
        self.model = build_model(self.path, verbose=False)
        self.dirty = False
        self.archived = False
        self.rerender()
        self.refresh_hint()
        self._update_merge_btn()
        self.refresh_status()

    # -- 采集 -----------------------------------------------------------------
    def open_collect(self):
        """开采集对话框。留住引用：Toplevel 由 tk 保活，但 Python 侧的实例只被闭包引用。"""
        self.editor.finish(True)
        if self.dirty and not messagebox.askyesno(
                "先保存", "当前有未保存的修改。采集结果会重新载入主数据，未保存的改动会丢。\n先保存？"):
            return
        if self.dirty:
            self.save()
            if self.dirty:          # save() 在「还有 ▲ 项」那一步被取消掉了
                return
        self._collect_win = CollectDialog(self)

    def collect_finished(self, mode, rc):
        """采集子进程刚退出：更新状态栏与主界面的合并按钮，返回补丁行（没有则 None）。

        只做「通知与刷新」，不弹窗也不合并 —— 那两件事由采集窗决定（它还要先问一句
        用户「现在进合并窗口吗」）。返回 None 时调用方留在采集窗里，别把没数据的窗口关掉。
        """
        rows = self._patch_rows()
        what = "全量" if mode == "full" else "增量"
        if rows:
            self.flash(f"{what}采集完成，补丁 {len(rows)} 条", "info")
        else:
            self.flash(f"{what}采集结束（返回码 {rc}），没有产生补丁 —— 可能没扫到角色或"
                       f"中途中断（{self.patch_path}）", "err")
        self._update_merge_btn()
        return rows

    def after_collect(self, mode, rc):
        """采集子进程退出、用户确认「进合并窗」后的收尾 —— 这一步就是「命令行跑完还得去
        图形界面」消失的地方。

        增量与全量走同一条路：采集都只产出补丁（主数据一个字节没动），这里接上合并。
        区别只剩扫了多少 —— 「全量」扫到最后一个角色，所以并入的条目多，而已。
        """
        rows = self._patch_rows()
        if not rows:
            return
        what = "全量" if mode == "full" else "增量"
        self.flash(f"{what}采集完成，补丁 {len(rows)} 条，正在合并…", "info")
        self.merge_patch(self.patch_path, rows)

    # -- 补丁合并 -------------------------------------------------------------
    def _patch_rows(self):
        """补丁文件存在且非空时返回它的行，否则 None。每次都重读：补丁是外部产物。"""
        if not Path(self.patch_path).exists():
            return None
        rows, _notes = load_rows(self.patch_path)
        return rows or None

    def _update_merge_btn(self):
        rows = self._patch_rows()
        self.merge_btn.configure(state="normal" if rows else "disabled",
                                 text=f"合并补丁 ({len(rows)}) …" if rows else "合并补丁 …")

    def open_merge(self):
        self.editor.finish(True)
        rows = self._patch_rows()
        if not rows:
            self.flash(f"没有补丁（{self.patch_path}）—— 先在「采集…」里跑一次增量扫描", "info")
            self._update_merge_btn()
            return
        if self.dirty:
            # 合并会把模型整份重写并重载，未保存的手改会一起没掉 —— 先落盘，别默默丢
            if not messagebox.askyesno("先保存", "当前有未保存的修改。先保存再合并？"):
                return
            self.save()
            if self.dirty:      # save() 在「还有 ▲ 项」那一步被取消掉了
                return
        self.merge_patch(self.patch_path, rows)

    def merge_patch(self, patch_path, patch_rows):
        """合并的**唯一入口**：先分类（图标四维定人），再开审阅窗。采集跑完也走这里。

        分类依据是 collect_cn 写进补丁顶层的图标四维（load_patch_icons）—— 与姓名完全无关的
        机器证据。**窗口里全部列出**：图标唯一确定的那些预选好、标绿，但同样可看可改；
        定不下来的带理由展开。不再有「全自动就跳过窗口」的旁路 —— 那正是以前
        「窗口瞒着你把几条直接并了」的来源。
        """
        icons = load_patch_icons(patch_path)
        attrs = load_roster_attrs()
        auto, ambiguous = plan_merge(self.model.rows, patch_rows, icons, attrs, self.model.roster)
        # 留住引用：Toplevel 本身由 tk 的对象树保活，但 Python 侧的 MergeWindow 实例
        # 只被按钮 command / 闭包引用着，显式留一份最不容易踩坑
        self._merge_win = MergeWindow(self, patch_path, patch_rows, auto, ambiguous)
        return None

    def merge_apply(self, patch_path, patch_rows, decisions):
        """执行合并：改模型 -> 按战力重排 -> 主数据留底后写入 -> 退掉补丁文件 -> 重载。

        返回是否成功。**先重排再写盘**：补丁换了战力，不重排就是造一处假的「违反降序」。
        """
        added, over, notes = apply_patch(self.model.rows, patch_rows, decisions)
        sort_by_cp(self.model.rows)
        try:
            self.backup = archive_previous(self.path)
            save_rows(self.model.rows, self.path)
        except Exception as e:
            messagebox.showerror("写盘失败", f"{type(e).__name__}: {e}\n"
                                             f"（模型已改但没落盘，重载即可回到磁盘上的状态）")
            return False
        self.archived = True
        retired = archive_previous(patch_path)   # 退掉补丁，防手滑再合并一次
        self.dirty = False
        self.load_from_disk()
        # 用可滚动窗，不用 messagebox：notes 逐条列（全量合并时几十行），
        # messagebox 会被撑得比屏幕还高，把「确定」顶出去。
        show_scrolled_text(self.root, "合并完成", "\n".join(notes) +
                           f"\n\n主数据：{self.path}（原文件留底：{self.backup or '无旧数据'}）"
                           f"\n补丁已退成：{retired or patch_path}"
                           f"\n（如需再合并，重新跑一次采集生成新补丁）")
        self.flash(f"合并完成：新增 {added} 覆盖 {over} "
                   f"跳过 {len(patch_rows) - added - over}", "info")
        return True

    def refresh_hint(self):
        """顶部那句提示：**有话说时变红**。

        灰字的提示实测会被当成装饰看漏 —— 缺 openpyxl 时它是「数值列下拉是空的」
        唯一线索，看过了也没意识到是故障。所以非空就上警色。
        """
        txt = self.hint_text()
        self.hint.configure(text=txt, foreground=FG_MARK if txt else "#666666")

    def hint_text(self):
        bits = []
        if self.model.roster is None:
            bits.append("图鉴未载入，姓名校验已关闭")
        if self.model.affix_table is None:
            bits.append("词条表未载入（缺 openpyxl？），数值列已退化成手输")
        return " / ".join(bits)

    def on_close(self):
        if self.dirty and not messagebox.askyesno("退出", "有未保存的修改，确定退出？"):
            return
        self.root.destroy()

    # -- 状态栏 ---------------------------------------------------------------
    def _set_status(self, msg, level):
        color = {"info": "#222222", "warn": "#8a6d00", "err": "#b00020"}.get(level, "#222222")
        self.stbar.configure(text=msg, foreground=color)

    def flash(self, msg, level="info"):
        self._set_status(msg, level)

    def refresh_status(self):
        rid = self.current_index()
        if rid is None:
            c = self.model.counts()
            self._set_status(f"共 {len(self.model.rows)} 行 · ▲{c['suspect']} !{c['warn']} ，"
                             f"双击单元格编辑，F3 跳下一个问题行", "info")
            return
        row = self.model.rows[rid]
        self._set_status(f"第 {rid + 1}/{len(self.model.rows)} 行 · {row.get('姓名') or '(空)'} · "
                         f"战力 {row.get('战力') if row.get('战力') is not None else '缺失'} · "
                         f"{self.model.detail_text(rid)}",
                         "warn" if self.model.problems[rid] else "info")

    def name_choices(self, view_r):
        """姓名下拉候选：本行的图鉴候选排最前，方便 ↓+Enter 直接选。"""
        roster = self.model.roster
        if not roster:
            return []
        row = self.model.rows[self.grid.row_ids[view_r]]
        cands = name_candidates(norm_roster_name(row.get("姓名") or ""), roster) \
            if row.get("姓名") else []
        rest = [roster["orig"][k] for k in roster["keys"]]
        return cands + [x for x in rest if x not in cands]


# ----------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="妮姬采集结果校正界面")
    ap.add_argument("--file", default=DEFAULT_JSON, help=f"要校正的 JSON（默认 {DEFAULT_JSON}）")
    ap.add_argument("--patch", default=DEFAULT_PATCH,
                    help=f"补丁文件（默认 {DEFAULT_PATCH}，有它才会出现「合并补丁」按钮）")
    ap.add_argument("--report", action="store_true", help="不开窗，只打印校验基线")
    ap.add_argument("--roundtrip", action="store_true", help="读入→原样写出→读回，验证零丢失")
    ap.add_argument("--merge-preview", action="store_true",
                    help="不开窗，只打印补丁会怎么合并（干跑，不写任何文件）")
    args = ap.parse_args()

    # 所有相对路径（output/、equipment.xlsx、docs/）都依赖 cwd，先切到脚本目录
    os.chdir(Path(__file__).resolve().parent)

    if args.report:
        report(args.file)
        return 0
    if args.roundtrip:
        return 0 if roundtrip(args.file) else 1
    if args.merge_preview:
        return merge_preview(args.file, args.patch)

    if tk is None:
        print("这台机器没有 tkinter，界面起不来。可以先用 --report 看校验结果。")
        return 1

    root = tk.Tk()
    root.withdraw()
    boot = tk.Toplevel(root)
    boot.title("载入中")
    ttk.Label(boot, text="正在载入数据与词条表 …", padding=24).pack()
    boot.update_idletasks()
    boot.geometry(f"+{boot.winfo_screenwidth() // 2 - 100}+{boot.winfo_screenheight() // 2 - 40}")
    try:
        model = build_model(args.file, verbose=False)
        boot.destroy()
        root.deiconify()
        App(root, model, args.file, args.patch)
        root.mainloop()
    finally:
        try:
            boot.destroy()
        except tk.TclError:
            pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
