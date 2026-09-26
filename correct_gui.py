#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""国服 NIKKE 采集结果校正界面（tkinter）。

跑法（Windows，项目根目录；也可以从 启动采集.bat 选 [5]）：
    python correct_gui.py                 打开界面，默认读写 output/cn_collect.json
    python correct_gui.py --file X.json   换文件
    python correct_gui.py --report        不开窗，只打印校验基线（无 tkinter 的机器也能跑）
    python correct_gui.py --roundtrip     读入→原样写出→读回，验证明细零丢失、格式零改动

界面是一张 31 列的大表，一行一个角色：
    姓名 | 战力 | 问题 | 头.等级 头1名 头1值 头2名 头2值 头3名 头3值 | 甲… | 手… | 脚…
双击单元格进入编辑。姓名/词条名是下拉框（可自由输入），等级是只读下拉，数值接受
`11.11%` 和 `0.1111` 两种写法（内部一律存小数）。

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
import difflib
import json
import os
import re
import sys
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

# 与 collect_cn.py 保持一致（见文件头说明）。空词条占位是游戏里的原文，
# 采集端也拿它当「这块是 T10」的铁证。
EMPTY_AFFIX = "未获得效果"
AFFIX_CANON = ["攻击力增加", "蓄力伤害增加", "防御力增加", "暴击伤害增加", "命中率增加",
               "暴击率增加", "最大装弹数增加", "蓄力速度增加", "优越代码伤害增加"]

SYM = {"suspect": "▲", "warn": "!"}
ROW_BG = {"suspect": "#ffd6a5", "warn": "#fff2cc", None: "#ffffff"}
F_SUSPECT, F_WARN = "suspect", "warn"
FG_NORMAL, FG_STALE, FG_MARK, FG_SEL = "#000000", "#9a9a9a", "#c00000", "#1a73e8"

CP_MIN, CP_MAX = 1000, 2_000_000  # 战力合理区间
CP_ORDER_RATIO = 0.03             # 战力违反降序的判定阈值，压掉 OCR 抖动


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
    ours = load_affix_table(EQUIP_XLSX)
    theirs = cc.load_affix_table(EQUIP_XLSX)
    if ours != theirs:
        diff = [k for k in set(ours) | set(theirs) if ours.get(k) != theirs.get(k)]
        ok = False
        print(f"  ✗ 词条表解析结果不一致：{diff}")
    print("  ✓ 与 collect_cn 一致（AFFIX_CANON / EMPTY_AFFIX / 词条表解析）" if ok else "  ↑ 需要修正")
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
        else:
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
        if spec.kind not in ("level", "affixname", "text"):
            ed.bind("<FocusOut>", lambda e: self.finish(commit=True))
        else:
            # ⚠️ 下拉框的弹出列表也会触发 FocusOut。直接提交会把刚弹出来的候选列表
            #    销毁掉 —— 表现就是「点下拉没反应、要点好几遍」（实测踩过）。
            #    所以延后一拍、确认焦点真的离开了自己和 popdown 再提交。
            ed.bind("<FocusOut>", lambda e, widget=ed: self.app.root.after(
                90, lambda: self._focus_left(widget)))
        self.w, self.view_r, self.spec, self.orig = ed, view_r, spec, display

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
# 主界面
# ----------------------------------------------------------------------------
class App:
    def __init__(self, root, model, path):
        self.root = root
        self.model = model
        self.path = path
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
        self.hint.configure(text=self.hint_text())
        self.rerender()
        self.refresh_status()

    # -- 构建 -----------------------------------------------------------------
    def _build(self):
        bar = ttk.Frame(self.root, padding=(6, 4))
        bar.grid(row=0, column=0, sticky="ew")
        ttk.Button(bar, text="保存 JSON", command=self.save).pack(side="left")
        ttk.Label(bar, text="(Ctrl+S)").pack(side="left", padx=(2, 10))
        ttk.Button(bar, text="重新载入", command=self.reload).pack(side="left", padx=(0, 12))
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
        self.hint.configure(text=self.hint_text())
        self.refresh_status()

    def hint_text(self):
        bits = []
        if self.model.roster is None:
            bits.append("图鉴未载入，姓名校验已关闭")
        if self.model.affix_table is None:
            bits.append("词条表未载入，档位校验已关闭")
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
    ap.add_argument("--report", action="store_true", help="不开窗，只打印校验基线")
    ap.add_argument("--roundtrip", action="store_true", help="读入→原样写出→读回，验证零丢失")
    args = ap.parse_args()

    # 所有相对路径（output/、equipment.xlsx、docs/）都依赖 cwd，先切到脚本目录
    os.chdir(Path(__file__).resolve().parent)

    if args.report:
        report(args.file)
        return 0
    if args.roundtrip:
        return 0 if roundtrip(args.file) else 1

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
        App(root, model, args.file)
        root.mainloop()
    finally:
        try:
            boot.destroy()
        except tk.TclError:
            pass
    return 0


if __name__ == "__main__":
    sys.exit(main())
