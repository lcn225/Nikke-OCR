#!/usr/bin/env python3
"""国服 NIKKE 角色装备采集脚本（独立命令行，无 GUI 依赖）。

流程（假设已停在角色信息页 ps.png）：
  循环：
    1. 读角色信息页：名字 / 等级(LV.x) / 战力
    2. 若等级 == 1，停止（按等级降序排列，LV.1 是最后一行）
    3. 依次点 4 个装备槽(头/甲/手/脚)进装备页，读装备名 / 等级 / T10 三行词条，再点同一槽位关闭
    4. 输出该角色一行 JSON
    5. 点「>>」翻页箭头切下一个角色

运行：
  在线(Windows 本机，pyautogui 截屏+点击)： python collect_cn.py
  离线(用截图文件模拟，验证逻辑)：   python collect_cn.py --offline

输出：output/cn_collect.json —— 每角色一行：姓名、战力、头/身/手/足 各{等级, 词条[3]}。
      开跑前会把上一轮的这个文件改名加时间戳留底（cn_collect.<YYYYMMDD-HHMMSS>.json），
      避免新一轮直接覆盖掉上一轮的结果。
词条规范名取 9 个游戏内名称（都带「增加」）；数值为百分比小数(11.81% → 0.1181)。
T9/T10 判定：装备页读到词条即 T10，无词条即 T9（词条留空）。
"""
import argparse
import json
import re
import sys
import time
from pathlib import Path

import numpy as np
import openpyxl
from PIL import Image
from rapidocr_onnxruntime import RapidOCR

# ----------------------------------------------------------------------------
# 坐标常量（1920x1080 全屏，官方 PC 客户端；像素分析 + 模板匹配标定）
# ----------------------------------------------------------------------------
# ---- 角色信息页 ps.png ----
PS_NAME = (1749, 233, 1840, 266)      # 角色名
PS_LEVEL = (1599, 240, 1725, 278)     # 等级 "LV.436/200"
PS_CP = (1726, 338, 1834, 385)        # 战斗力 "175222"（不含下方 BATTLE 行）

SLOTS = {                                # 4 装备槽点击中心（2x2 布局）
    "头": (1661, 775),                   # 左上
    "甲": (1793, 779),                   # 右上（身）
    "手": (1661, 889),                   # 左下
    "脚": (1793, 891),                   # 右下
}
NEXT_PAGE = (1892, 536)                # 「>>」翻页箭头（用户标定 1877,458~1906,614）
PREV_PAGE = (35, 582)                  # 「<<」左箭头（回上一角色 20,550~50,614，备用）

# ---- 装备页 equipment.png ----
# 词条名带 / 数值带。⚠️ 都要留足上下余量：词条区的行会整体上下移位 ~25px
# （同一角色不同槽位：头/甲 在 y768/804/840，手/脚 在 y743/779/815），
# 且「效果数值将在进入战斗时生效。」这句提示在有位移时会掉进词条区占掉一行。
# 名字带右边界要够到 x1005：空槽占位文字「未获得效果」实测在 x920-996，右边界卡 972
# 会把它截断成 '未获得效' 而被当成杂音丢掉，空槽的位置信息就没了（会被挤到末位）。
# 伸进数值列无妨——数字没有汉字，norm_affix 出来是 None，会被当杂音丢掉。
# 词条名带 / 数值带。三行词条实测占 ref 系 y734~824，但**不要把框收窄到刚好包住**：
# 实测收窄到 725~840 后，004_手 的 蓄力速度增加 从能读到 '1.98%' 变成完全读不出
# （裁剪高度会改变 OCR 的成行/上下文，不是"垃圾越少越干净"）。留到 880 是经验值。
AFFIX_NAME_BAND = (815, 700, 1005, 880)
AFFIX_VAL_BAND = (960, 700, 1085, 880)

# ---- 装备页对齐：两个锚点 ----
# ⚠️ 面板高度不是固定的：实测有 y33~989（高 956）、y59~964（高 905），还有 y85~939
# （高 854）第三种，顶部按 26px 一档差。所以面板内的固定绝对坐标都是在赌尺寸。
#
# 更要命的是：**面板顶对齐只对描述以上的区域成立**。
# 早先按「面板内一切居于面板顶」平移（24 张实测极差仅 1~2px）—— 那 24 张的描述恰好
# 行数相同，是个假规律。装备描述是 flavor text，**行数随装备不同**（实测 3~5 行），
# 它下面的「装备能力值 / 改造装备效果 / 词条三行」整体跟着描述行数平移，一行 ≈ 25.8px。
# 只按面板顶对齐时，描述行数不同的页面会整体错位 —— 实测 37 人真机采集里 21 人读崩：
# 词条顶行掉出框顶、剩下两行被当成第 1/2 行往上串位、能力值读空、等级反推 None。
#
# 于是拆成两个锚点（见 read_equip_page）：
#   描述**之上**（类型标签）-> panel_top
#   描述**之下**（能力值 + 词条）-> desc_bottom
PANEL_TOP_REF = 33          # 标准面板顶；描述之上的区域按这个顶标定
PANEL_PROBE = (700, 760)    # 面板左内侧一条竖条，用来测亮度剖面找面板顶
PANEL_BRIGHT = 150          # 竖条行均值 > 此值算面板内部

DESC_TOP_OFF = 260          # 从 面板顶+260 开始扫描述（首行文字顶实测 273~274，留 13px 余量；
                            # 上面的「持有数」实测在 +247，不会被扫进来）
DESC_PROBE = (720, 1200)    # 描述区横向范围（避开左侧装备图标列）
DESC_WIN = 440              # 描述区纵向搜索高度（够 17 行）
DESC_GAP = 47               # 切描述块的 gap 阈值：行距 13、段落空行 39 都 < 47，
                            # 而 描述块 ->「装备能力值」那一行 的 gap 是 55 > 47。
                            # ⚠️ 两侧只差 16px 余量，所以下面还有一道自校验兜底
DESC_TITLE_DY = 55          # 描述底 ->「装备能力值」标题行起点（实测 3 样本，差 ≤1px）
DESC_BOTTOM_REF = 424       # 标准版面（panel_top=33）下描述块底边的 y；描述之下的区域按它标定


def _shift(box, dy):
    """把一个区域框按锚点偏移量 dy 平移。

    dy 有两个来源（见 read_equip_page）：描述**之上**的区域（类型标签）跟面板顶，
    描述**之下**的区域跟描述块底边。别把两者混用 —— 类型标签在描述之上，
    用描述锚点会把推到面板外面去。
    """
    return (box[0], box[1] + dy, box[2], box[3] + dy)


def panel_top(img):
    """装备页白面板的上边界 y。探测失败退回标准值 PANEL_TOP_REF。"""
    g = np.array(img.convert("L")).astype(int)
    prof = g[:, PANEL_PROBE[0]:PANEL_PROBE[1]].mean(axis=1)
    ys = np.where(prof > PANEL_BRIGHT)[0]
    if len(ys) == 0:
        return PANEL_TOP_REF
    segs = [s for s in np.split(ys, np.where(np.diff(ys) > 3)[0] + 1) if len(s) > 50]
    return int(max(segs, key=len)[0]) if segs else PANEL_TOP_REF


def desc_bottom(img, top):
    """装备描述文字块的底边 y（纯像素、无 OCR，~5ms）。切不准返回 None。

    从 面板顶+DESC_TOP_OFF 起扫暗文字行段，按 DESC_GAP 合并成块：描述里的段落空行
    （gap 39）会被并进来，而描述块到「装备能力值」那一行的 gap 是 55，会把块切开。

    最后自校验一次：切点下方 DESC_TITLE_DY 处应当正好是「装备能力值」标题行。
    描述里万一出现连续两行空行（gap 会 > 47）合并就会提前停，这道校验会失败 ——
    此时返回 None 让调用方回退到面板顶：宁可退回旧的（有时会错的）行为，
    也不要用一个错的锚点把所有字段一起读歪。
    """
    g = np.array(img.convert("L")).astype(int)
    dark = (g[:, DESC_PROBE[0]:DESC_PROBE[1]] < 128).sum(axis=1)
    runs, start = [], None
    for y in range(top + DESC_TOP_OFF, min(top + DESC_TOP_OFF + DESC_WIN, len(dark))):
        if dark[y] >= 6:
            if start is None:
                start = y
        elif start is not None:
            if y - start >= 3:
                runs.append((start, y - 1))
            start = None
    if start is not None:
        runs.append((start, top + DESC_TOP_OFF + DESC_WIN - 1))
    if not runs:
        return None

    end = runs[0][1]
    for r0, r1 in runs[1:]:
        if r0 - end <= DESC_GAP:
            end = r1
        else:
            break  # 遇到 描述块 -> 「装备能力值」 那个 55px 的 gap，描述到此结束

    if not any(end + DESC_TITLE_DY - 7 <= r0 <= end + DESC_TITLE_DY + 7 for r0, _r1 in runs):
        return None
    return end
# 词条数值的读取参数 (放大倍数, 灰度阈值)。
# 快路径沿用旧的那一组；只有「有名字的行却没读出数值」时才把全部参数组都上，
# 凑齐候选重新投票。7 和 1 只差最上面那道横杠的深浅，单一阈值必然在部分行上读错
# （th=160 把 002_甲 的 7.59% 读成 1.59%；th=185 把 002_手 的 6.89% 读成 None）。
# 词条数值的读取参数 (放大倍数, 灰度阈值)。**配色比 OCR 可靠得多，先判色再决定要不要 OCR**
# —— 见 _row_style：暗底必然是第 15 档，直接查表连 OCR 都免了；蓝字/黑字也把候选从 15 档
# 缩到 3/11 档。所以这里只需要一路「白底」掩码，不需要反相档。
AFFIX_VAL_FAST = [(4, 160)]
# 慢路径用**规则网格**，不要手挑。曾经是一组组加上去的（每组都只因为恰好救回测试集里的
# 某一行），无可辩护：实测其中的 (6,200) 单用缺 24 条、几乎是废的，只为救一行而存在。
# 倍率要够高：002_脚 的 蓄力速度增加(4.92%) 在 scale 2~5 下只会读成乱码 '%26h'（各阈值
# 都如此，数量上占绝对多数），只有 scale ≥6 才读得出 '4.92%'。乱码 parse 不出数字、
# 不参与投票，所以扩到 6 之后它就是唯一的强票。倍率范围(2,3,4,6)与阈值(140,160,180)
# 都是系统性取值，不是针对某一行挑的。
AFFIX_VAL_ALL = [(s, t) for s in (2, 3, 4, 6) for t in (140, 160, 180)]

# 配色 -> 可能的档位下标区间（0 基，含两端）。暗底不在表里：那是第 15 档，唯一值，不用 OCR。
STYLE_TIERS = {"normal": (0, 10), "blue": (11, 13)}

# 词条规范名（9 种，游戏内名称，都带「增加」）
AFFIX_CANON = [
    "攻击力增加", "蓄力伤害增加", "防御力增加", "暴击伤害增加", "命中率增加",
    "暴击率增加", "最大装弹数增加", "蓄力速度增加", "优越代码伤害增加",
]

# 空词条槽的占位文字。读到它 = 这件装备有「改造装备效果」那一整块 = 铁证是 T10
# （T9 那块根本不存在，连这句占位都读不到）。
# 存字面量而不是 None：这样「只有 2 条词条」的空槽显示为「未获得效果」+0，
# 输出里的 None 就只剩「读失败/补位」一种含义，不再和空槽混淆。
EMPTY_AFFIX = "未获得效果"

SCALE = 2  # 整屏 OCR 放大倍数（2x 最稳；4x 小框反而把字切碎，实测教训）


# ----------------------------------------------------------------------------
# OCR 基础工具：整屏 2x 一次读全，按区域过滤取字段（比小框放大稳定得多）
# ----------------------------------------------------------------------------
def _ocr_screen(engine, img):
    """整屏 2x OCR，返回 [(text, x0, y0, x1, y1)]（原始 1920x1080 坐标）。"""
    w, h = img.size
    big = img.resize((w * SCALE, h * SCALE), Image.LANCZOS)
    result, _ = engine(np.array(big))
    if not result:
        return []
    out = []
    for box, text, _score in result:
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        out.append((text, min(xs) / SCALE, min(ys) / SCALE, max(xs) / SCALE, max(ys) / SCALE))
    return out


def _in_region(item, region):
    """item = (text, x0, y0, x1, y1)；判断 bbox 中心是否落在 region。"""
    _t, x0, y0, x1, y1 = item
    cx, cy = (x0 + x1) / 2, (y0 + y1) / 2
    rx0, ry0, rx1, ry1 = region
    return rx0 <= cx <= rx1 and ry0 <= cy <= ry1


def _ocr_crop(engine, img, box, scale=SCALE):
    """裁剪 box 放大 OCR，返回 [(text, x0, y0, x1, y1)]（原始 1920x1080 坐标）。

    整屏 2x OCR 的替代：只裁需要的小区域，实测比整屏快约 6x（7.4s → ~1.2s）。
    坐标换算回原图绝对坐标，供 _in_region 沿用原区域过滤逻辑。
    """
    bx0, by0, bx1, by1 = box
    crop = img.crop((bx0, by0, bx1, by1)).convert("RGB")
    w, h = crop.size
    big = crop.resize((w * scale, h * scale), Image.LANCZOS)
    result, _ = engine(np.array(big))
    if not result:
        return []
    out = []
    for b, text, _s in result:
        xs = [p[0] for p in b]
        ys = [p[1] for p in b]
        out.append((text, min(xs) / scale + bx0, min(ys) / scale + by0,
                    max(xs) / scale + bx0, max(ys) / scale + by0))
    return out


def _texts_in(engine, img, region):
    return [t for t, x0, y0, x1, y1 in _ocr_screen(engine, img) if _in_region((t, x0, y0, x1, y1), region)]


# 字段区域（bbox 中心落点范围）
NAME_REGION = (1740, 225, 1880, 280)      # 角色名（右侧面板）
LEVEL_REGION = (1580, 240, 1735, 280)     # 等级 "LV.xxx /200"
CP_REGION = (1700, 330, 1860, 395)        # 战力数字（区域过滤用；read_info_page 走大框+过滤，见下）
# 翻页判定专用战力裁条。⚠️ 不要拿紧贴的 CP_REGION 去裁图：它上切「战斗力」(y323-340)、
# 下切「BATTLE」(y385-400) 各一半，OCR 碰到半截残片会整块输出空 —— 003/004 实测小框裁出
# []，而同一张图走大框 + 区域过滤能读到 179418/177944，翻页判定因此 15s 超时。
# 留足余量把三行完整包住；parse_int 取最大数，多包进来的字无害。
CP_BOX = (1690, 315, 1875, 410)

# 联合裁剪框（外扩 15px 防切字；整屏 2x → 区域裁剪用）
INFO_PANEL_BOX = (1565, 210, 1895, 410)   # 信息页右侧面板：名字+等级+战力
# 装备页只裁词条名那一条带。**不再读装备名**：名字不参与任何判定（T9/T10 看有没有词条、
# 等级靠能力值反推），纯粹是输出信息；而读它的窄框在「整页内容下移」时会漏读
# （小红帽 头/甲 的 y245->y281），白引一类 bug。去掉名字后裁剪框能从 y210 收到 y690。

# ---- 页面状态判定（像素级，不走 OCR）----
# 装备页/装备选择页都是盖在信息页上的浮层，会把白面板 PANEL_BOX 之外的区域整体压暗。
# 判据 = 面板**外**那一圈里「亮像素(>=BRIGHT_TH)」的**占比**：
#   信息页：白菜单栏 / 右侧面板 / 底部导航等一大片亮 UI，实测 40.0~89.1%
#   浮层：外圈被整体压暗，实测 0.1~0.9%
# 分离度 44 倍，中间那一档留给转场动画 -> 'other'，调用方据此不盲目动作。
#
# ⚠️ 别改回「固定小框的白像素占比」。那版赌的是"某块区域永远不会被立绘碰到"，而
# 白雪公主的立绘一直伸到画面最左边缘、把左侧菜单条 y315 以下全盖了（89.6% -> 33.7%），
# 于是已经翻好的信息页被判成 other，翻页判定 15s 超时、全量跑到第 3 个人就中断。
# 现在量的是**占比**而不是"整块必须亮"：立绘盖掉外圈一部分，剩下的亮 UI 照样把比例顶住。
PANEL_BOX = (688, 35, 1228, 988)  # 浮层白面板；判据只统计它外面那一圈
BRIGHT_TH = 200                   # 判「亮」的灰度阈值（比纯白 240 宽容，仍是 UI 底色级别）
INFO_BRIGHT_MIN = 0.25            # 信息页：外圈亮像素 >=25%（实测最低 40.0%）
OVERLAY_BRIGHT_MAX = 0.05         # 浮层：外圈亮像素 <=5%（实测最高 0.9%）


# ----------------------------------------------------------------------------
# 字段解析
# ----------------------------------------------------------------------------
def parse_level(items):
    """同步等级（'LV.436' / 'LY.436' / '436/200' / 拆开的 '436'+'/200'）-> 436；失败 None。

    OCR 常把 'LV' 认成 'LY'，且把 'LV.436/200' 拆成两段；更糟的是标签侧可能只剩残片
    （002_info 实测 'LV.444' -> 'LY.4' + '444/200'，旧的"标签优先"会返回 4）。
    故取标签值与斜杠值中**位数更多**的那个（残片一定更短），同长优先标签；
    两者都没有再退化为最靠左的独立纯数字（同步等级在标签右侧，上限 '/200' 在最右）。
    """
    label = slash = None
    for t, *_ in items:
        if label is None:
            m = re.search(r"L[VY]\.?\s*(\d+)", t, re.IGNORECASE)
            if m:
                label = int(m.group(1))
        if slash is None:
            m = re.match(r"^(\d+)/", t)
            if m:
                slash = int(m.group(1))
    # OCR 会把 "LV.444" 截成 "LY.4"（标签侧残片）+ "444/200"（带斜杠锚定的完整值），
    # 例如 002_info 实测就是这两段。旧逻辑"标签优先"会返回残片 4。
    # 残片一定比真值短 -> 两者都在时信更长的那个（同长则信标签）。
    if slash is not None and (label is None or len(str(slash)) > len(str(label))):
        return slash
    if label is not None:
        return label
    if slash is not None:
        return slash
    pure = sorted((x0, int(t)) for t, x0, *_ in items if re.fullmatch(r"\d+", t.strip()))
    return pure[0][1] if pure else None


def parse_int(texts):
    """取最长的纯数字串（战力 181462 等），滤掉标签/BATTLE 等。"""
    best = None
    for t in texts:
        for m in re.finditer(r"\d[\d,]*", t):
            v = int(m.group(0).replace(",", ""))
            if best is None or v > best:
                best = v
    return best


def parse_name(texts):
    """角色名：只留汉字，取最长的候选（滤掉稀有度/数字等杂音）。"""
    cands = [re.sub(r"[^一-鿿]", "", t) for t in texts]
    cands = [c for c in cands if 1 <= len(c) <= 6]
    return max(cands, key=len) if cands else None


def norm_affix(texts):
    """OCR 文本 -> 9 规范词条名；不匹配返回 None。"""
    han = "".join(re.sub(r"[^一-鿿]", "", t) for t in texts)
    if not han:
        return None
    for canon in AFFIX_CANON:
        if han == canon or han in canon or canon in han:
            return canon
    return None


def parse_pct(texts):
    """词条数值文本 -> 百分比小数。'11.81%' -> 0.1181；'4.69' -> 0.0469。

    必须对单次 OCR 结果 parse（不能多阈值文本 join，否则拼出假数字）。
    优先带 % 的完整小数；纯小数要求至少两位（排除 "16." 半截）。
    """
    for t in texts:
        m = re.search(r"(\d+\.\d+)%", t)
        if m:
            return round(float(m.group(1)) / 100.0, 4)
    for t in texts:
        m = re.search(r"(\d+\.\d{2})", t)
        if m:
            return round(float(m.group(1)) / 100.0, 4)
    for t in texts:
        m = re.search(r"(\d+)%", t)
        if m:
            return round(float(m.group(1)) / 100.0, 4)
    return None


# ----------------------------------------------------------------------------
# 页面读取
# ----------------------------------------------------------------------------
def read_info_page(engine, img):
    """角色信息页 -> (名字, 等级, 战力)。右侧面板联合框单次裁剪 2x，按区域过滤。"""
    items = _ocr_crop(engine, img, INFO_PANEL_BOX)
    name = parse_name([t for t, x0, y0, x1, y1 in items if _in_region((t, x0, y0, x1, y1), NAME_REGION)])
    level = parse_level([(t, x0, y0, x1, y1) for t, x0, y0, x1, y1 in items if _in_region((t, x0, y0, x1, y1), LEVEL_REGION)])
    cp = parse_int([t for t, x0, y0, x1, y1 in items if _in_region((t, x0, y0, x1, y1), CP_REGION)])
    return name, level, cp


def _read_affix_names(engine, img, dy=0):
    """词条名 -> [(y中心, 规范名或 EMPTY_AFFIX)]，最多 3 行。

    只保留「能归一到 9 个规范名」或「未获得效果」的行，再取前 3 行 ——
    面板变矮时「效果数值将在进入战斗时生效。」这句提示会跟着掉进词条区占掉一行，
    不能把它当成一条词条（它归不到任何规范名，会被这里滤掉）。
    ⚠️ 别把「读不认识的行」也留成占位来标记读失败：那句提示（以及「效果变更」
    等按钮）同样归不到规范名，一旦留下就会顶掉后面一条真词条 —— 比丢一行更糟。
    区域位置由 dy（面板顶偏移）决定，见 panel_top。
    """
    items = _ocr_crop(engine, img, _shift(AFFIX_NAME_BAND, dy))
    rows = []
    for t, x0, y0, x1, y1 in items:
        name = norm_affix([t])
        if name is None:
            if "未获得效果" in t:
                name = EMPTY_AFFIX  # 空槽占位，位置要留住（见 AFFIX_CANON 下方说明）
            else:
                continue  # 提示文字「效果数值将在…」/「效果变更」「重新设定数值」等按钮 / 数值列残片
        rows.append(((y0 + y1) / 2, name))
    rows.sort(key=lambda z: z[0])
    return rows[:3]


def _read_affix_val_texts(engine, img, params, dy=0):
    """数值列按 params 逐组二值化各读一遍 -> [(y中心, 文本)]，不预先分箱。"""
    x0, y0, x1, y1 = _shift(AFFIX_VAL_BAND, dy)
    crop = img.crop((x0, y0, x1, y1))
    g = np.array(crop.convert("L")).astype(int)
    rgb = np.array(crop.convert("RGB")).astype(int)
    out = []

    def ocr_mask(mask, scale):
        h, w = mask.shape
        out_img = np.where(mask, 0, 255).astype("uint8")
        bimg = Image.fromarray(out_img).resize((w * scale, h * scale), Image.LANCZOS)
        result, _ = engine(np.array(bimg.convert("RGB")))
        for box, t, _s in result or []:
            out.append(((min(p[1] for p in box) + max(p[1] for p in box)) / 2 / scale + y0, t))

    for scale, th in params:
        ocr_mask(g < th, scale)
    ocr_mask((rgb[:, :, 2] - rgb[:, :, 0]) > 30, 2)  # 蓝字档兜底
    return out


def _row_style(img, y_row, dark_th=160, blue_th=0.05):
    """该词条行的配色档位 -> 'dark' | 'blue' | 'normal'。

    词条数值配色按档位分三种：1~11 白底黑字、12~14 白底蓝字、15 黑底蓝字。
    **底色直接确定了档位，比 OCR 可靠得多**，所以要先用它：
      - 'dark'   = 必然是第 15 档 —— 直接取档位表最后一项，连 OCR 都不用
      - 'blue'   = 只可能是 12~14 档 —— OCR 候选从 15 个缩到 3 个
      - 'normal' = 只可能是 1~11 档
    实测两种白底行的蓝像素占比两极分化（黑字行 0~2%、蓝字行 10~13%，无重叠），
    暗底行暗占比 83%（其余 ≤14%），两个阈值取中间都很安全。
    """
    a = np.array(img.crop((965, int(y_row) - 18, 1120, int(y_row) + 18)).convert("RGB")).astype(int)
    if (a.mean(axis=2) < dark_th).mean() > 0.5:
        return "dark"
    return "blue" if ((a[:, :, 2] - a[:, :, 0]) > 30).mean() > blue_th else "normal"


def _vote_row(cands, y_row, name, table, lo=0, hi=None, tol=18):
    """把候选里落在该词条行 y 附近的挑出来 -> 取值 -> 词条表校验 -> 按读取质量投票。

    **强票**是完整带 % 的读法（'14.63%'）；其余（掉小数点、被表硬凑回档位的）算**弱票**。
    有强票时只看强票，弱票只在完全没有强票时才轮到。
    ⚠️ 不分强弱会出错：实测 003_甲 蓄力伤害增加 有干净的 '14.63%'，但同时也冒出好几个
    乱码被表凑合成 0.0477 的候选，票数更多就把正确值投掉了 ——「候选更多」反而更糟。
    """
    strong, weak = {}, {}
    for yc, t in cands:
        if abs(yc - y_row) > tol:
            continue
        v = parse_pct([t])
        if v is None:
            continue
        aligned = align_affix_value(name, v, table, lo, hi)
        if aligned is None:
            continue
        bucket = strong if re.search(r"\d+\.\d+\s*%", t) else weak
        bucket[aligned] = bucket.get(aligned, 0) + 1
    if strong:
        return max(strong, key=strong.get)
    return max(weak, key=weak.get) if weak else None


def _read_affix_values(engine, img, name_rows, table, dy=0):
    """三行词条数值：按 y 把候选归到对应的词条行，逐行投票。

    按 y 对齐（而不是把三行切进固定分箱）顺带解掉「词条名/数值各走独立路径、各自漏读
    导致 zip 错位」的隐患。

    先走单组参数快路径；有行「有名字却没读出数值」再上全部参数重读。
    词条表在这里当裁判：落不到合法档位的读法直接作废（实测 002_甲 把 命中率增加 的
    7.59% 读成 1.59%，而该词条最低档是 0.0477，正好被表挡下）。
    ⚠️ 表只能挡「落不到档位」的错读，挡不住「落在别的档位上」的错读（0.5893 会被兜成
    0.6071），所以投票仍然必要。
    """

    styles = [_row_style(img, y) for y, _n in name_rows]  # 配色直接定档位，见 _row_style

    def run(params):
        cands = _read_affix_val_texts(engine, img, params, dy)
        out = []
        for i, (y, n) in enumerate(name_rows):
            if n == EMPTY_AFFIX:
                out.append(0)  # 空槽没有词条数值，直接 0：不是 OCR 失败，不需要读也不需要投票
                continue
            if styles[i] == "dark":
                # 暗底必然是第 15 档：直接取该词条档位表的最后一项，不 OCR
                levels = table.get(n) if n else None
                out.append(levels[-1] if levels else None)
                continue
            lo, hi = STYLE_TIERS[styles[i]]
            out.append(_vote_row(cands, y, n, table, lo, hi))
        return out

    out = run(AFFIX_VAL_FAST)
    # 空槽（EMPTY_AFFIX -> 0）和补位行（name 为 None）都不算漏读，只有真词条缺值才上慢路径
    if all(v is not None or name_rows[i][1] is None for i, v in enumerate(out)):
        return out
    return run(AFFIX_VAL_ALL)


# ----------------------------------------------------------------------------
# 装备等级反推：读装备能力值 -> (类型, 槽位, T10) 查表映射 0~5
# ----------------------------------------------------------------------------
# 装备页顶部类型标签（火力型/辅助型/防御型）。位置按面板顶标定：实测在 面板顶+37。
# （曾一度把 y 放宽到 135 去兜位移，那是治标；现在由 _shift 按面板顶平移解决。）
TYPE_BOX = (735, 45, 825, 92)
STAT_BOX = (860, 520, 1045, 600)       # 装备能力值面板（左标签列 右数值列）
EQUIP_XLSX = "equipment.xlsx"
ONLINE_OUT = "output/cn_collect.json"         # 在线采集结果
OFFLINE_OUT = "output/cn_collect_offline.json"  # 离线自检结果（独立文件，不覆盖真实结果）

_TYPES = ["火力型", "辅助型", "防御型"]
_SLOT_ALIAS = {"头": "头", "手": "手", "甲": "甲", "衣": "甲", "脚": "脚", "鞋": "脚"}

_STAT_TABLE = None


def load_stat_table(path=EQUIP_XLSX):
    """读 equipment.xlsx「装备加成结果表」-> {(T10, 类型, 槽位): {atk/hp/def: [0~5档]}}。全程只读一次。"""
    global _STAT_TABLE
    if _STAT_TABLE is not None:
        return _STAT_TABLE
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["装备加成结果表"]
    table = {}
    for r in range(3, ws.max_row + 1):
        tier = ws.cell(r, 1).value
        typ = ws.cell(r, 2).value
        slot = ws.cell(r, 3).value
        if tier != "T10" or not typ or not slot:
            continue
        slot = _SLOT_ALIAS.get(slot, slot)
        get = lambda i: 0 if ws.cell(r, i).value is None else int(ws.cell(r, i).value)
        table[(tier, typ, slot)] = {
            "atk": [get(c) for c in range(4, 10)],
            "hp": [get(c) for c in range(10, 16)],
            "def": [get(c) for c in range(16, 22)],
        }
    _STAT_TABLE = table
    return table


_AFFIX_TABLE = None


def load_affix_table(path=EQUIP_XLSX):
    """读 equipment.xlsx「词条表」-> {词条名: [15档数值]}。全程只读一次，名字归一化到 9 规范名。"""
    global _AFFIX_TABLE
    if _AFFIX_TABLE is not None:
        return _AFFIX_TABLE
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["词条表"]
    table = {}
    for r in range(2, ws.max_row + 1):
        name = ws.cell(r, 1).value
        if not name:
            continue
        name = name.replace("量", "数")  # 「最大装弹量增加」→ 规范名「最大装弹数增加」
        vals = [ws.cell(r, c).value for c in range(3, 18)]  # C..Q = 1~15 档
        table[name] = [float(v) for v in vals if v is not None]
    _AFFIX_TABLE = table
    return table


def align_affix_value(name, value, table, lo=0, hi=None):
    """词条数值 -> 该词条最近档位的标准值，只在 [lo, hi] 档位区间（0 基）里找。

    距离 ≤ 半档距才采纳，否则 None（错读/无对应名不硬猜）。
    档位区间由配色决定（见 _row_style），把「落在别的档位上」的错读也挡掉一大半。
    """
    if name is None or value is None:
        return None
    levels = table.get(name)
    if not levels:
        return None
    hi = len(levels) - 1 if hi is None else min(hi, len(levels) - 1)
    pool = levels[max(0, lo):hi + 1]
    if not pool:
        return None
    if value in pool:
        return value
    diffs = [abs(lv - value) for lv in pool]
    i = diffs.index(min(diffs))
    gap = (levels[-1] - levels[0]) / max(1, len(levels) - 1)
    return round(pool[i], 4) if diffs[i] <= gap * 0.5 + 1e-9 else None


def read_equip_type(engine, img, dy=0):
    """装备页顶部类型标签 -> 火力型/辅助型/防御型；匹配不上返回 None。"""
    texts = _crop_ocr(engine, img, _shift(TYPE_BOX, dy), scale=2)
    for t in texts:
        for canon in _TYPES:
            if canon in t:
                return canon
    joined = "".join(texts)
    if "火" in joined:
        return "火力型"
    if "辅" in joined or "助" in joined:
        return "辅助型"
    if "防" in joined:
        return "防御型"
    return None


def read_ability_stats(engine, img, dy=0):
    """装备能力值面板 -> {atk/hp/def: int|None}。左标签列(x<970)右数值列(x>=970)，按 y 同行配对。"""
    items = _ocr_crop(engine, img, _shift(STAT_BOX, dy), scale=2)
    labels, values = [], []
    for t, x0, y0, x1, y1 in items:
        cx = (x0 + x1) / 2
        cy = (y0 + y1) / 2
        if cx < 970:
            labels.append((t, cy))
        else:
            values.append((t, cy))

    stats = {"atk": None, "hp": None, "def": None}
    for vt, vy in values:
        m = re.search(r"\d[\d,]*", vt)
        if not m:
            continue
        val = int(m.group(0).replace(",", ""))
        lt = next((lt for lt, ly in labels if abs(ly - vy) < 15), None)
        if lt is None:
            continue
        han = re.sub(r"[^一-鿿]", "", lt)
        if "体力" in han:
            stats["hp"] = val
        elif "攻击力" in han:
            stats["atk"] = val
        elif "防御力" in han:
            stats["def"] = val
    return stats


def infer_level(stats, typ, slot, table):
    """能力值反推强化等级 0~5：所有非零能力列须指向同一级才采纳，否则 None（宁缺勿错）。"""
    row = table.get(("T10", typ, slot))
    if not row:
        return None
    votes = []
    for key in ("atk", "hp", "def"):
        v = stats.get(key)
        levels = row[key]
        if v is None or not any(levels):
            continue
        if v in levels:
            votes.append(levels.index(v))
        else:
            diffs = [abs(lv - v) for lv in levels]
            i = diffs.index(min(diffs))
            gap = max(1, (max(levels) - min(levels)) // 10)
            if diffs[i] <= gap:
                votes.append(i)
    return votes[0] if votes and all(x == votes[0] for x in votes) else None


def read_equip_page(engine, img, slot=None, stat_table=None):
    """装备页 -> (等级, [词条名x3], [数值x3])。

    两个锚点，别混用（见文件头「装备页对齐」那段）：
      dy_top —— 面板顶：只管描述**之上**的类型标签
      dy     —— 描述块底边：管描述**之下**的能力值和词条（描述行数会变，面板顶锚不住）
    """
    top = panel_top(img)
    dy_top = top - PANEL_TOP_REF
    bottom = desc_bottom(img, top)
    dy = dy_top if bottom is None else bottom - DESC_BOTTOM_REF

    name_rows = _read_affix_names(engine, img, dy)
    name_rows += [(-1e9, None)] * (3 - len(name_rows))  # 不足 3 行补空，保证长度对齐
    affix_names = [n for _y, n in name_rows]

    # T10 判定：三行至少一行读到词条名或空槽占位才算 T10；全空 = 无装备/T9（不读等级不读数值）。
    # 空槽占位「未获得效果」算 T10 是**正确**的：T9 没有「改造装备效果」那一整块，
    # 连占位文字都读不到。反过来说，「T10 但三行全空」以前会被判成 T9 而漏读等级，现在不会。
    if not any(affix_names):
        return None, [None, None, None], [None, None, None]

    # T10：读类型 + 能力值 -> 查表反推等级
    level = None
    if slot and stat_table is not None:
        typ = read_equip_type(engine, img, dy_top)  # 类型标签在描述之上，跟面板顶
        if typ:
            stats = read_ability_stats(engine, img, dy)
            level = infer_level(stats, typ, slot, stat_table)

    # 词条数值：按 y 对齐到 name_rows，多参数读 + 词条表校验 + 投票（仅 T10）
    affix_vals = _read_affix_values(engine, img, name_rows, load_affix_table(), dy)

    return level, affix_names, affix_vals


# ----------------------------------------------------------------------------
# 页面就绪判定 + 轮询（点击后不等固定时长，改等目标页特征出现）
# ----------------------------------------------------------------------------
def _crop_ocr(engine, img, region, scale=2):
    """裁剪 region 放大 OCR，返回 [text]。判定页面用，无需整屏（快得多）。"""
    x0, y0, x1, y1 = region
    crop = img.crop((x0, y0, x1, y1)).convert("RGB")
    w, h = crop.size
    big = crop.resize((w * scale, h * scale), Image.LANCZOS)
    result, _ = engine(np.array(big))
    return [t for _, t, _s in result] if result else []


_OUTSIDE_MASK = None


def _outside_mask(shape):
    """面板之外那圈的布尔掩码。截图尺寸不变时复用，避免每次轮询都重建 200 万像素的掩码。"""
    global _OUTSIDE_MASK
    if _OUTSIDE_MASK is None or _OUTSIDE_MASK.shape != shape:
        m = np.ones(shape, dtype=bool)
        x0, y0, x1, y1 = PANEL_BOX
        m[y0:y1, x0:x1] = False
        _OUTSIDE_MASK = m
    return _OUTSIDE_MASK


def _outside_bright_ratio(img):
    """面板之外那圈里「亮像素」的占比。浮层把外圈整体压暗，这个比例就塌下去。"""
    g = np.array(img.convert("L"), dtype=np.uint8)
    return float((g[_outside_mask(g.shape)] >= BRIGHT_TH).mean())


def page_state(img):
    """当前页面状态：'info'(角色信息页) | 'overlay'(信息页上盖了浮层) | 'other'。

    旧实现靠 is_equip_page「装备名区(110x30 小框)有没有汉字」判定，又脆又单向；
    而 is_info_page 是建立在它之上的（装备页也显示角色名，只能靠 is_equip_page 排除），
    于是 is_equip_page 一次假阴性会同时让「进装备页」超时、又让「回信息页」秒过 ——
    两个判定同向失效，状态机就会错位，在信息页/浮层之间来回翻转。

    这一版改判「浮层有没有打开」这个物理事实：浮层会把白面板之外的区域整体压暗。
    不依赖任何文字 OCR，也不依赖"某块区域不被立绘碰到"（见常量段的教训）。
    比例落在中间（转场动画中）-> 'other'，调用方据此不盲目动作。
    """
    r = _outside_bright_ratio(img)
    if r >= INFO_BRIGHT_MIN:
        return "info"
    if r <= OVERLAY_BRIGHT_MAX:
        return "overlay"
    return "other"


def wait_page(screen, want, timeout=8.0, interval=0.1):
    """轮询截图直到 page_state == want。返回 (是否等到, 最后看到的状态)。

    page_state 纯像素、~1ms，所以轮询可以比 OCR 判定密得多（截图 0.043s 才是大头）。
    """
    t0 = time.time()
    last = None
    while time.time() - t0 < timeout:
        last = page_state(screen.grab())
        if last == want:
            return True, last
        time.sleep(interval)
    return False, last


def _ensure_info(screen, xy, tries=3, timeout=5.0):
    """确保停在信息页。返回是否成功。

    只在**确认是浮层**时才点 xy 把它关掉（浮层背后就是信息页，点同一槽位即关）。
    'other'（转场动画、或跑到了完全陌生的画面）下不盲点，只等画面稳定再重判 ——
    不然会在不知道是什么的界面上乱点一气。
    """
    for _ in range(tries):
        st = page_state(screen.grab())
        if st == "info":
            return True
        if st == "overlay":
            screen.click(xy)
            if wait_page(screen, "info", timeout=timeout)[0]:
                return True
        else:
            time.sleep(0.5)
    return page_state(screen.grab()) == "info"


def wait_page_turn(screen, engine, prev_cp, prev_name, timeout=15.0, interval=0.3):
    """翻页完成判定：回到信息页，且（战力变化）或（名字稳定且 != 上一角色）。

    主判据是战力：长名有平行滚动，同一角色截两次可能读到不同片段，
    旧的 _name_changed 因此会假阳性（误判成已翻页）。
    名字只作兜底，且要求连续两次读到同一个名字才算稳定（滚动中读不到重复值）。
    """
    t0 = time.time()
    last_name = None
    trace = []  # 超时时回放最后几次轮询，看清是「页面没翻」还是「翻了但没读到」
    while time.time() - t0 < timeout:
        img = screen.grab()
        st = page_state(img)
        cp = None
        if st == "info":
            cp = parse_int(_crop_ocr(engine, img, CP_BOX))
            if cp is not None and cp != prev_cp:
                return True
            nm = parse_name(_crop_ocr(engine, img, NAME_REGION))
            if nm is not None and nm != prev_name and nm == last_name:
                return True
            last_name = nm
        trace.append((round(time.time() - t0, 1), st, cp))
        time.sleep(interval)
    print(f"    [翻页诊断] 上一角色 战力={prev_cp} 名字={prev_name}；最后 {min(6, len(trace))} 次轮询：")
    for ts, st, cp in trace[-6:]:
        print(f"      t={ts:>5}s  页面状态={st:<8} 读到的战力={cp}")
    return False


def _empty_slot():
    return {"等级": None, "词条": [{"名称": None, "数值": None} for _ in range(3)]}


# ----------------------------------------------------------------------------
# 采集主流程
# ----------------------------------------------------------------------------
def build_row(name, cp, slots):
    """组装输出行，对齐 Nikke.xlsx（A=姓名, B=战力, 头/甲/手/脚 各等级+3词条）。"""
    row = {"姓名": name, "战力": cp}
    for slot in ["头", "甲", "手", "脚"]:
        info = slots.get(slot, {})
        row[slot] = {
            "等级": info.get("等级"),
            "词条": info.get("词条", [None, None, None]),
        }
    return row


class Screen:
    """截图抽象：在线 pyautogui / 离线读文件。"""

    def __init__(self, offline=False):
        self.offline = offline
        self._pyautogui = None
        if not offline:
            import pyautogui  # 延迟导入，离线验证不依赖

            self._pyautogui = pyautogui

    def grab(self):
        if self.offline:
            return Image.open("ps.png").convert("RGB")
        return self._pyautogui.screenshot()

    def click(self, xy):
        if self.offline:
            print(f"    [离线] 模拟点击 {xy}")
            return
        self._pyautogui.click(*xy)


def check_state(paths):
    """离线校验页面状态分类器：打印每张图判成的状态 + 外圈亮像素占比（判据本身）。"""
    print("=== 页面状态分类器校验 ===")
    print(f"{'文件':<26}{'判定':<10}{'外圈亮%':>11}   (info>=25, overlay<=5)")
    print("-" * 62)
    for p in paths:
        p = Path(p)
        if not p.exists():
            print(f"{p.name:<26}{'不存在':<10}")
            continue
        img = Image.open(p).convert("RGB")
        print(f"{p.name:<26}{page_state(img):<10}{_outside_bright_ratio(img) * 100:>11.1f}")


def require_offline_shots(ps_path, eq_path):
    """离线模式的两张截图必须存在，缺失时报清楚怎么补。

    main() 在加载 OCR 引擎之前就会调一次 —— 路径打错不该先白等一次模型加载。
    """
    for p in (ps_path, eq_path):
        if not Path(p).exists():
            raise SystemExit(
                f"找不到 {p}。--offline 需要一张信息页和一张装备浮层页截图；"
                f"用 --ps/--eq 指定，或用 --save-shots 在线跑一次生成 output/shots/。"
            )


def collect_offline(engine, ps_path="ps.png", eq_path="equipment.png"):
    """离线验证：用两张截图跑一遍单角色采集。

    默认读根目录的本地标定样本（信息页 + 装备浮层页各一张），换机器/换账号时用
    --ps/--eq 指向自己的截图即可 —— 这里只要求能读出文字，不绑定特定账号。
    """
    require_offline_shots(ps_path, eq_path)  # main() 已提前查过；直接调用本函数时兜底
    print("=== 离线模式：单角色采集验证 ===")
    img_ps = Image.open(ps_path).convert("RGB")
    img_eq = Image.open(eq_path).convert("RGB")
    print(f"  {ps_path} 状态={page_state(img_ps)}（应为 info）  "
          f"{eq_path} 状态={page_state(img_eq)}（应为 overlay）")

    name, level, cp = read_info_page(engine, img_ps)
    print(f"信息页：名字={name} 等级={level} 战力={cp}")

    stat_table = load_stat_table()
    slots = {}
    for slot, _ in SLOTS.items():
        elevel, anames, avals = read_equip_page(engine, img_eq, slot=slot, stat_table=stat_table)
        affixes = [
            {"名称": n, "数值": v} for n, v in zip(anames, avals)
        ]
        slots[slot] = {"等级": elevel, "词条": affixes}
        print(f"  槽位[{slot}] 等级={elevel} 词条={affixes}")

    row = build_row(name, cp, slots)
    out = {"角色": [row]}
    # 写独立文件：离线是自检，不能覆盖在线采集的真实结果
    Path("output").mkdir(exist_ok=True)
    with open(OFFLINE_OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\n已写入 {OFFLINE_OUT}")
    return out


def _archive_previous(path):
    """开新采集前，把上一次的结果改名带时间戳留底。

    否则新一轮会直接覆盖 output/cn_collect.json —— 这个坑真实发生过：37 人的结果
    被一次 6 人试跑覆盖掉，而落盘用的是 os.replace（原地替换、不进回收站），找不回来。
    空壳/损坏的文件不值得留底，返回 None 让新结果直接覆盖它。
    """
    p = Path(path)
    if not p.exists():
        return None
    try:
        n = len(json.loads(p.read_text(encoding="utf-8")).get("角色", []))
    except Exception:
        n = 0  # 损坏或缺字段 -> 当成空壳，不占坑
    if n == 0:
        return None
    dst = p.with_name(f"{p.stem}.{time.strftime('%Y%m%d-%H%M%S')}{p.suffix}")
    p.replace(dst)
    return dst


def _save_rows(rows, out=ONLINE_OUT):
    """把已采的角色增量写盘：原子替换（先写临时文件再 rename），中途中断不丢、不坏。

    out 可覆盖，供 correct_gui.py 复用同一套落盘格式（默认值保证采集路径行为不变）。
    """
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    tmp = Path(str(out) + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump({"角色": rows}, f, ensure_ascii=False, indent=2)
    tmp.replace(out)


def collect_online(engine, screen, max_chars=500, save_shots=False):
    """在线循环采集：停在信息页，翻页直到 LV.1。save_shots=True 时每页截图存 output/shots/。"""
    rows = []
    stat_table = load_stat_table()
    shots_dir = Path("output/shots")
    if save_shots:
        shots_dir.mkdir(parents=True, exist_ok=True)
    print("=== 在线模式 ===")
    prev = _archive_previous(ONLINE_OUT)
    if prev:
        print(f"（上一轮结果已留底：{prev}）")
    print("5 秒后开始采集，请立即 Alt+Tab 切回游戏（角色信息页），期间别碰鼠标键盘 ...")
    for i in range(5, 0, -1):
        print(f"  {i} ...")
        time.sleep(1)
    prev_name, prev_cp = None, None
    for i in range(max_chars):
        # 角色起点：翻页刚结束可能在转场，多给几次机会（'other' 只等不点）
        if not _ensure_info(screen, SLOTS["头"], tries=4, timeout=5.0):
            print(f"⚠️ 起始画面不是信息页（判定={page_state(screen.grab())}），中断。")
            break
        time.sleep(0.2)  # 等画面稳定
        img = screen.grab()
        if save_shots:
            img.save(shots_dir / f"{i + 1:03d}_info.png")
        name, level, cp = read_info_page(engine, img)
        print(f"[{i + 1}] {name} LV.{level} 战力 {cp}")

        if level is None:
            print("⚠️ 读不到等级，画面可能不在信息页，中断。")
            break

        slots = {}
        for slot, xy in SLOTS.items():
            # 每次点击前先确认还在信息页：错位了当场纠正，而不是等某个布尔变真
            if not _ensure_info(screen, xy, tries=3, timeout=5.0):
                print(f"    {slot}: 无法回到信息页（判定={page_state(screen.grab())}），中断。")
                slots[slot] = _empty_slot()
                break

            # 点槽开浮层；偶发点击失效则重试一次
            screen.click(xy)
            ok, st = wait_page(screen, "overlay", timeout=8.0)
            if not ok:
                print(f"    {slot}: 第 1 次点击未生效（当前 {st}），重试 ...")
                screen.click(xy)
                ok, st = wait_page(screen, "overlay", timeout=8.0)
            if not ok:
                # 两次都失败：先截图存证，再记空槽并复位
                if save_shots:
                    screen.grab().save(shots_dir / f"{i + 1:03d}_{slot}_FAIL.png")
                print(f"    {slot}: 两次点击都未打开浮层（当前 {st}），跳过。")
                slots[slot] = _empty_slot()
                _ensure_info(screen, xy, tries=2, timeout=5.0)
                continue

            time.sleep(0.2)  # 等图标/等级/词条刷完
            eq_img = screen.grab()
            if save_shots:
                eq_img.save(shots_dir / f"{i + 1:03d}_{slot}.png")
            elevel, anames, avals = read_equip_page(engine, eq_img, slot=slot, stat_table=stat_table)
            if not any(anames):
                # 浮层开着但读不到任何词条。T9 本来就没有词条（属预期），也可能是该槽无装备/
                # 装备选择页 —— 三者输出都是空，这里不替它们下定论，记空槽并显式告警。
                print(f"    {slot}: ⚠️ 浮层已开但读不到词条（T9 属预期；也可能是无装备/装备选择页）")
                slots[slot] = _empty_slot()
            else:
                affixes = [{"名称": n, "数值": v} for n, v in zip(anames, avals)]
                slots[slot] = {"等级": elevel, "词条": affixes}
                print(f"    {slot}: Lv{elevel} {affixes}")

            # 关闭浮层：点同一槽位
            screen.click(xy)
            if not wait_page(screen, "info", timeout=8.0)[0]:
                print(f"    {slot}: 关闭后未回到信息页（判定={page_state(screen.grab())}），复位 ...")
                _ensure_info(screen, xy, tries=2, timeout=5.0)
            time.sleep(0.2)  # 返回后缓冲，避免紧接着点下一槽太快

        rows.append(build_row(name, cp, slots))
        _save_rows(rows)  # 每采完 1 角色就落盘，中断不丢

        if level == 1:
            print(f"到达 LV.1（{name}），采集完成，停止。")
            break

        # 翻页到下一个角色：等战力变化（名字有平行滚动，截两次可能不同，不可靠）
        prev_name, prev_cp = name, cp
        screen.click(NEXT_PAGE)
        if not wait_page_turn(screen, engine, prev_cp, prev_name, timeout=15.0):
            if save_shots:
                screen.grab().save(shots_dir / f"{i + 1:03d}_TURN_FAIL.png")
            print("⚠️ 翻页后未见战力变化，中断（可能已到末位或点击失效）。已存 NNN_TURN_FAIL.png。")
            break

    _save_rows(rows)
    print(f"\n采集 {len(rows)} 个角色，已写入 {ONLINE_OUT}")
    return {"角色": rows}


def main():
    ap = argparse.ArgumentParser(description="国服 NIKKE 角色装备采集")
    ap.add_argument("--offline", action="store_true", help="离线模式（用截图文件模拟，不依赖 pyautogui）")
    ap.add_argument("--ps", default="ps.png", metavar="PNG",
                    help="离线模式用的信息页截图（默认 ps.png）")
    ap.add_argument("--eq", default="equipment.png", metavar="PNG",
                    help="离线模式用的装备浮层页截图（默认 equipment.png）")
    ap.add_argument("--max", type=int, default=500, help="在线模式最大采集角色数（兜底）")
    ap.add_argument("--save-shots", action="store_true", help="在线采集时每页截图存 output/shots/（调试用）")
    ap.add_argument("--check-state", action="store_true",
                    help="只校验页面状态分类器（不采集、不加载 OCR）：根目录样本 + output/shots/*.png")
    args = ap.parse_args()

    if args.check_state:
        # 根目录这 6 张是本地标定样本（不入库），换机器时缺失会被跳过；
        # output/shots/ 才是可复现的来源，--save-shots 在线跑一次即可生成。
        paths = ["ps.png", "PS-NE.png", "equipment.png", "NE.png",
                 "equipment-blue.png", "equipment-black.png"]
        paths += sorted(Path("output/shots").glob("*.png"))
        check_state(paths)
        return

    if args.offline:
        # 先把截图查了再加载引擎：路径打错时不该先白等一次 OCR 模型加载
        require_offline_shots(args.ps, args.eq)

    print("加载 RapidOCR 引擎 ...")
    engine = RapidOCR()

    if args.offline:
        collect_offline(engine, args.ps, args.eq)
    else:
        screen = Screen(offline=False)
        collect_online(engine, screen, max_chars=args.max, save_shots=args.save_shots)

    del engine


if __name__ == "__main__":
    main()
