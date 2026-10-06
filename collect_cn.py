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

**采集只产出补丁。** 输出 output/cn_patch.json —— 每角色一行：姓名、战力、头/身/手/足
各{等级, 词条[3]}，外加顶层 `_图标` 兄弟键（每行的图标四维读数）。**主数据
output/cn_collect.json 一个字节都不动**，错名/误判因此永远污染不了它；主数据只有
correct_gui 的「合并」才写。日常不必从命令行走：界面里点「采集…」，跑完自动接上合并。
采集侧没有跳转能力（只有点「>>」逐格前进一条路），所以"翻到哪儿"由你在游戏里决定 ——
扫几个角色只是范围不同（「增量」填个数 / 「全量」扫到最后一个），后面的
「对比旧表 → 有更新就更新 / 没更新照旧 / 新角色追加」是同一件事。

  python collect_cn.py --max 3        从当前画面起，重扫 3 个（增量）
  python collect_cn.py --max 500      从当前画面起，扫到最后一个角色（全量；先翻到名单最上面）

**逃生口 `--replace`**：只有换号 / 砍掉重练这类"主数据整个不要了"的场合才用。它会
先把 output/cn_collect.json 改名加时间戳留底（cn_collect.<YYYYMMDD-HHMMSS>.json），再用
本次结果**整个替换**它 —— 正是上面那条不变量要躲开的破坏性路径，所以必须手打，界面不可达。

  python collect_cn.py --replace --max 500

姓名怎么定的：读到的名字是**滚动横幅**的截帧，实测约四成会截断读错，而且最险的一种错是
「截断后恰好等于另一个真实角色」——那种错在 correct_gui 里是**零告警**的。所以每行还会读
信息页下方**四个静态图标**（属性/武器/职业/企业，逐像素稳定），拿四维去图鉴里筛人：
图标唯一命中、或与姓名候选交成唯一，才敢用图标定下来的名字；其余一律保持 OCR 原名并告警。

`--as` 是废弃路径的遗留：它让人手打名字断言身份，打错还会中止整轮扫描，而图标本来就是
与姓名无关的机器证据（实测 31/31 全对）。仍保留为命令行逃生口（3 组四维相同的变体对、
以及图鉴外的国服特供才需要它），但已不在正常流程里。

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

# 复用 correct_gui 的「姓名 -> 图鉴候选」逻辑，不重写一份。它是纯函数、不依赖 tkinter
# （correct_gui 把 tkinter 的 import 包在 try 里），而且那边只在 --report 里**延迟** import
# 本模块，所以不会成环。反过来这里 import 它，只是拿模型层的两个工具函数。
from correct_gui import load_roster, name_candidates

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
CP_REGION = (1700, 330, 1860, 395)        # 战力数字（区域过滤用；战力只走 read_cp 一条路，见下）
# ⚠️ 曾经为了「翻页判定」另开过一个小裁剪框 CP_BOX(1690,315,1875,410)，**已经删掉**：
# 同一张信息页，大框+区域过滤读 146731，那个小框读 14673（丢末位）——21 张实测差 1 张。
# 裁剪框一动，OCR 的行切分就变，同一个字段能读出两个数；两处读数不同源就会假阳性，
# 详见 read_cp 的 docstring。

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
def _cp_from_items(items):
    """面板 OCR 结果 -> 战力（只取 bbox 中心落在 CP_REGION 里的那几条）。

    read_info_page 与 read_cp 共用，保证「存进数据的战力」和「判翻页/复核用的战力」
    永远是同一个数 —— 分两条路读会出现两个值，见 read_cp。
    """
    return parse_int([t for t, x0, y0, x1, y1 in items if _in_region((t, x0, y0, x1, y1), CP_REGION)])


def read_cp(engine, img):
    """信息页战力。**全流程只走这一条路**（read_info_page / wait_page_turn / _confirm_same_char）。

    ⚠️ 别为了省一点裁剪面积另开小框：**裁剪框一变，OCR 的行切分就变，同一张图能读出不同的数**
    —— 实测 21 张真机信息页里，`021`（桃乐丝 146731）走大框读对，走旧的小框 CP_BOX 读成
    14673（丢末位）。两处读数一旦不同源，「翻页完成」和「换人复核」就都会假阳性：

      * `wait_page_turn` 拿小框读数去和 read_info_page 的 prev_cp 比，**点完「>>」的第一帧
        就成立**（页面根本没动），于是这一行的身份读的是旧角色、装备点的是新角色 ——
        莱伊·甲 那个错位就是这么来的（那次两行战力一字不差，正是"没翻页却宣布翻页成功"）。
      * `_confirm_same_char` 同理，会把好行判成"屏幕换人了"而中止整轮（桃乐丝那一行实测复现）。
    """
    return _cp_from_items(_ocr_crop(engine, img, INFO_PANEL_BOX))


def read_info_page(engine, img):
    """角色信息页 -> (名字, 等级, 战力)。右侧面板联合框单次裁剪 2x，按区域过滤。"""
    items = _ocr_crop(engine, img, INFO_PANEL_BOX)
    name = parse_name([t for t, x0, y0, x1, y1 in items if _in_region((t, x0, y0, x1, y1), NAME_REGION)])
    level = parse_level([(t, x0, y0, x1, y1) for t, x0, y0, x1, y1 in items if _in_region((t, x0, y0, x1, y1), LEVEL_REGION)])
    return name, level, _cp_from_items(items)


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
# 数值像素校验：OCR 之外的第二条证据
# ----------------------------------------------------------------------------
# 动机：align_affix_value 会把读不准的结果**吸附到最近的合法档位**，于是"没把握"被静默
# 转成"自信的错值"，而词条表校验**按设计拦不住**（错值也是合法档位）。2026-10-07 全量
# 离线重放实测：主数据 636 格里有 8 格这么错（把 11.81% 记成 11.11%、把 6.18% 记成 6.88%）。
#
# 手段：数值文字的字形宽度是稳的（亚像素渲染动不了它，实测同字符极差 ≤1px），
# 所以可以按记录值逐字符比宽度，对不上就报警。
#
# **能做什么**：检出 **'1' 与宽字形之间**的误读 —— 实测 8 个真错（11.81% 被记成 11.11%、
#   6.18% 被记成 6.88%）全属此类；在 output/shots/ 265 张带标签语料上命中 8/8、误报 0。
# **不能做什么**（别把话说大）：说不出正确值是多少（'6' 与 '8' 同宽），所以**只能报警、
#   不能纠正**；也检不出同宽数字之间（'6'↔'8'、'0'↔'2'）的误读。
#   '7'(10px)↔'0'(11px) 也检不出：只差 1px，而 GLYPH_W_TOL 是留给字体/网格漂移的余量
#   （实测同字符极差 1px）。为了多抓这 1px 而收到 TOL=1 不划算 —— 换来的假警报要人
#   一条条看，比漏掉那个尚未被观测到的错法代价更大。
#
# ⚠️ 宽度表只对当前这个字体/字号 + **当前裁框**（AFFIX_VAL_BAND、行 y±18）成立：
#    同字符极差 1px，换裁框就会整体偏 1px。游戏改 UI 后要连同裁框一起重标。
GLYPH_W = {"1": 4, "7": 10, "0": 12, "2": 12, "3": 12, "4": 12, "5": 12,
           "6": 12, "8": 12, "9": 12, ".": 1, "%": 15}   # 130 串标定，众数占比 95~100%
GLYPH_W_TOL = 3     # 实测同字符极差 ≤1px（'1' 3~4、'0' 11~12、'%' 14~15），取 3 留余量


def _val_glyph_widths(img, y_row, dy=0):
    """该词条行的数值文字 -> 各字形宽度 [px]（左起）。量不到返回 []。"""
    x0, y0, x1, y1 = _shift(AFFIX_VAL_BAND, dy)
    crop = img.crop((x0, int(y_row) - 18, x1, int(y_row) + 18)).convert("L")
    m = np.array(crop).astype(int) < 140
    # 裁框里有一条贯穿全宽的 UI 分隔线（行边界）。不剔掉，紧包围盒会被它撑满整幅宽度，
    # 后面按列切分就全乱 —— 实测每个裁框都恰好命中 1 行。
    m[m.mean(axis=1) > 0.8] = False

    cols = m.any(axis=0)
    runs, start = [], None
    for x, v in enumerate(cols):
        if v and start is None:
            start = x
        if not v and start is not None:
            runs.append((start, x - 1)); start = None
    if start is not None:
        runs.append((start, len(cols) - 1))
    if len(runs) < 2:
        return []
    # 直接取**全部**列段：裁框用的是项目自己的数值列取值域 AFFIX_VAL_BAND，左侧那些
    # 与本题无关的孤立墨点本来就在窗外。
    # ⚠️ 别在这里加"按最大间隙切一刀丢掉左边"的启发式 —— 那是在宽窗（x940~1140）上
    # 才需要的补丁；换成窄窗后杂点消失，最大间隙落到了 '.' 与 '%' 之间，
    # 反而会把 11.81% 切成 [4, 15] 两段（实测踩过）。
    return [b - a + 1 for a, b in runs]


def verify_affix_values(img, name_rows, values, dy=0):
    """按像素字形宽度复核每行的数值读数 -> [True(可疑) / False(相符) / None(不适用)] × 3。

    **它不是在识别，是在找矛盾**：说不出正确值是多少，只能说"这串字的形状配不上这个数"。

    None 的三种情形都是**无从校验**、不是"校验失败"，所以不报警：
      * 该行没有数值（空槽 / 补位 / 漏读）
      * 该行是**暗底**（第 15 档）—— 那一档的值由**背景颜色**直接定（见 _row_style），
        根本不经过 OCR，也就没有"读错"这回事
      * 量不到墨迹，或分量数与字符数对不上 —— **宁漏勿扰**：实测这类只占 2%，
        且已知样本全是暗底格；判可疑只会凭空造出假警报
    """
    out = []
    for i, (y, name) in enumerate(name_rows):
        v = values[i] if i < len(values) else None
        if y is None or y < 0 or v is None or not name or name == EMPTY_AFFIX:
            out.append(None); continue
        if _row_style(img, y) == "dark":
            out.append(None); continue   # 值由背景色定，不经过 OCR
        expect = f"{v * 100:.2f}%"
        got = _val_glyph_widths(img, y, dy)
        if len(got) != len(expect):
            out.append(None); continue
        out.append(any(abs(GLYPH_W[c] - w) > GLYPH_W_TOL for c, w in zip(expect, got)))
    return out


# ----------------------------------------------------------------------------
# 装备等级反推：读装备能力值 -> (类型, 槽位, T10) 查表映射 0~5
# ----------------------------------------------------------------------------
# 装备页顶部类型标签（火力型/辅助型/防御型）。位置按面板顶标定：实测在 面板顶+37。
# （曾一度把 y 放宽到 135 去兜位移，那是治标；现在由 _shift 按面板顶平移解决。）
# 框比标签本身大一圈是**故意的**：标签像素没问题（「火力型」实测在 x758..806、y60..78），
# 但把原来那个 90×47 的小框抠出来放大 2 倍再 OCR，2026-10-03 装的那版 rapidocr 检测不出来
# —— 20 个采样只命中 8 个，换成整条横带一起 OCR 就正常。下面这个尺寸实测 20/20，
# 再放宽（(700,40,900,105) / (680,35,960,110)）结果完全相同，不是踩某个魔法尺寸。
# 这**不是**回到上面那句「放宽 y 兜位移」的老路：坐标仍由 _shift 按面板顶平移，
# 放宽的只是「给检测模型的上下文」，因为失败原因出在框太小而不是框跑偏。
TYPE_BOX = (720, 40, 900, 100)
STAT_BOX = (860, 520, 1045, 600)       # 装备能力值面板（左标签列 右数值列）
EQUIP_XLSX = "equipment.xlsx"
ONLINE_OUT = "output/cn_collect.json"         # 主数据（角色表）；**只由 correct_gui 的合并写入**
OFFLINE_OUT = "output/cn_collect_offline.json"  # 离线自检结果（独立文件，不覆盖真实结果）
PATCH_OUT = "output/cn_patch.json"              # 采集产出（补丁）；**绝不直接改主数据**
# 补丁文件的**顶层兄弟键**：与「角色」同下标，每条是 {维度: 值 或 None}（整条也可为 None）。
# 存在的理由：图标是唯一的机器身份证据，但原先只打印到控制台、扫完就没了，于是合并侧手里
# 只剩一个可能截断的姓名，只能拿名字猜。带过去之后合并侧才能按四维定人。
# **刻意放顶层而不是塞进行里**：行的键集是对下游的契约（数据说明.md），且 correct_gui
# 的 normalize_row 会丢掉行里不认识的键 —— 塞进去既带不过去又会漏进主数据。
ICONS_KEY = "_图标"
# 第二个顶层兄弟键：数值的**像素校验**结果（见 verify_affix_values）。
# 与「角色」同下标，每条是 {寻址键: 该格被标时的数值} 或 None。寻址键对齐 correct_gui 的
# Prob.col（1 基，如 "头1值"）。
# **值随标记一起存**是这个设计的要点：界面只在「当前数值 == 标记里的数值」时才显示可疑，
# 于是**改对之后标记自动失效**，不需要任何清标记的钩子。
# 同样刻意放顶层不进行里（理由见上面 ICONS_KEY），且这条要跟着主数据长期存活，
# 不像 _图标 用完即弃 —— 所以 correct_gui 侧另有持久化处理。
SUSPECT_KEY = "_可疑"
ROSTER_PATH = "docs/chacters.json"              # 图鉴（姓名 -> 属性/企业/武器/职业），只拿姓名当名单

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
    """装备页 -> (等级, [词条名x3], [数值x3], [可疑x3])。

    第 4 个元素来自 verify_affix_values（像素字形宽度复核）：
    True=读数与字形宽度矛盾、False=相符、None=无从校验。**只报警，不改数值。**

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
        return None, [None, None, None], [None, None, None], [None, None, None]

    # T10：读类型 + 能力值 -> 查表反推等级
    level = None
    if slot and stat_table is not None:
        typ = read_equip_type(engine, img, dy_top)  # 类型标签在描述之上，跟面板顶
        if typ:
            stats = read_ability_stats(engine, img, dy)
            level = infer_level(stats, typ, slot, stat_table)

    # 词条数值：按 y 对齐到 name_rows，多参数读 + 词条表校验 + 投票（仅 T10）
    affix_vals = _read_affix_values(engine, img, name_rows, load_affix_table(), dy)
    # 第二条证据：OCR 读数照旧，另外用像素字形宽度复核一遍（只标记，不改值）
    suspect = verify_affix_values(img, name_rows, affix_vals, dy)

    return level, affix_names, affix_vals, suspect


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
    """翻页完成判定：回到信息页，且**战力连续两次读到同一个新值**（或名字连续稳定且 != 上一角色）。

    主判据是战力：长名有平行滚动，同一角色截两次可能读到不同片段，
    旧的 _name_changed 因此会假阳性（误判成已翻页）。
    名字只作兜底，且要求连续两次读到同一个名字才算稳定（滚动中读不到重复值）。

    ⚠️ 战力**必须连着两帧读到同一个值**才算数，只信一帧会在动画没走完时就放行：
    翻页只是面板在滑动，面板外的亮度不变，所以 page_state 全程都判 'info'（这道判据在这里
    完全失效），而战力裁剪框在滑动中会读到**混叠数字** —— 那一帧读到的"新战力"既不是旧的
    也不是新的。放行早了，这一行的姓名/战力/图标读的是**旧页**、四次装备点击却落在**新页**上，
    于是把下一个角色的装备记到这一行头上（莱伊·甲 = 桑迪·头 就是这么来的）。
    实测：停米卡跑 3 行增量，第 2 行读到的还是米卡（战力与第 1 行一字不差）。
    """
    t0 = time.time()
    last_name, last_cp = None, None
    trace = []  # 超时时回放最后几次轮询，看清是「页面没翻」还是「翻了但没读到」
    while time.time() - t0 < timeout:
        img = screen.grab()
        st = page_state(img)
        cp = None
        if st == "info":
            # 战力必须走 read_cp（= read_info_page 同一条路）：换个小框读会得到另一个数，
            # 那样 prev_cp 永远比不上，第一帧就假阳性，见 read_cp。
            cp = read_cp(engine, img)
            # 同一个 ≠prev_cp 的值连着两帧 → 认定翻页完成（滑页里的混叠读数撑不过第二帧）
            if cp is not None and cp != prev_cp and cp == last_cp:
                return True
            nm = parse_name(_crop_ocr(engine, img, NAME_REGION))
            # 名字兜底：**prev_name 读不到时这一路直接作废**。否则「页面上还是旧角色」的那两帧
            # 天然同名同值，动画还没开始动就秒返回 —— 旧实现正是这么在 prev_name=None 上栽的。
            if (prev_name is not None and nm is not None
                    and nm != prev_name and nm == last_name):
                return True
            last_name, last_cp = nm, cp
        else:
            last_name, last_cp = None, None   # 状态不明/转场：清掉，连续帧重新攒
        trace.append((round(time.time() - t0, 1), st, cp))
        time.sleep(interval)
    print(f"    [翻页诊断] 上一角色 战力={prev_cp} 名字={prev_name}；最后 {min(6, len(trace))} 次轮询：")
    for ts, st, cp in trace[-6:]:
        print(f"      t={ts:>5}s  页面状态={st:<8} 读到的战力={cp}")
    return False


def _confirm_same_char(screen, engine, cp, tries=4, interval=0.3):
    """点第一个装备槽之前问一句：屏幕上**还是战力 cp 的那个角色**吗。-> (是否还是, 最后读到的战力)

    这是「身份与装备不许错位」的最后一道闸。`wait_page_turn` 判「翻页完成」若放行早了，
    这一行的名字/战力来自旧页、而下面点开的是新页的装备槽 —— 实测 莱伊·甲 就是这么变成
    桑迪·头的。判负 → 整行作废 + 立即中止（不写进补丁）。

    读的是**战力值**、不是像素：信息页上有会动的东西（滚动横幅/立绘/高亮），像素级比对
    实测每一行都对不上（见 read_icons_stable 那段教训）。该问的是「再独立读一次，还是同一个人吗」。

    只做**一次**（第一槽之前），不是每槽都做：翻页动画的尾巴只可能落在开头这一下，而复核
    一次要花一次面板 OCR —— 实测机器上一次 OCR 要 1~2 秒，每槽都做等于给每个角色 +4 次
    （50s/角色的扫描再多 8s），不值。

    判定写成「**读到过 cp 就放行，读满 tries 次都没读到 cp 才判负**」：
      * 屏幕上就是这个角色时，哪怕 OCR 偶发读空，后面的重试总能读回 cp，不会冤枉；
      * 真换人了的话，滑动中读空也好、读出混叠值也好，等动画停下来只会读到**新角色**的
        战力，永远等不回 cp —— 所以重试不是"宽容"，而是把动画尾巴等完。
    """
    last = None
    for _ in range(tries):
        now = read_cp(engine, screen.grab())
        if now == cp:
            return True, cp
        if now is not None:
            last = now
        time.sleep(interval)
    return False, last


def _empty_slot():
    return {"等级": None, "词条": [{"名称": None, "数值": None} for _ in range(3)]}


# ----------------------------------------------------------------------------
# 身份断言（--as）：人就在游戏里看着屏幕，比任何机器判定都可靠
# ----------------------------------------------------------------------------
def norm_name(s):
    """姓名归一化：去空白 + 半角 ':' 统一成全角 '：'。

    图鉴和游戏横幅都用全角，但手打 --as 常打半角；不统一的话「红莲:暗影」永远对不上。
    """
    return re.sub(r"\s+", "", str(s or "")).replace(":", "：")


def name_matches(ocr_name, want):
    """OCR 读到的名与断言的名是否「不算明显不符」。**读不到名字时返回 True**（无从比对，不冤枉）。

    姓名区是滚动横幅，截帧天然会截断，所以**子串和子序列都算合理**
    （「红莲」/「红暗」都可能是「红莲：暗影」的截帧）。真正要抓的是「停错格子 / 翻页没完成」——
    那种情况下两个名字几乎不可能有包含关系。
    """
    if not ocr_name or not want:
        return True
    o, w = norm_name(ocr_name), norm_name(want)
    if o == w or o in w or w in o:
        return True
    it = iter(w)
    return all(c in it for c in o)  # o 是 w 的子序列


def load_roster_names(path=ROSTER_PATH):
    """图鉴里的姓名集合。**只取 key 当名单**，值（属性/企业/武器/职业）这里用不到。

    与 correct_gui.load_roster 同口径。读不到返回 None（不阻断采集，只是关掉这道校验）。
    """
    try:
        raw = json.loads(Path(path).read_text(encoding="utf-8"))
    except Exception:
        return None
    return {norm_name(k) for k in raw} if isinstance(raw, dict) and raw else None


# ----------------------------------------------------------------------------
# 信息页那四个图标：抠图 -> 匹配 -> 级联定名
# ----------------------------------------------------------------------------
# 为什么要有这一套：姓名读的是游戏里的**滚动横幅**，截帧天然会截断，实测约四成读错。
# 最险的一种错是「截断后恰好等于另一个真实角色」—— 那种错在 correct_gui 里是**零告警**的
# （`check_name` 第一句就是「名字在图鉴里就返回无问题」），于是会静默把别人的装备并到
# 另一个人头上。信息页下方那四个图标是静态资源、逐像素稳定，是**和姓名完全无关**的独立判据。
ICON_DIMS = ["属性", "武器", "职业", "企业"]   # 同时也是图鉴里的字段名（游戏里叫「妮姬类型」的那维，图鉴叫「职业」）
ICON_TILE = (24, 20)                          # 模板与样本都降采样到这个尺寸再比（实测零损失）
ICON_TEMPLATES = "docs/icon_templates.png"    # 模板图：4 行 x 6 列，每格 ICON_TILE

# 抠图框（1920x1080 绝对坐标，实测定死）。四条都别随手放宽：
#   * 四个图标的 x 在 31 张样本里**逐像素一致**，所以框可以贴紧字形。
#   * **武器框 y0 必须 >= 608** —— 正上方 y593~607 有一行**随角色变化**的文字。切进去就等于把
#     一个「随名字变的东西」当成了图标特征，判别力被静默拉低，而且看不出是谁的错。
#   * 职业/企业的 y1 停在字形底（658 / 660）：再往下一行是 `LV.xxx`，那个数字每次都会变。
#   * 企业框贴到字形（32px 宽），不是原来凭感觉给的 55px —— 原框六成是空白面板，把类间距离
#     稀释掉三分之二（实测最优/次优间隔因此从 30x 掉到 7~9x）。
ICON_BOXES = {
    "属性": (1595, 621, 1647, 675),
    "武器": (1659, 620, 1713, 665),
    "职业": (1740, 622, 1766, 658),
    "企业": (1810, 625, 1838, 660),
}
# 每个维度的取值。**顺序必须与模板图里的列一致**（重建模板用 build_icon_templates.py）。
# 「超规格极乐净土」是游戏里真实存在的第 6 种企业图标（同一家企业的超规格版，中心十字是空心
# 的，与普通版只差 4.6 —— 而不同企业之间差 14~19），但图鉴的 `企业` 字段里没有这回事，
# 所以匹配图鉴时靠 ICON_ALIAS 折回去。
ICON_VALUES = {
    "属性": ["水冷", "燃烧", "电击", "铁甲", "风压"],
    "武器": ["冲锋枪", "发射器", "机枪", "步枪", "狙击步枪", "霰弹枪"],
    "职业": ["火力型", "辅助型", "防御型"],
    "企业": ["反常", "朝圣者", "极乐净土", "泰特拉", "米西利斯", "超规格极乐净土"],
}
ICON_ALIAS = {"超规格极乐净土": "极乐净土"}   # 图标值 -> 图鉴值；不在表里的原样用
# 「这次匹配算不算可信」的判据：最优距离必须 < 0.75 x 次优。用**相对**值而不是绝对阈值 ——
# 四个维度的距离尺度差好几倍（职业的类间距离 20+，属性只有 5 上下），一个绝对阈值必然在
# 某一维上过松、另一维上过紧。0.75 是拿 31 张真实截图回归出来的：**一个正确的匹配都没误杀**。
ICON_MARGIN = 0.75

_ICON_TPL = None


def load_icon_templates(path=ICON_TEMPLATES, tile=ICON_TILE):
    """读模板图 -> {维度: {值: 灰度 ndarray}}。首次调用解析一次，之后走缓存。

    ⚠️ 读不到就**直接报错退出**，绝不照抄 load_roster_names 那种「返回 None、不阻断采集」的
    容错 —— 模板缺失的后果是**每一行都静默退回 OCR 原名**，而那正是这套东西要防的事。
    一次扫描 30 分钟，必须在开跑前就报错。

    路径走 `__file__` 而不是 cwd：本模块的 main() 从不 chdir，`python collect_cn.py` 从别处
    调用时 cwd 是调用者目录，相对路径会找不到（equipment.xlsx 那几个老相对路径同样脆弱，
    但那是既有行为，这里不跟着踩）。
    """
    global _ICON_TPL
    if _ICON_TPL is not None:
        return _ICON_TPL
    p = Path(__file__).resolve().parent / path
    if not p.exists():
        raise SystemExit(f"找不到图标模板：{p}\n"
                         f"它随仓库提供；要从截图重建请跑 build_icon_templates.py。")
    img = np.asarray(Image.open(p).convert("L"), dtype=float)
    tw, th = tile
    need = (tw * max(len(v) for v in ICON_VALUES.values()), th * len(ICON_DIMS))
    if img.shape[0] < need[1] or img.shape[1] < need[0]:
        raise SystemExit(f"图标模板尺寸不对：{p} 是 {img.shape[1]}x{img.shape[0]}，"
                         f"至少要 {need[0]}x{need[1]}（{len(ICON_DIMS)} 行 x 每行最多 "
                         f"{max(len(v) for v in ICON_VALUES.values())} 列，每格 {tw}x{th}）。")
    _ICON_TPL = {dim: {val: img[r * th:(r + 1) * th, c * tw:(c + 1) * tw]
                       for c, val in enumerate(ICON_VALUES[dim])}
                 for r, dim in enumerate(ICON_DIMS)}
    return _ICON_TPL


def read_icon_tiles(img):
    """信息页截图 -> {维度: 降采样后的灰度块}。纯裁剪，不比对、不 OCR（整轮约 1ms）。"""
    out = {}
    for dim, box in ICON_BOXES.items():
        out[dim] = np.asarray(img.crop(box).convert("L").resize(ICON_TILE, Image.BOX), dtype=float)
    return out


def match_icons(tiles, tpl):
    """{维度: 灰度块} -> {维度: (值 或 None, 是否可信)}。

    逐格算平均绝对差取最小，再用上面那条相对判据定「可信」。**认不准就返回 None，绝不硬猜** ——
    认错的代价是把别人的装备并到另一个人头上，比「认不出、退回原名」严重得多。
    """
    out = {}
    for dim, a in tiles.items():
        scored = sorted((np.abs(a - t).mean(), v) for v, t in tpl[dim].items())
        best, dist = scored[0][1], scored[0][0]
        second = scored[1][0] if len(scored) > 1 else float("inf")
        out[dim] = (best, True) if dist < ICON_MARGIN * second else (None, False)
    return out


def read_icons_stable(screen, img, tpl, verify=False, tries=3, wait=0.15):
    """读四个图标 -> {维度: (值, 可信)}。verify=True 时要求**两帧的判定结论一致**才认。

    ⚠️ 比对的是**匹配出来的值**，不是像素。一开始写成比像素（容差 1.0），**真机上每一行都被判
    「两帧不一致」**——信息页上有会动的东西（滚动横幅、立绘、高亮），像素级永远对不上，等于
    把这套功能整个关掉。而真正该问的是「两次独立读数会不会得出同一个人」：结论一致就说明画面
    已经定下来了，那点像素噪声根本不影响匹配（同值类内距离 ~1.5，不同值之间 14~19）。

    为什么还要这一道：`page_state` 判「信息页」只看外圈亮度、**不看内容**，而 `wait_page_turn`
    用战力（CP_REGION 在 y330~395）判翻页 —— 图标在 y620，比它低 300px。面板若自上而下重绘，
    会出现「战力已经是新角色的、图标还是上一个人的」那种帧，两个判据都拦不住。而图标正是拿来
    **定身份**的，读错一帧就等于把上一个人的身份安到了这一行上 —— 那比不读还糟。
    """
    prev_tiles = read_icon_tiles(img)
    first = match_icons(prev_tiles, tpl)
    if not verify:
        return first
    for _ in range(max(1, tries)):
        time.sleep(wait)
        cur_tiles = read_icon_tiles(screen.grab())
        again = match_icons(cur_tiles, tpl)
        if all(first[d][0] == again[d][0] for d in ICON_DIMS):
            return first
        first, prev_tiles = again, cur_tiles
    # 三次结论都不一样：画面一直在动。把两帧读到的东西和像素差都打出来 ——
    # 分不清「真是转场」还是「某个图标本身有动画」，得靠这两个数判断。
    show = lambda g: "、".join(f"{d}={g[d][0] or '?'}" for d in ICON_DIMS)
    gap = "、".join(f"{d}{np.abs(prev_tiles[d] - cur_tiles[d]).mean():.1f}" for d in ICON_DIMS)
    print(f"      ⚠️ 图标两次读数结论不同：前一帧 {show(first)}；像素差 {gap}")
    return None


def icons_agree(name, icons, attrs):
    """图标读数与该姓名在图鉴里的四维是否一致。-> True / 对不上的维度名 / None（无从判断）。

    「认不出」的维度不参与（返回 None 的那一维没有证据）；图鉴里查无此人（国服特供）也返回 None。
    """
    a = attrs.get(name)
    if not a:
        return None
    for d in ICON_DIMS:
        v = icons[d][0]
        if v is not None and ICON_ALIAS.get(v, v) != a.get(d):
            return d
    return True


def load_roster_attrs(path=ROSTER_PATH):
    """图鉴 -> {姓名: {属性,企业,武器,职业}}。读不到返回 None（调用方据此关掉图标定名并告警）。"""
    try:
        raw = json.loads((Path(__file__).resolve().parent / path).read_text(encoding="utf-8"))
    except Exception:
        return None
    if not isinstance(raw, dict) or not raw:
        return None
    return {k: v for k, v in raw.items() if isinstance(v, dict) and all(d in v for d in ICON_DIMS)}


def resolve_icons(ocr_name, icons, roster, attrs):
    """图标读数 + OCR 名 -> (要采用的名字 或 None, 说明)。None = 保持 OCR 原名不动。

    判据分档，**只有前两档敢改名**：
      ① 图标唯一命中（|C| == 1）且不是 OCR 名 —— 图标单独就把人认死了。
      ② C 与姓名候选 N 的交集唯一、且不是 OCR 名 —— 图标和姓名两路互相印证。
      ③ 其余一律**保持原名 + 告警**。特别地，「C ∩ N 恰好等于 OCR 名、但 |C| > 1」是**伪确认**：
        图标其实一个候选都没排除掉（`红莲` 的四维同元组里还有 `红莲：暗影`），只是恰好包含
        这个名字而已。把它记成「按图标定为红莲」会让人误以为拿到了双证据 —— 必须单独拎出来。
    """
    bad = [d for d, (v, _ok) in icons.items() if v is None]
    if bad:
        return None, "图标认不出（" + "、".join(bad) + "），退回 OCR 名"
    want = {d: ICON_ALIAS.get(v, v) for d, (v, _ok) in icons.items()}
    shown = "、".join(f"{d}={icons[d][0]}" for d in ICON_DIMS)
    C = {n for n, a in attrs.items() if all(a.get(d) == want[d] for d in ICON_DIMS)}
    if not C:
        return None, (f"图标（{shown}）在图鉴里找不到四维全等的人 "
                      f"—— 国服特供（如 婴宁/画皮）属正常，其余要人工看")
    nm = norm_name(ocr_name)
    # 改名时把「原来读成什么」一并报出来 —— 否则日志里只有「按图标定为 X」，看不出级联修了什么
    was = f"屏幕读到「{ocr_name}」" if ocr_name else "屏幕没读到名字"
    N = {norm_name(x) for x in name_candidates(nm, roster, limit=10)} if nm else set()
    if len(C) == 1:
        only = next(iter(C))
        if norm_name(only) == nm:
            return None, f"图标（{shown}）唯一命中「{only}」，与屏幕读到的名字一致"
        if not N:
            # 屏幕上读到的名字在图鉴里**连一个相近的都没有**。这只有两种可能：这人是国服特供
            # （如 婴宁/画皮，全网无图鉴），或者这名字已经碎到不成形。**两种分不开** ——
            # 但代价不对称：特供被「唯一命中」的那个人**一定是错的**，而且改名之后名字就合法了，
            # correct_gui 反而不再告警（现状至少会标 ▲）；而读崩的名字留着，现状就会把它标出来。
            # 所以这一档**不改名**，只把图标看到的喊出来。
            return None, (f"图标唯一命中「{only}」（{shown}），但屏幕上读到的「{ocr_name}」"
                          f"在图鉴里连相近的名字都没有 —— 可能是国服特供（如 婴宁/画皮），"
                          f"也可能只是读崩了。**不改名**，请人工确认")
        return only, f"按图标定为「{only}」（{was}）：四维（{shown}）在图鉴里唯一命中"
    inter = {n for n in C if norm_name(n) in N}
    if len(inter) == 1:
        only = next(iter(inter))
        if norm_name(only) != nm:
            return only, (f"按图标 + 姓名候选定为「{only}」（{was}）：图标留下 {len(C)} 个候选 "
                          f"（{'、'.join(sorted(C))}），姓名候选再筛到它一个")
        return None, (f"图标与姓名一致，但**图标没能排除别人**：四维同组的还有 "
                      f"{'、'.join(sorted(x for x in C if x != only))} —— 需人工确认")
    return None, (f"图标留下 {len(C)} 个候选（{'、'.join(sorted(C))}），"
                  f"与姓名候选交不出唯一结果，退回 OCR 名")


def _icon_record(icons):
    """{维度: (值 或 None, 是否可信)} -> {维度: 值 或 None}，供补丁文件携带（见 ICONS_KEY）。

    **无损**：match_icons 只在不可信时才给 None（可信时给的一定是个值），所以「值」这一个
    字段就同时表达了「读到了什么」和「可不可信」—— 某一维是 None 就等于那一维没读准。
    icons 本身为 None（read_icons_stable 两帧结论不一致）时整条记 None，与「读到了但都不可信」区分开。
    """
    if icons is None:
        return None
    return {d: icons[d][0] for d in ICON_DIMS}


def _suspect_record(items):
    """[(槽, 词条下标, 该格数值)] -> {"头1值": 0.1111, ...}，供补丁携带（见 SUSPECT_KEY）。

    寻址键刻意对齐 correct_gui 的 `Prob.col`（1 基，如 "头1值"），这样合并侧不必再做一层
    位置换算就能直接拿去标记那一格。**值一起存**：界面只在当前数值仍等于它时才显示可疑，
    于是改对之后标记自动失效。
    """
    if not items:
        return None
    return {f"{slot}{i + 1}值": v for slot, i, v in items}


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

    # 图标级联自检：只读不改、几毫秒，比真机跑一趟快得多。用来确认模板与抠图框在
    # 这台机器/这个分辨率上是通的 —— 认不出会明说，而不是让姓名静默退回 OCR。
    try:
        icons = match_icons(read_icon_tiles(img_ps), load_icon_templates())
        print("  图标：" + "、".join(f"{d}={icons[d][0] or '认不出'}" for d in ICON_DIMS))
        attrs = load_roster_attrs()
        roster_full, rerr = load_roster(str(Path(__file__).resolve().parent / ROSTER_PATH))
        if attrs and roster_full:
            got, why = resolve_icons(name, icons, roster_full, attrs)
            print(f"  图标定名：{why}" + (f" → 「{got}」" if got else "（保持原名）"))
        else:
            print(f"  图标定名：跳过（图鉴读不到：{rerr}）")
    except SystemExit as e:
        print(f"  图标自检跳过：{e}")

    stat_table = load_stat_table()
    slots = {}
    for slot, _ in SLOTS.items():
        elevel, anames, avals, asuspect = read_equip_page(
            engine, img_eq, slot=slot, stat_table=stat_table)
        affixes = [
            {"名称": n, "数值": v} for n, v in zip(anames, avals)
        ]
        slots[slot] = {"等级": elevel, "词条": affixes}
        flag = [f"{slot}{k + 1}值" for k, b in enumerate(asuspect) if b]
        print(f"  槽位[{slot}] 等级={elevel} 词条={affixes}"
              + (f"  ⚠️ 字形宽度不符：{flag}" if flag else ""))

    row = build_row(name, cp, slots)
    out = {"角色": [row]}
    # 写独立文件：离线是自检，不能覆盖在线采集的真实结果
    Path("output").mkdir(exist_ok=True)
    with open(OFFLINE_OUT, "w", encoding="utf-8") as f:
        json.dump(out, f, ensure_ascii=False, indent=2)
    print(f"\n已写入 {OFFLINE_OUT}")
    return out


def _archive_previous(path):
    """--replace 要整表覆盖主数据前，把旧文件改名带时间戳留底。

    否则新一轮会直接覆盖 output/cn_collect.json —— 这个坑真实发生过：37 人的结果
    被一次 6 人试跑覆盖掉，而落盘用的是 os.replace（原地替换、不进回收站），找不回来。
    空壳/损坏的文件不值得留底，返回 None 让新结果直接覆盖它。

    默认的补丁路径**不调它**：补丁本来就一个字节都不动主数据，留底无从谈起。
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


def _json_rows(path):
    """这个结果文件里有多少行；读不到/空壳返回 0。"""
    try:
        return len(json.loads(Path(path).read_text(encoding="utf-8")).get("角色", []))
    except Exception:
        return 0


def _save_rows(rows, out=ONLINE_OUT, extra=None):
    """把已采的角色增量写盘：原子替换（先写临时文件再 rename），中途中断不丢、不坏。

    out 可覆盖，供 correct_gui.py 复用同一套落盘格式（默认值保证采集路径行为不变）。

    extra 是**顶层兄弟键**，补丁路径用（见 ICONS_KEY）；--replace 路径 extra=None。
    刻意不写进行里：行的键集是给下游的契约（`数据说明.md`），而且 correct_gui.normalize_row
    会把行里不认识的键静默丢掉 —— 塞进行里既带不过去、又有漏进主数据的风险。
    extra 为 None 时输出与旧版逐字节相同。
    """
    data = {"角色": rows}
    if extra:
        data.update(extra)
    Path(out).parent.mkdir(parents=True, exist_ok=True)
    tmp = Path(str(out) + ".tmp")
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
    tmp.replace(out)


def collect_online(engine, screen, max_chars=500, save_shots=False, replace=False, asserts=None,
                   with_icons=True):
    """在线循环采集：停在信息页，翻页直到 LV.1。save_shots=True 时每页截图存 output/shots/。

    **默认写补丁**（PATCH_OUT + 顶层 `_图标`）：主数据 ONLINE_OUT **一个字节都不动** ——
    补丁要经 correct_gui 合并后才并进去，错名/误判因此永远污染不了主数据。扫多少只是范围
    不同（max_chars 小 = 增量重扫，大 = 全量，遇 LV.1 自动停），后面「对比旧表 → 更新/照旧/
    追加」是同一件事，所以这里没有「增量/全量」两套代码。

    replace=True 是**逃生口**（换号 / 砍掉重练）：归档主数据后用本次结果整个替换它。
    界面不可达，只能手打 --replace。这条路径不写 `_图标`（extra=None，输出与旧版逐字节相同），
    也正因为它是唯一会**整表覆盖**的路径，才必须显式。
    asserts 是 --as 给的身份断言，**按扫描顺序一一对应**（见 name_matches）。
    with_icons=True 时用信息页那四个图标独立定名（见 resolve_icons），--no-icons 可关掉。
    """
    rows = []
    icon_records = []          # 与 rows 同下标；随补丁落盘（见 ICONS_KEY）
    suspect_records = []       # 同上；随补丁落盘（见 SUSPECT_KEY）
    out = ONLINE_OUT if replace else PATCH_OUT
    # 顶层兄弟键。字典里放的是两个 list 的**引用**，所以它们随循环增长、无需重建。
    # replace 路径 extra=None —— 主数据输出与旧版逐字节相同（逃生口，不携带这两项）。
    extra = None if replace else {ICONS_KEY: icon_records, SUSPECT_KEY: suspect_records}
    asserts = [a for a in (asserts or []) if a]
    roster = load_roster_names() if asserts else None
    stat_table = load_stat_table()

    # 图标定名的三个前置件。模板读不到直接 SystemExit（见 load_icon_templates）；
    # 图鉴读不到则**关掉整条图标定名并响亮告警** —— 那种情况下每一行都会退回 OCR 原名，
    # 约四成是错的，绝不能让它在没人知道的情况下发生。
    icon_tpl = roster_full = attrs = None
    if with_icons:
        icon_tpl = load_icon_templates()
        roster_full, rerr = load_roster(str(Path(__file__).resolve().parent / ROSTER_PATH))
        attrs = load_roster_attrs()
        if not roster_full or not attrs:
            icon_tpl = None
            print(f"⚠️⚠️ 图鉴读不到（{rerr}），图标定名整条关闭 —— "
                  f"本次每一行都会退回 OCR 原名，实测约四成会读错。")
        else:
            print(f"图标定名已就绪：{sum(len(v) for v in ICON_VALUES.values())} 个模板，图鉴 {len(attrs)} 人。")
    shots_dir = Path("output/shots")
    if save_shots:
        shots_dir.mkdir(parents=True, exist_ok=True)
    print("=== 在线模式（替换主数据） ===" if replace else "=== 在线模式（补丁） ===")
    if replace:
        n_before = _json_rows(ONLINE_OUT)
        prev = _archive_previous(ONLINE_OUT)
        if prev:
            print(f"（上一轮结果已留底：{prev}）")
        if n_before and max_chars < n_before:
            # --replace 是唯一会整表覆盖的路径，界面不可达，但手打时仍要拦一下：
            # `--max 3` 看着人畜无害，一跑就把 52 行的主数据换成 3 行。
            print(f"⚠️ --replace 会整个替换主数据：原有 {n_before} 行，本次只采 {max_chars} 行的量，"
                  f"其余 {max(n_before - max_chars, 0)} 行会消失（旧文件已留底，可还原）。\n"
                  f"   只想更新个别角色的话去掉 --replace —— 默认就写补丁，一个字节都不动主数据，"
                  f"跑完在 correct_gui 里合并。")
    else:
        print(f"结果写 {out}；主数据 {ONLINE_OUT} 不动，留给 correct_gui 合并（也不会被归档）。")
    if asserts:
        # 手打错字在这里就要报出来：等扫完再发现，白等几分钟
        if roster is not None:
            unknown = [a for a in asserts if norm_name(a) not in roster]
            if unknown:
                print(f"⚠️ 断言里有图鉴查无的名字：{'、'.join(unknown)}"
                      f"（国服特供如「婴宁」「画皮」属正常；否则大概是打错了）")
        print(f"身份断言 {len(asserts)} 个：{'、'.join(asserts)}"
              + ("" if len(asserts) >= max_chars
                 else f"（不足 --max {max_chars}，第 {len(asserts) + 1} 行起没有断言）"))
    print("5 秒后开始采集，请立即 Alt+Tab 切回游戏（角色信息页），期间别碰鼠标键盘 ...")
    for i in range(5, 0, -1):
        print(f"  {i} ...")
        time.sleep(1)
    prev_name, prev_cp = None, None
    # 第一行也按「刚翻页」处理：那是操作员手动把画面摆到这一格的，同样没被任何内容判据校验过
    turned = True
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

        # ---- 图标：四个静态图标 -> 四维读数（不 OCR，~1ms）----
        icons, icon_str, verdict_note = None, "", ""
        if icon_tpl is not None:
            icons = read_icons_stable(screen, img, icon_tpl, verify=turned)
            turned = False
            if icons is None:
                icon_str = "图标读数不稳定，这一行不做图标判定（退回 OCR 名）"
            else:
                icon_str = "图标 " + "、".join(f"{d}={icons[d][0] or '?'}" for d in ICON_DIMS)

        want = asserts[i] if i < len(asserts) else None
        if want:
            if name_matches(name, want):
                # 相符（含横幅截断）-> 采用断言的姓名。这正是 --as 的主要价值：
                # 把「红莲」这种截断名修回「红莲：暗影」——姓名是主键，截断名会撞上真实角色。
                if name != want:
                    print(f"    （姓名按断言修正：{name!r} → 「{want}」）")
                name = want
            else:
                # 不符 -> **保留屏幕读到的名字**，不让断言盖掉它。
                # 盖掉就等于「停错格子也照写断言名」，会拿着别人的装备静默覆盖「{want}」那一行；
                # 留着 OCR 名，correct_gui 会把它标成「图鉴查无此人」逼人看一眼。
                print(f"    ❌ 身份不符：屏幕读到「{name}」，你断言的是「{want}」——"
                      f"很可能停错格子/翻页没完成。这一行保留屏幕读到的名字，请务必核对！")
            # 图标与断言冲突：断言是**按下标一一对应**的，错一格后面每一行都会错位，
            # 而且错位后的名字仍然可能「看着对」（子序列规则很松）。图标的读数与下标无关，
            # 是这里唯一能当场戳穿错位的硬证据 —— 所以不停下来只会把后面全部污染。
            if icons:
                bad = icons_agree(name, icons, attrs)
                if bad not in (None, True):
                    print(f"    ❌❌ 图标与断言不符：「{bad}」这一维图标读的是「{icons[bad][0]}」，"
                          f"而「{name}」在图鉴里是「{attrs[name][bad]}」。"
                          f"断言错一格后面全会错位 —— 立即中止，请重跑。")
                    break
        elif icons and attrs:
            fixed, why = resolve_icons(name, icons, roster_full, attrs)
            verdict_note = why
            if fixed and fixed != name:
                name = fixed

        print(f"[{i + 1}] {name} LV.{level} 战力 {cp}")
        if icon_str:
            print(f"    {icon_str}")
        if verdict_note:
            print(f"    → {verdict_note}")

        if level is None:
            print("⚠️ 读不到等级，画面可能不在信息页，中断。")
            break

        slots = {}
        suspects = []   # [(槽, 词条下标, 数值)]，本行检出「读数与字形宽度矛盾」的格
        drifted = False     # 采到一半屏幕换人了：这一行整个作废，别再往下点
        checked = False     # 换人复核只做第一槽之前那一处（成本与理由见 _confirm_same_char）
        for slot, xy in SLOTS.items():
            # 每次点击前先确认还在信息页：错位了当场纠正，而不是等某个布尔变真
            if not _ensure_info(screen, xy, tries=3, timeout=5.0):
                print(f"    {slot}: 无法回到信息页（判定={page_state(screen.grab())}），中断。")
                slots[slot] = _empty_slot()
                break

            if not checked:
                checked = True
                if cp is not None:
                    same, now_cp = _confirm_same_char(screen, engine, cp)
                    if not same:
                        print(f"    ❌❌ 点「{slot}」之前屏幕就换人了：这一行是「{name}」"
                              f"（战力 {cp}），现在读到的是 {now_cp if now_cp is not None else '读不出'}。"
                              f"翻页动画在采集途中才走完 —— 再点下去会把**下一个角色的装备**"
                              f"记到「{name}」头上。本行作废，立即中止，请重跑。")
                        drifted = True
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
            elevel, anames, avals, asuspect = read_equip_page(
                engine, eq_img, slot=slot, stat_table=stat_table)
            if not any(anames):
                # 浮层开着但读不到任何词条。T9 本来就没有词条（属预期），也可能是该槽无装备/
                # 装备选择页 —— 三者输出都是空，这里不替它们下定论，记空槽并显式告警。
                print(f"    {slot}: ⚠️ 浮层已开但读不到词条（T9 属预期；也可能是无装备/装备选择页）")
                slots[slot] = _empty_slot()
            else:
                affixes = [{"名称": n, "数值": v} for n, v in zip(anames, avals)]
                slots[slot] = {"等级": elevel, "词条": affixes}
                # 像素校验只报警、不改值；这里连带把字面值记进标记（见 _suspect_record）
                for k, bad in enumerate(asuspect):
                    if bad and affixes[k]["名称"] not in (None, EMPTY_AFFIX) \
                            and affixes[k]["数值"] is not None:
                        suspects.append((slot, k, affixes[k]["数值"]))
                print(f"    {slot}: Lv{elevel} {affixes}"
                      + (f"  ⚠️ 字形宽度不符："
                         f"{[f'{s}{k + 1}值' for s, k, _ in suspects if s == slot]}"
                         if any(bad for bad in asuspect) else ""))

            # 关闭浮层：点同一槽位
            screen.click(xy)
            if not wait_page(screen, "info", timeout=8.0)[0]:
                print(f"    {slot}: 关闭后未回到信息页（判定={page_state(screen.grab())}），复位 ...")
                _ensure_info(screen, xy, tries=2, timeout=5.0)
            time.sleep(0.2)  # 返回后缓冲，避免紧接着点下一槽太快

        if drifted:
            # 身份与装备对不上的一行**整个丢掉**：补丁里宁可少一行，也不能多一行错的
            # （错行一旦并进主数据，会靠合并的「空槽保留旧值」一直活着，见莱伊·甲）。
            break

        rows.append(build_row(name, cp, slots))
        icon_records.append(_icon_record(icons))   # 必须与 rows 同步 append，否则下标错位
        suspect_records.append(_suspect_record(suspects))   # 同样必须锁步
        _save_rows(rows, out, extra=extra)  # 每采完 1 角色就落盘，中断不丢

        if level == 1:
            print(f"到达 LV.1（{name}），采集完成，停止。")
            break

        if i + 1 >= max_chars:
            # 别在最后一格再翻一次页：翻了还要白等最多 15 秒，而且游戏会停在**下一个**角色上 ——
            # 增量扫描正靠「停在原地、下次接着往后扫」，起点一漂就串行。
            print(f"已达 --max {max_chars}，不再翻页（游戏停在「{name}」这一格）。")
            break

        # 翻页到下一个角色：等战力变化（名字有平行滚动，截两次可能不同，不可靠）
        prev_name, prev_cp = name, cp
        screen.click(NEXT_PAGE)
        if not wait_page_turn(screen, engine, prev_cp, prev_name, timeout=15.0):
            if save_shots:
                screen.grab().save(shots_dir / f"{i + 1:03d}_TURN_FAIL.png")
            print("⚠️ 翻页后未见战力变化，中断（可能已到末位或点击失效）。已存 NNN_TURN_FAIL.png。")
            break
        turned = True   # 翻页后图标读数要双帧确认，见 read_icons_stable

    _save_rows(rows, out, extra=extra)
    print(f"\n采集 {len(rows)} 个角色，已写入 {out}")
    if replace:
        print(f"⚠️ --replace：主数据 {ONLINE_OUT} 已被本次结果整个替换（旧文件已留底）。")
    else:
        print(f"这是补丁，主数据没动。界面里点「采集…」跑完会自动接上合并"
              f"（每一条都列在合并窗里，四维唯一的已预选好、可改）；命令行老路可走「合并补丁」。")
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
    ap.add_argument("--patch", action="store_true",
                    help=f"兼容别名：结果写 {PATCH_OUT}。现在不写它也是这个行为，"
                         f"主数据 {ONLINE_OUT} 不动也不归档")
    ap.add_argument("--replace", action="store_true",
                    help=f"⚠️ 逃生口：归档后用本次结果整个替换主数据 {ONLINE_OUT}。"
                         f"只有换号 / 砍掉重练才用；默认（写补丁）安全得多，界面里也不可达")
    ap.add_argument("--as", dest="asserts", default="", metavar="名字[,名字…]",
                    help="按扫描顺序断言身份，如 --as '红莲：暗影,桃乐丝'。人就在游戏里看着屏幕，"
                         "比 OCR 可靠。已废弃，只在图标定不下来时当逃生口用")
    ap.add_argument("--no-icons", action="store_true",
                    help="关掉「用信息页那四个图标独立定名」。逃生开关，正常情况下不用它 —— "
                         "关掉之后姓名就只剩 OCR 一条路，实测约四成会读错")
    ap.add_argument("--check-state", action="store_true",
                    help="只校验页面状态分类器（不采集、不加载 OCR）：根目录样本 + output/shots/*.png")
    args = ap.parse_args()

    # 全角/半角逗号都认：手打的时候没人会去想用的是哪一种
    asserts = [s.strip() for s in re.split(r"[,，]", args.asserts) if s.strip()]
    if args.replace and (args.patch or asserts):
        # 语义打架：--replace 要整表覆盖，而 --patch / --as 是补丁语义。手打错了就直说。
        # 放在最前面：打错字不该先白等一次 OCR 引擎加载才知道。
        print("✗ --replace 不能和 --patch / --as 一起用 —— --replace 是整表替换，"
              "那两者是补丁语义。二选一。")
        sys.exit(2)

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
        collect_online(engine, screen, max_chars=args.max, save_shots=args.save_shots,
                       replace=args.replace, asserts=asserts, with_icons=not args.no_icons)

    del engine


if __name__ == "__main__":
    main()
