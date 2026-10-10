"""_vote_row 第二层（精确票压吸附票）的离线条烟：全部合成候选，不碰 OCR。

跑法（仓库根目录）：python tests/test_vote_exact.py
被测的是 collect_cn.py 里**真代码**；old_vote 是改动前的逻辑副本，用来对比。
"""
import re
import sys
from pathlib import Path

REPO = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(REPO))

import collect_cn as cc  # noqa: E402

sys.stdout.reconfigure(encoding="utf-8")
TABLE = cc.load_affix_table(REPO / "equipment.xlsx")
AFFIX = "攻击力增加"
POOL = TABLE[AFFIX][:11]          # normal 配色：第 1~11 档
T_EXACT = POOL[10]
# 吸附票用的文本故意**不是**中点：中点两侧等距，align_affix_value 会判成"两边都不像"而返回
# None（半档距正好卡在边界上）。真实的吸附错读是「档位值的轻微错读」，如 '11.01%' -> 0.1111、
# '11.95%' -> 0.1234，所以照这个形态造：偏一点点，落回最近的档位。
T_SNAP_TEXT = "11.01%"
T_SNAP_ALT = "10.35%"             # 第三个值，用来测"没有精确票时按票数"


def old_vote(cands, y_row, name, table, lo=0, hi=None, tol=18):
    """改动前的 _vote_row（单一强/弱桶，按票数取最大）。"""
    strong, weak = {}, {}
    for yc, t in cands:
        if abs(yc - y_row) > tol:
            continue
        v = cc.parse_pct([t])
        if v is None:
            continue
        aligned = cc.align_affix_value(name, v, table, lo, hi)
        if aligned is None:
            continue
        bucket = strong if re.search(r"\d+\.\d+\s*%", t) else weak
        bucket[aligned] = bucket.get(aligned, 0) + 1
    if strong:
        return max(strong, key=strong.get)
    return max(weak, key=weak.get) if weak else None


def snap_of(v):
    return cc.align_affix_value(AFFIX, v, TABLE, 0, 10)


ok = 0
fail = []


def check(label, got, want):
    global ok
    if got == want:
        ok += 1
        print(f"  ✓ {label}")
    else:
        fail.append(label)
        print(f"  ✗ {label}: got={got} want={want}")


TXT_EXACT = f"{T_EXACT * 100:.2f}%"
SNAPPED = cc.align_affix_value(AFFIX, cc.parse_pct([T_SNAP_TEXT]), TABLE, 0, 10)
SNAP_ALT = cc.align_affix_value(AFFIX, cc.parse_pct([T_SNAP_ALT]), TABLE, 0, 10)

print(f"档位（前 11 档）: {POOL}")
print(f"精确票 {TXT_EXACT} -> {T_EXACT}（在表里）")
print(f"吸附票 {T_SNAP_TEXT} -> {SNAPPED}（不在表里，被拉过去）；"
      f"另一档 {T_SNAP_ALT} -> {SNAP_ALT}\n")
assert SNAPPED != T_EXACT and SNAP_ALT not in (SNAPPED, T_EXACT)

TXT_SNAP = T_SNAP_TEXT

print("① 雪子型：吸附票 5 : 精确票 1 —— 多数票本身是错的")
cands = [(100.0, TXT_SNAP)] * 5 + [(101.0, TXT_EXACT)]
check("旧逻辑投给了吸附值（复现当时的错）", old_vote(cands, 100, AFFIX, TABLE, 0, 10), SNAPPED)
check("新逻辑翻回精确票", cc._vote_row(cands, 100, AFFIX, TABLE, 0, 10), T_EXACT)

print("\n② 003_甲型：3 个被表凑合的乱码 vs 1 个干净读数（乱码是强票，同为吸附）")
garble = [(100.0, TXT_SNAP)] * 3 + [(101.0, TXT_EXACT)]
check("旧逻辑把正确值投掉", old_vote(garble, 100, AFFIX, TABLE, 0, 10), SNAPPED)
check("新逻辑仍然保住正确值", cc._vote_row(garble, 100, AFFIX, TABLE, 0, 10), T_EXACT)

print("\n③ 没有精确票时：行为必须与旧逻辑逐字一致（回归闸）")
only_snap = [(100.0, TXT_SNAP)] * 5 + [(101.0, T_SNAP_ALT)] * 3
check("旧（5 票压 3 票）", old_vote(only_snap, 100, AFFIX, TABLE, 0, 10), SNAPPED)
check("新（与旧同）", cc._vote_row(only_snap, 100, AFFIX, TABLE, 0, 10), SNAPPED)

print("\n④ 强/弱分层没被新层打乱：强吸附票仍压弱精确票")
weak_exact = (100.0, f"{POOL[9] * 100:.2f}")      # 掉 % -> 弱票，但精确（0.1111）
strong_snap = (101.0, T_SNAP_ALT)                  # 带 % -> 强票，但吸附（0.104）
check("旧", old_vote([strong_snap, weak_exact], 100, AFFIX, TABLE, 0, 10), SNAP_ALT)
check("新（强在精确之前，刻意）",
      cc._vote_row([strong_snap, weak_exact], 100, AFFIX, TABLE, 0, 10), SNAP_ALT)

print("\n⑤ 边界：无视候选")
check("空", cc._vote_row([], 100, AFFIX, TABLE, 0, 10), None)
check("y 不在行附近", cc._vote_row([(300.0, TXT_EXACT)], 100, AFFIX, TABLE, 0, 10), None)
check("乱码 parse 不出", cc._vote_row([(100.0, "%26h"), (100.0, "效果变更")], 100, AFFIX, TABLE, 0, 10),
      None)
check("词条名不在表里", cc._vote_row([(100.0, TXT_EXACT)], 100, "不存在的词条", TABLE, 0, 10), None)

print("\n⑥ 第三层：吸附票胜出 + 像素宽度签名唯一 -> 改判（#020 头[1] 型）")
DEF = "防御力增加"
D_SNAP, D_TRUE = TABLE[DEF][9], TABLE[DEF][10]      # 0.1111 / 0.1181
assert (D_SNAP, D_TRUE) == (0.1111, 0.1181), (D_SNAP, D_TRUE)
c_snap = [(100.0, "11.01%")]                        # 吸附 -> 0.1111
sig_true_def = [cc.GLYPH_W.get(c, -1) for c in f"{D_TRUE * 100:.2f}%"]
check("不传 widths（= 改动前行为）：保持吸附值", cc._vote_row(c_snap, 100, DEF, TABLE, 0, 10), D_SNAP)
check("传 widths：宽度签名翻回真值", cc._vote_row(c_snap, 100, DEF, TABLE, 0, 10, widths=sig_true_def), D_TRUE)

print("\n⑦ 同宽歧义（'2'/'5'/'8' 都是 12px）-> margin=0，必须弃权")
amb = [cc.GLYPH_W.get(c, -1) for c in "6.88%"]      # 与 4.77/5.47/6.18/8.29/9.00 同签名，谁也压不过谁
check("弃权，保持 OCR 的吸附值", cc._vote_row(c_snap, 100, DEF, TABLE, 0, 10, widths=amb), D_SNAP)

print("\n⑧ widths 段数与任何档位都对不上 -> 弃权（宁漏勿扰）")
check("弃权", cc._vote_row(c_snap, 100, DEF, TABLE, 0, 10, widths=[4, 4, 1, 12, 4, 15, 7]), D_SNAP)

print("\n⑨ 优越代码型（真实病例 #010 头[0]）：0.1234 -> 0.1795")
SUP = "优越代码伤害增加"
S_SNAP, S_TRUE = TABLE[SUP][2], TABLE[SUP][6]        # 0.1234 / 0.1795
assert (S_SNAP, S_TRUE) == (0.1234, 0.1795), (S_SNAP, S_TRUE)
c_sup = [(100.0, "11.95%")]                          # 吸附 -> 0.1234（真实错读形态）
sig_true_sup = [cc.GLYPH_W.get(c, -1) for c in f"{S_TRUE * 100:.2f}%"]
check("OCR 吸附成 0.1234", cc._vote_row(c_sup, 100, SUP, TABLE, 0, 10), S_SNAP)
check("宽度签名翻回 0.1795（margin 只有 2，是最紧的一例）",
      cc._vote_row(c_sup, 100, SUP, TABLE, 0, 10, widths=sig_true_sup), S_TRUE)

print(f"\n=== {ok} 通过 / {len(fail)} 失败 ===")
for f in fail:
    print(f"    ✗ {f}")
sys.exit(1 if fail else 0)
