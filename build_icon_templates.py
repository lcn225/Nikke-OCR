#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""从一整套信息页截图重建 docs/icon_templates.png（图标模板库）。

**什么时候需要重跑**：游戏更新了那四个图标的美术资源（换字形、换配色）。症状是
collect_cn.py 开始报「图标认不出」。这时用**新**截图重跑本脚本即可，不用改代码 ——
抠图框和取值表都在 collect_cn.py 里统一定义，本脚本直接 import 它们，不会各写一份。

用法（在项目根目录）：
    python build_icon_templates.py
    python build_icon_templates.py --shots output/shots --json output/cn_collect.json
    python build_icon_templates.py --check        # 只校验现有模板，不重建

原理：**同一个值在不同角色上渲染得逐像素一致**（31 张实测），所以按像素聚类后，
每一簇取成员的中位数就是该值的模板。标签不用手工标 ——
截图顺序 == 扫描顺序，把扫出来的行按名字对回图鉴就得到每张截图是谁；某一簇里若有个别
成员对不上（横幅截断读错名），取多数票即可，聚类本身会把它暴露成「少数派」。

⚠️ 前提是那份 JSON 已人工校正过。全新扫描的 JSON 名字约四成是错的，但**多数票能扛住**：
一个值通常有 3~16 个样本，错名是分散的、不会同时错到同一个值上。
"""
import argparse
import collections
import itertools
import json
import sys
from pathlib import Path

import numpy as np
from PIL import Image

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT))

import collect_cn as cc                                   # noqa: E402
from correct_gui import load_roster, name_candidates, norm_roster_name   # noqa: E402

CLUSTER_DIST = 3.0      # 同值像素簇的阈值（实测同值内 ~1.5、不同值 14+，中间很空）


def cluster(arrs, th=CLUSTER_DIST):
    """单链聚类：两两平均绝对差 < th 就并到一起。返回 [成员下标列表]，按人数降序。"""
    n = len(arrs)
    lab = list(range(n))
    for i, j in itertools.combinations(range(n), 2):
        if np.abs(arrs[i] - arrs[j]).mean() < th:
            a, b = lab[i], lab[j]
            if a != b:
                lab = [a if x == b else x for x in lab]
    out = collections.defaultdict(list)
    for i, l in enumerate(lab):
        out[l].append(i)
    return sorted(out.values(), key=len, reverse=True)


def resolve_names(rows, roster):
    """扫出来的每一行 -> 图鉴里的姓名（精确命中优先，否则取相似度第一的候选）。"""
    out = []
    for r in rows:
        nm = norm_roster_name(r.get("姓名") or "")
        if nm in roster["orig"]:
            out.append(roster["orig"][nm])
            continue
        cands = name_candidates(nm, roster, limit=1)
        out.append(roster["orig"][norm_roster_name(cands[0])] if cands else None)
    return out


def build(shots_dir, json_path, out_path):
    attrs = json.loads((ROOT / cc.ROSTER_PATH).read_text(encoding="utf-8"))
    roster, err = load_roster(str(ROOT / cc.ROSTER_PATH))
    if not roster:
        raise SystemExit(f"图鉴读不到（{err}），没法给截图打标签。")
    rows = json.loads(Path(json_path).read_text(encoding="utf-8"))["角色"]
    shots = sorted(Path(shots_dir).glob("*_info.png"))
    if len(shots) != len(rows):
        raise SystemExit(f"截图 {len(shots)} 张、JSON {len(rows)} 行，对不上 —— "
                         f"两者必须是同一次扫描（顺序即扫描顺序）。")
    names = resolve_names(rows, roster)

    tw, th = cc.ICON_TILE
    sheet = np.zeros((th * len(cc.ICON_DIMS),
                      tw * max(len(v) for v in cc.ICON_VALUES.values())), dtype="uint8")
    print(f"截图 {len(shots)} 张 -> 模板 {out_path}")
    for r, dim in enumerate(cc.ICON_DIMS):
        tiles = [np.asarray(Image.open(p).convert("RGB").crop(cc.ICON_BOXES[dim])
                            .convert("L").resize(cc.ICON_TILE, Image.BOX), dtype=float)
                 for p in shots]
        # 先给每一簇投出它「图鉴认为」的值
        voted = []          # [(票数, 值, 成员下标)]
        for mem in cluster(tiles):
            votes = collections.Counter(
                attrs[names[i]][dim] if names[i] and dim in attrs.get(names[i], {}) else "?"
                for i in mem)
            if "?" in votes:
                print(f"  ⚠ {dim}：有 {votes['?']} 张的名字对不上图鉴，已从多数票里剔除")
                del votes["?"]
            if not votes:
                print(f"  ⚠ {dim}：这一簇没有一个能确定的名字，整簇跳过（{len(mem)} 张）")
                continue
            val, n = votes.most_common(1)[0]
            if len(votes) > 1:
                others = "、".join(f"{k}x{v}" for k, v in votes.items() if k != val)
                print(f"  · {dim} {val}：{n} 票通过，{len(mem) - n} 张对不上（{others}）"
                      f" —— 多半是横幅截断读错了名，聚类本身没受影响")
            voted.append((len(mem), val, mem))

        # ⚠️ 同一个值可能聚出**多个**视觉簇。这不是噪声 —— 是游戏里真有变体，而图鉴的字段
        # （比如「企业」）装不下它。实测：`拉毗：小红帽` 的企业图标比同企业的其余 7 人
        # 差 4.6（不同企业之间差 14~19），就是「超规格」版。**绝不能拿小簇覆盖大簇**
        # （那会把 7 个样本的模板换成 1 个的），也绝不能把两簇合并（那就把变体抹掉了）。
        by_val, extra = {}, []
        for n, val, mem in sorted(voted, key=lambda z: -z[0]):
            if val in cc.ICON_VALUES[dim] and val not in by_val:
                by_val[val] = [tiles[i] for i in mem]
            else:
                extra.append((n, val, mem))
        # 多出来的簇，按人数从多到少去认领 ICON_VALUES 里**没有任何簇投票**的空档 ——
        # 那些空档正是为变体留的名字。认领是启发式的，所以把话说明白，别装作知道。
        empty = [v for v in cc.ICON_VALUES[dim] if v not in by_val]
        for (n, val, mem), slot in zip(extra, empty):
            print(f"  ⚠ {dim}：另有 {n} 张聚成一簇，票投「{val}」，但那一格已被更大的簇占了。"
                  f"按 ICON_VALUES 的空档认领为「{slot}」—— 图鉴里没有这个值，是启发式指派，"
                  f"请核对它是不是这个变体。")
            by_val[slot] = [tiles[i] for i in mem]
        for n, val, mem in extra[len(empty):]:
            print(f"  ⚠ {dim}：还有 {n} 张的簇票投「{val}」，但 ICON_VALUES 里没有空档可放，"
                  f"整簇丢弃。游戏加了新图标？把它加进 collect_cn.ICON_VALUES 再来。")
        missing = [v for v in cc.ICON_VALUES[dim] if v not in by_val]
        if missing:
            print(f"  ⚠ {dim}：这些值一个样本都没有 -> {missing}（模板会是空白，匹配时认不出）")
        for c, val in enumerate(cc.ICON_VALUES[dim]):
            if val not in by_val:
                continue
            med = np.median(np.stack(by_val[val]), axis=0)
            sheet[r * th:(r + 1) * th, c * tw:(c + 1) * tw] = np.round(med).astype("uint8")
            print(f"  {dim:<3} {val:<10} 样本 {len(by_val[val]):>2}")
    Image.fromarray(sheet).save(ROOT / out_path)
    print(f"已写入 {ROOT / out_path}（{sheet.shape[1]}x{sheet.shape[0]}）")


def check(json_path=None):
    """用现成模板回放一遍截图，看每个维度认得出多少、间隔够不够。"""
    tpl = cc.load_icon_templates()
    if not json_path:
        print("模板已载入：")
        for d in cc.ICON_DIMS:
            print(f"  {d:<3} {len(tpl[d])} 个值  {list(tpl[d])}")
        return 0
    rows = json.loads(Path(json_path).read_text(encoding="utf-8"))["角色"]
    roster, _ = load_roster(str(ROOT / cc.ROSTER_PATH))
    attrs = json.loads((ROOT / cc.ROSTER_PATH).read_text(encoding="utf-8"))
    names = resolve_names(rows, roster)
    shots = sorted(Path("output/shots").glob("*_info.png"))
    unsure = collections.Counter()
    clash = []          # 「图标认出 X、按名字推出来是 Y」——分不清是谁错，所以只列出来
    for i, p in enumerate(shots[:len(rows)]):
        got = cc.match_icons(cc.read_icon_tiles(Image.open(p).convert("RGB")), tpl)
        for d in cc.ICON_DIMS:
            v, ok = got[d]
            if not ok or v is None:
                unsure[d] += 1
                continue
            exp = attrs[names[i]][d] if names[i] else None
            if exp and cc.ICON_ALIAS.get(v, v) != exp:
                clash.append(f"  · {p.name} {d}: 图标认出「{v}」，按名字推是「{exp}」"
                             f"（这一行的名字读错了？还是模板不对？）")
    n = len(shots[:len(rows)])
    print(f"回放 {n} 张 x {len(cc.ICON_DIMS)} 维 = {n * len(cc.ICON_DIMS)} 次匹配")
    for d in cc.ICON_DIMS:
        print(f"  {d:<3} 认不出 {unsure[d]} 次" + ("  ✅" if not unsure[d] else "  ⚠️"))
    if clash:
        print(f"\n图标与「按名字推出来的值」对不上 {len(clash)} 处。**这分不清是图标错还是名字错** ——"
              f"本项目里名字本来就有约四成会读错，所以先怀疑名字：")
        print("\n".join(clash))
    print("\n结论：认不出 = 0 且上面没有对不上的，模板才算干净。")
    return 1 if (unsure or clash) else 0


def main():
    ap = argparse.ArgumentParser(description="重建图标模板库")
    ap.add_argument("--shots", default="output/shots", help="信息页截图目录")
    ap.add_argument("--json", default="output/cn_collect.json",
                    help="与截图同一次扫描的 JSON（用于给截图打标签）")
    ap.add_argument("--out", default=cc.ICON_TEMPLATES, help="输出模板图")
    ap.add_argument("--check", action="store_true", help="只回放校验，不重建")
    args = ap.parse_args()
    if args.check:
        return check(args.json)
    build(args.shots, args.json, args.out)
    return 0


if __name__ == "__main__":
    sys.exit(main())
