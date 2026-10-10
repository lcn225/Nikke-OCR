"""离线条烟：翻页判定 + 换人复核（莱伊·甲=桑迪·头 那个 bug 的回归测试）。

不需要 OCR、不需要游戏：screen 按帧喂脚本化的「战力 / 名字」读数，read_cp 与 _crop_ocr
都换成桩。唯一照搬实测事实的是 page_state —— 滑页时面板外亮度不变，它**全程判 'info'**，
所以直接 patch 掉（这正是原 bug 的前提之一）。
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import collect_cn as C

C.page_state = lambda img: "info"
C.read_cp = lambda engine, img: img.cp
C._crop_ocr = lambda engine, img, region, scale=2: [img.name] if img.name else []


class FakeImg:
    def __init__(self, cp, name=None):
        self.cp, self.name = cp, name


class FakeScreen:
    """按帧喂数据；重复最后一帧。i 记录吃掉了多少帧。"""

    def __init__(self, frames):
        self.frames, self.i = [FakeImg(c, n) for c, n in frames], 0

    def grab(self):
        f = self.frames[min(self.i, len(self.frames) - 1)]
        self.i += 1
        return f


OK = True


def check(label, got, want):
    global OK
    good = got == want
    OK &= good
    print(f"{'PASS' if good else 'FAIL'}  {label}: got={got!r} want={want!r}")


# ---------------------------------------------------------------- wait_page_turn
# 场景 1（原 bug 现场）：停米卡，prev_name 读失败过（None），滑动中出现一帧混叠读数。
#   旧逻辑会在第 2 帧（名字兜底）或第 3 帧（混叠战力）就放行 → 后续整行错位。
frames = [
    (88763, "米卡：雪地伙伴"),   # 1 还没开始动
    (88763, "米卡：雪地伙伴"),   # 2 旧逻辑的名字兜底在这里就放行了
    (1, "米卡：雪地伙伴"),        # 3 滑动中的混叠读数，旧逻辑在这里放行
    (78729, "莱伊"),             # 4 新页读到了
    (78729, "莱伊"),             # 5 连续第二帧 → 才算翻页完成
]
sc = FakeScreen(frames)
check("场景1 中途不误判、连续两帧才放行",
      (C.wait_page_turn(sc, None, 88763, None, timeout=15.0), sc.i), (True, 5))

# 对照：把旧逻辑照抄一遍，确认它真的会在第 2 帧放行（证明这测试抓的是真行为）
def old_wait(frames):
    sc, last_name = FakeScreen(frames), None
    for _ in range(50):
        img = sc.grab()
        cp = img.cp
        if cp is not None and cp != 88763:
            return sc.i
        nm = img.name
        if nm is not None and nm != None and nm == last_name:  # noqa: E711  (prev_name=None)
            return sc.i
        last_name = nm
    return None


check("场景1 对照：旧逻辑第 2 帧就放行", old_wait(frames), 2)

# 场景 2：正常翻页（旧页一帧、新页两帧）→ 只多等一帧
sc = FakeScreen([(88763, "米卡"), (78729, "莱伊"), (78729, "莱伊")])
check("场景2 正常翻页 3 帧完成", (C.wait_page_turn(sc, None, 88763, "米卡"), sc.i), (True, 3))

# 场景 3：一直没翻（战力一直等于 prev_cp）→ 超时返回 False
sc = FakeScreen([(88763, "米卡")])
check("场景3 页面没翻 -> 超时 False",
      (C.wait_page_turn(sc, None, 88763, "米卡", timeout=0.7), sc.i > 1), (False, True))

# ------------------------------------------------------------ _confirm_same_char
sc = FakeScreen([(19394, "桑迪")])
check("复核1 屏幕已是下一个角色 -> 判负并交出读到的战力",
      C._confirm_same_char(sc, None, 78729), (False, 19394))

sc = FakeScreen([(None, None), (78729, "莱伊")])
check("复核2 偶发读空不该冤枉 -> 重试读回 cp", C._confirm_same_char(sc, None, 78729), (True, 78729))

sc = FakeScreen([(78729, "莱伊")])
check("复核3 命中即放行（只花一帧）", (C._confirm_same_char(sc, None, 78729), sc.i), ((True, 78729), 1))

sc = FakeScreen([(None, None)])
check("复核4 一直读不出 -> 判负但不谎报战力",
      C._confirm_same_char(sc, None, 78729, tries=2), (False, None))

print("\n全部通过" if OK else "\n有 FAIL")
sys.exit(0 if OK else 1)
