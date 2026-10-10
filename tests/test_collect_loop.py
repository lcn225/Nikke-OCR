"""离线条烟：把「身份读到 A、装备却点到了 B」灌进真 collect_online，看它是不是当场中止、丢掉那一行。

只桩掉 I/O 与 OCR：屏幕（grab/click）、read_info_page、read_cp、read_equip_page、写盘。
循环控制流、drifted 标志、追加时机都是真代码。
"""
import io
import sys
import contextlib
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import collect_cn as C

C.time.sleep = lambda s: None          # 别真等
C.page_state = lambda img: "info"
C._ensure_info = lambda *a, **k: True
C.wait_page = lambda screen, want, timeout=8.0, interval=0.1: (True, want)
C._crop_ocr = lambda *a, **k: []        # wait_page_turn 的姓名兜底：不给读数

STATE = {"screen_cp": 88763, "saved": []}
C.read_cp = lambda engine, img: STATE["screen_cp"]
C._save_rows = lambda rows, out, extra=None: STATE["saved"].append([r["姓名"] for r in rows])

# 第 1 行米卡（无装备），第 2 行莱伊（读到身份后画面立刻滑到桑迪）；LV.1 用 400 挡掉
INFO = [("米卡：雪地伙伴", 436, 88763), ("莱伊", 400, 78729)]
info_i = [0]


def fake_read_info_page(engine, img):
    r = INFO[min(info_i[0], len(INFO) - 1)]
    info_i[0] += 1
    if info_i[0] >= 2:
        STATE["screen_cp"] = 19394     # ← 身份刚读完，翻页动画的尾巴才把画面带到桑迪
    return r


C.read_info_page = fake_read_info_page
C.read_equip_page = lambda engine, img, slot=None, stat_table=None: (
    None, [None] * 3, [None] * 3, [None] * 3)


class FakeScreen:
    def grab(self):
        return None

    def click(self, xy):
        if xy == C.NEXT_PAGE:
            STATE["screen_cp"] = 78729     # 翻到莱伊


buf = io.StringIO()
with contextlib.redirect_stdout(buf):
    res = C.collect_online(None, FakeScreen(), max_chars=2, with_icons=False)
out = buf.getvalue()

names = [r["姓名"] for r in res["角色"]]
msg_hit = "之前屏幕就换人了" in out
ok = names == ["米卡：雪地伙伴"] and msg_hit

print(f"采集到的行：{names}")
print(f"中止告警出现：{msg_hit}")
print(f"落盘快照（每次追加时的行名）：{STATE['saved']}")
print("\nPASS：第 2 行被整个丢掉，且中止信息打了出来" if ok else "\nFAIL")
sys.exit(0 if ok else 1)
