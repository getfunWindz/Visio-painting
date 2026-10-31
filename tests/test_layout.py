# -*- coding: utf-8 -*-
"""布局回归测试：不需要 Visio，用一个记录型 Draw 替身截获配方里的每次绘图调用，
然后对几何关系做断言。

覆盖的已知缺陷：
  · 系统边界标题压在边框线上（SKILL.md 要求「置于框上方，不压线」）
  · 元素贴边（EMF 导出会裁掉四周 1.4%，要求四周留白 ≥0.35cm）

运行：python tests/test_layout.py      （也可用 pytest 收集）
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "scripts"))
sys.path.insert(0, os.path.join(HERE, "..", "scripts", "recipes"))

# 本测试只检查几何关系，不启动 Visio；非 Windows 上桩掉 win32com，方便在任意平台跑 pytest。
try:
    import win32com.client  # noqa: F401
except ImportError:
    import types
    _pkg = types.ModuleType("win32com")
    _client = types.ModuleType("win32com.client")
    _client.Dispatch = lambda *a, **kw: None
    _pkg.client = _client
    sys.modules["win32com"] = _pkg
    sys.modules["win32com.client"] = _client

PT2CM = 2.54 / 72.0        # 1pt = 0.03528cm，与 visio_engine.fit_size 的换算一致
MARGIN = 0.35              # 页面四周最小留白（cm）


class Shape(object):
    __slots__ = ("kind", "cx", "cy", "w", "h", "text", "size")

    def __init__(self, kind, cx, cy, w, h, text="", size=9.0):
        self.kind, self.cx, self.cy = kind, cx, cy
        self.w, self.h, self.text, self.size = w, h, text, size

    def top(self):
        return self.cy - self.h / 2.0

    def bottom(self):
        return self.cy + self.h / 2.0

    def ink_top(self):
        """文字实际墨迹范围（字形高度按字号算，不用文本框高）。"""
        return self.cy - self.size * PT2CM / 2.0

    def ink_bottom(self):
        return self.cy + self.size * PT2CM / 2.0


class Recorder(object):
    """替代 visio_engine.Draw，仅记录几何信息。"""

    def __init__(self, app, width_cm, height_cm, name=""):
        self.w, self.h = float(width_cm), float(height_cm)
        self.shapes = []
        self.lines = []

    def _add(self, kind, cx, cy, w, h, text="", **kw):
        s = Shape(kind, cx, cy, w, h, text, float(kw.get("size", 9.0)))
        self.shapes.append(s)
        return s

    def rect(self, cx, cy, w, h, text="", **kw):
        return self._add("rect", cx, cy, w, h, text, **kw)

    def rrect(self, cx, cy, w, h, text="", **kw):
        return self._add("rect", cx, cy, w, h, text, **kw)

    def oval(self, cx, cy, w, h, text="", **kw):
        return self._add("oval", cx, cy, w, h, text, **kw)

    def diamond(self, cx, cy, w, h, text="", **kw):
        return self._add("diamond", cx, cy, w, h, text, **kw)

    def text(self, cx, cy, s, w=6.0, h=0.6, **kw):
        return self._add("text", cx, cy, w, h, s, **kw)

    def line(self, x1, y1, x2, y2, **kw):
        self.lines.append((x1, y1, x2, y2))

    def actor(self, cx, cy, label, scale=1.0, size=9, label_w=4.2):
        self._add("actor", cx, cy, 0.0, 0.0, label, size=float(size))


def _build(module):
    module.Draw = Recorder
    d, stem = module.build(None)
    return d


def _border(d):
    rects = [s for s in d.shapes if s.kind == "rect" and not s.text]
    assert rects, "没有找到系统边界矩形"
    return max(rects, key=lambda s: s.w * s.h)


def _check_usecase(d):
    """用例图的通用几何断言：标题不压线、四周留白、用例全在边界内且居中。"""
    border = _border(d)
    title = [s for s in d.shapes if s.kind == "text" and s.text == "智能创作"][0]
    gap = border.top() - title.ink_bottom()
    assert gap >= 0.05, (
        "系统边界标题压线：标题墨迹底 %.2fcm，边框上沿 %.2fcm，间距 %.2fcm（要求 ≥0.05）"
        % (title.ink_bottom(), border.top(), gap))

    for s in d.shapes:
        if s.kind in ("actor", "text"):
            top, bottom = s.ink_top(), s.ink_bottom()
        else:
            top, bottom = s.top(), s.bottom()
        left, right = s.cx - s.w / 2, s.cx + s.w / 2
        assert top >= MARGIN - 1e-6, "%s 顶部留白不足：%.2fcm" % (s.kind + s.text, top)
        assert d.h - bottom >= MARGIN - 1e-6, "%s 底部留白不足：%.2fcm" % (s.kind + s.text, d.h - bottom)
        assert left >= MARGIN - 1e-6, "%s 左侧留白不足：%.2fcm" % (s.kind + s.text, left)
        assert d.w - right >= MARGIN - 1e-6, "%s 右侧留白不足：%.2fcm" % (s.kind + s.text, d.w - right)

    ovals = [s for s in d.shapes if s.kind == "oval"]
    assert len(ovals) == 5, "用例数量应为 5，实际 %d" % len(ovals)
    for s in ovals:
        assert s.top() > border.top(), "%s 越过边界上沿" % s.text
        assert s.bottom() < border.bottom(), "%s 越过边界下沿" % s.text
        assert s.cx == border.cx, "%s 未与边界水平居中" % s.text


def test_boundary_title_not_on_border_line():
    import usecase
    _check_usecase(_build(usecase))


def test_page_margins():
    import usecase
    d = _build(usecase)
    _check_usecase(d)


def test_usecases_inside_boundary():
    import usecase
    _check_usecase(_build(usecase))


class FakeSession(object):
    """供 examples/paper_figure_set.py 使用：它有 S.draw(...) 入口。"""

    def draw(self, width_cm, height_cm, name=""):
        return Recorder(None, width_cm, height_cm, name)


def test_example_usecase_layout():
    """examples/paper_figure_set.py 里的用例图是同一套布局，必须一起受保护。"""
    sys.path.insert(0, os.path.join(HERE, "..", "examples"))
    import paper_figure_set as ex
    _check_usecase(ex.fig_usecase(FakeSession()))


if __name__ == "__main__":
    failed = 0
    for name, fn in sorted(globals().items()):
        if name.startswith("test_") and callable(fn):
            try:
                fn()
                print("PASS  %s" % name)
            except AssertionError as e:
                failed += 1
                print("FAIL  %s\n      %s" % (name, e))
    sys.exit(1 if failed else 0)
