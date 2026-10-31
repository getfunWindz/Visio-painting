# -*- coding: utf-8 -*-
"""陈氏 E-R 图骨架：实体属性图 + 局部 E-R 图。

要点：
  · 连线必须画到图形边界（rect / ellipse / diamond 三种求交），否则线穿字；
  · 属性单行横向排开、等宽，实体居中下方；
  · 主键属性名加下划线（Char.Style = 4）。
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from visio_engine import *

ENT = dict(fill=C_BLUE_L, line_color=C_BLUE_D, line_w=1.0)
ATT = dict(fill=WHITE, line_color=C_GREEN_D, line_w=0.9)
REL = dict(line_color=C_ORANGE_D, line_w=0.9)


def _edge(node, tx, ty):
    cx, cy, w, h, kind = node[0], node[1], node[2], node[3], node[5]
    dx, dy = tx - cx, ty - cy
    if abs(dx) < 1e-9 and abs(dy) < 1e-9:
        return cx, cy
    if kind == "ent":
        t = min((w / 2.0) / abs(dx) if abs(dx) > 1e-9 else 1e9,
                (h / 2.0) / abs(dy) if abs(dy) > 1e-9 else 1e9)
    elif kind == "att":
        t = 1.0 / (((dx / (w / 2.0)) ** 2 + (dy / (h / 2.0)) ** 2) ** 0.5)
    else:
        t = 1.0 / (abs(dx) / (w / 2.0) + abs(dy) / (h / 2.0))
    return cx + dx * t, cy + dy * t


def draw_nodes(d, nodes, links, cards=None):
    for a, b in links:
        p1 = _edge(nodes[a], nodes[b][0], nodes[b][1])
        p2 = _edge(nodes[b], nodes[a][0], nodes[a][1])
        d.line(p1[0], p1[1], p2[0], p2[1], line_w=0.8)
    for v in nodes.values():
        cx, cy, w, h, text, kind = v[0], v[1], v[2], v[3], v[4], v[5]
        pk = v[6] if len(v) > 6 else False
        if kind == "ent":
            d.rect(cx, cy, w, h, text, size=fit_size(text, w, 9.2), **ENT)
        elif kind == "att":
            d.oval(cx, cy, w, h, text, size=fit_size(text, w, 8.5),
                   align_underline=pk, **ATT)
        else:
            d.diamond(cx, cy, w, h, text, size=8.5, **REL)
    for a, b, txt in (cards or []):
        na, nb = nodes[a], nodes[b]
        px, py = _edge(na, nb[0], nb[1])
        dx, dy = nb[0] - na[0], nb[1] - na[1]
        L = max((dx * dx + dy * dy) ** 0.5, 1e-6)
        ux, uy = dx / L, dy / L
        nx, ny = -uy, ux
        if ny > 0:
            nx, ny = -nx, -ny
        d.text(px + ux * 0.45 + nx * 0.28, py + uy * 0.45 + ny * 0.28, txt,
               w=0.62, h=0.40, size=8.5)


def attrs_diagram(app, w_cm, entity, attributes, name):
    n = len(attributes)
    gap = 0.10
    aw = min((w_cm - 0.40 - gap * (n - 1)) / n, 1.95)
    total = aw * n + gap * (n - 1)
    x0 = (w_cm - total) / 2.0
    xs = [x0 + aw / 2.0 + i * (aw + gap) for i in range(n)]
    d = Draw(app, w_cm, 4.6, name)
    ay, ey = 1.30, 3.55
    for x in xs:                                  # 先连线，再画图形，接点干净
        d.line(x, ay, w_cm / 2.0, ey, line_w=0.8)
    for x, (t, pk) in zip(xs, attributes):
        d.oval(x, ay, aw, 0.92, t, size=fit_size(t, aw, 8.5),
               align_underline=pk, **ATT)
    d.rect(w_cm / 2.0, ey, 2.80, 1.00, entity, size=10.0, **ENT)
    return d


def local_er(app):
    d = Draw(app, 14.4, 3.0, "局部E-R图示例")
    cy = 1.50
    N = {
        "u": (1.50, cy, 2.20, 0.98, "用户", "ent"),
        "c": (4.60, cy, 1.80, 1.20, "创作", "rel"),
        "w": (7.70, cy, 2.20, 0.98, "文稿", "ent"),
        "s": (10.75, cy, 1.75, 1.16, "适用", "rel"),
        "g": (13.25, cy, 2.10, 0.98, "文体类型", "ent"),
    }
    draw_nodes(d, N, [("u", "c"), ("c", "w"), ("w", "s"), ("s", "g")],
               cards=[("u", "c", "1"), ("w", "c", "n"),
                      ("w", "s", "n"), ("g", "s", "1")])
    return d


if __name__ == "__main__":
    with Session() as S:
        d1 = attrs_diagram(S, 14.5, "用户",
                           [("用户编号", True), ("用户名", False), ("口令散列", False),
                            ("昵称", False), ("电子邮箱", False), ("注册时间", False),
                            ("账户状态", False)], "实体属性图示例")
        print(S.save(d1, "er_attrs"))
        d2 = local_er(S)
        print(S.save(d2, "er_local"))
