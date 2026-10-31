# -*- coding: utf-8 -*-
"""网络架构图骨架：使用 Visio 内置网络图元（PC / 防火墙 / 服务器 / 云）。

要点：
  · 模具要从 Office 目录复制到可写目录再 OpenEx；
  · 图元不带自带文字，标签另绘于图元下方 0.42cm；
  · 按包围盒左对齐、等间距；连线取图标垂直中心；
  · NETLOC 的「云」不渲染，用 Stencils.cloud()（修订云形）。
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from visio_engine import *


def place_row(st, d, nodes, cy, h, gap, x0, max_w):
    placed = []
    for nd in nodes:
        if nd.get("cloud"):
            s = st.cloud(d, x0, cy, nd.get("w", 2.45), nd.get("h", 1.60), "")
        else:
            s = st.drop(d, nd["kind"], nd["master"], x0, cy, None, h_cm=h)
            bb = Stencils.bbox_cm(d, s)
            if bb[2] - bb[0] > max_w:          # 过宽图元等比收窄
                k = max_w / (bb[2] - bb[0])
                s.Cells("Width").FormulaU = "%.4f in" % (s.Cells("Width").ResultIU * k)
                s.Cells("Height").FormulaU = "%.4f in" % (s.Cells("Height").ResultIU * k)
                s.Cells("PinX").FormulaU = "%.4f in" % d.X(x0)
                s.Cells("PinY").FormulaU = "%.4f in" % d.Y(cy)
        placed.append([s, Stencils.bbox_cm(d, s), nd["label"]])
    x = x0
    for it in placed:
        dx = x - it[1][0]
        it[0].Cells("PinX").FormulaU = "%.5f in" % (it[0].Cells("PinX").ResultIU + dx / CM)
        it[1] = Stencils.bbox_cm(d, it[0])
        x = it[1][2] + gap
    for it in placed:
        bb = it[1]
        d.text((bb[0] + bb[2]) / 2.0, bb[3] + 0.42, it[2], w=2.10, h=0.30, size=8)
    return [(it[0], it[1], it[2]) for it in placed]


def build(app):
    st = Stencils(app)
    d = Draw(app, 15.4, 8.4, "网络架构示例")
    CY, H = 2.70, 2.00

    d.rrect(1.80, 2.98, 2.45, 4.55, fill=C_GREY_L, line_color=C_GREY_D,
            line_w=0.7, dash=LINE_PAT_DASH, radius=0.10)
    d.text(1.80, 0.52, "用户终端", w=2.3, h=0.36, size=9)
    for cy, master, label in ((1.78, "PC", "桌面浏览器"),
                              (3.85, "笔记本电脑", "移动端浏览器")):
        s = st.drop(d, "pc", master, 1.80, cy, None, h_cm=1.58)
        bb = Stencils.bbox_cm(d, s)
        d.text(1.80, bb[3] + 0.34, label, w=2.0, h=0.30, size=8)

    nodes = [
        dict(cloud=True, label="Internet", w=2.45, h=1.60),
        dict(kind="per", master="防火墙", label="防火墙"),
        dict(kind="srv", master="代理服务器", label="Nginx 代理"),
        dict(kind="srv", master="应用程序服务器", label="应用服务器"),
        dict(cloud=True, label="大模型平台", w=2.45, h=1.60),
    ]
    row = place_row(st, d, nodes, CY, H, 0.80, 3.28, 2.30)

    LY = (row[0][1][1] + row[0][1][3]) / 2.0
    for i in range(len(row) - 1):
        d.line(row[i][1][2], LY, row[i + 1][1][0], LY,
               line_w=0.9, arrow_end=ARROW_FILL)
    for i, t in enumerate(("HTTPS", "过滤", "HTTP", "API")):
        d.text((row[i][1][2] + row[i + 1][1][0]) / 2.0, LY - 0.27, t,
               w=0.80, h=0.30, size=7)

    app_bb = row[3][1]
    app_cx = (app_bb[0] + app_bb[2]) / 2.0
    dbs = [dict(kind="srv", master="数据库服务器", label="MySQL 8.0"),
           dict(kind="srv", master="服务器", label="Redis 7.0"),
           dict(kind="srv", master="文件服务器", label="MinIO")]
    drow = place_row(st, d, dbs, 6.10, 1.80, 0.95, 8.60, 2.30)

    bus_y = 5.00
    d.line(app_cx, app_bb[3] + 0.62, app_cx, bus_y, line_w=0.9)
    d.line(8.60, bus_y, drow[-1][1][2], bus_y, line_w=0.9)
    for _, bb, _l in drow:
        cx = (bb[0] + bb[2]) / 2.0
        d.line(cx, bus_y, cx, bb[1], line_w=0.9, arrow_end=ARROW_FILL)
    d.text(7.35, bus_y, "数据服务器组", w=2.00, h=0.38, size=8.5)
    st.close()
    return d, "network"


if __name__ == "__main__":
    # 产物统一落在 <当前项目>/图片输出/ ：png + emf + vsdx 源文件
    with Session() as S:
        d, stem = build(S)
        print(S.save(d, stem))
