# -*- coding: utf-8 -*-
"""示例：一次会话产出「一整册」论文插图。

用法：
    cd <你的论文项目目录>
    python <skill>/examples/paper_figure_set.py

产物（全部落在 <当前项目>/图片输出/，一图一档）：
    fig3-1_系统用例图.png    + fig3-1_系统用例图.vsdx
    fig4-1_逻辑架构图.png    + fig4-1_逻辑架构图.vsdx
    fig4-15_局部E-R图.png    + fig4-15_局部E-R图.vsdx
另外多出 .emf（本示例开启了 emf，用于插 Word）；不需要就删掉 fmts 参数。
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, "..", "scripts"))

from visio_engine import *   # noqa: E402


# ------------------------------------------------------------------ 图1：用例图
def fig_usecase(S):
    d = S.draw(14.1, 9.2, "图3-1 系统用例图")
    BX, BW = 7.55, 5.45
    d.rect(BX, 4.55, BW, 7.78, fill=None, line_w=1.0)
    d.text(BX, 0.62, "智能创作", w=BW, h=0.5, size=9.5)

    names = ["文章智能续写", "全文润色优化", "文本内容提炼",
             "内容精简改写", "创作大纲生成"]
    P = {n: (BX, 1.55 + i * 1.50) for i, n in enumerate(names)}
    for n, (cx, cy) in P.items():
        d.oval(cx, cy, 2.75, 1.02, n, fill=C_BLUE_L, line_color=C_BLUE_D,
               line_w=0.9, size=9)

    UA, AC = (1.90, 4.55), (12.55, 4.55)
    for n in names:                                  # 从躯干中线引出关联线
        d.line(UA[0], UA[1], BX - 1.375, P[n][1], line_w=0.8)
    d.line(AC[0], AC[1], BX + 1.375, P["文本内容提炼"][1], line_w=0.8)

    d.actor(UA[0], UA[1], "普通用户", label_w=3.2)
    d.actor(AC[0], AC[1], "大语言模型服务", label_w=3.0)
    return d


# ------------------------------------------------------------------ 图2：分层架构
def fig_layered(S):
    d = S.draw(15.4, 7.2, "图4-1 逻辑架构图")
    rows = [
        ("表现层", ["Web 创作工作台", "移动端 H5 页面", "平台管理后台"]),
        ("应用层", ["接口网关", "身份认证与鉴权", "参数校验与限流"]),
        ("业务层", ["用户服务", "文稿服务", "智能创作服务", "文体模板服务"]),
        ("数据层", ["MySQL 业务数据库", "Redis 会话缓存", "对象存储"]),
    ]
    y, h = 0.55, 1.35
    for name, items in rows:
        d.rrect(8.0, y + h / 2, 14.6, h, fill=C_GREY_L, line_color=C_GREY_D,
                line_w=0.6, dash=LINE_PAT_DASH, radius=0.10)
        d.rrect(1.42, y + h / 2, 1.85, h - 0.18, name, fill=C_BLUE_M,
                line_color=C_BLUE_D, line_w=0.7, radius=0.08, size=9)
        n, gap = len(items), 0.14
        bw = ((14.2 - 2.42) - gap * (n - 1)) / n
        for i, it in enumerate(items):
            d.rect(2.42 + bw / 2 + i * (bw + gap), y + h / 2, bw, h - 0.44, it,
                   fill=WHITE if n > 3 else C_BLUE_L, line_color=BLACK,
                   line_w=0.7, size=fit_size(it, bw, 9))
        y += h + 0.14
    return d


# ------------------------------------------------------------------ 图3：局部 E-R
def fig_er(S):
    d = S.draw(14.4, 3.0, "图4-15 局部E-R图")

    def edge(node, tx, ty):
        cx, cy, w, h, kind = node[0], node[1], node[2], node[3], node[5]
        dx, dy = tx - cx, ty - cy
        if kind == "ent":
            t = min((w / 2) / abs(dx) if dx else 1e9, (h / 2) / abs(dy) if dy else 1e9)
        else:
            t = 1.0 / (abs(dx) / (w / 2) + abs(dy) / (h / 2))
        return cx + dx * t, cy + dy * t

    cy = 1.50
    N = {
        "u": (1.50, cy, 2.20, 0.98, "用户", "ent"),
        "c": (4.60, cy, 1.80, 1.20, "创作", "rel"),
        "w": (7.70, cy, 2.20, 0.98, "文稿", "ent"),
    }
    for a, b in (("u", "c"), ("c", "w")):
        p1 = edge(N[a], N[b][0], N[b][1])
        p2 = edge(N[b], N[a][0], N[a][1])
        d.line(p1[0], p1[1], p2[0], p2[1], line_w=0.8)
    d.rect(*N["u"][:4], N["u"][4], fill=C_BLUE_L, line_color=C_BLUE_D, line_w=1.0, size=9.5)
    d.diamond(*N["c"][:4], N["c"][4], line_color=C_ORANGE_D, line_w=0.9, size=8.5)
    d.rect(*N["w"][:4], N["w"][4], fill=C_BLUE_L, line_color=C_BLUE_D, line_w=1.0, size=9.5)
    d.text(2.95, cy - 0.42, "1", w=0.6, h=0.4, size=8.5)
    d.text(3.75, cy - 0.42, "n", w=0.6, h=0.4, size=8.5)
    return d


# ------------------------------------------------------------------ 主流程
def main():
    with Session(fmts=("png", "emf")) as S:   # 需要插 Word 才加 emf
        print(S.save(fig_usecase(S), "fig3-1_系统用例图"))
        print(S.save(fig_layered(S), "fig4-1_逻辑架构图"))
        print(S.save(fig_er(S), "fig4-15_局部E-R图"))
    out = output_dir()
    print("\n输出目录：", out)
    for f in sorted(os.listdir(out)):
        print("  %-34s %8d B" % (f, os.path.getsize(os.path.join(out, f))))


if __name__ == "__main__":
    main()
