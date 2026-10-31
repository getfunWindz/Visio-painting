# -*- coding: utf-8 -*-
"""分层架构图与模块结构图骨架。

分层架构要点：层名带宽度 ≥1.85cm 才能放下 5 个汉字；组件宽度按可用宽度均分，
字号用 fit_size 反算，避免六字名称折行。
模块结构要点：只做两级（系统→子系统），三级交给正文叙述，否则功能点框必然过窄。
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from visio_engine import *


def layered(app):
    d = Draw(app, 15.4, 9.7, "分层架构示例")
    MX0, MX1 = 0.45, 13.30
    rows = [
        ("表现层", 1.38, ["Web 创作工作台", "移动端 H5 页面", "平台管理后台"], C_BLUE_L),
        ("应用层", 1.38, ["接口网关", "身份认证与鉴权", "参数校验与限流"], WHITE),
        ("业务层", 1.62, ["用户服务", "文稿服务", "智能创作服务", "文体模板服务",
                          "平台管理服务"], C_BLUE_L),
        ("智能能力层", 1.40, ["提示词模板", "上下文构建", "模型路由与降级",
                              "结果规范化", "任务留痕"], WHITE),
        ("数据层", 1.38, ["MySQL 业务数据库", "Redis 会话缓存", "对象存储"], C_BLUE_L),
        ("基础设施层", 1.30, ["Linux 服务器", "JDK 17 运行时", "Nginx 与 Docker",
                              "日志与监控"], WHITE),
    ]
    y = 0.55
    for name, h, items, fill in rows:
        d.rrect((MX0 + MX1) / 2.0, y + h / 2.0, MX1 - MX0, h, fill=C_GREY_L,
                line_color=C_GREY_D, line_w=0.6, dash=LINE_PAT_DASH, radius=0.10)
        d.rrect(1.42, y + h / 2.0, 1.85, h - 0.18, name, fill=C_BLUE_M,
                line_color=C_BLUE_D, line_w=0.7, radius=0.08, size=9.0)
        n = len(items)
        gap = 0.14
        bw = ((MX1 - 0.10) - 2.42 - gap * (n - 1)) / n
        for i, it in enumerate(items):
            d.rect(2.42 + bw / 2.0 + i * (bw + gap), y + h / 2.0,
                   bw, min(0.86, h - 0.44), it, fill=fill,
                   line_color=BLACK, line_w=0.7, size=fit_size(it, bw, 9.0))
        y += h + 0.14
    d.rrect(14.14, (0.55 + y - 0.14) / 2.0, 1.16, y - 0.14 - 0.55,
            "安\n全\n与\n运\n维\n保\n障\n体\n系", fill=C_GREEN_L, line_color=C_GREEN_D,
            line_w=0.8, radius=0.10, size=8.5)
    return d


def structure(app):
    """两级模块结构：根结点 + 一行子系统 + 正交连线。"""
    d = Draw(app, 15.4, 4.3, "模块结构示例")
    subs = ["用户服务\n子系统", "文稿管理\n子系统", "智能创作\n子系统",
            "文体模板\n子系统", "内容治理\n子系统", "系统管理\n子系统"]
    n, gap = len(subs), 0.16
    bw = (15.4 - 0.40 - gap * (n - 1)) / n
    d.rect(15.4 / 2.0, 0.62, min(15.4 - 1.2, bw * 2.2), 0.88,
           "面向多文体的智能文章辅助创作平台",
           fill=C_BLUE_M, line_color=C_BLUE_D, line_w=1.1, size=10.5)
    xs = [0.20 + bw / 2.0 + i * (bw + gap) for i in range(n)]
    cy = 2.85
    for x, s in zip(xs, subs):
        d.rect(x, cy, bw, 1.45, s, fill=C_BLUE_L, line_color=C_BLUE_D,
               line_w=0.9, size=9.2)
    bus = 1.06 + (cy - 0.725 - 1.06) * 0.55
    d.line(15.4 / 2.0, 1.06, 15.4 / 2.0, bus, line_w=0.8)
    d.line(min(xs + [15.4 / 2.0]), bus, max(xs + [15.4 / 2.0]), bus, line_w=0.8)
    for x in xs:
        d.line(x, bus, x, cy - 0.725, line_w=0.8, arrow_end=ARROW_FILL)
    return d


if __name__ == "__main__":
    with Session() as S:
        print(S.save(layered(S), "layered"))
        print(S.save(structure(S), "structure"))
