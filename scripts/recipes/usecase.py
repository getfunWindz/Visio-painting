# -*- coding: utf-8 -*-
"""用例图骨架：参与者 + 系统边界 + 用例列（居中）+ «include» 关系。

要点：
  · 用例列中心必须等于系统边界中心，否则右侧会留大片空白；
  · 关联线从参与者躯干中线引出（不是从图形旁的空白点或手臂端点引出）；
  · 用例一律单列：双列时右列用例无处置线，极易漏连（本项目就漏了“文体模板应用”）；
  · 同侧有两个参与者时，两角色纵向间距须 ≥3.9cm（各自身高约 3.5cm），否则头脚相撞。
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from visio_engine import *


def build(app):
    d = Draw(app, 14.1, 9.6, "用例图示例")
    BX, BW = 7.55, 5.45                     # 边界中心与宽度
    # 页面高度要同时容纳：上留白 0.35 + 标题字高 0.335 + 间距 0.05 + 边框 + 下留白 0.35。
    # 边框纵向 0.92–8.92（高 8.0），标题墨迹 0.45–0.79，与边框上沿留 0.13cm。
    d.rect(BX, 4.92, BW, 8.00, fill=None, line_w=1.0)
    d.text(BX, 0.62, "智能创作", w=BW, h=0.5, size=9.5)

    names = ["文章智能续写", "全文润色优化", "文本内容提炼", "内容精简改写", "创作大纲生成"]
    UC = dict(fill=C_BLUE_L, line_color=C_BLUE_D, line_w=0.9, size=9)
    # 用例列在边框内上下居中：(8.00 - (4*1.50 + 1.02)) / 2 = 0.49 → 首行中心 1.92
    P = {n: (BX, 1.92 + i * 1.50) for i, n in enumerate(names)}   # 列中心 = 边界中心
    for n, (cx, cy) in P.items():
        d.oval(cx, cy, 2.75, 1.02, n, **UC)

    # 关联线必须从参与者躯干中线（cx, cy）引出，不能从图形外侧的空白点
    # 或手臂端点引出；先画线后画角色，线头会被躯干压住，接点更干净。
    UA, AC = (1.90, 4.92), (12.55, 4.92)
    for n in names:
        d.line(UA[0], UA[1], BX - 1.375, P[n][1], line_w=0.8)
    d.line(AC[0], AC[1], BX + 1.375, P["文本内容提炼"][1], line_w=0.8)
    d.actor(UA[0], UA[1], "普通用户", label_w=3.2)
    d.actor(AC[0], AC[1], "大语言模型服务", label_w=3.0)
    return d, "usecase"


if __name__ == "__main__":
    # 产物统一落在 <当前项目>/图片输出/ ：png + emf + vsdx 源文件
    with Session() as S:
        d, stem = build(S)
        print(S.save(d, stem))
