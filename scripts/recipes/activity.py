# -*- coding: utf-8 -*-
"""活动图骨架：主流程单列下行 + 异常分支右列 + 起止结点。

要点：
  · 菱形与相邻活动之间留 0.45cm 以上，给「是/否」标签腾位置；
  · 右列中心距主列 5cm 以上，回退连线沿右列外侧上行，不横穿主流程；
  · 四周留白 ≥0.5cm（EMF 会裁掉 1.4%）。
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from visio_engine import *


def build(app):
    d = Draw(app, 11.9, 16.9, "活动图示例")
    CX, RX = 3.20, 8.40                     # 主流程 / 右分支中心线
    W, H = 3.40, 0.86
    DW, DH = 4.40, 2.00                     # 菱形：每行 4 汉字
    ACT = dict(fill=C_BLUE_L, line_color=C_BLUE_D, line_w=0.8)
    DEC = dict(line_color=C_ORANGE_D, line_w=0.8)
    A = dict(line_w=0.8, arrow_end=ARROW_FILL)
    LBL = dict(size=7.5, w=0.52, h=0.30)

    def act(cy, text):
        d.rrect(CX, cy, W, H, text, size=fit_size(text, W, 8.6), **ACT)

    d.start_node(CX, 0.90)
    d.line(CX, 1.18, CX, 1.45, **A)
    act(1.88, "选择文稿与续写位置")
    d.line(CX, 2.31, CX, 2.59, **A)
    act(3.02, "设置续写要求")
    d.line(CX, 3.45, CX, 3.73, **A)
    act(4.16, "提交续写请求")

    D1 = 5.86
    d.diamond(CX, D1, DW, DH, "参数校验\n通过？", size=8.4, **DEC)
    d.line(CX, 4.59, CX, D1 - 1.00, **A)

    # 否 → 提示错误 → 回到上一活动（沿右列外侧上行）
    d.line(CX + DW / 2.0, D1, RX - 1.50, D1, **A)
    d.text((CX + DW / 2.0 + RX - 1.50) / 2.0, D1 - 0.26, "否", **LBL)
    d.rrect(RX, D1, 3.00, 0.90, "提示校验错误信息",
            size=fit_size("提示校验错误信息", 3.00, 8.4), **ACT)
    d.poly([(RX, D1 - 0.45), (RX, 3.02), (CX + W / 2.0, 3.02)], **A)

    # 是 → 继续主流程
    d.line(CX, D1 + 1.00, CX, D1 + 1.47, **A)
    d.text(CX + 0.42, D1 + 1.23, "是", **LBL)
    act(7.76, "组装上下文与提示词")
    d.line(CX, 8.19, CX, 8.47, **A)
    act(8.90, "调用大语言模型接口")

    D2 = 10.68
    d.diamond(CX, D2, DW, DH, "调用返回\n成功？", size=8.4, **DEC)
    d.line(CX, 9.33, CX, D2 - 1.00, **A)

    d.line(CX + DW / 2.0, D2, RX - 1.50, D2, **A)
    d.text((CX + DW / 2.0 + RX - 1.50) / 2.0, D2 - 0.26, "否", **LBL)
    d.rrect(RX, D2, 3.00, 0.90, "记录失败并提示",
            size=fit_size("记录失败并提示", 3.00, 8.4), **ACT)
    d.line(RX, D2 + 0.45, RX, 11.74, **A)
    d.end_node(RX, 12.10)
    d.text(RX + 1.70, 12.10, "异常结束", w=1.60, h=0.34, size=8)

    d.line(CX, D2 + 1.00, CX, D2 + 1.47, **A)
    d.text(CX + 0.42, D2 + 1.23, "是", **LBL)
    act(12.58, "流式展示生成结果")
    d.line(CX, 13.01, CX, 13.29, **A)
    act(13.72, "用户确认并采纳")
    d.line(CX, 14.15, CX, 14.43, **A)
    act(14.86, "插入文稿并保存")
    d.line(CX, 15.29, CX, 15.44, **A)
    d.end_node(CX, 15.80)
    d.text(CX + 1.70, 15.80, "正常结束", w=1.60, h=0.34, size=8)
    return d, "activity"


if __name__ == "__main__":
    # 产物统一落在 <当前项目>/图片输出/ ：png + emf + vsdx 源文件
    with Session() as S:
        d, stem = build(S)
        print(S.save(d, stem))
