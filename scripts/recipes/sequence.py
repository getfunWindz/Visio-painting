# -*- coding: utf-8 -*-
"""时序图骨架：对象头 + 生命线 + 激活条 + 同步/返回消息 + 自消息。

要点：
  · 第一条消息 y 起点 = 对象头底部 + 0.55cm，否则标签压住头框；
  · 生命线间距 ≥2.6cm；
  · 返回消息用虚线 + 开口箭头，同步消息用实线 + 实心箭头。
"""
import os, sys
sys.path.insert(0, os.path.join(os.path.dirname(__file__), ".."))
from visio_engine import *


class Seq:
    def __init__(self, app, w, h, name, xs, names, widths=None):
        self.d = Draw(app, w, h, name)
        self.xs, self.top, self.head_h = xs, 0.75, 0.95
        self.bot = h - 0.55
        for x, nm, bw in zip(xs, names, widths or [2.15] * len(xs)):
            self.d.rect(x, self.top + self.head_h / 2.0, bw, self.head_h, nm,
                        fill=C_BLUE_L, line_color=C_BLUE_D, line_w=0.9,
                        size=fit_size(nm, bw, 9.0))
            self.d.line(x, self.top + self.head_h, x, self.bot,
                        line_w=0.7, dash=LINE_PAT_DASH)

    def bar(self, i, y0, y1):
        self.d.rect(self.xs[i], (y0 + y1) / 2.0, 0.26, abs(y1 - y0),
                    fill=C_GREY_M, line_color=C_GREY_D, line_w=0.6)

    def msg(self, i, j, y, text, ret=False, num=None):
        x1, x2 = self.xs[i], self.xs[j]
        self.d.line(x1, y, x2, y, line_w=0.8,
                    arrow_end=ARROW_OPEN if ret else ARROW_FILL,
                    dash=LINE_PAT_DASH if ret else LINE_PAT_SOLID)
        lbl = ("%s. %s" % (num, text)) if num else text
        bw = max(abs(x2 - x1) - 0.05, 1.2)
        self.d.text((x1 + x2) / 2.0, y - 0.30, lbl, w=bw, h=0.45,
                    size=fit_size(lbl, bw, 8.0))

    def selfmsg(self, i, y, text, num=None):
        x = self.xs[i]
        self.d.poly([(x, y), (x + 0.90, y), (x + 0.90, y + 0.56), (x, y + 0.56)],
                    line_w=0.8, arrow_end=ARROW_OPEN)
        lbl = ("%s. %s" % (num, text)) if num else text
        self.d.text(x + 1.62, y - 0.26, lbl, w=2.75, h=0.46,
                    size=fit_size(lbl, 2.75, 8.0))


def build(app):
    xs = [1.55, 4.10, 7.10, 10.10, 13.15]
    s = Seq(app, 15.0, 10.0, "时序图示例", xs,
            ["用户", "登录页面", "用户服务", "Redis 缓存", "MySQL 数据库"],
            [2.0, 2.3, 2.2, 2.4, 2.6])
    Y = [2.32, 3.05, 3.80, 4.55, 5.30, 6.32, 7.08, 7.84, 8.60]
    s.bar(1, 2.22, 3.05)
    s.bar(2, 3.05, 7.84)
    s.bar(4, 3.80, 4.55)
    s.bar(3, 6.32, 7.08)

    s.msg(0, 1, Y[0], "输入账号与口令", num=1)
    s.msg(1, 2, Y[1], "提交登录请求", num=2)
    s.msg(2, 4, Y[2], "查询账户与口令散列", num=3)
    s.msg(4, 2, Y[3], "返回账户记录", ret=True, num=4)
    s.selfmsg(2, Y[4], "校验口令与状态", num=5)
    s.msg(2, 3, Y[5], "写入会话令牌", num=6)
    s.msg(3, 2, Y[6], "写入成功", ret=True, num=7)
    s.msg(2, 1, Y[7], "返回令牌与用户信息", ret=True, num=8)
    s.msg(1, 0, Y[8], "进入创作工作台", ret=True, num=9)
    return s.d, "sequence"


if __name__ == "__main__":
    # 产物统一落在 <当前项目>/图片输出/ ：png + emf + vsdx 源文件
    with Session() as S:
        d, stem = build(S)
        print(S.save(d, stem))
