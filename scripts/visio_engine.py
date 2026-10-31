# -*- coding: utf-8 -*-
"""
Visio 软件工程绘图引擎
------------------------------------------------------------------
面向本科毕业设计（软件工程类）插图的 Visio COM 绘图封装。
所有坐标以厘米为单位、以左上角为原点，导出时按 1:1 物理尺寸生成，
保证插入 Word 后图内文字字号即为设定字号。

支持的图形语义（严格遵循软件工程制图规范）：
  · UML 用例图      —— 参与者（火柴人）、用例（椭圆）、系统边界、
                       关联（实线无箭头）、«include»/«extend»（虚线开箭头）、泛化（空心三角）
  · 陈氏 E-R 图     —— 实体（矩形）、属性（椭圆）、联系（菱形）、
                       主键（属性名下加下划线）、基数标注（1/n/m）
  · 分层架构图      —— 层容器 + 层带 + 组件框
  · 模块结构图      —— 正交树形结构 + 实心箭头
  · UML 活动图      —— 初始结点（实心圆）、活动（圆角矩形）、判定（菱形）、终止结点（牛眼）
  · UML 时序图      —— 生命线（虚线）、激活条、同步消息（实心箭头）、返回消息（虚线开箭头）
"""
import os
import win32com.client as win32

# ---------------------------------------------------------------- 输出约定
# 本 skill 的统一产物目录：<当前工作目录>/图片输出/
# 一图一档：每张图产出「位图 png/jpg + 同名 vsdx 源文件」（可选再加 emf 矢量）
OUT_DIR_NAME = "图片输出"


def output_dir(base=None):
    """返回（并创建）项目根目录下的“图片输出”文件夹。"""
    base = os.path.abspath(base or os.getcwd())
    p = os.path.join(base, OUT_DIR_NAME)
    os.makedirs(p, exist_ok=True)
    return p


# ---------------------------------------------------------------- 常量
CM = 2.54                      # 1 英寸 = 2.54 厘米
SONG = "宋体"
HEI = "黑体"
BLACK = "RGB(0,0,0)"
WHITE = "RGB(255,255,255)"

# 论文插图常用低饱和配色（深色描边 + 浅色填充），保证黑白打印仍可辨识
C_BLUE_L = "RGB(228,236,248)"
C_BLUE_M = "RGB(198,215,240)"
C_BLUE_D = "RGB(43,74,125)"
C_GREEN_L = "RGB(228,242,231)"
C_GREEN_M = "RGB(198,224,204)"
C_GREEN_D = "RGB(52,106,68)"
C_GREY_L = "RGB(242,242,242)"
C_GREY_M = "RGB(226,226,226)"
C_GREY_D = "RGB(90,90,90)"
C_ORANGE_L = "RGB(252,238,226)"
C_ORANGE_M = "RGB(246,214,186)"
C_ORANGE_D = "RGB(150,88,40)"
C_PURPLE_L = "RGB(238,233,246)"
C_PURPLE_M = "RGB(215,205,232)"
C_PURPLE_D = "RGB(80,62,122)"

# 箭头样式：Visio EndArrow 索引
ARROW_NONE = 0
ARROW_FILL = 4     # 实心三角箭头（UML 控制流／消息）
ARROW_OPEN = 1     # 开口箭头（依赖、返回消息）
LINE_PAT_SOLID = 1
LINE_PAT_DASH = 2
LINE_PAT_DOT = 3


def text_em_width(text):
    """按 em 为单位估算单行文本宽度（中文 1em，西文约 0.52em）。"""
    best = 0.0
    for ln in str(text).split("\n"):
        zh = sum(1 for c in ln if ord(c) > 127)
        asc = len(ln) - zh
        best = max(best, zh * 1.0 + asc * 0.56)
    return best


def fit_size(text, box_w_cm, base=9.0, min_size=6.0):
    """求使单行不折行的最大字号（pt）。"""
    em = text_em_width(text)
    if em <= 0:
        return base
    usable = box_w_cm - 0.22
    return max(min_size, min(base, usable / (em * 0.03528)))


# 当设置 master 文档时，所有 Draw 实例都在该文档中新建页面，
# 便于把全部插图汇总成一个可编辑的 .vsdx 源文件。
_MASTER = {"doc": None}


def set_master(doc):
    _MASTER["doc"] = doc


class Draw:
    """一页一图。"""

    def __init__(self, app, width_cm, height_cm, name=""):
        # 第一个参数既可以是 Visio.Application，也可以是 Session（会话会把新图登记进去）
        sess = app if hasattr(app, "_by_id") else None
        self.session = sess
        self.app = sess.app if sess is not None else app
        self.w = float(width_cm)
        self.h = float(height_cm)
        if _MASTER["doc"] is not None:
            self.doc = _MASTER["doc"]
            self.page = self.doc.Pages.Add()
        else:
            self.doc = self.app.Documents.Add("")
            self.page = self.doc.Pages.Item(1)
        if name:
            try:
                self.page.Name = name
            except Exception:
                pass
        ps = self.page.PageSheet
        ps.Cells("PageWidth").FormulaU = "%.4f in" % (self.w / CM)
        ps.Cells("PageHeight").FormulaU = "%.4f in" % (self.h / CM)
        for cell, val in (("PageLeftMargin", "0 in"), ("PageRightMargin", "0 in"),
                          ("PageTopMargin", "0 in"), ("PageBottomMargin", "0 in"),
                          ("PrintPageOrientation", "1")):
            try:
                ps.Cells(cell).FormulaU = val
            except Exception:
                pass
        self.page.Background = 0
        self._n = 0
        try:
            self.page.Name = name or ("图%d" % (self.doc.Pages.Count))
        except Exception:
            pass
        if sess is not None:
            sess._items.append([name, None])
            sess._by_id[id(self)] = sess._items[-1]
        # 画布框：铺满整页的白色无边矩形，置底。
        # 作用：使导出图像的边界严格等于页面边界，从而保证插入 Word 后 1:1 缩放、字号不变。
        frame = self.page.DrawRectangle(0, 0, self.X(self.w), self.Y(0))
        self._style(frame, fill=WHITE, no_line=True)
        try:
            frame.SendToBack()
        except Exception:
            pass

    # ------------------------------------------------ 坐标
    def X(self, cm):
        return cm / CM

    def Y(self, cm):
        return (self.h - cm) / CM

    # ------------------------------------------------ 基础样式
    def _style(self, shp, line_color=BLACK, line_w=0.75, fill=None,
               dash=LINE_PAT_SOLID, arrow_end=ARROW_NONE, arrow_begin=ARROW_NONE,
               rounding=None, text="", font=SONG, size=9, bold=False,
               halign=1, valign=1, align_underline=False, italic=False,
               no_line=False):
        c = shp.Cells
        c("LineColor").FormulaU = line_color
        c("LineWeight").FormulaU = "%.2f pt" % line_w
        c("LinePattern").FormulaU = "0" if no_line else str(dash)
        if fill is None:
            c("FillPattern").FormulaU = "0"
        else:
            c("FillPattern").FormulaU = "1"
            c("FillForegnd").FormulaU = fill
        c("EndArrow").FormulaU = str(arrow_end)
        c("BeginArrow").FormulaU = str(arrow_begin)
        c("EndArrowSize").FormulaU = "2"
        c("BeginArrowSize").FormulaU = "2"
        if rounding is not None:
            try:
                c("Rounding").FormulaU = "%.3f in" % (rounding / CM)
            except Exception:
                pass
        # 文字
        try:
            c("TextBkgnd").FormulaU = "0"
            c("LeftMargin").FormulaU = "0.02 in"
            c("RightMargin").FormulaU = "0.02 in"
            c("TopMargin").FormulaU = "0.01 in"
            c("BottomMargin").FormulaU = "0.01 in"
        except Exception:
            pass
        if text:
            shp.Text = text
            try:
                c("Char.Font").FormulaU = '"%s"' % font
                c("Char.Size").FormulaU = "%.2f pt" % size
                c("Char.Color").FormulaU = BLACK
                c("Char.Style").FormulaU = str((4 if align_underline else 0) + (2 if italic else 0) + (1 if bold else 0))
                c("Para.HorzAlign").FormulaU = str(halign)
                c("VerticalAlign").FormulaU = str(valign)
            except Exception:
                pass
        return shp

    # ------------------------------------------------ 图元
    def rect(self, cx, cy, w, h, text="", **kw):
        s = self.page.DrawRectangle(self.X(cx - w / 2.), self.Y(cy + h / 2.),
                                    self.X(cx + w / 2.), self.Y(cy - h / 2.))
        return self._style(s, text=text, **kw)

    def rrect(self, cx, cy, w, h, text="", radius=0.18, **kw):
        return self.rect(cx, cy, w, h, text, rounding=radius, **kw)

    def oval(self, cx, cy, w, h, text="", **kw):
        s = self.page.DrawOval(self.X(cx - w / 2.), self.Y(cy + h / 2.),
                               self.X(cx + w / 2.), self.Y(cy - h / 2.))
        return self._style(s, text=text, **kw)

    def diamond(self, cx, cy, w, h, text="", **kw):
        """菱形（判定结点）：以四条线段闭合绘制，保证线宽与线型一致。"""
        pts = [(cx, cy - h / 2.), (cx + w / 2., cy), (cx, cy + h / 2.), (cx - w / 2., cy),
               (cx, cy - h / 2.)]
        for i in range(4):
            self.line(pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1],
                      line_color=kw.get("line_color", BLACK),
                      line_w=kw.get("line_w", 0.75),
                      dash=kw.get("dash", LINE_PAT_SOLID))
        if text:
            t = self.text(cx, cy, text, w=w * 0.56, h=h * 0.52,
                          size=kw.get("size", 8.2))
            return t
        return None

    def line(self, x1, y1, x2, y2, **kw):
        s = self.page.DrawLine(self.X(x1), self.Y(y1), self.X(x2), self.Y(y2))
        return self._style(s, **kw)

    def poly(self, pts, arrow_last=True, **kw):
        """正交折线；仅最后一段带箭头。"""
        out = []
        for i in range(len(pts) - 1):
            last = (i == len(pts) - 2)
            k = dict(kw)
            if not (last and arrow_last):
                k["arrow_end"] = ARROW_NONE
                k["arrow_begin"] = ARROW_NONE
            out.append(self.line(pts[i][0], pts[i][1], pts[i + 1][0], pts[i + 1][1], **k))
        return out

    def text(self, cx, cy, s, w=6.0, h=0.6, **kw):
        kw.setdefault("size", 9)
        kw["fill"] = None
        kw["no_line"] = True
        return self.rect(cx, cy, w, h, s, **kw)

    # ------------------------------------------------ 复合图元
    def actor(self, cx, cy, label, scale=1.0, size=9, label_w=4.2):
        """UML 参与者（火柴人）。cy 为图形（不含姓名）的垂直中心，姓名绘制在其下方。"""
        r = 0.40 * scale
        head_cy = cy - 1.10 * scale
        self.oval(cx, head_cy, 2 * r, 2 * r, fill=None, line_w=0.9)
        body_top = head_cy + r
        body_bot = cy + 0.45 * scale
        self.line(cx, body_top, cx, body_bot, line_w=0.9)
        self.line(cx - 0.70 * scale, body_top + 0.30 * scale,
                  cx + 0.70 * scale, body_top + 0.30 * scale, line_w=0.9)
        self.line(cx, body_bot, cx - 0.62 * scale, body_bot + 0.85 * scale, line_w=0.9)
        self.line(cx, body_bot, cx + 0.62 * scale, body_bot + 0.85 * scale, line_w=0.9)
        self.text(cx, body_bot + 1.30 * scale, label, w=label_w, h=0.6, size=size)
        return body_bot + 1.62 * scale

    def start_node(self, cx, cy, d=0.55):
        self.oval(cx, cy, d, d, fill=BLACK, line_color=BLACK, line_w=0.75)

    def end_node(self, cx, cy, d=0.72):
        self.oval(cx, cy, d, d, fill=WHITE, line_w=1.0)
        self.oval(cx, cy, d * 0.55, d * 0.55, fill=BLACK, line_color=BLACK, line_w=0.5)

    # ------------------------------------------------ 关系线
    @staticmethod
    def ellipse_edge(cx, cy, w, h, tx, ty):
        """椭圆边界上朝向 (tx,ty) 的点。"""
        dx, dy = tx - cx, ty - cy
        if dx == 0 and dy == 0:
            return cx, cy
        a, b = w / 2.0, h / 2.0
        t = 1.0 / ((dx / a) ** 2 + (dy / b) ** 2) ** 0.5
        return cx + dx * t, cy + dy * t

    @staticmethod
    def rect_edge(cx, cy, w, h, tx, ty, pad=0.0):
        """矩形边界上朝向 (tx,ty) 的点。"""
        dx, dy = tx - cx, ty - cy
        if dx == 0 and dy == 0:
            return cx, cy
        a, b = w / 2.0 + pad, h / 2.0 + pad
        tx_ = a / abs(dx) if dx else 1e9
        ty_ = b / abs(dy) if dy else 1e9
        t = min(tx_, ty_)
        return cx + dx * t, cy + dy * t

    def include(self, a, b, label="«include»", dashed=True, size=7.5,
                aw=3.2, ah=1.05, bw=3.2, bh=1.05, off=0.42, swap=False):
        """a、b 为用例椭圆中心；从 a 指向 b 的依赖/包含关系，带构造型标签。"""
        x1, y1 = self.ellipse_edge(a[0], a[1], aw, ah, b[0], b[1])
        x2, y2 = self.ellipse_edge(b[0], b[1], bw, bh, a[0], a[1])
        self.line(x1, y1, x2, y2, dash=LINE_PAT_DASH if dashed else LINE_PAT_SOLID,
                  arrow_end=ARROW_OPEN, line_w=0.8)
        if label:
            mx, my = (x1 + x2) / 2.0, (y1 + y2) / 2.0
            # 沿法向偏移，避免压线
            dx, dy = x2 - x1, y2 - y1
            L = max((dx * dx + dy * dy) ** 0.5, 1e-6)
            nx, ny = -dy / L, dx / L
            if ny > 0:
                nx, ny = -nx, -ny
            lw = 1.85
            self.text(mx + nx * off, my + ny * off, label, w=lw, h=0.42, size=size)

    # ------------------------------------------------ 导出
    def export(self, outdir=None, stem="fig", fmts=("png", "emf")):
        """按 fmts 导出；outdir 缺省为 <当前项目>/图片输出（首次使用时自动创建）。

        默认同时给出 png（便于查看/插入 PPT 与 Markdown）与 emf（插入 Word 的矢量），
        源文件 .vsdx 由 Session.save() 或 Session 退出时统一写出。
        """
        outdir = outdir or output_dir()
        os.makedirs(outdir, exist_ok=True)
        paths = {}
        for ext in fmts:
            p = os.path.join(outdir, "%s.%s" % (stem, ext))
            self.page.Export(p)
            paths[ext] = p
        return paths

    def close(self, keep=False):
        try:
            if keep:
                return
            self.doc.Close()
        except Exception:
            pass


# ---------------------------------------------------------------- Visio 内置模具
# Visio 自带的网络/服务器/外设模具可以显著提升网络架构图的可读性。
# 注意：Office 安装目录下的 .VSSX 直接用 OpenEx 打开会报“文件未找到”，
# 需先复制到可写目录再打开。
STENCIL_SRC = [
    "C:/Program Files/Microsoft Office/root/Office16/Visio Content/2052",
    "C:/Program Files (x86)/Microsoft Office/root/Office16/Visio Content/2052",
    "C:/Program Files/Microsoft Office/root/Office15/Visio Content/2052",
    "C:/Program Files (x86)/Microsoft Office/root/Office15/Visio Content/2052",
]
STENCIL_FILES = {
    "loc": "NETLOC_M.VSSX",     # 云、建筑物、城市
    "srv": "SERVER_M.VSSX",     # 服务器、Web/数据库/应用程序服务器
    "per": "PERIPH_M.VSSX",     # 防火墙、路由器、交换机、用户
    "pc": "COMPME_M.VSSX",      # PC、笔记本电脑、平板设备
    "sym": "NETSYM_M.VSSX",     # 路由器、交换机、主机、数据库
    "annot": "ANNOT_M.VSSX",    # 修订云形（用作“网络/云”节点，渲染稳定）
}


class Stencils:
    """加载并缓存 Visio 内置模具，供 drop() 使用。"""

    def __init__(self, app, cache_dir=None):
        import shutil
        self.app = getattr(app, "app", app)     # 允许直接传 Session
        self.docs = {}
        self.cache = cache_dir or os.path.join(
            os.environ.get("TEMP", "."), "se_visio_stencils")
        os.makedirs(self.cache, exist_ok=True)
        src_dir = next((d for d in STENCIL_SRC if os.path.isdir(d)), None)
        if src_dir is None:
            return
        for key, fn in STENCIL_FILES.items():
            src = os.path.join(src_dir, fn)
            if not os.path.exists(src):
                continue
            dst = os.path.join(self.cache, fn)
            try:
                if (not os.path.exists(dst)
                        or os.path.getsize(dst) != os.path.getsize(src)):
                    shutil.copyfile(src, dst)
                self.docs[key] = self.app.Documents.OpenEx(dst, 2)   # visOpenRO
            except Exception:
                pass

    def available(self):
        return list(self.docs)

    def master(self, key, name):
        doc = self.docs.get(key)
        if doc is None:
            raise KeyError("模具未加载: %s" % key)
        for i in range(1, doc.Masters.Count + 1):
            m = doc.Masters.Item(i)
            if m.Name == name:
                return m
        raise KeyError("模具 %s 中无图元: %s" % (key, name))

    @staticmethod
    def bbox_cm(draw, shape):
        """返回图元包围盒（以左上角为原点的厘米坐标）：(x1, y1, x2, y2)。"""
        x1, y1, x2, y2 = shape.BoundingBox(1)   # visBBoxUprightWH
        return (x1 * CM, draw.h - y2 * CM, x2 * CM, draw.h - y1 * CM)

    def drop(self, draw, key, name, cx, cy, label=None, h_cm=None,
             font_size=8.0, font=None, align_center=True):
        """投放一个模具图元；h_cm 指定图形总高（含下方文字），按原比例缩放。

        部分模具（如“云”）的局部原点不在几何中心，投放后按包围盒重新对齐，
        保证 cx / cy 始终是图形的视觉中心。
        """
        m = self.master(key, name)
        s = draw.page.Drop(m, draw.X(cx), draw.Y(cy))
        if h_cm:
            oh = s.Cells("Height").ResultIU
            ow = s.Cells("Width").ResultIU
            if oh > 1e-6:
                k = (h_cm / CM) / oh
                s.Cells("Height").FormulaU = "%.4f in" % (oh * k)
                s.Cells("Width").FormulaU = "%.4f in" % (ow * k)
        if label is not None:
            s.Text = label
        try:
            s.Cells("Char.Size").FormulaU = "%.2f pt" % font_size
            if font:
                s.Cells("Char.Font").FormulaU = '"%s"' % font
            s.Cells("Para.HorzAlign").FormulaU = "1"
        except Exception:
            pass
        if align_center:
            try:
                bx1, by1, bx2, by2 = self.bbox_cm(draw, s)
                dx = cx - (bx1 + bx2) / 2.0
                dy = cy - (by1 + by2) / 2.0
                s.Cells("PinX").FormulaU = "%.5f in" % (s.Cells("PinX").ResultIU + dx / CM)
                s.Cells("PinY").FormulaU = "%.5f in" % (s.Cells("PinY").ResultIU - dy / CM)
            except Exception:
                pass
        return s

    def cloud(self, draw, cx, cy, w_cm, h_cm, label="", size=8.5,
              fill=None, line=None):
        """以 Visio“修订云形”绘制网络云节点（NETLOC 的“云”在部分环境下不渲染，
        故统一改用几何云形，可稳定导出）。"""
        m = self.master("annot", "修订云形")
        s = draw.page.Drop(m, draw.X(cx), draw.Y(cy))
        s.Cells("Width").FormulaU = "%.4f in" % (w_cm / CM)
        s.Cells("Height").FormulaU = "%.4f in" % (h_cm / CM)
        s.Cells("FillPattern").FormulaU = "1"
        s.Cells("FillForegnd").FormulaU = fill or C_BLUE_L
        s.Cells("LineColor").FormulaU = line or C_BLUE_D
        s.Cells("LineWeight").FormulaU = "1 pt"
        if label:
            s.Text = label
            s.Cells("Char.Size").FormulaU = "%.2f pt" % size
            s.Cells("Char.Font").FormulaU = '"%s"' % SONG
            s.Cells("Para.HorzAlign").FormulaU = "1"
            s.Cells("VerticalAlign").FormulaU = "1"
        return s

    def close(self):
        for d in self.docs.values():
            try:
                d.Close()
            except Exception:
                pass
        self.docs = {}


def open_app():
    app = win32.Dispatch("Visio.Application")
    app.Visible = False
    app.AlertResponse = 7
    try:
        app.ScreenUpdating = False
    except Exception:
        pass
    return app


def close_app(app):
    try:
        app.AlertResponse = 7
        while app.Documents.Count > 0:
            app.Documents.Item(1).Close()
    except Exception:
        pass
    try:
        app.Quit()
    except Exception:
        pass


# ---------------------------------------------------------------- 会话（推荐入口）
class Session:
    """一次绘图会话：统一管理 Visio 进程、“图片输出”目录与源文件。

    默认「一图一档」——
        每张图 = 一个独立 Visio 文档 → 输出「1 张位图 + 1 个同名 .vsdx」，
        用户拿到 vsdx 可以单独打开、单独修改，不影响其他图。

        with Session() as S:
            d = S.draw(15.4, 9.7, "图4-1 逻辑架构图")
            ...                              # d.rect / d.oval / d.line 作图
            S.save(d, "fig4-1_逻辑架构图")

        产物：
            图片输出/fig4-1_逻辑架构图.png    ← 可视图
            图片输出/fig4-1_逻辑架构图.vsdx   ← 对应的 Visio 源文件

    需要插 Word 的矢量时传 fmts=("png", "emf")，会多产出一个 .emf。
    需要「一册多页」的合集时传 bundle=True，额外写出一个含全部页面的图集 vsdx。
    """

    def __init__(self, name="图集", outdir=None, fmts=("png",), bundle=False, notes=None):
        self.name = name                  # bundle=True 时的图集名
        self.outdir = outdir or output_dir()
        self.fmts = tuple(fmts)
        self.bundle = bundle
        self.notes = notes or []
        self.app = None
        self.doc = None                   # bundle 模式下的合集文档
        self._docs = []                   # 一图一档时各图自己的文档
        self._items = []
        self._by_id = {}

    # ---- 生命周期
    def __enter__(self):
        self.app = open_app()
        if self.bundle:
            self.doc = self.app.Documents.Add("")
            set_master(self.doc)
        return self

    def __exit__(self, exc_type, exc, tb):
        try:
            if exc_type is None and self.bundle:
                self.finalize()
        finally:
            set_master(None)
            for doc in ([self.doc] if self.doc is not None else []) + self._docs:
                try:
                    if doc is not None:
                        doc.Close()
                except Exception:
                    pass
            close_app(self.app)
        return False

    # ---- 作图与导出
    def draw(self, width_cm, height_cm, name=""):
        d = Draw(self.app, width_cm, height_cm, name)   # 非 bundle 时各建独立文档
        if not self.bundle:
            self._docs.append(d.doc)
        self._items.append([name, None])
        self._by_id[id(d)] = self._items[-1]
        return d

    def save(self, draw, stem, fmts=None):
        """输出该图：位图（默认 png）+ 同名 vsdx 源文件。

        返回 {'png': 路径, 'vsdx': 路径, ['emf': 路径]}。
        """
        fmts = tuple(fmts or self.fmts)
        paths = draw.export(self.outdir, stem, fmts)

        if not self.bundle:
            vpath = os.path.join(self.outdir, stem + ".vsdx")
            try:
                draw.doc.SaveAs(vpath)
                paths["vsdx"] = vpath
            except Exception as e:
                print("保存 vsdx 失败：", e)

        rec = self._by_id.get(id(draw))
        if rec is None:
            rec = [stem, None]
            self._items.append(rec)
            self._by_id[id(draw)] = rec
        rec[0] = rec[0] or stem
        rec[1] = stem
        return paths

    # ---- 合集模式收尾：写说明页并保存整册 vsdx
    def finalize(self, vsdx_name=None):
        if self.doc is None:
            return None
        self._write_cover()
        path = os.path.join(self.outdir, (vsdx_name or self.name) + ".vsdx")
        try:
            self.doc.SaveAs(path)
        except Exception as e:
            print("保存图集 vsdx 失败：", e)
            return None
        return path

    def _write_cover(self):
        """把新建文档的空白首页写成“图集说明”，列出每页对应的文件。"""
        try:
            pg = self.doc.Pages.Item(1)
            if pg.Shapes.Count > 0:
                return
            pg.Name = "图集说明"
            ps = pg.PageSheet
            ps.Cells("PageWidth").FormulaU = "8.27 in"
            ps.Cells("PageHeight").FormulaU = "11.69 in"
            lines = ["%s" % self.name, ""]
            lines.append("本图集共 %d 幅插图；每幅另有独立的 .vsdx 与位图，可单独编辑。" % len(self._items))
            lines.append("本文件为合集，也可在 Visio 中继续编辑。")
            lines.append("")
            for i, (nm, stem) in enumerate(self._items):
                lines.append("%d. %s%s    ···· 第 %d 页" % (
                    i + 1, nm or "", ("  [%s.png]" % stem) if stem else "", i + 2))
            if self.notes:
                lines.append("")
                lines.extend(self.notes)
            sh = pg.DrawRectangle(0.6, 0.6, 7.67, 11.09)
            sh.Text = "\n".join(lines)
            sh.Cells("LinePattern").FormulaU = "0"
            sh.Cells("FillPattern").FormulaU = "0"
            sh.Cells("Char.Font").FormulaU = '"宋体"'
            sh.Cells("Char.Size").FormulaU = "11 pt"
            sh.Cells("Para.HorzAlign").FormulaU = "0"
            sh.Cells("VerticalAlign").FormulaU = "0"
        except Exception:
            pass


def quick_page(width_cm, height_cm, name="", outdir=None, fmts=("png",)):
    """一次性出图：返回 (Draw, outdir)，调用方画完后自行 d.export(...)。"""
    app = open_app()
    d = Draw(app, width_cm, height_cm, name)
    return d, outdir or output_dir()
