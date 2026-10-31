---
name: visio-painting
description: 用本机 Visio（COM 自动化）画软件工程类插图：UML 用例图/活动图/时序图、陈氏 E-R 图、分层架构图、模块结构图、网络拓扑图。每张图产出可视图（png/svg/emf）+ 同名单页 .vsdx 源文件，统一落在项目「图片输出/」。当用户说「画个图/出图/软件工程绘图/visio 自动化绘图」，或需要画用例图、活动图、时序图、类图、状态图、E-R 图、数据流图、架构图、部署图、拓扑图、流程图、模块结构图，或需要可编辑的 vsdx / 插 Word 不糊的矢量图时使用。仅限 Windows 且已安装 Visio。
license: MIT
compatibility: Windows + Microsoft Visio 2016/2019/2021/365 + Python 3.9+ + pywin32
summary: 用本机 Visio（COM 自动化）绘制软件工程类图表——UML 用例图/活动图/时序图、陈氏 E-R 图、分层架构图、模块结构图、网络拓扑图；输出「一图一档」：每张图给一张 png/jpg 可视图 + 一个同名 vsdx 源文件，统一落在当前项目的「图片输出」文件夹
read_when:
  - 用户说「软件工程绘图」「visio 自动化绘图」「帮我画图」「画个图」「出图」「画张示意图」
  - 需要画用例图、活动图、时序图、状态图、类图、E-R 图、数据流图、架构图、部署图、拓扑图、流程图、模块结构图
  - 出现 UML、Visio、vsdx、C4、陈氏表示法、Chen ER 等字样
  - 论文/需求规格说明书/设计文档需要成体系的工程插图
  - 需要可编辑的图源文件，或需要矢量图插入 Word 不糊
  - 需要把图批量导出并放进项目目录
agent_created: true
---

# Visio-painting　软件工程绘图 / Visio 自动化绘图

用 Visio COM 把「图」写成代码：厘米坐标精确摆放，**一张图 → 一张可视图（png/jpg）+ 一个同名 .vsdx 源文件**，
用户拿到 vsdx 就能单独打开、单独改，互不影响。

---

## 一、什么时候该用它（被动调用判定）

只要下面这些字眼出现，**先做一次 3 秒判定，再决定是否调用本 skill**——不要直接开画，也不要直接推给图像生成。

**命中词**：软件工程绘图、visio 自动化绘图、帮我画图、画个图、出图、画用例图 / 活动图 / 时序图 / 状态图 /
类图 / E-R 图 / 数据流图 / 架构图 / 部署图 / 网络拓扑图 / 流程图 / 模块结构图 / 功能结构图、UML、Visio、vsdx。

**判为「符合」→ 直接调用本 skill（满足任一即可）**

1. 要表达的是**系统的结构、流程、数据关系、对象交互**（不是画一个人、一只猫）；
2. 用户要**可编辑的源文件**（.vsdx / 能改文字和连线）；
3. 图要插进 Word / PPT / 论文，需要**矢量、放大不糊**；
4. 要求**符号规范**：UML、陈氏 E-R、GB/T 图符、分层架构的惯用画法。

**判为「不符合」→ 不要用本 skill**

| 用户其实想要 | 应该转给 |
|---|---|
| AI 插画、海报、写实图像、风格化美术、logo | `ImageGen` |
| 数据图表（折线/柱状/饼图） | 直接用图表库画，不必开 Visio |
| 界面截图、手绘涂鸦、手机拍照 | 说明做不到，给出替代方案 |

**判定后的默认动作**：符合就先画**一张样图**（第一个最核心的图）给用户看风格，
确认后再批量出剩下的；不要一上来闷头画 20 张。

---

## 二、产物规范（硬性要求）

**输出目录**：当前项目根目录下新建 `图片输出/`（引擎自动创建，不要手写到别处）。

**一图一档**：每画一张图，产出**一个可视文件 + 一个同名 vsdx**，两者文件名严格对应。

```
<当前项目>/
└── 图片输出/
    ├── fig3-1_系统用例图.png      ← 可视图：能直接看、能插 PPT / Markdown / 微信
    ├── fig3-1_系统用例图.vsdx     ← 对应的 Visio 源文件：单页，只有这一张图，可单独改
    ├── fig4-1_逻辑架构图.png
    ├── fig4-1_逻辑架构图.vsdx
    └── ...
```

规则：

1. **每张图必须成对**：`<stem>.png` + `<stem>.vsdx`，stem 完全一致。**vsdx 不可省**——用户就是要拿它自己改。
2. **每个 vsdx 只含一页**：页面尺寸 = 该图尺寸，页名 = 图名。改这张图不会影响其他图。
3. 命名：`fig{章}-{序}_{中文名}`，如 `fig3-1_系统用例图`；无章节时用 `fig01_xxx`。
4. 需要 jpg 就把 `fmts` 改成 `("jpg",)`；需要 jpg+png 就 `("png","jpg")`。
5. **要插 Word**（矢量不糊）时加 `emf`：`Session(fmts=("png","emf"))`，会多产出一个 `.emf`，vsdx 照旧。
6. 需要「一册多页」合集时用 `Session(..., bundle=True)`，额外写出一个含全部页面的图集 vsdx（可选功能，**不是默认**）。
7. 交付时必须主动报账：输出目录绝对路径、图的数量、图与 vsdx 的对应关系（一张图 = 1 张图 + 1 个 vsdx）。

> 页面尺寸 = 插入尺寸。`Draw(15.4, 9.7, ...)` 意味着这张图在 Word 里就占 15.4×9.7cm，图内 9pt 的字就是 9pt。

## 三、用法

### 3.1 推荐写法

```python
import sys; sys.path.insert(0, r"<skill>/scripts")
from visio_engine import *

with Session() as S:                          # 默认一图一档
    d = S.draw(15.4, 9.7, "图4-1 逻辑架构图")   # 页面尺寸 = 插入尺寸
    d.rrect(7.7, 1.2, 14.8, 1.3, "表现层", fill=C_BLUE_L, line_color=C_BLUE_D)
    d.rect(4.0, 3.0, 3.2, 0.9, "Web 创作工作台", size=9)
    d.line(6.0, 1.9, 6.0, 2.5, arrow_end=ARROW_FILL)
    print(S.save(d, "fig4-1_逻辑架构图"))
```

跑完 `图片输出/` 里就是一对：

```
fig4-1_逻辑架构图.png     ← 看
fig4-1_逻辑架构图.vsdx    ← 改（单页，只有这一张图）
```

想同时拿到插 Word 的矢量：

```python
with Session(fmts=("png", "emf")) as S:       # 多出一个 .emf，vsdx 照旧
    ...
```

### 3.2 一次画一组（每张各出一个 vsdx）

```python
with Session() as S:
    for stem, (w, h, title) in {
        "fig3-1_系统用例图":  (15.0, 10.2, "图3-1 系统用例图"),
        "fig4-1_逻辑架构图": (15.4, 9.7,  "图4-1 逻辑架构图"),
    }.items():
        d = S.draw(w, h, title)
        ...                                   # 画这一张
        S.save(d, stem)                       # 该图自己的 png + vsdx
```

产物是 2 张 png + 2 个 vsdx，一一对应；改其中一张不影响另一张。

### 3.3 一册合集（可选）

```python
with Session("系统设计图集", bundle=True) as S:   # 多页合集模式
    ...
# 额外写出 图片输出/系统设计图集.vsdx（含说明页 + 每图一页）
```

只有需要「把 24 张图装进一个文件方便通览」时才用；日常交付用默认的一图一档。

### 3.4 直接调用配方（最快的起点）

`scripts/recipes/` 下有六类图的完整骨架，改内容即可用：

```bash
cd <你的项目目录>
python <skill>/scripts/recipes/usecase.py        # 用例图
python <skill>/scripts/recipes/activity.py       # 活动图
python <skill>/scripts/recipes/sequence.py       # 时序图
python <skill>/scripts/recipes/er.py             # E-R 图（属性图 + 局部图）
python <skill>/scripts/recipes/architecture.py   # 分层架构 + 模块结构
python <skill>/scripts/recipes/network.py        # 网络拓扑（Visio 内置图元）
```

产物同样落在 `图片输出/`。

### 3.5 插进 Word（矢量、带裁边补偿）

```python
from word_emf import add_emf
p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
add_emf(p, "图片输出/fig4-1_逻辑架构图.emf", w_cm=15.4, h_cm=9.7)
```

`w_cm/h_cm` 必须等于该图的 Visio 页面尺寸；函数内部会裁掉 1.4% 的导出边框并等比补回外框。

### 3.6 常用 API 速查

| 需求 | 调用 |
|---|---|
| 矩形 / 圆角矩形 / 椭圆 / 菱形 | `d.rect / d.rrect / d.oval / d.diamond` |
| 直线 / 正交折线 | `d.line(..., arrow_end=ARROW_FILL)` / `d.poly([(x,y),...])` |
| 纯文字（无框） | `d.text(cx, cy, "文字", w=6, h=0.5, size=9)` |
| 火柴人参与者 | `d.actor(cx, cy, "普通用户")` |
| 起止结点（活动图） | `d.start_node` / `d.end_node` |
| `«include»` / `«extend»` | `d.include(a, b, label="«extend»")` |
| 字号自适应（防文字撑破框） | `size=fit_size(文字, 框宽cm, 9)` |
| Visio 内置图元（服务器/防火墙/PC） | `Stencils(app).drop(d, "srv", "应用程序服务器", cx, cy, "应用服务器")` |
| 云形（网络图） | `Stencils(app).cloud(d, cx, cy, 2.45, 1.6, "Internet")` |

---

## 四、六类图的规范要点

详细符号与间距经验值见 `reference/notation.md`，这里只列最容易翻车的点：

| 图型 | 必守规则 |
|---|---|
| **用例图** | 用例列的中心 = 系统边界中心；**一律单列**（双列会让右列用例无处引线而漏连）；关联线从参与者**躯干中线**引出；同侧两个角色纵向间距 ≥3.9cm |
| **活动图** | 菱形取 4.4×2.0cm（每行 4 汉字）；菱形与相邻活动留 ≥0.45cm 给「是/否」；主流程单列、异常分支右列且距 5cm 以上；终止结点标签放在右侧 |
| **时序图** | 首条消息距对象头 ≥0.55cm；生命线间距 ≥2.6cm；同步=实线实心箭头，返回=虚线开口箭头；自消息回环必须带箭头 |
| **E-R 图** | 连线画到**图形边界**（不能连中心，否则线穿字）；属性单行等宽排开；主键名下加下划线；基数标在线段上方靠端处 |
| **分层架构** | 层名带 ≥1.85cm（否则「基础设施层」折行）；组件宽度按可用宽度均分 + `fit_size` 定字号；层间不画箭头 |
| **模块结构** | 只做两级（系统→子系统→模块）；三级功能点会挤成单字换行，改在正文叙述 |
| **网络拓扑** | 一律用 Visio 内置模具；图元不带自带文字，标签另画在图元下方 0.42cm；云形用 `Stencils.cloud()` |

---

## 五、Visio 内置模具（网络拓扑必备）

```python
st = Stencils(app)                      # 首次会自动把 .VSSX 复制到 %TEMP% 再打开
st.drop(d, "pc",  "PC",             cx, cy, "桌面浏览器", h_cm=1.6)
st.drop(d, "per", "防火墙",          cx, cy, "防火墙",   h_cm=2.0)
st.drop(d, "srv", "应用程序服务器",  cx, cy, "应用服务器", h_cm=2.0)
st.drop(d, "srv", "数据库服务器",    cx, cy, "MySQL 8.0", h_cm=1.8)
```

可用键：`pc`（PC/笔记本/平板）、`per`（防火墙/路由器/交换机）、`srv`（各类服务器）、
`sym`（路由器/主机/数据库）、`loc`、`annot`（修订云形）。
完整图元名清单见 `reference/stencils.md`。

---

## 六、踩过的坑（照做可以少走两天）

**坐标与页面**
- `DrawRectangle` 用英寸、左下原点，**不受吸附网格影响**，坐标是准的，别怀疑 snapping。
- 页面尺寸必须等于插入尺寸，否则字号被等比缩放，和正文对不上。

**导出与灰边框**
- `Page.Export` 的 EMF **只取内容范围**，要等于页面须在底层铺一个填充矩形当画布框。
- 但**有填充的框一定会在 EMF 里留一圈灰色发丝边框**（`LinePattern=0`、白色描边都挡不住）。
- 解法：插图时用 `<a:srcRect l/t/r/b="1400">` 四边各裁 1.4%，并把外框按 `1/(1-2×1.4%)` 放大——
  裁掉边框且图内仍是 1:1。`word_emf.add_emf()` 已封装。
- 因此**图内四周留白 ≥0.35cm**，否则裁边会切掉贴边元素。

**EMF 插 Word**
- python-docx 不支持 EMF（`add_picture` 依赖 PIL），必须手写 DrawingML 并自建 image part。
- 拼 XML 时**必须声明 `xmlns:r`**，否则 lxml 报 `Namespace prefix r for embed is not defined`。

**Visio 模具**
- 安装目录下的 `.VSSX` 直接 `OpenEx` 会报「文件未找到」，**必须先复制到可写目录**。
- `Masters.ItemU('中文名')` 取的是通用名，会报「未找到对象名」；遍历 `Masters.Item(i).Name` 才可靠。
- `Shape.BoundingBox(0)` 返回 `(0,0,-1,-1)` 这种垃圾值，**必须用 `BoundingBox(1)`**。
- NETLOC 的「云」放置后**不渲染**（导出是空白），改用 `Stencils.cloud()`（修订云形）。
- 模具自带文字位置各不相同，会压住图形 → **关掉图元文字，标签单独画**。

**画图本身**
- `Page.DrawPolyline` 需要 SAFEARRAY，传 list 会报错 → 菱形用四条线段画（`d.diamond` 已封装）。
- 下划线用 `Char.Style = 4`（位掩码：1 粗、2 斜、4 下划线）。
- 中文在 Visio 里**逐字断行**，框窄一个四字词就变两行 → 用 `fit_size()` 反算字号或加宽。

---

## 七、交付前自检

1. 打开 `图片输出/*.png` 逐张看：有没有元素碰撞、超出画布、文字被菱形/边框切割。
2. 核对**每张图都成对**：`xxx.png` 与 `xxx.vsdx` 同名同数，且每个 vsdx 打开只有 1 页、页面尺寸等于该图尺寸。
3. 图题在图下方、表题在表上方；编号连续，正文每幅图都有「如图 X-Y 所示」。
4. 插进 Word 后导一次 PDF，确认没有灰边框、没有被裁边。
5. 图的四周留白 ≥0.35cm。
6. 交付说明里写清输出目录、图数、文件清单。

---

## 八、目录说明

```
Visio-painting/
├── SKILL.md                  ← 本文件
├── PORTABLE.md               ← 装到其他 harness（Claude Code / Cursor / 通用）的说明
├── scripts/
│   ├── visio_engine.py       ← 引擎：图元、样式、模具、Session、导出
│   ├── word_emf.py           ← EMF 矢量插入 docx（裁边补偿）
│   └── recipes/              ← 六类图的骨架，改内容即用
├── reference/
│   ├── notation.md           ← 各图型符号规范与间距经验值
│   └── stencils.md           ← Visio 中文模具图元名清单
├── examples/                 ← 一份完整的论文图集脚本（可当模板复制）
└── tests/                    ← 布局回归测试（不需要 Visio）
```
