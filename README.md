# Visio-painting

用 **Visio COM 自动化**把软件工程图「写成代码」：厘米坐标精确摆放，**一张图 = 一张可视图（png / emf）+ 一个同名 `.vsdx` 源文件**，统一输出到当前项目的 `图片输出/` 文件夹。

面向本科毕业设计、课程设计、需求规格说明书、设计文档等需要「成体系工程插图 + 可编辑源文件」的场景。

## 支持的图型

| 类别 | 具体图型 |
|---|---|
| UML | 用例图、活动图、时序图（含状态图/类图的基础图元） |
| 数据库 | 陈氏 E-R 图（实体属性图、局部 E-R、汇总 E-R） |
| 架构 | 分层架构图、模块（功能）结构图 |
| 网络 | 网络拓扑图 / 部署图（基于 Visio 内置模具） |

## 一图一档

```
<当前项目>/
└── 图片输出/
    ├── fig3-1_系统用例图.png      ← 可视图：看、插 PPT / Markdown
    ├── fig3-1_系统用例图.vsdx     ← 源文件：单页、可单独改
    ├── fig4-1_逻辑架构图.png
    └── fig4-1_逻辑架构图.vsdx
```

- 每个 `.vsdx` **只含一页**，页面尺寸 = 该图尺寸，改一张不影响其他图；
- 需要插 Word 的矢量时用 `Session(fmts=("png", "emf"))`，额外得到一个 `.emf`（`scripts/word_emf.py` 已封装裁边补偿）；
- 页面尺寸 = 插入尺寸，图内 9pt 的字插进 Word 就是 9pt。

## 环境要求

| 项 | 要求 | 检查命令 |
|---|---|---|
| 系统 | **Windows**（Visio 只有 Windows 版） | — |
| Visio | 2016 / 2019 / 2021 / 365 | 能新建文档即可 |
| Python | 3.9+ | `python -V` |
| pywin32 | COM 调用必需 | `python -c "import win32com.client"` |

```bash
pip install pywin32
python -c "import win32com.client as w; print(w.Dispatch('Visio.Application').Version)"
```

能打印出版本号（如 `16.0`）即可使用。

## 快速开始

```python
import sys; sys.path.insert(0, r"<路径>/Visio-painting/scripts")
from visio_engine import *

with Session() as S:                          # 默认一图一档
    d = S.draw(15.4, 9.7, "图4-1 逻辑架构图")   # 页面尺寸（cm）
    d.rrect(7.7, 1.2, 14.8, 1.3, "表现层", fill=C_BLUE_L, line_color=C_BLUE_D)
    d.rect(4.0, 3.0, 3.2, 0.9, "Web 创作工作台", size=9)
    d.line(6.0, 1.9, 6.0, 2.5, arrow_end=ARROW_FILL)
    print(S.save(d, "fig4-1_逻辑架构图"))
```

产物：`图片输出/fig4-1_逻辑架构图.png` + `图片输出/fig4-1_逻辑架构图.vsdx`。

现成骨架（改内容即用）：

```bash
cd <你的项目目录>
python <路径>/Visio-painting/scripts/recipes/usecase.py       # 用例图
python <路径>/Visio-painting/scripts/recipes/activity.py      # 活动图
python <路径>/Visio-painting/scripts/recipes/sequence.py      # 时序图
python <路径>/Visio-painting/scripts/recipes/er.py            # E-R 图
python <路径>/Visio-painting/scripts/recipes/architecture.py  # 分层架构 + 模块结构
python <路径>/Visio-painting/scripts/recipes/network.py       # 网络拓扑
```

完整论文插图流程见 [`examples/paper_figure_set.py`](examples/paper_figure_set.py)。

## 目录结构

```
Visio-painting/
├── SKILL.md                  技能说明：适用场景、API 速查、六类图规范要点、踩坑记录
├── PORTABLE.md               装到 Claude Code / Cursor / 其他 Agent 的说明与自检步骤
├── scripts/
│   ├── visio_engine.py       引擎：图元 / 样式 / 模具 / Session / 导出
│   ├── word_emf.py           EMF 矢量插入 docx（1.4% 裁边补偿）
│   └── recipes/              六类图的骨架
├── reference/
│   ├── notation.md           各图型符号规范与间距经验值
│   └── stencils.md           Visio 中文模具图元名清单
└── examples/paper_figure_set.py
```

## 文档

- [SKILL.md](SKILL.md) — 什么时候用它、产物规范、用法、六类图必守规则、踩坑记录
- [reference/notation.md](reference/notation.md) — 符号与间距的硬指标（UML / 陈氏 E-R / 分层架构 / 拓扑）
- [reference/stencils.md](reference/stencils.md) — `Stencils.drop()` 可用的中文图元名
- [PORTABLE.md](PORTABLE.md) — 安装到其他 harness、 30 秒自检

## 已知限制

- **必须在 Windows 且安装 Visio**；macOS / Linux 请改用 Graphviz、draw.io CLI 或 mermaid。
- EMF 导出会带一圈浅灰发丝边框：插 Word 时四边各裁 1.4%（`add_emf()` 已处理），因此图内四周留白需 ≥ 0.35cm。
- 中文在 Visio 中逐字断行，框宽不足时需用 `fit_size()` 反算字号或加宽。

## 许可

仓库暂未声明开源许可（All rights reserved）。如需复用或二次分发，请先联系仓库作者。
