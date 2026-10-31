# PORTABLE — 装到别的 harness

这个目录是 `Visio-painting` 的**独立可分发副本**。整个文件夹拷走即可，无外部依赖（除系统环境要求）。

## 一、硬性环境

| 项 | 要求 | 检查命令 |
|---|---|---|
| 系统 | **Windows**（Visio 只有 Windows 版） | — |
| Visio | 2016 / 2019 / 2021 / 365 任一款 | 打开 Visio 能新建文档即可 |
| Python | 3.9+ | `python -V` |
| pywin32 | COM 调用必需 | `python -c "import win32com.client"` |

缺 pywin32：`pip install pywin32`。

一条命令自检环境：

```bash
python -c "import win32com.client as w; print(w.Dispatch('Visio.Application').Version)"
```

能打印出版本号（如 `16.0`）就说明可用。

> ⚠️ 没有 Visio 的机器**跑不起来**。这是本 skill 的前提，不是可绕过的配置项。
> 若对方是 macOS / Linux，只能改用 Graphviz、draw.io CLI 或 mermaid，不适合本 skill。

## 二、安装到各 harness

原则：**把整个 `Visio-painting/` 目录放进该 harness 读取 skill 的位置**，让它能读到 `SKILL.md`。

### Claude Code
```bash
cp -r Visio-painting ~/.claude/skills/
```
项目级（只对当前仓库生效）：放进 `<repo>/.claude/skills/`。

### Cursor / Windsurf / 其他支持 skills 或 rules 的 IDE
复制到其 skills 目录；若只支持 rules，把 `SKILL.md` 内容粘进 `.cursor/rules/visio-painting.mdc`，
并在 rules 开头加一行路径提示：
```
绘图脚本与参考文档位于：<绝对路径>/Visio-painting/scripts、/reference
```

### Pi（pi coding agent）
```bash
cp -r Visio-painting ~/.pi/agent/skills/visio-painting      # 全局
cp -r Visio-painting <repo>/.pi/skills/visio-painting       # 项目级（只对当前仓库生效）
```
目录名建议与 frontmatter 里的 `name: visio-painting` 保持一致。装好后可用 `/skill:visio-painting`
强制加载，也会按 `description` 自动匹配；改了 `SKILL.md` 在会话里用 `/reload` 生效。
> Agent Skills 规范要求 `description` 字段（本技能的 `summary` / `read_when` 是非标准字段，
> 仅作兼容保留），缺 `description` 时 Pi 不会加载该技能。

### 通用（任何能执行 Python 的 Agent）
不需要"安装"，只要告诉它两件事：

1. 读 `<路径>/Visio-painting/SKILL.md` 作为工作规程；
2. 脚本目录加入 `sys.path`：
   ```python
   import sys; sys.path.insert(0, r"<路径>/Visio-painting/scripts")
   from visio_engine import *
   ```

可直接粘给对方的开场提示：

```
你有本地 Visio。请先阅读 <路径>/Visio-painting/SKILL.md，
按其中的 Session 用法作图，产物统一输出到当前项目的「图片输出」文件夹，
每张图要同时产出 png 与 emf，并保存 vsdx 源文件。
```

## 三、装完自检（30 秒）

在一个空目录里跑：

```bash
cd /tmp/visio-smoke
python - <<'PY'
import sys
sys.path.insert(0, r"<路径>/Visio-painting/scripts")
from visio_engine import *
with Session() as S:                     # 默认一图一档
    d = S.draw(10, 5, "冒烟测试")
    d.rrect(5, 2, 4, 1, "Visio-painting OK", fill=C_BLUE_L, line_color=C_BLUE_D, size=10)
    print(S.save(d, "smoke"))
PY
ls 图片输出
```

期望得到**一对**文件：

```
smoke.png      smoke.vsdx
```

- 两个都在 → 安装成功（`smoke.vsdx` 打开应只有 1 页、尺寸 10×5cm）。
- 只有 png 没有 vsdx → Visio 保存路径无权限，或该目录不可写。
- 报 `Invalid class string` → Visio 未安装或未注册 COM。

需要插 Word 的矢量时，把 `Session()` 改成 `Session(fmts=("png","emf"))`，会多出一个 `smoke.emf`。

## 四、版本与目录

```
Visio-painting/
├── SKILL.md              技能说明与用法（Agent 读这个）
├── PORTABLE.md           本文件
├── scripts/              visio_engine.py / word_emf.py / recipes/
├── reference/            notation.md（图型规范）/ stencils.md（模具图元名）
├── examples/             论文插图全流程示例
└── tests/                布局回归测试（无需 Visio）
```

约定：`脚本只依赖 scripts/ 内部文件 + 系统 Visio`，不读取任何项目数据，
因此可以整目录打包分发，也可以同时装进多个 harness。
