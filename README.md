# dsh-pptx-anim

给 `.pptx` 批量注入**真正的 PowerPoint 进入动画**与幻灯片切换 —— 补上 `python-pptx` 没有动画 API 的缺口。

> `python-pptx` 不实现 `<p:timing>`，所以市面上绝大多数"用代码生成 PPT"的方案都只能产出**静态页**。
> 本插件直接写 OOXML 的时间轴树，动画是文件里的数据：**目标机器不需要装 PowerPoint / WPS**。

## 提供的工具

| 工具 | 作用 |
|---|---|
| `pptx_animate` | 给 pptx 注入进入动画 + 页面切换（不修改源文件） |
| `pptx_anim_validate` | 结构校验：元素顺序、空条件列表、动画目标是否指向真实形状 |
| `pptx_anim_env` | 探测哪个 Python 能 `import pptx` |

## 依赖

- Node.js ≥ 20（DSH 本身自带）
- **Python 3 且能 `import pptx`**（即 `python-pptx`）——工具会在 `python` / `python3` / `py` 里自动探测；
  找不到就显式传 `python` 参数，或设置环境变量 `DSH_PPTX_ANIM_PYTHON`。

没有 Python 也能用 DSH，只是这个插件会返回"找不到解释器"的明确提示，不会静默失败。

## 用法

```
pptx_animate({ input: "C:/x/deck.pptx" })
  → 生成 C:/x/deck_anim.pptx，返回每页动画节点数

pptx_anim_validate({ input: "C:/x/deck_anim.pptx" })
  → 结构校验通过 / 列出问题
```

指定预设：

```
pptx_animate({ input: "...", preset: "cards", transition: "fade" })
```

## 预设

| preset | 适用 | 交错 | 时长 | 图片 | 文字 |
|---|---|---|---|---|---|
| `auto`（默认） | 按每页构成自动判定 | — | — | — | — |
| `title` | 封面 | 180ms | 450ms | 淡出 | 淡出 |
| `proposition` | 命题大字页 | 150ms | 420ms | 淡出 | 淡出 |
| `image` | 图文页 | 200ms | 500ms | 淡出 | 擦除（上） |
| `cards` | 三/四栏卡片页 | 110ms | 320ms | 擦除 | 擦除 |
| `data` | 数据密集页 | 90ms | 320ms | 淡出 | 淡出 |
| `none` | 只加切换、不加动画 | — | — | — | — |

只使用 `animEffect` 型效果（淡出 / 擦除 / 盒状 / 圆形 / 菱形 / 十字）——它们是最稳的单节点编码。

## 形状命名（精准控制播放顺序）

生成 deck 时给形状命名 `s01`、`s02`……（按阅读顺序）：

```python
shape = slide.shapes.add_textbox(...)
shape.name = 's01'
```

有名字按序号播；没名字按 z-order 播（自动生成的 deck 通常一致，手改过会错位）。

## 已验证 / 未验证

**已验证**（2026-10-29，Windows 11 + DSH 0.2.1）：
- 15 页一次性注入 193 个 `animEffect` 节点 + 全页切换
- 修掉两个会让 PowerPoint 提示"需要修复"的坑：
  `transition` 必须排在 `timing` **之前**（CT_Slide 子元素顺序）、`stCondLst` 不能为空
- 注入后的文件能被 LibreOffice 正常打开（15 页、无缺字体）
- `validate.py` 对故意打乱顺序的文件能报错

**未验证**：
- PowerPoint / WPS 的**实机播放效果**（开发机只装了 ONLYOFFICE，无 PowerPoint COM）
- 单页 25+ 动画节点的极端情况

⇒ 交付时请同时保留「无动画版」，并提示首次放映前用 F5 预览一次。

## 禁区

- **不做** Morph 变形、运动路径（`animMotion`）、触发器交互链、按段落逐条动画
  —— 这些要么需要 PowerPoint 专有数据，要么手写风险远高于收益
- 一页动画 > 25 个节点会让放映发卡，应拆页
- 不要把 `p:timing` 追加在 `p:transition` 之前

## 工作原理

```
lib/index.js        Host 半：注册 3 个工具，探测 Python，spawn 助手脚本
scripts/pptx_anim.py   读 pptx → 按预设生成 <p:timing> 树 → 写回新文件
scripts/validate.py    直接解 zip 检查 slide XML 的结构合法性
```

动画树的形状（简化）：

```
p:timing → p:tnLst → p:par(tmRoot) → p:seq(mainSeq)
          → p:par(点击组) → p:par(延迟0) → p:par(delay=交错) × N
                                            └ presetClass="entr" + p:set + p:animEffect
```

## License

MIT
