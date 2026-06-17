---
name: 腾讯云架构师南京城市沙龙PPT模板适配Skill
version: "4.0-nanjing"
description: 腾讯云架构师南京城市沙龙 PPT 模板适配 Skill。定位仍然是 PPT 生命周期的“最后一公里”模板适配层，不创作内容，只在不破坏讲师/嘉宾原始内容的前提下统一背景、字体、配色与南京模板页眉/视觉资产。适用于将已有 PPT 迁移到南京城市沙龙模板，或将 HTML/Markdown 内容直接转换为南京模板风格 PPT。南京限定色板：#3272DC、#08194B、#FFFFFF、#00C8D8、#01A4FF、#44474F、#8B8C8C。
agent_created: true
---

# 腾讯云架构师南京城市沙龙 PPT 模板适配 Skill

## 一、Skill 定位

这是一个 **PPT 模板适配 Skill**，不是内容创作器。

唯一职责：在不破坏讲师/嘉宾创作成果的前提下，让最终 PPT 继承腾讯云架构师南京城市沙龙模板的品牌资产：

1. **背景**：替换为南京模板背景图。
2. **字体**：统一为 `TencentSans W7` / `TencentSans W3`。
3. **配色**：严格限定南京 7 色色板。
4. **页眉/Logo**：不再自动插入旧模板横版 Logo；保留南京模板自身页眉与背景视觉，未来若提供南京官方 Logo 再开启。
5. **文字对比度修复**：浅蓝背景/白卡上自动转为深蓝或深灰，深蓝/主蓝色块上用白字。
6. **元素溢出检测**：检测是否溢出页面或侵入南京模板顶部页眉安全区。

**只锁品牌适配层，不主动改写内容、版式、图表数据、图片风格、动画或信息架构。**

---

## 二、核心铁律

1. 不得擅自重构页面内容、改写讲师文案、删减原始信息。
2. 输出必须是 `.pptx`。
3. 背景必须使用南京模板背景，不允许出现旧模板米黄/红色城市线稿背景。
4. 配色只能使用南京限定色板：`#3272DC`、`#08194B`、`#FFFFFF`、`#00C8D8`、`#01A4FF`、`#44474F`、`#8B8C8C`。
5. 禁止继续使用原旧红/暖色体系：`#D80C01`、`#FF0000`、`#FAF6EE`、`#D4A574`、`#FAD16A`、`#F19D19`。
6. 字体名锁定为英文 family name：`TencentSans W7` / `TencentSans W3`；代码块可使用等宽字体。
7. 南京模板默认不强制插入横版 Logo；不得误用旧 Logo 或头像占位图作为 Logo。
8. 内容页顶部左侧为页眉/标题安全区，正文内容建议从 `y >= 1.5"` 开始。
9. HTML/Markdown 转 PPT 后必须运行 `scripts/verify_output.py` 做验证。

---

## 三、模板资产清单

```text
assets/
├── templates/
│   └── nanjing-architect-salon-template.pptx
├── backgrounds/
│   ├── bg-cover.jpeg        # 南京封面/标题页背景
│   ├── bg-hero-cover.jpeg   # 南京强视觉首页背景，保留备用
│   ├── bg-section.jpeg      # 南京章节/过渡页背景
│   ├── bg-content.jpeg      # 南京内容页背景
│   └── bg-end.jpeg          # 南京结尾页背景
├── logos/                   # 当前为空；南京模板默认不强制插入横版 Logo
└── fonts/
    ├── TencentSans-W3.ttf
    ├── TencentSans-W7.ttf
    └── README.md
```

页面尺寸：20.00 × 11.25 inch，16:9。

---

## 四、南京品牌规范

### 4.1 背景映射

| 页面类型 | 背景文件 | 说明 |
|---|---|---|
| cover | `bg-cover.jpeg` | 南京标题/封面页浅蓝科技背景 |
| section | `bg-section.jpeg` | 南京章节/过渡页背景 |
| content | `bg-content.jpeg` | 南京内容页浅蓝背景 |
| end | `bg-end.jpeg` | 南京结尾页背景 |

所有背景图必须铺满整页，z-order 置底，不加滤镜、不变色、不叠加旧模板背景。

### 4.2 字体

| 用途 | family name |
|---|---|
| 标题 / 强调 / 字号 >= 24pt | `TencentSans W7` |
| 正文 | `TencentSans W3` |
| 代码块 | `JetBrains Mono` / `Consolas` / `Courier New` |

### 4.3 南京限定色板

| HEX | 用途 |
|---|---|
| `#3272DC` | 主蓝：标题强调、表头、流程箭头、关键数字 |
| `#08194B` | 深蓝：正文、深色数据卡、代码块背景 |
| `#FFFFFF` | 卡片底、深色块文字、留白区域 |
| `#00C8D8` | 青蓝强调：图表、步骤节点、少量强调 |
| `#01A4FF` | 亮蓝强调：图表、状态高亮 |
| `#44474F` | 深灰辅助：副文本、图例、说明 |
| `#8B8C8C` | 中灰辅助：注释、弱提示、边线 |

图表色板顺序：

```text
#3272DC → #08194B → #00C8D8 → #01A4FF → #44474F → #8B8C8C
```

### 4.4 对比度规则

| 背景 | 文字 |
|---|---|
| 浅蓝背景 / 白色卡片 | `#08194B` 或 `#44474F` |
| 主蓝 `#3272DC` | `#FFFFFF` |
| 深蓝 `#08194B` | `#FFFFFF` |
| 青蓝 `#00C8D8` | `#08194B` |
| 亮蓝 `#01A4FF` | `#FFFFFF` 或 `#08194B`，按可读性判断 |

---

## 五、统一入口

```bash
python scripts/apply_template.py --input <input.pptx|html|md> --output <out.pptx>
```

| 输入格式 | 路由 |
|---|---|
| `.pptx` | 保留内容并迁移到南京模板背景、字体与配色 |
| `.html/.htm/.xhtml` | 解析 HTML 结构，从零生成南京模板 PPT |
| `.md/.markdown/.mdown` | Markdown 转 HTML 后生成南京模板 PPT |

---

## 六、推荐执行流程

1. 判断输入类型：PPTX / HTML / Markdown。
2. PPTX 模式：执行 `apply_template.py --mode full`。
3. HTML/Markdown 模式：直接用 `html_to_pptx.py` 生成南京模板 PPT。
4. 生成后运行：

```bash
python scripts/verify_output.py --pptx output.pptx --strict
```

5. 人工抽查：背景是否是南京浅蓝模板、是否有旧红/暖色残留、是否误插旧 Logo。

---

## 七、异常处理

| 问题 | 处理 |
|---|---|
| 背景文件缺失 | 检查 `assets/backgrounds/` 是否含 4 类南京背景 |
| 仍有旧红/暖色 | 运行色彩合规，确认没有关闭 `--color-compliance` |
| 文字发白不可读 | 运行默认对比度修复；浅蓝/白卡文字应为 `#08194B` 或 `#44474F` |
| 出现旧 Logo | 检查 `assets/logos/`，南京模板默认不应自动插入旧横版 Logo |
| 字体显示异常 | 安装 `assets/fonts/` 下 TencentSans W3/W7 后重启 Office |

---

**记住：本 Skill 不创作 PPT，它只让 PPT 看起来属于腾讯云架构师南京城市沙龙。**
