---
name: 腾讯云架构师南京城市沙龙PPT模板适配Skill
version: "4.1-nanjing-adaptive"
description: 腾讯云架构师南京城市沙龙 PPT 模板适配 Skill。用于将已有 PPT、HTML slide deck 或 Markdown 迁移为南京城市沙龙模板风格。执行时必须先分析用户给到的源文件和模板语义，不得固化套用某一份 PPT 的判断；根据页面类型、组件复杂度和内容密度自适应选择背景、字体大小、渲染模式与验证策略。南京限定色板：#3272DC、#08194B、#FFFFFF、#00C8D8、#01A4FF、#44474F、#8B8C8C。
agent_created: true
---

# 腾讯云架构师南京城市沙龙 PPT 模板适配 Skill

## 一、Skill 定位

这是一个 **PPT 模板适配 Skill**，不是内容创作器。

目标是在不改写讲师/嘉宾内容的前提下，将 PPT、HTML 或 Markdown 迁移为腾讯云架构师南京城市沙龙视觉体系。

核心适配层：

1. **背景**：按输入页面语义动态映射为 `cover / section / content / end`，不得简单全页套同一张背景。
2. **字体**：优先使用 `TencentSans W7` / `TencentSans W3`；HTML 高保真模式必须注入字体文件渲染。
3. **字号**：按内容密度自适应放大或微调，不得固定使用一套字号导致空页过小、密集页溢出。
4. **配色**：严格限定南京 7 色色板。
5. **组件**：保留用户源文件中的卡片、网格、流程、时间线、产品卡、问答卡等结构语义。
6. **Logo/页眉**：南京模板默认不强制插入旧横版 Logo；保留南京背景自身视觉资产。
7. **验证**：生成后必须检查页数、背景类型、字体、颜色、组件完整性。

**只锁品牌适配层，不主动改写内容、删减信息、篡改图表数据或改变演讲逻辑。**

---

## 二、执行铁律

1. **先分析输入，再动手转换**：每次读取用户给到的 PPT/HTML/Markdown，判断页面类型、组件种类、内容密度、文字长度、是否已有封面/结尾/目录。
2. **不得固化某一份案例**：不要把某个客户 PPT 的页数、组件、标题、背景顺序写死到规则中。
3. **不得只看验证 PASS**：`verify_output.py` 只说明技术项通过，不代表视觉质量合格；必须抽查封面、复杂组件页、结尾页。
4. **复杂 HTML slide deck 优先高保真模式**：当检测到大量自定义 CSS 组件时，优先使用 Chrome 渲染保留组件结构，而不是用 python-pptx 抽文本重排。
5. **可编辑性与保真度要说明取舍**：高保真模式输出视觉更准，但内容层为图片；可编辑模式输出文本可编辑，但组件还原有限。
6. **字体大小必须自适应**：根据页面内容密度设置字体缩放。稀疏页可明显放大；中等内容页适度放大；高密度页只微增以防溢出。
7. **背景必须多类型映射**：封面、目录/章节、内容、结尾至少分开判断；不能因为解析错误全部落到 content 背景。
8. **南京色板不可突破**：`#3272DC`、`#08194B`、`#FFFFFF`、`#00C8D8`、`#01A4FF`、`#44474F`、`#8B8C8C`。
9. **旧红/暖色禁用**：禁止继续使用 `#D80C01`、`#FF0000`、`#FAF6EE`、`#D4A574`、`#FAD16A`、`#F19D19`。

---

## 三、模板资产清单

```text
assets/
├── templates/
│   └── nanjing-architect-salon-template.pptx
├── backgrounds/
│   ├── bg-cover.jpeg        # 南京封面/标题页背景
│   ├── bg-hero-cover.jpeg   # 强视觉首页备用背景
│   ├── bg-section.jpeg      # 章节/目录/过渡页背景
│   ├── bg-content.jpeg      # 内容页背景
│   └── bg-end.jpeg          # 结尾页背景
├── logos/                   # 当前为空；南京模板默认不强制插入旧横版 Logo
└── fonts/
    ├── TencentSans-W3.ttf
    ├── TencentSans-W7.ttf
    └── README.md
```

页面尺寸：20.00 × 11.25 inch，16:9。

---

## 四、推荐执行流程

### 4.1 统一入口

```bash
python scripts/apply_template.py --input <input.pptx|html|md> --output <out.pptx>
```

### 4.2 输入路由

| 输入格式 | 推荐模式 | 说明 |
|---|---|---|
| `.pptx` | PPTX 迁移适配 | 保留内容，替换背景、字体、配色 |
| HTML slide deck | `--html-render-mode auto` | 优先高保真 Chrome 渲染；失败回退可编辑模式 |
| HTML 文章型 | 可编辑结构化生成 | 按标题、段落、列表、表格生成 PPT |
| Markdown | 可编辑结构化生成 | 先转 HTML，再生成 PPT |

### 4.3 HTML 渲染模式

```bash
# 默认：自动判断，复杂 slide deck 优先高保真
python scripts/apply_template.py --input input.html --output out.pptx

# 强制高保真：适合复杂 CSS 卡片/网格/流程/时间线页面
python scripts/apply_template.py --input input.html --output out.pptx --html-render-mode high-fidelity

# 强制可编辑：适合简单文章/列表，后续要在 PPT 中继续编辑文字
python scripts/apply_template.py --input input.html --output out.pptx --html-render-mode editable
```

高保真模式依赖：

- Chrome / Edge / Chromium
- Node.js
- `puppeteer-core`（可执行 `npm install` 安装 optional dependency）

---

## 五、动态分析规则

### 5.1 页面类型识别

按以下优先级判断：

1. `cover`：第一页；或含 `.cover-layout`、`.cover-title`、`.cover-eyebrow`、`title-page`、`slide-title`。
2. `end`：最后一页且包含“感谢 / 谢谢 / 结语 / Thanks / Thank you / Q&A”；或 class 含 `slide-end`。
3. `section`：含 `.section-num`、`.section-title`、`.section-heading`；或标题包含“目录 / 议程 / 分享结构 / Agenda”。
4. `content`：其余页面。

不得只靠单一 class 判断，必须结合页码、文本和组件结构。

### 5.2 内容密度与字体缩放

根据每页文本长度和组件数量动态设置字体缩放：

| 密度 | 参考条件 | 字体策略 |
|---|---|---|
| cover | 封面 | 标题/副标题放大，突出主题 |
| sparse | 文本少、组件少 | 明显放大，避免空、弱、小 |
| medium | 常规内容页 | 适度放大 |
| rich | 多卡片/多列表 | 小幅放大，保持可读 |
| dense | 文本极多或组件很多 | 仅微增，优先防溢出 |
| end | 结尾页 | 适度放大，增强收束感 |

原则：**字体大小随着内容丰富度自适应，而不是全局固定。** 页面越空，越应放大；页面越密，越应谨慎放大并保持层级清晰。

### 5.3 组件识别

高保真模式必须保留以下组件的视觉结构：

```text
.card
.stat-block
.product-card
.scene-card
.question-card
.sales-step
.timeline-item
.tech-list-item
.info-card
.banner
.flow-step
.summary-card
.policy-card
.highlight-box
table / img / svg / canvas
```

可编辑模式若无法覆盖这些组件，必须在最终说明中提示“该页建议使用高保真模式”。

---

## 六、南京品牌规范

### 6.1 背景映射

| 页面类型 | 背景文件 |
|---|---|
| cover | `bg-cover.jpeg` |
| section | `bg-section.jpeg` |
| content | `bg-content.jpeg` |
| end | `bg-end.jpeg` |

### 6.2 字体

| 用途 | family name |
|---|---|
| 标题 / 强调 / 大字号 | `TencentSans W7` |
| 正文 | `TencentSans W3` |
| 代码块 | `JetBrains Mono` / `Consolas` / `Courier New` |

### 6.3 色板

| HEX | 用途 |
|---|---|
| `#3272DC` | 主蓝：标题强调、表头、流程箭头、关键数字 |
| `#08194B` | 深蓝：正文、深色数据卡、代码块背景 |
| `#FFFFFF` | 卡片底、深色块文字、留白区域 |
| `#00C8D8` | 青蓝强调：图表、步骤节点、少量强调 |
| `#01A4FF` | 亮蓝强调：图表、状态高亮 |
| `#44474F` | 深灰辅助：副文本、图例、说明 |
| `#8B8C8C` | 中灰辅助：注释、弱提示、边线 |

---

## 七、验证要求

生成后至少执行：

```bash
python scripts/verify_output.py --pptx output.pptx --strict
```

并人工抽查：

1. 页数是否和源 slide deck 一致。
2. 封面是否使用封面背景。
3. 目录/章节页是否使用 section 背景。
4. 结尾页是否使用 end 背景。
5. 复杂卡片/组件是否保留视觉结构。
6. 字体是否显著比默认可读，是否存在过小文本。
7. 是否有旧红/暖色残留。
8. 是否误插旧 Logo。

---

## 八、异常处理

| 问题 | 处理 |
|---|---|
| 高保真模式提示缺少 `puppeteer-core` | 执行 `npm install` 或在 WorkBuddy 隔离 Node workspace 安装 `puppeteer-core` |
| 高保真模式提示找不到 Chrome | 安装 Chrome/Edge/Chromium，或设置 `CHROME_PATH` |
| 字体显示异常 | 高保真模式会注入 `assets/fonts/` 字体；可编辑模式需本机安装 TencentSans |
| 页数比源 HTML 多 | 检查是否误匹配 `slide-header/slide-body/slide-footer`，必须精确匹配独立 `slide` class token |
| 组件变形严重 | 改用 `--html-render-mode high-fidelity` |
| 需要后续编辑文字 | 改用 `--html-render-mode editable`，但接受组件还原度下降 |

---

**记住：每次都根据用户给到的文件重新分析，不固化页面数量、组件结构、标题位置或字体大小。**
