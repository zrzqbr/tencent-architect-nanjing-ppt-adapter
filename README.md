<div align="center">

<img src="https://img.shields.io/badge/Python-3.9+-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python">
<img src="https://img.shields.io/badge/License-MIT-2ea44f?style=for-the-badge" alt="License">
<img src="https://img.shields.io/badge/python--pptx-1.0+-3272DC?style=for-the-badge" alt="python-pptx">
<img src="https://img.shields.io/badge/WorkBuddy-Skill-08194B?style=for-the-badge" alt="Platform">

<br><br>

# 🏛️ Tencent Architect Nanjing PPT Adapter

**腾讯云架构师南京城市沙龙 · PPT 模板适配 Skill**

<sub>一键将 PPT / HTML slide deck 适配为腾讯云架构师南京城市沙龙风格 — 动态分析输入，只锁品牌，不碰内容</sub>

<br>

[🚀 快速开始](#-快速开始) · [✨ 核心能力](#-四大核心能力) · [📋 使用模式](#-两种使用模式) · [💬 提示词](#-提示词模板) · [📁 结构](#-目录结构)

<br>

---

</div>

## 💡 Skill 定位

> **这是一个 PPT 模板适配 Skill，不是内容创作器。**
>
> 它的唯一职责：在不破坏讲师/嘉宾创作成果的前提下，让最终 PPT 继承统一模板的品牌资产。
> 适用于沙龙分享、技术讲座、架构评审等场景。
>
> **核心原则：** 每次先分析用户给到的 PPT / HTML，而不是固化某一份案例；只锁品牌资产（背景、字体、配色），其他一切不强制。

<table>
  <tr>
    <td>✅ 锁定</td>
    <td><strong>背景</strong> · <strong>字体</strong> · <strong>配色</strong></td>
  </tr>
  <tr>
    <td>🚫 不碰</td>
    <td>内容文案 · 页面版式 · 图表数据 · 动画效果</td>
  </tr>
</table>

---

## 🎨 南京限定色板

只能使用以下 7 个颜色：

```text
#3272DC  南京主蓝
#08194B  南京深蓝
#FFFFFF  纯白
#00C8D8  青蓝强调
#01A4FF  亮蓝强调
#44474F  深灰辅助
#8B8C8C  中灰辅助
```

图表配色顺序：

```text
#3272DC → #08194B → #00C8D8 → #01A4FF → #44474F → #8B8C8C
```

禁止使用旧红/暖色体系：`#D80C01`、`#FF0000`、`#FAF6EE`、`#D4A574`、`#FAD16A`、`#F19D19`。

---

## ✨ 四大核心能力

<table>
  <tr>
    <th align="center" width="25%">🧠 动态分析</th>
    <th align="center" width="25%">🎨 背景适配</th>
    <th align="center" width="25%">🔤 字体自适应</th>
    <th align="center" width="25%">🧩 组件保真</th>
  </tr>
  <tr>
    <td align="center">
      每次读取源文件<br>
      分析页数 / 组件 / 内容密度<br>
      不固化某一份 PPT<br>
      动态选择策略
    </td>
    <td align="center">
      自动识别页面类型<br>
      封面 / 章节 / 内容 / 结尾<br>
      分别映射南京背景<br>
      避免全页单一背景
    </td>
    <td align="center">
      TencentSans W7 / W3<br>
      按内容丰富度调节字号<br>
      稀疏页放大<br>
      密集页防溢出
    </td>
    <td align="center">
      Chrome 高保真渲染<br>
      保留卡片 / 网格 / 流程<br>
      支持可编辑回退模式<br>
      兼顾质量与可维护
    </td>
  </tr>
</table>

---

## 🚀 快速开始

```bash
# 1️⃣ 克隆仓库
git clone https://github.com/zrzqbr/tencent-architect-nanjing-ppt-adapter.git
cd tencent-architect-nanjing-ppt-adapter

# 2️⃣ 安装 Python 依赖
pip install -r requirements.txt

# 可选：HTML slide deck 高保真模式需要 puppeteer-core + Chrome/Edge
npm install

# 3️⃣ 运行适配（auto 会对复杂 HTML 优先使用高保真模式）
python scripts/apply_template.py --input your.pptx --output branded.pptx
python scripts/apply_template.py --input your.html --output branded.pptx --html-render-mode auto

# 4️⃣ 验证输出
python scripts/verify_output.py --pptx branded.pptx --strict
```

<table>
  <tr>
    <td><strong>环境要求</strong></td>
    <td>Python 3.9+ &nbsp;·&nbsp; pip；高保真 HTML 模式需 Node.js + Chrome/Edge</td>
  </tr>
  <tr>
    <td><strong>核心依赖</strong></td>
    <td><code>python-pptx</code> · <code>lxml</code> · <code>Pillow</code> · <code>beautifulsoup4</code></td>
  </tr>
</table>

---

## 📋 两种使用模式

<table>
  <tr>
    <th width="50%">⭐ 模式 A：常用 PPT 生成 Skill + 适配 Skill 搭配使用</th>
    <th width="50%">🔄 模式 B：迁移已有 PPT</th>
  </tr>
  <tr>
    <td align="center">
      <br>
      <code>输入主题 + 大纲</code><br>
      &nbsp;&nbsp;⬇️<br>
      <code>常用 PPT 生成 Skill + 本适配 Skill 协同工作</code><br>
      &nbsp;&nbsp;⬇️<br>
      <code>输出 .pptx</code>
      <br><br>
    </td>
    <td align="center">
      <br>
      <code>输入原始 .pptx</code><br>
      &nbsp;&nbsp;⬇️<br>
      <code>模板适配</code><br>
      &nbsp;&nbsp;⬇️<br>
      <code>输出适配后 .pptx</code>
      <br><br>
    </td>
  </tr>
  <tr>
    <td align="center"><sub>内容创作与品牌风格适配各司其职，没有先后顺序限制</sub></td>
    <td align="center"><sub>嘉宾/讲师已有 PPT，直接由本 Skill 完成适配</sub></td>
  </tr>
</table>

---

## 💬 提示词模板（直接复制使用）

> 💡 **使用建议**：将你常用的 PPT 生成 Skill 与本适配 Skill 搭配使用，内容创作与品牌风格适配各司其职，没有先后顺序限制。

<details>
<summary><strong>📝 场景一：配合其他 Skill 适配专属腾讯云架构师南京城市沙龙 PPT 风格</strong></summary>

<br>

> 💡 适用场景：将你常用的 PPT 生成 Skill 与本适配 Skill 搭配使用

```text
请帮我制作一份腾讯云架构师城市沙龙 PPT。

主题：[在此填写你的主题，例如：云原生架构落地实践]
大纲：
1. [章节一标题]
2. [章节二标题]
3. [章节三标题]
4. [章节四标题]

要求：
- 搭配我常用的 PPT 生成 Skill 与腾讯云架构师南京城市沙龙 PPT 模板适配 Skill 协同完成
- 内容详实，结构清晰
- 包含完整页面结构：封面页、目录页、章节扉页、结尾页
- 统一南京品牌配色规范（南京限定 7 色板）
```

</details>

<details>
<summary><strong>📝 场景二：迁移已有 PPT 到南京品牌模板</strong></summary>

<br>

> 💡 适用场景：嘉宾/讲师已有 PPT，需要套用统一模板，直接由本 Skill 完成适配

```text
请帮我将附件中的 PPT 迁移到腾讯云架构师南京城市沙龙模板。

要求：
- 调用腾讯云架构师南京城市沙龙 PPT 模板适配 Skill 完成迁移
- 保留原始 PPT 的所有内容、版式和图表，不要修改文案
- 替换背景为南京品牌模板背景
- 字体统一为 TencentSans
- 检查配色合规性，替换禁用色为南京限定色板
```

</details>

---

## 🏷️ 触发场景

以下关键词会触发本 Skill：

<p>
  <img src="https://img.shields.io/badge/腾讯云架构师南京城市沙龙模板-3272DC?style=flat-square" alt="">
  <img src="https://img.shields.io/badge/南京沙龙%20PPT-3272DC?style=flat-square" alt="">
  <img src="https://img.shields.io/badge/把%20PPT%20套到南京模板-01A4FF?style=flat-square" alt="">
  <img src="https://img.shields.io/badge/嘉宾%20PPT%20换统一模板-01A4FF?style=flat-square" alt="">
  <img src="https://img.shields.io/badge/南京沙龙模板适配-08194B?style=flat-square" alt="">
  <img src="https://img.shields.io/badge/保留内容只替换背景%2F字体%2F配色-08194B?style=flat-square" alt="">
</p>

> 💡 **提示：** 可以直接把上面的提示词模板复制到 WorkBuddy 对话框中使用，填入你的具体内容即可一键完成 PPT 制作或迁移。

---

## ⚠️ 注意事项

| | 说明 |
|:---:|------|
| 🔒 | 本 Skill **只做风格适配**，不会修改你的文案内容和页面布局 |
| 📂 | 支持多种输入格式，包括 `.pptx`、`HTML`、`Markdown` 等 |
| 📎 | 迁移模式需将原始 PPT 作为附件上传到对话中 |
| ⚖️ | TencentSans 字体为腾讯品牌字体，请确认使用授权 |
| 🚫 | 南京模板默认不强制插入旧横版 Logo |
| 🧩 | 复杂 HTML 默认优先高保真渲染，组件还原更好，但文字层是图片；如需继续编辑文字，可使用 `--html-render-mode editable` |

---

## 📁 目录结构

```
tencent-architect-nanjing-ppt-adapter/
│
├── 📋 SKILL.md                         # 主文档（完整迁移经验 & 品牌规范）
├── 📖 README.md                        # 本文件
├── 📜 LICENSE                          # MIT 开源协议
├── 📦 requirements.txt                 # Python 依赖清单
│
├── 🔧 scripts/                         # 核心脚本
│   ├── apply_template.py               # 统一入口 · 模板适配引擎
│   ├── html_to_pptx.py                 # HTML → PPTX 可编辑结构化转换
│   ├── high_fidelity_html_to_pptx.py   # HTML slide deck 高保真渲染转换
│   ├── verify_output.py                # 输出质量验证
│   ├── brand_palette.py                # 南京限定色板校验 + 上游约束注入
│   ├── extract_template_assets.py      # 模板资产提取工具
│   └── test_v5.py                      # 自动化测试套件
│
├── 🎨 assets/                          # 品牌资产
│   ├── backgrounds/                    # 南京品牌背景图（封面/章节/内容/结尾）
│   ├── logos/                          # 默认不强制插入横版 Logo
│   ├── fonts/                          # TencentSans W3 / W7
│   └── templates/                      # 南京城市沙龙模板源文件
│
└── 📚 references/                      # 参考文档
    ├── brand-rules.md                  # 南京品牌规范详细补充
    └── changelog-v1-v5.md              # 版本演进 & 实战经验沉淀
```

---

## 🤝 贡献

欢迎提交 Issue 和 Pull Request！无论是 Bug 反馈、功能建议还是代码贡献，都非常感谢。

---

## 📬 反馈

- **GitHub Issues：** [提交问题](https://github.com/zrzqbr/tencent-architect-nanjing-ppt-adapter/issues)

---

<div align="center">

<sub>

**本 Skill 不创作 PPT，它让 PPT 形成专属腾讯云架构师南京城市沙龙的统一风格。**

</sub>

</div>
