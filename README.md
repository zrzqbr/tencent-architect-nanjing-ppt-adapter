<div align="center">

<img src="https://img.shields.io/badge/Python-3.9+-3776AB?style=for-the-badge&logo=python&logoColor=white" alt="Python">
<img src="https://img.shields.io/badge/License-MIT-2ea44f?style=for-the-badge" alt="License">
<img src="https://img.shields.io/badge/python--pptx-1.0+-3272DC?style=for-the-badge" alt="python-pptx">
<img src="https://img.shields.io/badge/WorkBuddy-Skill-08194B?style=for-the-badge" alt="Platform">

# Tencent Architect Nanjing PPT Adapter

**腾讯云架构师南京城市沙龙 · PPT 模板适配 Skill**

<sub>一键将 PPT / HTML / Markdown 适配为南京城市沙龙模板风格 —— 只锁品牌，不碰内容</sub>

</div>

---

## 定位

这是一个 **PPT 模板适配 Skill**，不是内容创作器。

它保留讲师/嘉宾原始内容、版式和图表，只统一：

- 南京模板背景
- TencentSans 字体
- 南京限定 7 色色板
- 南京页眉安全区与输出验证

南京模板默认不强制插入旧横版 Logo；当前 `assets/logos/` 不再保留旧模板 Logo。

---

## 南京限定色板

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

图表顺序：

```text
#3272DC → #08194B → #00C8D8 → #01A4FF → #44474F → #8B8C8C
```

禁止继续使用旧红/暖色体系：`#D80C01`、`#FF0000`、`#FAF6EE`、`#D4A574`、`#FAD16A`、`#F19D19`。

---

## 快速开始

```bash
pip install -r requirements.txt

# 迁移已有 PPT
python scripts/apply_template.py --input your.pptx --output nanjing.pptx

# HTML / Markdown 转南京模板 PPT
python scripts/apply_template.py --input article.html --output nanjing.pptx

# 验证输出
python scripts/verify_output.py --pptx nanjing.pptx --strict
```

---

## 目录结构

```text
assets/
├── templates/
│   └── nanjing-architect-salon-template.pptx
├── backgrounds/
│   ├── bg-cover.jpeg
│   ├── bg-hero-cover.jpeg
│   ├── bg-section.jpeg
│   ├── bg-content.jpeg
│   └── bg-end.jpeg
├── logos/                   # 默认不强制插入横版 Logo
└── fonts/
    ├── TencentSans-W3.ttf
    └── TencentSans-W7.ttf
scripts/
├── apply_template.py         # 统一入口
├── html_to_pptx.py           # HTML/Markdown 生成南京模板 PPT
├── brand_palette.py          # 南京限定色板工具
├── verify_output.py          # 输出验证
└── extract_template_assets.py
```

---

## 使用模式

### 1. 迁移已有 PPT

```bash
python scripts/apply_template.py --input guest.pptx --output guest_nanjing.pptx --mode full
```

会执行：背景替换、字体统一、南京色板合规、文字对比度修复、页眉安全区检测。

### 2. HTML / Markdown 生成 PPT

```bash
python scripts/apply_template.py --input article.html --output article_nanjing.pptx --title "主题"
```

会根据 H1/H2、段落、列表、表格、代码块、图片、卡片、流程图等结构生成南京模板 PPT。

---

## 给上游 PPT 生成器的配色提示词

```bash
python scripts/brand_palette.py --prompt
```

校验颜色：

```bash
python scripts/brand_palette.py --validate "#3272DC"
python scripts/brand_palette.py --validate "#D80C01"
```

---

## 设计原则

- 不改文案，不删信息。
- 不强制重排页面。
- 只在品牌资产层做适配。
- 颜色必须收敛到南京 7 色板。
- 不再出现旧模板背景和旧红/暖色系统。

---

## License

MIT License。字体、模板、品牌资产请按原权利归属和授权范围使用。
