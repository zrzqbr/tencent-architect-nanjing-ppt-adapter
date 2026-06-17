# Changelog

本文件遵循 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/) 规范。

## [v4.0-nanjing] - 2026-06-17

### Changed

- 将项目完全切换为腾讯云架构师南京城市沙龙 PPT 模板适配器。
- 移除旧模板的横版 Logo 自动插入逻辑，南京模板默认保留自身页眉与背景视觉。
- 将配色体系收敛到南京限定 7 色：`#3272DC`、`#08194B`、`#FFFFFF`、`#00C8D8`、`#01A4FF`、`#44474F`、`#8B8C8C`。
- 背景映射从 cover/content 两类扩展为 cover/section/content/end，并保留 hero cover 备用背景。
- 更新 `SKILL.md`、`README.md`、`references/brand-rules.md`、`CONTRIBUTING.md` 为南京模板规范。
- 更新 `brand_palette.py`，旧红/暖色不再视为安全色。

### Removed

- 移除旧模板文件 `assets/templates/changsha-architect-salon-template.pptx`。
- 移除旧横版 Logo 资产 `assets/logos/logo-main.png` 与 `assets/logos/logo-corner.png`。
- 去除旧红/暖色体系作为默认配色的文档说明与生成规则。

---

## 版本命名规则

- **主版本号 (Major)**：模板城市或品牌资产体系发生重大切换。
- **次版本号 (Minor)**：新增功能或文档大幅更新。
- **修订号 (Patch)**：Bug 修复、文档小修、项目规范化。
