# 腾讯云架构师南京城市沙龙 PPT 模板品牌规范（Brand Rules）

本文档记录南京模板的背景、字体、配色、页眉/Logo 规则。当前项目已去除旧模板模板资产与红色暖色体系。

---

## 一、背景规范

| 文件 | 用途 | 分辨率 | 说明 |
|---|---|---:|---|
| `assets/backgrounds/bg-cover.jpeg` | 封面 / 标题页 | 1920×1080 | 南京浅蓝科技风封面背景 |
| `assets/backgrounds/bg-hero-cover.jpeg` | 强视觉首页备用 | 1920×1080 | 南京首页直接全屏图背景，备用 |
| `assets/backgrounds/bg-section.jpeg` | 章节 / 过渡页 | 1920×1080 | 南京章节页背景 |
| `assets/backgrounds/bg-content.jpeg` | 内容页 | 1920×1080 | 南京内容页浅蓝背景 |
| `assets/backgrounds/bg-end.jpeg` | 结尾页 | 1920×1080 | 南京结尾页背景 |

规则：

1. 背景必须铺满整页，z-order 置底。
2. 不加滤镜、不变色、不叠加其他城市背景。
3. 不允许再出现旧米黄红色城市线稿背景。
4. 全屏代码/截图页如必须使用纯色背景，只允许 `#FFFFFF` 或 `#08194B`，并保证文字对比度。

---

## 二、字体规范

| 用途 | family name |
|---|---|
| 标题 / 强调 | `TencentSans W7` |
| 正文 | `TencentSans W3` |
| 代码 | `JetBrains Mono` / `Consolas` / `Courier New` |

字号建议：

| 元素 | 推荐字号 | 最小字号 |
|---|---:|---:|
| 封面主标题 | 56–72pt | 48pt |
| 封面副标题 | 24–32pt | 20pt |
| 章节标题 | 48–60pt | 40pt |
| 内容页标题 | 32–40pt | 28pt |
| 正文 | 18–22pt | 16pt |
| 注释 | 12–14pt | 12pt |

---

## 三、南京限定色板

只能使用以下 7 个颜色：

| 色名 | HEX | 用途 |
|---|---|---|
| 南京主蓝 | `#3272DC` | 标题强调、表头、流程箭头、关键数字 |
| 南京深蓝 | `#08194B` | 正文、深色块、代码块背景 |
| 纯白 | `#FFFFFF` | 卡片底、深色块文字 |
| 青蓝强调 | `#00C8D8` | 图表系列、步骤节点、少量强调 |
| 亮蓝强调 | `#01A4FF` | 图表系列、状态高亮 |
| 深灰辅助 | `#44474F` | 副文本、说明、图例 |
| 中灰辅助 | `#8B8C8C` | 注释、弱提示、边线 |

图表顺序：

```text
#3272DC → #08194B → #00C8D8 → #01A4FF → #44474F → #8B8C8C
```

禁止：

- 旧红/暖色体系：`#D80C01`、`#FF0000`、`#FAF6EE`、`#D4A574`、`#FAD16A`、`#F19D19`
- 绿色、粉色、紫色、彩色渐变、玻璃拟态、霓虹荧光色
- 不在南京 7 色板中的任意大面积色块

---

## 四、色彩对比规则

| 背景 | 推荐文字 |
|---|---|
| 浅蓝模板背景 / 白卡 | `#08194B` / `#44474F` |
| 主蓝 `#3272DC` | `#FFFFFF` |
| 深蓝 `#08194B` | `#FFFFFF` |
| 青蓝 `#00C8D8` | `#08194B` |
| 亮蓝 `#01A4FF` | `#FFFFFF` 或 `#08194B`，取更清晰者 |

正文文字必须 100% 不透明。卡片可以使用纯白，不做玻璃拟态。

---

## 五、页眉 / Logo 规则

南京模板当前不强制插入旧横版 Logo。`assets/logos/` 目录不再放置旧模板 Logo。

顶部左侧为南京模板页眉/标题安全区：

```text
x = 1.20" – 6.00"
y = 0.35" – 1.15"
```

正文内容建议从 `y >= 1.50"` 开始，避免侵入页眉。

如果未来提供南京官方横版 Logo，应先实测模板 XML 坐标，再启用 `add_logo()`。

---

## 六、页面类型映射

| 页面类型 | 背景 |
|---|---|
| cover | `bg-cover.jpeg` |
| section | `bg-section.jpeg` |
| content | `bg-content.jpeg` |
| end | `bg-end.jpeg` |

识别优先级：

1. Slide Layout 名称。
2. 首页/尾页规则。
3. 文本数量、字号、图表/表格启发式。

---

## 七、验证重点

1. 所有页背景是否为南京浅蓝模板。
2. 是否仍残留旧红/暖色。
3. 是否误插旧横版 Logo。
4. 字体是否统一为 TencentSans W7/W3。
5. 浅蓝背景与白卡上的文字是否足够清晰。
6. 内容是否侵入顶部页眉安全区。

---

## 八、架构级规范（v5.0+）

### 8.1 执行层次

适配操作必须严格按以下层次从上到下执行，影响范围从大到小：

```
第一层：主题层（Theme）
  → 覆写 theme1.xml 的 clrScheme 全部 12 色槽
  → 覆写 theme1.xml 的 majorFont / minorFont
  → 将 TTF 字体嵌入 ppt/fonts/ 并注册 Content_Types

第二层：母版层（SlideMaster）
  → 设置 slideMaster 的 bgPr 为品牌背景图片
  → 设置 slideMaster 的默认文字颜色

第三层：页面层（Slide）
  → 按页面类型（cover/section/content/end）覆盖背景
  → 清除全屏遮罩、修复对比度

第四层：元素层（Shape/Run）
  → 逐 run 字体兜底（处理未继承主题字体的显式 typeface）
  → 逐形状色块合规替换（处理未引用 schemeClr 的硬编码 srgbClr）
```

### 8.2 主题色槽映射（12 槽 → 南京品牌色）

| 色槽 | 品牌色 | 语义 |
|---|---|---|
| dk1 (tx1) | `#08194B` | 正文默认色 |
| lt1 (bg1) | `#FFFFFF` | 背景色 |
| dk2 | `#44474F` | 二级深色 |
| lt2 | `#8B8C8C` | 二级浅色 |
| accent1 | `#3272DC` | 南京主蓝 |
| accent2 | `#08194B` | 南京深蓝 |
| accent3 | `#00C8D8` | 青蓝强调 |
| accent4 | `#01A4FF` | 亮蓝强调 |
| accent5 | `#44474F` | 深灰辅助 |
| accent6 | `#8B8C8C` | 中灰辅助 |
| hlink | `#3272DC` | 超链接 |
| folHlink | `#08194B` | 已访问超链接 |

### 8.3 字体嵌入规范

**只写 typeface 名称不等于字体可用。** 必须同时完成：

1. 修改 `theme1.xml` 的 `majorFont` / `minorFont` → `TencentSans W7` / `TencentSans W3`
2. 将 `assets/fonts/TencentSans-W7.ttf` 和 `TencentSans-W3.ttf` 写入 `ppt/fonts/` 目录
3. 在 `presentation.xml` 中添加 `<p:embeddedFontLst>` 声明
4. 在 `presentation.xml.rels` 中添加 font relationship
5. 确保 `[Content_Types].xml` 包含 `application/x-fontdata` 类型

### 8.4 颜色还原要求

适配后颜色不能"一刀切"：

- 不得将所有文字统一染成同一个深蓝色，必须保留原始的颜色层次
- 颜色替换后，按语义重新分配层次：标题用 `#3272DC`，正文用 `#08194B`，辅助用 `#44474F`，弱化用 `#8B8C8C`
- 深色背景上的文字保持白色 `#FFFFFF`，不要反转
