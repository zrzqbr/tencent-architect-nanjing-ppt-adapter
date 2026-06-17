# 贡献指南

欢迎通过 Issue 和 Pull Request 参与改进南京城市沙龙 PPT 模板适配器。

## 提交 Issue

请包含以下信息：

- Python 版本和操作系统
- 输入文件类型（PPTX / HTML / Markdown）
- 完整错误信息、截图或最小复现文件
- 预期行为与实际行为

## 提交 Pull Request

1. Fork 本仓库
2. 创建功能分支：`git checkout -b feature/your-feature`
3. 提交修改：`git commit -m "feat: add your feature"`
4. 推送分支：`git push origin feature/your-feature`
5. 创建 Pull Request，说明改动内容和原因

## 代码风格

- Python 代码遵循 PEP 8
- 函数和类应有 docstring
- 变量命名使用 snake_case
- 不要引入不必要的新依赖

## 品牌资产修改禁区

以下文件受南京模板品牌规范约束，不可随意修改：

| 文件 / 目录 | 说明 |
|---|---|
| `assets/backgrounds/` | 南京模板背景图 |
| `assets/templates/` | 南京基准模板 .pptx |
| `assets/fonts/` | TencentSans 字体文件 |
| `references/brand-rules.md` | 南京品牌配色/排版硬规范 |

当前南京模板默认不强制插入横版 Logo，`assets/logos/` 不应放入旧模板 Logo。若未来接入官方南京 Logo，必须先实测模板 XML 坐标并更新验证规则。

## 测试

提交前请确保：

```bash
python -m py_compile scripts/*.py
python scripts/brand_palette.py --validate "#3272DC"
python scripts/brand_palette.py --validate "#D80C01"  # 应返回 unsafe
python scripts/verify_output.py --pptx your-output.pptx --strict
```

## 许可

提交贡献即表示你同意将代码以 MIT License 授权发布。字体、模板、品牌资产请按原权利归属和授权范围使用。
