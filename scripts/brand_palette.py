#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
brand_palette.py — 南京模板配色协作辅助模块

提供给上游 PPT Skill 使用：
  1) 输出南京限定安全色板 prompt 片段
  2) 校验给定 HEX 是否属于南京限定色板
  3) 提供 Python dict 形式的色板，便于程序化生成图表配色
"""

import argparse
import re
import sys

# 用户指定的南京限定色板：只允许这 7 个颜色作为品牌表达色。
CORE_COLORS = {
    "primary_blue": "#3272DC",
    "deep_navy": "#08194B",
    "white": "#FFFFFF",
    "cyan": "#00C8D8",
    "bright_blue": "#01A4FF",
    "dark_gray": "#44474F",
    "mid_gray": "#8B8C8C",
}

CHART_PALETTE = [
    "#3272DC",  # 系列 1 — 南京主蓝
    "#08194B",  # 系列 2 — 南京深蓝
    "#00C8D8",  # 系列 3 — 青蓝强调
    "#01A4FF",  # 系列 4 — 亮蓝强调
    "#44474F",  # 系列 5 — 深灰辅助
    "#8B8C8C",  # 系列 6 — 中灰辅助
]

CARD_RULES = {
    "default": {"bg": "#FFFFFF", "text": "#08194B"},
    "emphasis": {"bg": "#3272DC", "text": "#FFFFFF"},
    "data": {"bg": "#08194B", "text": "#FFFFFF"},
    "accent": {"bg": "#00C8D8", "text": "#08194B"},
    "secondary": {"bg": "#01A4FF", "text": "#FFFFFF"},
    "muted": {"bg": "#44474F", "text": "#FFFFFF"},
}

ALLOWED_LIST = set(CORE_COLORS.values())


def normalize_hex(hex_color: str) -> str:
    """标准化 HEX 字符串。"""
    hex_color = hex_color.strip().upper()
    if not hex_color.startswith("#"):
        hex_color = "#" + hex_color
    return hex_color


def validate_color(hex_color: str) -> dict:
    """校验颜色是否属于南京限定色板。返回 {'safe': bool, 'reason': str}。"""
    hex_color = normalize_hex(hex_color)
    if not re.fullmatch(r"#[A-F0-9]{6}", hex_color):
        return {"safe": False, "reason": "格式非法（应为 #RRGGBB）"}

    if hex_color in ALLOWED_LIST:
        return {"safe": True, "reason": "在南京限定色板中"}

    return {
        "safe": False,
        "reason": "不在南京限定色板中；请改用 #3272DC、#08194B、#FFFFFF、#00C8D8、#01A4FF、#44474F、#8B8C8C",
    }


def get_prompt_snippet() -> str:
    """返回注入给上游 PPT Skill 的南京配色 + 背景可见性约束 prompt。"""
    return """配色约束（南京模板，强制遵守，不可偏离）：
- 只能使用以下 7 个颜色：#3272DC、#08194B、#FFFFFF、#00C8D8、#01A4FF、#44474F、#8B8C8C
- 主色：#3272DC（南京主蓝），用于标题强调、流程箭头、表头、关键数字
- 深色：#08194B（南京深蓝），用于正文、深色数据卡、深色背景代码块
- 白色：#FFFFFF，用于卡片底、深色块文字、留白区域
- 辅助强调：#00C8D8、#01A4FF，用于图表、步骤节点、少量强调，不做大面积铺底
- 中性辅助：#44474F、#8B8C8C，用于副文本、注释、边线、弱化说明
- 图表多系列按顺序：#3272DC → #08194B → #00C8D8 → #01A4FF → #44474F → #8B8C8C
- 禁止继续使用旧红/暖色体系：#D80C01、#FF0000、#FAF6EE、#D4A574、#FAD16A、#F19D19
- 禁止使用不在色板内的绿色、粉色、紫色、彩色渐变、玻璃拟态、霓虹荧光色

背景可见性约束：
- 不要给 slide 另设背景色，保持透明，让南京浅蓝模板背景完整显示
- 内容页优先使用 #FFFFFF 卡片，卡片间保留 0.2-0.4 inch 间隙
- 浅蓝背景/白卡上的正文用 #08194B 或 #44474F；主蓝/深蓝色块上的文字用 #FFFFFF
- 不再自动插入旧横版 Logo；保留南京模板自身页眉与背景视觉
"""


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--prompt", action="store_true", help="输出南京配色约束 prompt（用于注入上游 Skill）")
    ap.add_argument("--validate", help="校验某个 HEX 颜色是否属于南京限定色板")
    ap.add_argument("--chart", action="store_true", help="输出图表色板")
    ap.add_argument("--list", action="store_true", help="列出全部南京安全色")
    args = ap.parse_args()

    if args.prompt:
        print(get_prompt_snippet())
    elif args.validate:
        r = validate_color(args.validate)
        flag = "[SAFE]" if r["safe"] else "[UNSAFE]"
        print(f"{flag}：{args.validate} — {r['reason']}")
        sys.exit(0 if r["safe"] else 1)
    elif args.chart:
        print("南京图表色板（按系列顺序）：")
        for i, c in enumerate(CHART_PALETTE, 1):
            print(f"  系列 {i}: {c}")
    elif args.list:
        print("Nanjing Colors:")
        for k, v in CORE_COLORS.items():
            print(f"  {v}  {k}")
        print("\nCard Rules:")
        for k, v in CARD_RULES.items():
            print(f"  {k}: bg={v['bg']}, text={v['text']}")
    else:
        ap.print_help()


if __name__ == "__main__":
    main()
