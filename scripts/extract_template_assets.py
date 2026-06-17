#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_template_assets.py — 从南京城市沙龙模板 PPTX 中提取背景资产。

用法：
  python scripts/extract_template_assets.py --pptx "/path/to/南京模板.pptx" --dry-run
  python scripts/extract_template_assets.py --pptx "/path/to/南京模板.pptx"

说明：
- 南京模板默认不强制插入横版 Logo，因此本脚本不再自动提取 Logo。
- 根据实测模板媒体关系，优先使用 image1~image5 作为背景资产：
  image1 → bg-cover.jpeg
  image2 → bg-section.jpeg
  image3 → bg-content.jpeg
  image4 → bg-end.jpeg
  image5 → bg-hero-cover.jpeg
"""

import argparse
import shutil
import sys
import zipfile
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
SKILL_DIR = SCRIPT_DIR.parent
ASSETS = SKILL_DIR / "assets"

BACKGROUND_MAP = {
    "image1.jpeg": "bg-cover.jpeg",
    "image2.jpeg": "bg-section.jpeg",
    "image3.jpeg": "bg-content.jpeg",
    "image4.jpeg": "bg-end.jpeg",
    "image5.jpeg": "bg-hero-cover.jpeg",
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--pptx", required=True, help="南京城市沙龙模板 .pptx 路径")
    ap.add_argument("--dry-run", action="store_true", help="只打印计划，不写入")
    args = ap.parse_args()

    src = Path(args.pptx)
    if not src.exists():
        print(f"[ERROR] 不存在：{src}", file=sys.stderr)
        sys.exit(1)

    target_tpl = ASSETS / "templates" / "nanjing-architect-salon-template.pptx"
    print(f"[PLAN] 模板 → {target_tpl}")
    if not args.dry_run:
        target_tpl.parent.mkdir(parents=True, exist_ok=True)
        shutil.copy(src, target_tpl)

    bg_dir = ASSETS / "backgrounds"
    if not args.dry_run:
        bg_dir.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(src) as z:
        names = set(z.namelist())
        for media_name, output_name in BACKGROUND_MAP.items():
            member = f"ppt/media/{media_name}"
            if member not in names:
                print(f"[WARN] 未找到 {member}")
                continue
            target = bg_dir / output_name
            print(f"[PLAN] {member} → {target}")
            if not args.dry_run:
                target.write_bytes(z.read(member))

    print("[DONE]")


if __name__ == "__main__":
    main()
