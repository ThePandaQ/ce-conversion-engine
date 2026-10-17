#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 pdfconv 快捷方式图标（多尺寸 .ico）。

用法: python make_icon.py [output.ico]
"""
import sys
from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

BG = (23, 70, 110, 255)
WHITE = (255, 255, 255, 255)
RED = (200, 32, 32, 255)


def make(out: Path):
    # 优先使用已抠好的透明 PNG（背景 100% 透明），其次原参考图，最后才用脚本绘制旧版图标。
    for candidate in ("icon-source-clean.png", "icon-source.webp"):
        source = Path(__file__).with_name(candidate)
        if source.exists():
            img = Image.open(str(source)).convert("RGBA")
            out.parent.mkdir(parents=True, exist_ok=True)
            img.save(str(out), format="ICO",
                     sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
            print(f"icon written from {source.name}: {out}")
            return

    size = 256
    img = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    d = ImageDraw.Draw(img)
    # 深蓝圆角底
    d.rounded_rectangle([6, 6, 250, 250], radius=44, fill=BG)
    # 白色纸张
    d.rounded_rectangle([64, 36, 192, 220], radius=12, fill=WHITE)
    # 红色 PDF 标签
    d.rounded_rectangle([82, 156, 174, 206], radius=10, fill=RED)
    # 箭头
    d.line([(100, 66), (150, 66)], fill=BG, width=10)
    d.polygon([(150, 46), (190, 66), (150, 86)], fill=BG)
    # PDF 文字
    font = None
    for name in ("arialbd.ttf", "segoeuib.ttf", "arial.ttf"):
        try:
            font = ImageFont.truetype(name, 34)
            break
        except Exception:
            continue
    if font is None:
        font = ImageFont.load_default()
    text = "PDF"
    bbox = d.textbbox((0, 0), text, font=font)
    tw = bbox[2] - bbox[0]
    th = bbox[3] - bbox[1]
    d.text(((size - tw) / 2 - bbox[0], 181 - th / 2 - bbox[1]), text, fill=WHITE, font=font)

    out.parent.mkdir(parents=True, exist_ok=True)
    img.save(str(out), sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
    print(f"icon written: {out}")


if __name__ == "__main__":
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("pdfconv.ico")
    make(target)
