#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""生成 pdfconv 回归测试用 PDF：多页、中文、表格线、内嵌图片。

用法: python make_samples.py <output_dir>
"""
import sys
from pathlib import Path

try:
    import pymupdf as fitz  # PyMuPDF 新包名
except ImportError:  # pragma: no cover
    import fitz


def make(path: Path):
    doc = fitz.open()
    # 第 1 页：中文 + 说明
    page = doc.new_page(width=595, height=842)  # A4
    page.insert_text((60, 80), "pdfconv 本地离线转换测试", fontsize=22, fontname="china-s")
    page.insert_text((60, 120), "这是一个用于离线回归测试的 PDF。", fontsize=13, fontname="china-s")
    page.insert_text((60, 145), "English line: offline conversion smoke test.", fontsize=12)
    page.insert_text((60, 175), "第二行中文：扫描件识别、可搜索 PDF、表格提取。", fontsize=13, fontname="china-s")

    # 第 2 页：带线表格
    page2 = doc.new_page(width=595, height=842)
    page2.insert_text((50, 60), "表格页 / Table page", fontsize=18, fontname="china-s")
    x0, y0, col_w, row_h = 50, 100, 120, 30
    rows, cols = 5, 4
    for r in range(rows + 1):
        y = y0 + r * row_h
        page2.draw_line(fitz.Point(x0, y), fitz.Point(x0 + cols * col_w, y))
    for c in range(cols + 1):
        x = x0 + c * col_w
        page2.draw_line(fitz.Point(x, y0), fitz.Point(x, y0 + rows * row_h))
    header = ["编号", "名称", "数量", "备注"]
    for c, t in enumerate(header):
        page2.insert_text((x0 + c * col_w + 6, y0 + 20), t, fontsize=11, fontname="china-s")
    data = [["1", "苹果", "10", "红色"], ["2", "香蕉", "20", "黄色"],
            ["3", "橙子", "15", "橙色"], ["4", "葡萄", "30", "紫色"]]
    for r, row in enumerate(data, start=1):
        for c, t in enumerate(row):
            page2.insert_text((x0 + c * col_w + 6, y0 + r * row_h + 20), t, fontsize=11, fontname="china-s")

    # 第 3 页：内嵌图片（用 PyMuPDF 生成一张小 PNG 再插入）
    import io
    from PIL import Image, ImageDraw
    im = Image.new("RGB", (360, 200), (240, 245, 255))
    d = ImageDraw.Draw(im)
    d.rectangle([20, 20, 340, 180], outline=(30, 90, 200), width=4)
    d.text((40, 90), "pdfconv inline image", fill=(30, 60, 140))
    buf = io.BytesIO()
    im.save(buf, format="PNG")
    page3 = doc.new_page(width=595, height=842)
    page3.insert_text((60, 70), "图片页 / Image page", fontsize=18, fontname="china-s")
    page3.insert_image(fitz.Rect(60, 100, 535, 390), stream=buf.getvalue())

    out = path / "sample-multi.pdf"
    doc.save(str(out))
    doc.close()
    print(f"wrote {out} ({out.stat().st_size} bytes)")
    return out


if __name__ == "__main__":
    target = Path(sys.argv[1] if len(sys.argv) > 1 else ".")
    target.mkdir(parents=True, exist_ok=True)
    make(target)
