#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""pdfconv 1.0.1 — CE转换引擎（Conversion Engine）：本地离线 PDF/OFD/WPS 转换工具（开源引擎实现，不依赖 WPS）。
作者：温酒会知音    版权所有 © 2026 温酒会知音

用法示例：
  pdfconv info report.pdf
  pdfconv pdf2png report.pdf --dpi 150 --pages 1-3
  pdfconv pdf2txt report.pdf -o report.txt
  pdfconv pdf2docx report.pdf -o report.docx
  pdfconv pdf2pptx report.pdf -o report.pptx
  pdfconv pdf2xlsx report.pdf -o report.xlsx
  pdfconv ocr scan.pdf -o scan.ocr.txt
  pdfconv pdf2searchable scan.pdf -o scan-searchable.pdf
  pdfconv img2pdf a.png b.jpg -o photos.pdf
  pdfconv merge a.pdf b.pdf -o merged.pdf
  pdfconv split big.pdf --every 10 -o split/
  pdfconv extract-pages big.pdf --pages 2-5,9 -o pages.pdf
  pdfconv extract-images big.pdf -o images/
  pdfconv ofd2pdf notice.ofd -o notice.pdf
  pdfconv ofd2docx notice.ofd
  pdfconv ofd2xlsx notice.ofd
  pdfconv ofd2pptx notice.ofd
  pdfconv ofd2png notice.ofd --dpi 150
  pdfconv pdf2ofd report.pdf -o report.ofd
  pdfconv wps2docx legacy.wps
  pdfconv wps2pdf legacy.wps
  pdfconv wps2txt legacy.wps
  pdfconv doc2wps report.docx
  pdfconv pdf2wps report.pdf
  pdfconv doctor

所有转换均在本机完成，不联网。
许可证：PyMuPDF 为 AGPL-3.0；其余依赖以各自安装包元数据为准，个人本机使用前请自行确认。
"""
import argparse
import contextlib
import glob
import io
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import warnings
from pathlib import Path

VERSION = "1.0.1"
AUTHOR = "温酒会知音"
COPYRIGHT = "Copyright © 2026 温酒会知音"

# 强制 UTF-8 输出：避免在 GBK 控制台/管道下打印中文符号（如 OFD↔PDF）时报 UnicodeEncodeError。
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# 第三方库（pdf2docx 等）仍使用旧别名 fitz，会向 stderr 打印弃用告警；日志里不需要这条噪音。
warnings.filterwarnings("ignore", message=r".*fitz.*deprecated.*")

RESULTS = []


class ConvError(Exception):
    pass


def eprint(*a):
    print(*a, file=sys.stderr)


def need(module, hint=""):
    try:
        return __import__(module)
    except Exception as e:  # pragma: no cover
        raise ConvError(f"缺少依赖 {module}（{e}）。{hint} 请运行: pdfconv doctor 查看环境，或重跑 install.ps1")


def need_fitz(hint="pip install pymupdf"):
    """优先使用新包名 pymupdf，回退旧别名 fitz（避免弃用告警）。"""
    try:
        return __import__("pymupdf")
    except Exception:
        try:
            return __import__("fitz")
        except Exception as e:  # pragma: no cover
            raise ConvError(f"缺少依赖 PyMuPDF（{e}）。{hint} 请运行: pdfconv doctor 查看环境，或重跑 install.ps1")


def rel(p):
    try:
        return str(Path(p).resolve())
    except Exception:
        return str(p)


def ensure_dir(p):
    Path(p).mkdir(parents=True, exist_ok=True)
    return Path(p)


def out_path(default, out):
    if out:
        return Path(out)
    return Path(default)


def guard_out(path, overwrite=False):
    path = Path(path)
    if path.exists() and not overwrite:
        raise ConvError(f"输出已存在: {path}（加 --overwrite 覆盖）")
    if path.parent and str(path.parent) not in ("", "."):
        path.parent.mkdir(parents=True, exist_ok=True)
    return path


def parse_pages(spec, npages):
    if not spec or str(spec).strip().lower() in ("all", "*"):
        return list(range(npages))
    out = []
    for part in str(spec).split(","):
        part = part.strip()
        if not part:
            continue
        m = re.fullmatch(r"(\d*)-(\d*)", part)
        if m:
            a = int(m.group(1)) if m.group(1) else 1
            b = int(m.group(2)) if m.group(2) else npages
            if a < 1 or b > npages or a > b:
                raise ConvError(f"页码越界: {part}（文档共 {npages} 页）")
            out.extend(range(a - 1, b))
            continue
        if part.isdigit():
            i = int(part)
            if i < 1 or i > npages:
                raise ConvError(f"页码越界: {part}（文档共 {npages} 页）")
            out.append(i - 1)
            continue
        raise ConvError(f"无法解析页码: {part}")
    return sorted(set(out))


def page_slices(text, per_page=False, max_chars=0):
    if per_page:
        return [text]
    return [text]


def input_pdf(p):
    path = Path(p)
    if not path.is_file():
        raise ConvError(f"输入不存在: {p}")
    return path


def open_fitz(path):
    fitz = need_fitz()
    try:
        return fitz.open(str(path))
    except Exception as e:
        raise ConvError(f"无法打开 PDF: {path} ({e})")


def collect(args, command, func):
    """执行一个转换任务并记录结果。"""
    t0 = time.time()
    rec = {"command": command, "input": rel(args.input), "output": None,
           "ok": False, "seconds": 0, "error": None, "dry_run": bool(getattr(args, "dry_run", False))}
    try:
        if getattr(args, "dry_run", False):
            rec.update(func(dry_run=True) or {})
        else:
            meta = func() or {}
            rec.update(meta)
        rec["ok"] = True
    except Exception as e:
        rec["error"] = f"{type(e).__name__}: {e}"
        if not getattr(args, "quiet", False):
            eprint(f"[FAIL] {command} {rel(args.input)} :: {rec['error']}")
    rec["seconds"] = round(time.time() - t0, 2)
    RESULTS.append(rec)
    if rec["ok"] and not getattr(args, "quiet", False):
        out = rec.get("output")
        print(f"[OK] {command}: {rel(args.input)}" + (f" -> {out}" if out else "") + f" ({rec['seconds']}s)")
    return rec["ok"]


# ---------------------------------------------------------------- PDF -> image
def cmd_pdf2image(args, fmt):
    pdf = input_pdf(args.input)
    fitz = need_fitz()
    doc = open_fitz(pdf)

    def work(dry_run=False):
        pages = parse_pages(args.pages, doc.page_count)
        outdir = Path(args.out) if args.out else Path(f"{pdf.with_suffix('')}-{fmt}")
        outdir.mkdir(parents=True, exist_ok=True)
        files = []
        for p in pages:
            target = outdir / f"{pdf.stem}-p{p + 1:03d}.{fmt}"
            if not dry_run:
                page = doc[p]
                pix = page.get_pixmap(matrix=fitz.Matrix(args.dpi / 72.0, args.dpi / 72.0), alpha=False)
                pix.save(str(target))
            files.append(rel(target))
        return {"output": rel(outdir), "files": files, "pages": len(pages), "dpi": args.dpi}

    ok = collect(args, f"pdf2{fmt}", work)
    doc.close()
    return ok


# ---------------------------------------------------------------- PDF -> txt
def cmd_pdf2txt(args):
    pdf = input_pdf(args.input)
    fitz = need_fitz()
    doc = open_fitz(pdf)

    def work(dry_run=False):
        pages = parse_pages(args.pages, doc.page_count)
        chunks = []
        for p in pages:
            t = doc[p].get_text("text")
            if args.per_page:
                target = Path(args.out) if args.out and Path(args.out).suffix == "" else Path(
                    f"{pdf.with_suffix('')}-pages")
                target.mkdir(parents=True, exist_ok=True)
                per = target / f"{pdf.stem}-p{p + 1:03d}.txt"
                if not dry_run:
                    per.write_text(t, encoding="utf-8")
                chunks.append(rel(per))
            else:
                chunks.append(f"\n===== Page {p + 1} =====\n{t}")
        if args.per_page:
            return {"output": rel(target), "files": chunks, "pages": len(pages)}
        out = guard_out(out_path(str(pdf.with_suffix(".txt")), args.out), args.overwrite)
        if not dry_run:
            out.write_text("".join(chunks), encoding="utf-8")
        return {"output": rel(out), "pages": len(pages), "chars": sum(len(c) for c in chunks)}

    ok = collect(args, "pdf2txt", work)
    doc.close()
    return ok


# ---------------------------------------------------------------- PDF -> docx
def cmd_pdf2docx(args):
    pdf = input_pdf(args.input)
    fitz = need_fitz()
    doc = open_fitz(pdf)
    npages = doc.page_count
    doc.close()

    def work(dry_run=False):
        need("pdf2docx", "pip install pdf2docx")
        from pdf2docx import Converter
        pages = parse_pages(args.pages, npages)
        out = guard_out(out_path(str(pdf.with_suffix(".docx")), args.out), args.overwrite)
        note = None
        contiguous = pages == list(range(min(pages), max(pages) + 1))
        if not contiguous:
            note = "pdf2docx 仅支持连续页范围，已按全文档转换"
        if not dry_run:
            cv = Converter(str(pdf))
            try:
                if contiguous and pages != list(range(npages)):
                    cv.convert(str(out), start=min(pages), end=max(pages) + 1)
                else:
                    cv.convert(str(out))
            finally:
                cv.close()
        return {"output": rel(out), "pages": npages, "note": note}

    return collect(args, "pdf2docx", work)


# ---------------------------------------------------------------- PDF -> pptx
def cmd_pdf2pptx(args):
    pdf = input_pdf(args.input)
    fitz = need_fitz()
    doc = open_fitz(pdf)

    def work(dry_run=False):
        pptx = need("pptx", "pip install python-pptx")
        from pptx import Presentation
        from pptx.util import Inches
        pages = parse_pages(args.pages, doc.page_count)
        out = guard_out(out_path(str(pdf.with_suffix(".pptx")), args.out), args.overwrite)
        prs = Presentation()
        p0 = doc[pages[0]]
        w_in = p0.rect.width / 72.0
        h_in = p0.rect.height / 72.0
        prs.slide_width = Inches(w_in)
        prs.slide_height = Inches(h_in)
        blank = prs.slide_layouts[6]
        for p in pages:
            page = doc[p]
            if dry_run:
                continue
            pix = page.get_pixmap(matrix=fitz.Matrix(args.dpi / 72.0, args.dpi / 72.0), alpha=False)
            slide = prs.slides.add_slide(blank)
            slide.shapes.add_picture(io.BytesIO(pix.tobytes("png")), 0, 0,
                                     width=prs.slide_width, height=prs.slide_height)
        if not dry_run:
            prs.save(str(out))
        return {"output": rel(out), "slides": len(pages), "dpi": args.dpi}

    ok = collect(args, "pdf2pptx", work)
    doc.close()
    return ok


# ---------------------------------------------------------------- PDF -> xlsx
def cmd_pdf2xlsx(args):
    pdf = input_pdf(args.input)
    fitz = need_fitz()
    doc = open_fitz(pdf)

    def work(dry_run=False):
        openpyxl = need("openpyxl", "pip install openpyxl")
        from openpyxl import Workbook
        from openpyxl.styles import Font
        from openpyxl.utils import get_column_letter
        pages = parse_pages(args.pages, doc.page_count)
        out = guard_out(out_path(str(pdf.with_suffix(".xlsx")), args.out), args.overwrite)
        tables_by_page = {}
        plumber = None
        mode = args.mode
        if mode in ("auto", "tables"):
            try:
                import pdfplumber
                plumber = pdfplumber
            except Exception:
                if mode == "tables":
                    raise ConvError("缺少 pdfplumber（pip install pdfplumber）")
        if plumber is not None:
            with plumber.open(str(pdf)) as pl:
                for p in pages:
                    tables_by_page[p] = pl.pages[p].extract_tables() or []
        has_table = any(tables_by_page.values())
        wb = Workbook()
        wb.remove(wb.active)
        for p in pages:
            ws = wb.create_sheet(title=f"P{p + 1}")
            tables = tables_by_page.get(p, [])
            if mode == "auto" and not has_table:
                tables = []
            if tables:
                row = 1
                for t in tables:
                    for r in t:
                        for c, v in enumerate(r, start=1):
                            cell = ws.cell(row=row, column=c, value=(v or "").strip() if isinstance(v, str) else v)
                            if row == 1:
                                cell.font = Font(bold=True)
                        row += 1
                    row += 1
                for col in range(1, max((len(r) for t in tables for r in t), default=1) + 1):
                    ws.column_dimensions[get_column_letter(col)].width = 18
            else:
                lines = [ln for ln in doc[p].get_text("text").splitlines() if ln.strip()]
                for i, ln in enumerate(lines, start=1):
                    ws.cell(row=i, column=1, value=ln)
                ws.column_dimensions["A"].width = 80
            ws.freeze_panes = "A2"
        if not dry_run:
            wb.save(str(out))
        return {"output": rel(out), "pages": len(pages), "mode": mode,
                "table_pages": sum(1 for v in tables_by_page.values() if v), "fallback_text": not has_table}

    ok = collect(args, "pdf2xlsx", work)
    doc.close()
    return ok


# ---------------------------------------------------------------- OCR
def ocr_pages(args, doc, fitz, dpi):
    """返回 [(page_index, [(box, text, score), ...])]"""
    rapidocr = need("rapidocr_onnxruntime", "pip install rapidocr_onnxruntime")
    engine = rapidocr.RapidOCR()
    pages = parse_pages(args.pages, doc.page_count)
    out = []
    for p in pages:
        page = doc[p]
        pix = page.get_pixmap(matrix=fitz.Matrix(dpi / 72.0, dpi / 72.0), alpha=False)
        png = pix.tobytes("png")
        try:
            result, _ = engine(png)
        except Exception as e:
            raise ConvError(f"OCR 第 {p + 1} 页失败: {e}")
        items = []
        for row in result or []:
            box, text, score = row[0], row[1], (row[2] if len(row) > 2 else 0)
            if text and str(text).strip():
                items.append((box, str(text).strip(), float(score)))
        out.append((p, items))
    return out


def _insert_invisible_text(page, items, scale):
    """把 OCR 结果写成不可见文字层（可行时）。"""
    fitz = need_fitz()
    font = None
    for name in ("china-s", "cjk", "helv"):
        try:
            font = fitz.Font(name)
            break
        except Exception:
            font = None
    if font is None:
        return False
    try:
        tw = fitz.TextWriter(page.rect)
        for box, text, _score in items:
            xs = [pt[0] for pt in box]
            ys = [pt[1] for pt in box]
            x0, x1 = min(xs), max(xs)
            y0, y1 = min(ys), max(ys)
            h = max(1.0, (y1 - y0))
            fs = max(2.0, h * 0.82 / scale)
            tw.append((x0 / scale, (y1 - h * 0.15) / scale), text, font=font, fontsize=fs)
        tw.write_text(page, render_mode=3)
        return True
    except Exception:
        return False


def cmd_ocr(args):
    pdf = input_pdf(args.input)
    fitz = need_fitz()
    doc = open_fitz(pdf)

    def work(dry_run=False):
        out = guard_out(out_path(str(pdf.with_suffix(".ocr.txt")), args.out), args.overwrite)
        if dry_run:
            return {"output": rel(out), "pages": len(parse_pages(args.pages, doc.page_count))}
        data = ocr_pages(args, doc, fitz, args.dpi)
        lines = []
        for p, items in data:
            lines.append(f"\n===== Page {p + 1} =====")
            lines.extend(t for _b, t, _s in items)
        out.write_text("\n".join(lines), encoding="utf-8")
        side = out.with_suffix(".ocr.json")
        side.write_text(json.dumps(
            [{"page": p + 1, "items": [{"box": b, "text": t, "score": s} for b, t, s in items]}
             for p, items in data], ensure_ascii=False, indent=1), encoding="utf-8")
        return {"output": rel(out), "sidecar": rel(side),
                "pages": len(data), "chars": sum(len(t) for _p, it in data for _b, t, _s in it)}

    ok = collect(args, "ocr", work)
    doc.close()
    return ok


def cmd_pdf2searchable(args):
    pdf = input_pdf(args.input)
    fitz = need_fitz()
    doc = open_fitz(pdf)

    def work(dry_run=False):
        out = guard_out(out_path(str(pdf.with_name(pdf.stem + "-searchable.pdf")), args.out), args.overwrite)
        if dry_run:
            return {"output": rel(out), "pages": len(parse_pages(args.pages, doc.page_count))}
        scale = args.dpi / 72.0
        data = ocr_pages(args, doc, fitz, args.dpi)
        layer = 0
        for p, items in data:
            page = doc[p]
            if _insert_invisible_text(page, items, scale):
                layer += 1
        doc.save(str(out))
        return {"output": rel(out), "pages": len(data), "text_layer_pages": layer,
                "note": None if layer == len(data) else "部分页无法写入中文不可见文字层（仍已生成 OCR 文本）"}

    ok = collect(args, "pdf2searchable", work)
    doc.close()
    return ok


# ---------------------------------------------------------------- images -> PDF
def _fit_image(img, pagesize):
    from PIL import Image
    if pagesize == "auto":
        return img.convert("RGB")
    w, h = img.size
    if pagesize == "a4":
        pw, ph = 1240, 1754  # 150 dpi
    elif pagesize == "letter":
        pw, ph = 1275, 1650
    else:
        raise ConvError(f"未知 pagesize: {pagesize}")
    if w > h:
        pw, ph = ph, pw
    canvas = Image.new("RGB", (pw, ph), (255, 255, 255))
    ratio = min(pw / w, ph / h)
    im2 = img.convert("RGB").resize((max(1, int(w * ratio)), max(1, int(h * ratio))), Image.LANCZOS)
    canvas.paste(im2, ((pw - im2.width) // 2, (ph - im2.height) // 2))
    return canvas


def cmd_img2pdf(args):
    files = []
    for pattern in args.input:
        hits = sorted(glob.glob(pattern))
        if hits:
            files.extend(hits)
        elif Path(pattern).is_file():
            files.append(pattern)
    if not files:
        raise ConvError("没有找到输入图片（支持通配符）")
    files = sorted(files, key=lambda s: [int(t) if t.isdigit() else t.lower() for t in re.split(r"(\d+)", s)])

    def work(dry_run=False):
        from PIL import Image
        out = guard_out(out_path(str(Path(files[0]).with_suffix(".pdf")), args.out), args.overwrite)
        images = []
        for f in files:
            with Image.open(f) as im:
                images.append(_fit_image(im, args.pagesize))
        if not dry_run:
            images[0].save(str(out), "PDF", save_all=True, append_images=images[1:], resolution=150)
        return {"output": rel(out), "pages": len(images), "inputs": [rel(f) for f in files]}

    return collect(args, "img2pdf", work)


def cmd_pdf2imgpdf(args):
    """PDF -> 纯图/扫描型 PDF：逐页渲染成图后重建 PDF（对应 WPS 的“转为扫描型PDF”）。"""
    pdf = input_pdf(args.input)
    fitz = need_fitz()
    doc = open_fitz(pdf)

    def work(dry_run=False):
        from PIL import Image
        pages = parse_pages(args.pages, doc.page_count)
        out = guard_out(out_path(str(pdf.with_name(pdf.stem + "-imgpdf.pdf")), args.out), args.overwrite)
        images = []
        for p in pages:
            if dry_run:
                continue
            pix = doc[p].get_pixmap(matrix=fitz.Matrix(args.dpi / 72.0, args.dpi / 72.0), alpha=False)
            images.append(Image.open(io.BytesIO(pix.tobytes("png"))).convert("RGB"))
        if not dry_run:
            images[0].save(str(out), "PDF", save_all=True, append_images=images[1:], resolution=args.dpi)
        return {"output": rel(out), "pages": len(pages), "dpi": args.dpi}

    ok = collect(args, "pdf2imgpdf", work)
    doc.close()
    return ok


# ---------------------------------------------------------------- merge / split / extract
def cmd_merge(args):
    pypdf = need("pypdf", "pip install pypdf")
    from pypdf import PdfWriter, PdfReader
    inputs = []
    for pattern in args.input:
        hits = sorted(glob.glob(pattern))
        inputs.extend(hits if hits else ([pattern] if Path(pattern).is_file() else []))
    if not inputs:
        raise ConvError("没有找到输入 PDF")

    def work(dry_run=False):
        out = guard_out(out_path("merged.pdf", args.out), args.overwrite)
        w = PdfWriter()
        npages = 0
        for f in inputs:
            r = PdfReader(str(input_pdf(f)))
            npages += len(r.pages)
            if not dry_run:
                for pg in r.pages:
                    w.add_page(pg)
        if not dry_run:
            with open(out, "wb") as fh:
                w.write(fh)
        return {"output": rel(out), "inputs": len(inputs), "pages": npages}

    return collect(args, "merge", work)


def cmd_split(args):
    pypdf = need("pypdf", "pip install pypdf")
    from pypdf import PdfReader, PdfWriter
    pdf = input_pdf(args.input)
    reader = PdfReader(str(pdf))
    npages = len(reader.pages)

    def work(dry_run=False):
        outdir = Path(args.out) if args.out else Path(f"{pdf.with_suffix('')}-split")
        outdir.mkdir(parents=True, exist_ok=True)
        jobs = []
        if args.ranges:
            for part in args.ranges.split(","):
                pages = parse_pages(part.strip(), npages)
                if pages:
                    jobs.append(pages)
        else:
            step = max(1, args.every)
            jobs = [list(range(i, min(i + step, npages))) for i in range(0, npages, step)]
        files = []
        for pages in jobs:
            a, b = pages[0] + 1, pages[-1] + 1
            target = outdir / (f"{pdf.stem}-p{a:03d}-{b:03d}.pdf" if len(pages) > 1 else f"{pdf.stem}-p{a:03d}.pdf")
            if target.exists() and not args.overwrite:
                raise ConvError(f"输出已存在: {target}（加 --overwrite 覆盖）")
            if not dry_run:
                w = PdfWriter()
                for i in pages:
                    w.add_page(reader.pages[i])
                with open(target, "wb") as fh:
                    w.write(fh)
            files.append(rel(target))
        return {"output": rel(outdir), "files": files, "chunks": len(jobs), "pages": npages}

    return collect(args, "split", work)


def cmd_extract_pages(args):
    pypdf = need("pypdf", "pip install pypdf")
    from pypdf import PdfReader, PdfWriter
    pdf = input_pdf(args.input)
    reader = PdfReader(str(pdf))

    def work(dry_run=False):
        pages = parse_pages(args.pages, len(reader.pages))
        out = guard_out(out_path(str(pdf.with_name(pdf.stem + "-pages.pdf")), args.out), args.overwrite)
        if not dry_run:
            w = PdfWriter()
            for i in pages:
                w.add_page(reader.pages[i])
            with open(out, "wb") as fh:
                w.write(fh)
        return {"output": rel(out), "pages": len(pages), "source_pages": len(reader.pages)}

    return collect(args, "extract-pages", work)


def cmd_extract_images(args):
    pdf = input_pdf(args.input)
    fitz = need_fitz()
    doc = open_fitz(pdf)

    def work(dry_run=False):
        pages = parse_pages(args.pages, doc.page_count)
        outdir = Path(args.out) if args.out else Path(f"{pdf.with_suffix('')}-images")
        outdir.mkdir(parents=True, exist_ok=True)
        seen = set()
        files = []
        for p in pages:
            for img in doc[p].get_images(full=True):
                xref = img[0]
                if xref in seen:
                    continue
                seen.add(xref)
                info = doc.extract_image(xref)
                data = info.get("image")
                if not data or len(data) < args.min_size:
                    continue
                ext = info.get("ext", "png")
                target = outdir / f"{pdf.stem}-img-{xref:04d}.{ext}"
                if not dry_run:
                    target.write_bytes(data)
                files.append(rel(target))
        return {"output": rel(outdir), "files": files, "images": len(files)}

    ok = collect(args, "extract-images", work)
    doc.close()
    return ok


# ---------------------------------------------------------------- OFD
def input_ofd(p):
    path = Path(p)
    if not path.is_file():
        raise ConvError(f"输入不存在: {p}")
    if path.suffix.lower() != ".ofd":
        raise ConvError(f"不是 .ofd 文件: {p}")
    return path


def _silence_easyofd():
    """easyofd 使用 loguru，默认 DEBUG 日志会污染输出；这里关闭其日志。"""
    try:
        from loguru import logger
        logger.remove()
    except Exception:
        pass


@contextlib.contextmanager
def _easyofd_workspace():
    """easyofd 会在当前目录创建临时文件/目录（如 ./test），这里给它一个独立临时工作目录。"""
    old = os.getcwd()
    td = tempfile.mkdtemp(prefix="pdfconv-easyofd-")
    try:
        os.chdir(td)
        yield td
    finally:
        os.chdir(old)
        shutil.rmtree(td, ignore_errors=True)


def _ofd_to_pdf(src, dst):
    """OFD → PDF（easyofd）。返回生成的 PDF 页数。"""
    captured = io.StringIO()
    with _easyofd_workspace():
        with contextlib.redirect_stdout(captured), contextlib.redirect_stderr(captured):
            need("easyofd", "pip install easyofd")
            from easyofd import OFD
            _silence_easyofd()
            ofd = OFD()
            try:
                ofd.read(str(src), fmt="path")
            except Exception as e:
                raise ConvError(f"OFD 解析失败: {e}")
            data = ofd.to_pdf()
    if not data:
        raise ConvError("OFD 转 PDF 失败：未生成数据")
    Path(dst).write_bytes(data)
    fitz = need_fitz()
    doc = fitz.open(str(dst))
    try:
        pages = doc.page_count
    finally:
        doc.close()
    if pages < 1:
        raise ConvError("OFD 转 PDF 后为 0 页")
    return pages


def cmd_ofd2pdf(args):
    src = input_ofd(args.input)

    def work(dry_run=False):
        out = guard_out(out_path(str(src.with_suffix(".pdf")), args.out), args.overwrite)
        if dry_run:
            return {"output": rel(out), "note": "OFD → PDF"}
        pages = _ofd_to_pdf(src, out)
        return {"output": rel(out), "pages": pages}

    return collect(args, "ofd2pdf", work)


_OFD_CHAIN = {
    "docx": ("pdf2docx", ".docx", "file"),
    "xlsx": ("pdf2xlsx", ".xlsx", "file"),
    "pptx": ("pdf2pptx", ".pptx", "file"),
    "txt": ("pdf2txt", ".txt", "file"),
    "png": ("pdf2png", "-png", "dir"),
    "jpg": ("pdf2jpg", "-jpg", "dir"),
}


def _ofd_chain(args, command, kind):
    """OFD → 目标格式：先用 easyofd 转 PDF，再调用 pdfconv 自身的 PDF 转换命令。"""
    src = input_ofd(args.input)
    sub_cmd, suffix, out_kind = _OFD_CHAIN[kind]

    def work(dry_run=False):
        if out_kind == "file":
            default = str(src.with_suffix(suffix))
        else:
            default = str(src.with_name(src.stem + suffix))
        out = guard_out(out_path(default, args.out), args.overwrite)
        if dry_run:
            return {"output": rel(out), "note": f"OFD → {kind.upper()}（经 PDF）"}
        with tempfile.TemporaryDirectory(prefix="pdfconv-ofd-") as td:
            tmp_pdf = Path(td) / "source.pdf"
            ofd_pages = _ofd_to_pdf(src, tmp_pdf)
            argv = [sys.executable, str(Path(__file__).resolve()), sub_cmd, str(tmp_pdf),
                    "-o", str(out), "--quiet", "--overwrite"]
            if getattr(args, "pages", None):
                argv += ["--pages", str(args.pages)]
            if kind in ("png", "jpg", "pptx") and getattr(args, "dpi", None):
                argv += ["--dpi", str(args.dpi)]
            if kind == "xlsx" and getattr(args, "mode", None):
                argv += ["--mode", str(args.mode)]
            proc = subprocess.run(argv, capture_output=True, text=True,
                                  encoding="utf-8", errors="replace")
            if proc.returncode != 0:
                msg = ((proc.stdout or "") + "\n" + (proc.stderr or "")).strip()
                raise ConvError(f"OFD→{kind.upper()} 失败: {msg[-1200:]}")
            if out_kind == "dir":
                # 子进程输出的文件名基于临时 PDF（source-pXXX），改回 OFD 源文件名
                for f in sorted(Path(out).glob("source-p*")):
                    target = f.with_name(f.name.replace("source-p", f"{src.stem}-p", 1))
                    if target.exists():
                        target.unlink()
                    f.rename(target)
        return {"output": rel(out), "ofd_pages": ofd_pages, "via": f"OFD→PDF→{kind.upper()}"}

    return collect(args, command, work)


def cmd_ofd2docx(args):
    return _ofd_chain(args, "ofd2docx", "docx")


def cmd_ofd2xlsx(args):
    return _ofd_chain(args, "ofd2xlsx", "xlsx")


def cmd_ofd2pptx(args):
    return _ofd_chain(args, "ofd2pptx", "pptx")


def cmd_ofd2txt(args):
    return _ofd_chain(args, "ofd2txt", "txt")


def cmd_ofd2image(args, fmt):
    return _ofd_chain(args, f"ofd2{fmt}", fmt)


def cmd_pdf2ofd(args):
    pdf = input_pdf(args.input)

    def work(dry_run=False):
        out = guard_out(out_path(str(pdf.with_suffix(".ofd")), args.out), args.overwrite)
        if dry_run:
            return {"output": rel(out), "note": "PDF → OFD"}
        need("easyofd", "pip install easyofd")
        from easyofd import OFD
        _silence_easyofd()
        captured = io.StringIO()
        with _easyofd_workspace():
            with contextlib.redirect_stdout(captured), contextlib.redirect_stderr(captured):
                ofd = OFD()
                data = ofd.pdf2ofd(pdf.read_bytes())
        if not data:
            raise ConvError("PDF 转 OFD 失败：未生成数据")
        Path(out).write_bytes(data)
        return {"output": rel(out), "bytes": len(data)}

    return collect(args, "pdf2ofd", work)


# ---------------------------------------------------------------- WPS（金山 WPS 文字格式 .wps）
WPS_PROGID = "KWPS.Application"
WPS_DOC_EXTS = {".doc", ".docx", ".rtf", ".txt", ".wps"}


def input_wps(p):
    path = Path(p)
    if not path.is_file():
        raise ConvError(f"输入不存在: {p}")
    if path.suffix.lower() != ".wps":
        raise ConvError(f"不是 .wps 文件: {p}")
    return path


def input_document(p):
    path = Path(p)
    if not path.is_file():
        raise ConvError(f"输入不存在: {p}")
    if path.suffix.lower() not in WPS_DOC_EXTS:
        raise ConvError(f"不支持的文档格式: {p}（支持 doc/docx/rtf/txt/wps）")
    return path


def _wps_app():
    """启动本机 WPS 文字的 COM 自动化（KWPS.Application）。"""
    need("win32com.client", "pip install pywin32")
    import win32com.client
    try:
        app = win32com.client.DispatchEx(WPS_PROGID)
    except Exception:
        raise ConvError("未检测到 WPS Office（COM: KWPS.Application）。请安装 WPS Office 后重试；或改用其它输出格式。")
    for attr, val in (("Visible", False), ("DisplayAlerts", 0), ("ScreenUpdating", False)):
        try:
            setattr(app, attr, val)
        except Exception:
            pass
    return app


class _WpsSession:
    def __enter__(self):
        self.app = _wps_app()
        return self.app

    def __exit__(self, *exc):
        try:
            self.app.Quit()
        except Exception:
            pass
        return False


def _wps_open_save(app, src, dst, fmt):
    doc = app.Documents.Open(str(src), ReadOnly=True)
    try:
        doc.SaveAs2(str(dst), fmt)
    finally:
        try:
            doc.Close(0)
        except Exception:
            pass


def _wps_export_pdf(app, src, dst):
    doc = app.Documents.Open(str(src), ReadOnly=True)
    try:
        doc.ExportAsFixedFormat(str(dst), 17)  # wdExportFormatPDF
    finally:
        try:
            doc.Close(0)
        except Exception:
            pass


def _wps_extract_text(app, src):
    doc = app.Documents.Open(str(src), ReadOnly=True)
    try:
        return doc.Content.Text or ""
    finally:
        try:
            doc.Close(0)
        except Exception:
            pass


def cmd_wps2docx(args):
    src = input_wps(args.input)

    def work(dry_run=False):
        out = guard_out(out_path(str(src.with_suffix(".docx")), args.out), args.overwrite)
        if dry_run:
            return {"output": rel(out), "note": "WPS → DOCX"}
        with _WpsSession() as app:
            _wps_open_save(app, src, out, 12)  # wdFormatXMLDocument
        if not out.is_file():
            raise ConvError("WPS → DOCX 未生成文件")
        return {"output": rel(out), "bytes": out.stat().st_size}

    return collect(args, "wps2docx", work)


def cmd_wps2pdf(args):
    src = input_wps(args.input)

    def work(dry_run=False):
        out = guard_out(out_path(str(src.with_suffix(".pdf")), args.out), args.overwrite)
        if dry_run:
            return {"output": rel(out), "note": "WPS → PDF"}
        with _WpsSession() as app:
            _wps_export_pdf(app, src, out)
        if not out.is_file():
            raise ConvError("WPS → PDF 未生成文件")
        return {"output": rel(out), "bytes": out.stat().st_size}

    return collect(args, "wps2pdf", work)


def cmd_wps2txt(args):
    src = input_wps(args.input)

    def work(dry_run=False):
        out = guard_out(out_path(str(src.with_suffix(".txt")), args.out), args.overwrite)
        if dry_run:
            return {"output": rel(out), "note": "WPS → TXT"}
        with _WpsSession() as app:
            text = _wps_extract_text(app, src)
        out.write_text(text, encoding="utf-8")
        return {"output": rel(out), "chars": len(text)}

    return collect(args, "wps2txt", work)


def _run_pdfconv_sub(argv):
    proc = subprocess.run([sys.executable, str(Path(__file__).resolve())] + [str(x) for x in argv],
                          capture_output=True, text=True, encoding="utf-8", errors="replace")
    if proc.returncode != 0:
        msg = ((proc.stdout or "") + "\n" + (proc.stderr or "")).strip()
        raise ConvError(f"内部转换失败: {msg[-1200:]}")


def cmd_wps2ofd(args):
    src = input_wps(args.input)

    def work(dry_run=False):
        out = guard_out(out_path(str(src.with_suffix(".ofd")), args.out), args.overwrite)
        if dry_run:
            return {"output": rel(out), "note": "WPS → PDF → OFD"}
        with tempfile.TemporaryDirectory(prefix="pdfconv-wps-") as td:
            tmp_pdf = Path(td) / "wps.pdf"
            with _WpsSession() as app:
                _wps_export_pdf(app, src, tmp_pdf)
            _run_pdfconv_sub(["pdf2ofd", tmp_pdf, "-o", out, "--overwrite", "--quiet"])
        return {"output": rel(out), "via": "WPS→PDF→OFD"}

    return collect(args, "wps2ofd", work)


def cmd_doc2wps(args):
    src = input_document(args.input)
    if src.suffix.lower() == ".wps":
        raise ConvError("输入已经是 .wps；如需转换请使用 wps2docx / wps2pdf / wps2txt 等命令")

    def work(dry_run=False):
        out = guard_out(out_path(str(src.with_suffix(".wps")), args.out), args.overwrite)
        if dry_run:
            return {"output": rel(out), "note": "文档 → WPS"}
        with _WpsSession() as app:
            _wps_open_save(app, src, out, 0)  # wdFormatDocument -> 真 WPS 格式（OLE）
        if not out.is_file():
            raise ConvError("文档 → WPS 未生成文件")
        return {"output": rel(out), "bytes": out.stat().st_size}

    return collect(args, "doc2wps", work)


def cmd_pdf2wps(args):
    pdf = input_pdf(args.input)

    def work(dry_run=False):
        out = guard_out(out_path(str(pdf.with_suffix(".wps")), args.out), args.overwrite)
        if dry_run:
            return {"output": rel(out), "note": "PDF → DOCX → WPS"}
        with tempfile.TemporaryDirectory(prefix="pdfconv-pdf2wps-") as td:
            tmp_docx = Path(td) / "source.docx"
            _run_pdfconv_sub(["pdf2docx", pdf, "-o", tmp_docx, "--overwrite", "--quiet"])
            with _WpsSession() as app:
                _wps_open_save(app, tmp_docx, out, 0)
        return {"output": rel(out), "via": "PDF→DOCX→WPS"}

    return collect(args, "pdf2wps", work)


def cmd_ofd2wps(args):
    src = input_ofd(args.input)

    def work(dry_run=False):
        out = guard_out(out_path(str(src.with_suffix(".wps")), args.out), args.overwrite)
        if dry_run:
            return {"output": rel(out), "note": "OFD → PDF → DOCX → WPS"}
        with tempfile.TemporaryDirectory(prefix="pdfconv-ofd2wps-") as td:
            tmp_pdf = Path(td) / "ofd.pdf"
            _ofd_to_pdf(src, tmp_pdf)
            tmp_docx = Path(td) / "ofd.docx"
            _run_pdfconv_sub(["pdf2docx", tmp_pdf, "-o", tmp_docx, "--overwrite", "--quiet"])
            with _WpsSession() as app:
                _wps_open_save(app, tmp_docx, out, 0)
        return {"output": rel(out), "via": "OFD→PDF→DOCX→WPS"}

    return collect(args, "ofd2wps", work)


# ---------------------------------------------------------------- info / doctor
def cmd_info(args):
    pdf = input_pdf(args.input)
    fitz = need_fitz()
    doc = open_fitz(pdf)

    def work(dry_run=False):
        pages = []
        for i in range(doc.page_count):
            pg = doc[i]
            pages.append({"page": i + 1, "width_pt": round(pg.rect.width, 1),
                          "height_pt": round(pg.rect.height, 1),
                          "text_chars": len(pg.get_text("text")), "images": len(pg.get_images())})
        return {"output": None, "pages": doc.page_count, "metadata": doc.metadata,
                "page_info": pages, "file_size": Path(pdf).stat().st_size}

    ok = collect(args, "info", work)
    doc.close()
    return ok


def cmd_doctor(args):
    deps = [("pymupdf", "PyMuPDF", "pdf 渲染/文本/图片"), ("pypdf", "pypdf", "合并/拆分/提取页"),
            ("pdfplumber", "pdfplumber", "表格提取"), ("pdf2docx", "pdf2docx", "PDF→DOCX"),
            ("docx", "python-docx", "DOCX 写入"), ("pptx", "python-pptx", "PPTX 写入"),
            ("openpyxl", "openpyxl", "XLSX 写入"), ("PIL", "Pillow", "图片→PDF"),
            ("rapidocr_onnxruntime", "RapidOCR", "本地 OCR"),
            ("easyofd", "easyofd", "OFD 读写（OFD↔PDF）"),
            ("win32com", "pywin32", "WPS .wps 格式桥接")]
    rows = []
    for mod, name, use in deps:
        try:
            m = __import__(mod)
            ver = getattr(m, "__version__", "")
            rows.append({"module": mod, "package": name, "version": str(ver), "ok": True, "use": use})
        except Exception as e:
            rows.append({"module": mod, "package": name, "version": "", "ok": False, "use": use,
                         "error": str(e)})
    # WPS Office COM 可用性（.wps 转换依赖本机 WPS）
    wps_ok = False
    try:
        import winreg
        winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "KWPS.Application"))
        wps_ok = True
    except Exception:
        pass
    rows.append({"module": "KWPS.Application", "package": "WPS Office COM", "version": "",
                 "ok": wps_ok, "use": ".wps 格式转换（需本机安装 WPS Office）"})
    if args.json:
        print(json.dumps({"tool": "pdfconv", "version": VERSION, "python": sys.version,
                          "executable": sys.executable, "deps": rows}, ensure_ascii=False, indent=1))
        return True
    print(f"pdfconv {VERSION} — CE转换引擎（Conversion Engine）")
    print(f"作者: {AUTHOR}    {COPYRIGHT}")
    print(f"python: {sys.version.split()[0]}  ({sys.executable})")
    print("依赖:")
    for r in rows:
        print(f"  [{'OK ' if r['ok'] else 'MISS'}] {r['package']:<12} {r['version']:<12} {r['use']}")
    return True


# ---------------------------------------------------------------- CLI
def add_common(sp, pdf=True):
    sp.add_argument("input", nargs="+" if pdf is False else None, help="输入文件（PDF；img2pdf/merge 可多个/通配符）" if pdf else "输入文件（可多个/通配符）")
    sp.add_argument("-o", "--out", default=None, help="输出路径（文件或目录）")
    sp.add_argument("--overwrite", action="store_true", help="允许覆盖已存在输出")
    sp.add_argument("--json", action="store_true", help="以 JSON 输出结果")
    sp.add_argument("--quiet", action="store_true", help="不打印进度")
    sp.add_argument("--dry-run", action="store_true", help="只显示将要做什么，不写文件")


def build_parser():
    p = argparse.ArgumentParser(
        prog="pdfconv",
        description="本地离线 PDF 转换工具（开源引擎实现，不依赖 WPS，不联网）",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__.split("用法示例：", 1)[-1])
    p.add_argument("--version", action="version", version=f"pdfconv {VERSION} — CE转换引擎（Conversion Engine）— 作者：{AUTHOR}")

    sub = p.add_subparsers(dest="cmd", required=True)

    sp = sub.add_parser("info", help="查看 PDF 页数/尺寸/文字量/图片数")
    add_common(sp)
    sp.set_defaults(func=cmd_info)

    for fmt in ("png", "jpg"):
        sp = sub.add_parser(f"pdf2{fmt}", help=f"PDF → {fmt.upper()} 图片（逐页）")
        add_common(sp)
        sp.add_argument("--dpi", type=int, default=150)
        sp.add_argument("--pages", default="all")
        sp.set_defaults(func=lambda a, f=fmt: cmd_pdf2image(a, f))

    sp = sub.add_parser("pdf2txt", help="PDF → TXT")
    add_common(sp)
    sp.add_argument("--pages", default="all")
    sp.add_argument("--per-page", action="store_true", help="每页一个 txt（输出为目录）")
    sp.set_defaults(func=cmd_pdf2txt)

    sp = sub.add_parser("pdf2docx", help="PDF → Word（pdf2docx）")
    add_common(sp)
    sp.add_argument("--pages", default="all")
    sp.set_defaults(func=cmd_pdf2docx)

    sp = sub.add_parser("pdf2pptx", help="PDF → PPT（逐页渲染整页入幻灯片）")
    add_common(sp)
    sp.add_argument("--pages", default="all")
    sp.add_argument("--dpi", type=int, default=150)
    sp.set_defaults(func=cmd_pdf2pptx)

    sp = sub.add_parser("pdf2xlsx", help="PDF → Excel（表格提取；无表格时回退文字）")
    add_common(sp)
    sp.add_argument("--pages", default="all")
    sp.add_argument("--mode", choices=["auto", "tables", "text"], default="auto")
    sp.set_defaults(func=cmd_pdf2xlsx)

    sp = sub.add_parser("ocr", help="本地 OCR → TXT/JSON")
    add_common(sp)
    sp.add_argument("--pages", default="all")
    sp.add_argument("--dpi", type=int, default=200)
    sp.set_defaults(func=cmd_ocr)

    sp = sub.add_parser("pdf2searchable", help="生成可搜索 PDF（本地 OCR + 不可见文字层）")
    add_common(sp)
    sp.add_argument("--pages", default="all")
    sp.add_argument("--dpi", type=int, default=200)
    sp.set_defaults(func=cmd_pdf2searchable)

    sp = sub.add_parser("img2pdf", help="图片 → PDF")
    add_common(sp, pdf=False)
    sp.add_argument("--pagesize", choices=["auto", "a4", "letter"], default="auto")
    sp.set_defaults(func=cmd_img2pdf)

    sp = sub.add_parser("pdf2imgpdf", help="PDF → 纯图/扫描型 PDF（逐页渲染后重建）")
    add_common(sp)
    sp.add_argument("--pages", default="all")
    sp.add_argument("--dpi", type=int, default=150)
    sp.set_defaults(func=cmd_pdf2imgpdf)

    sp = sub.add_parser("merge", help="合并多个 PDF")
    add_common(sp, pdf=False)
    sp.set_defaults(func=cmd_merge)

    sp = sub.add_parser("split", help="拆分 PDF")
    add_common(sp)
    sp.add_argument("--every", type=int, default=1, help="每 N 页一个文件（默认 1）")
    sp.add_argument("--ranges", default=None, help='按范围拆，如 "1-3,4-6"')
    sp.set_defaults(func=cmd_split)

    sp = sub.add_parser("extract-pages", help="提取指定页为单个 PDF")
    add_common(sp)
    sp.add_argument("--pages", required=True)
    sp.set_defaults(func=cmd_extract_pages)

    sp = sub.add_parser("extract-images", help="提取 PDF 内嵌图片")
    add_common(sp)
    sp.add_argument("--pages", default="all")
    sp.add_argument("--min-size", type=int, default=512, help="小于该字节数的图片忽略")
    sp.set_defaults(func=cmd_extract_images)

    # ---------- OFD ----------
    sp = sub.add_parser("ofd2pdf", help="OFD → PDF")
    add_common(sp)
    sp.set_defaults(func=cmd_ofd2pdf)

    for name, func, help_text in (
        ("ofd2docx", cmd_ofd2docx, "OFD → Word（经 PDF）"),
        ("ofd2xlsx", cmd_ofd2xlsx, "OFD → Excel（经 PDF）"),
        ("ofd2pptx", cmd_ofd2pptx, "OFD → PPT（经 PDF）"),
        ("ofd2txt", cmd_ofd2txt, "OFD → TXT（经 PDF）"),
    ):
        sp = sub.add_parser(name, help=help_text)
        add_common(sp)
        sp.add_argument("--pages", default="all")
        if name == "ofd2xlsx":
            sp.add_argument("--mode", choices=["auto", "tables", "text"], default="auto")
        if name == "ofd2pptx":
            sp.add_argument("--dpi", type=int, default=150)
        sp.set_defaults(func=func)

    for fmt in ("png", "jpg"):
        sp = sub.add_parser(f"ofd2{fmt}", help=f"OFD → {fmt.upper()} 图片（经 PDF）")
        add_common(sp)
        sp.add_argument("--pages", default="all")
        sp.add_argument("--dpi", type=int, default=150)
        sp.set_defaults(func=lambda a, f=fmt: cmd_ofd2image(a, f))

    sp = sub.add_parser("pdf2ofd", help="PDF → OFD")
    add_common(sp)
    sp.set_defaults(func=cmd_pdf2ofd)

    for name, func, help_text in (
        ("wps2docx", cmd_wps2docx, "WPS (.wps) → Word (.docx)"),
        ("wps2pdf", cmd_wps2pdf, "WPS (.wps) → PDF"),
        ("wps2txt", cmd_wps2txt, "WPS (.wps) → 文本 TXT"),
        ("wps2ofd", cmd_wps2ofd, "WPS (.wps) → OFD"),
        ("doc2wps", cmd_doc2wps, "Word/RTF/TXT → WPS (.wps)"),
        ("pdf2wps", cmd_pdf2wps, "PDF → WPS (.wps)"),
        ("ofd2wps", cmd_ofd2wps, "OFD → WPS (.wps)"),
    ):
        sp = sub.add_parser(name, help=help_text)
        add_common(sp)
        sp.set_defaults(func=func)

    sp = sub.add_parser("doctor", help="检查依赖与版本")
    sp.add_argument("--json", action="store_true")
    sp.set_defaults(func=lambda a: cmd_doctor(a))
    return p


def main(argv=None):
    parser = build_parser()
    args = parser.parse_args(argv)
    if args.cmd == "doctor":
        ok = cmd_doctor(args)
        return 0 if ok else 1
    try:
        ok = args.func(args)
    except ConvError as e:
        eprint(f"[ERROR] {e}")
        ok = False
    if getattr(args, "json", False):
        print(json.dumps({"tool": "pdfconv", "version": VERSION, "results": RESULTS},
                         ensure_ascii=False, indent=1))
    elif not getattr(args, "quiet", False):
        failed = [r for r in RESULTS if not r["ok"]]
        if failed:
            eprint(f"完成: {len(RESULTS) - len(failed)} 成功 / {len(failed)} 失败")
    return 0 if ok and all(r["ok"] for r in RESULTS) else 1


if __name__ == "__main__":
    sys.exit(main())
