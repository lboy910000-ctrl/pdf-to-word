"""Offline PDF -> Word (.docx) extractor: Bengali / Hindi / English text, equations, graphs, tables.

Built for big files (100 MB+):
- one page at a time (the PDF is never loaded whole), memory is released as it goes;
- output is split into several Word files (default 200 pages or ~60 MB each) so Word can open them;
- OCR of scanned pages runs in parallel on all CPU cores;
- a page that fails never stops the run: it is replaced by an image of that page and listed in the notes;
- a Cancel hook saves what is done so far.

Per page:
- Digital PDFs: exact text from the PDF's text layer (bold / italic / size / colour / superscripts kept).
- Scanned or garbled pages: Tesseract OCR (ben+hin+eng) + the original page image under the text.
- Tables -> real Word tables.  Equations -> exact picture + copyable text (+ LaTeX if pix2tex is installed).
- Graphs -> pictures at their page position, with text inside the chart kept as a caption.
"""
import io
import os
import re
import statistics
import sys
from concurrent.futures import ProcessPoolExecutor
from pathlib import Path

import fitz  # PyMuPDF
from docx import Document
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt, RGBColor

CS_FONT = "Nirmala UI"  # Windows font with Bengali + Devanagari; Word falls back elsewhere
LATIN_FONT = "Calibri"
MAX_PART_IMAGE_BYTES = 60 * 1024 * 1024

STRONG_MATH = set("∑∫√∞≤≥≠≈±×÷∂∇∈∉⊂⊆∪∩→⇒⇔∀∃∝∴∵°′″ⁿ²³⁴⁵⁶⁷⁸⁹½¼¾")
GREEK = set("αβγδεζηθικλμνξπρστυφχψωΔΘΛΞΠΣΦΨΩ")
MATH_RE = re.compile(r"(\d\s*[=+\-*/^]\s*[\dA-Za-z(])|([A-Za-z]\s*=\s*[\dA-Za-z(−-])|(\^|_\{|\\frac|\bsin\b|\bcos\b|\btan\b|\blog\b|\blim\b)")

_ocr_ready = None
_latex_model = None

# Bengali / Hindi / English language data is bundled in ./tessdata, so only the Tesseract program itself
# has to be installed. Set at import time so parallel OCR worker processes inherit it too.
# BASE is the app folder, or PyInstaller's unpack folder when running as the packaged Windows app.
BASE = Path(getattr(sys, "_MEIPASS", Path(__file__).resolve().parent))
_BUNDLED_TESSDATA = BASE / "tessdata"
if all((_BUNDLED_TESSDATA / f"{l}.traineddata").exists() for l in ("ben", "hin", "eng")):
    os.environ["TESSDATA_PREFIX"] = str(_BUNDLED_TESSDATA)
# The packaged app also carries the Tesseract program itself, so nothing needs installing.
_BUNDLED_EXE = BASE / "tesseract" / ("tesseract.exe" if os.name == "nt" else "tesseract")
if _BUNDLED_EXE.exists():
    try:
        import pytesseract as _pt
        _pt.pytesseract.tesseract_cmd = str(_BUNDLED_EXE)
    except Exception:
        pass


# ----------------------------------------------------------------- OCR / LaTeX
def _ocr_available():
    global _ocr_ready
    if _ocr_ready is None:
        try:
            import pytesseract
            langs = set(pytesseract.get_languages())
            _ocr_ready = {l for l in ("ben", "hin", "eng") if l in langs} or False
        except Exception:
            _ocr_ready = False
    return _ocr_ready


def _vov(a, b):
    """Vertical overlap of two (y0, y1) ranges as a share of the shorter one."""
    inter = min(a[1], b[1]) - max(a[0], b[0])
    return inter / max(1, min(a[1] - a[0], b[1] - b[0]))


def _script_counts(text):
    """(bengali_letters, devanagari_letters). The shared danda marks U+0964/0965 and digits are ignored."""
    ben = sum(1 for c in text if 0x0980 <= ord(c) <= 0x09FF and not 0x09E6 <= ord(c) <= 0x09EF)
    dev = sum(1 for c in text if 0x0900 <= ord(c) <= 0x097F and ord(c) not in (0x964, 0x965)
              and not 0x966 <= ord(c) <= 0x96F)
    return ben, dev


def _repair_mixed_line(pil_img, ln, langs):
    """The combined ben+hin+eng model sometimes reads part of a Hindi line as Bengali (or the reverse).
    Such a line mixes two scripts, which real text almost never does, so re-read just that line with the
    single script it mostly is. Keeps the original if the re-read is not cleaner."""
    text = " ".join(t for _, t in sorted(ln["w"]))
    ben, dev = _script_counts(text)
    if not (ben and dev) or pil_img is None:
        return
    want = "hin" if dev >= ben else "ben"
    if want not in langs:
        return
    import pytesseract
    lang = want + ("+eng" if "eng" in langs else "")
    pad = 6
    box = (max(ln["x0"] - pad, 0), max(ln["y0"] - pad, 0), min(ln["x1"] + pad, pil_img.width), min(ln["y1"] + pad, pil_img.height))
    try:
        new = pytesseract.image_to_string(pil_img.crop(box), lang=lang, config="--oem 1 --psm 7").strip()
    except Exception:
        return
    b2, d2 = _script_counts(new)
    if new and not (b2 and d2):
        ln["w"] = [(ln["x0"], new)]


def _text_from_data(d, pil_img=None, langs=()):
    """Rebuild reading order from Tesseract word boxes. Tesseract's own layout step sometimes cuts one
    text line into a left and a right 'column' (tables, forms, regular lists) and reads all the left
    parts first. Blocks that share baselines and sit close together are merged back into whole lines;
    genuinely separate columns (wide gutter) stay separate."""
    lines = {}
    for k, t in enumerate(d["text"]):
        if d["level"][k] != 5 or not str(t).strip():
            continue
        key = (d["block_num"][k], d["par_num"][k], d["line_num"][k])
        x, y, w, h = d["left"][k], d["top"][k], d["width"][k], d["height"][k]
        ln = lines.setdefault(key, {"b": key[0], "w": [], "x0": x, "y0": y, "x1": x + w, "y1": y + h})
        ln["w"].append((x, str(t)))
        ln["x0"], ln["y0"] = min(ln["x0"], x), min(ln["y0"], y)
        ln["x1"], ln["y1"] = max(ln["x1"], x + w), max(ln["y1"], y + h)
    if not lines:
        return ""
    for ln in lines.values():
        _repair_mixed_line(pil_img, ln, langs)
    allv = list(lines.values())
    med_h = statistics.median(l["y1"] - l["y0"] for l in allv)
    gap = 0.9 * med_h
    by_block = {}
    for l in allv:
        by_block.setdefault(l["b"], []).append(l)
    box = {b: (min(l["x0"] for l in ls), min(l["y0"] for l in ls), max(l["x1"] for l in ls), max(l["y1"] for l in ls))
           for b, ls in by_block.items()}
    parent = {b: b for b in by_block}

    def find(b):
        while parent[b] != b:
            parent[b] = parent[parent[b]]
            b = parent[b]
        return b

    ids = sorted(by_block)
    for ia, A in enumerate(ids):
        for B in ids[ia + 1:]:
            a, b = box[A], box[B]
            if _vov((a[1], a[3]), (b[1], b[3])) <= 0:
                continue
            left, right = (A, B) if a[0] <= b[0] else (B, A)
            if box[right][0] - box[left][2] > gap:
                continue
            pairs = sum(1 for lr in by_block[right]
                        if any(_vov((ll["y0"], ll["y1"]), (lr["y0"], lr["y1"])) >= 0.5
                               and 0 <= lr["x0"] - ll["x1"] <= gap for ll in by_block[left]))
            if pairs >= 0.5 * min(len(by_block[left]), len(by_block[right])):
                parent[find(B)] = find(A)
    groups = {}
    for b in ids:
        groups.setdefault(find(b), []).extend(by_block[b])
    out = []
    for root in sorted(groups, key=lambda r: min(l["b"] for l in groups[r])):
        rows = []
        for l in sorted(groups[root], key=lambda l: (l["y0"] + l["y1"]) / 2):
            if rows and _vov((rows[-1]["y0"], rows[-1]["y1"]), (l["y0"], l["y1"])) >= 0.5:
                rows[-1]["parts"].append(l)
                rows[-1]["y0"], rows[-1]["y1"] = min(rows[-1]["y0"], l["y0"]), max(rows[-1]["y1"], l["y1"])
            else:
                rows.append({"y0": l["y0"], "y1": l["y1"], "parts": [l]})
        text_rows = []
        for r in rows:
            parts = sorted(r["parts"], key=lambda l: l["x0"])
            text_rows.append(" ".join(" ".join(t for _, t in sorted(p["w"])) for p in parts))
        out.append("\n".join(text_rows))
    return "\n\n".join(out)


def ocr_image(pil_img):
    langs = _ocr_available()
    if not langs:
        return ""
    import pytesseract
    lang, cfg = "+".join(sorted(langs)), "--oem 1 --psm 3"
    try:
        d = pytesseract.image_to_data(pil_img, lang=lang, config=cfg, output_type=pytesseract.Output.DICT)
        return _text_from_data(d, pil_img, langs)
    except Exception:
        return pytesseract.image_to_string(pil_img, lang=lang, config=cfg)


_w_docs = {}


def _ocr_worker(args):
    """Runs in a worker process: OCR one page. Returns (page_no, text, error)."""
    path, pno, dpi = args
    try:
        d = _w_docs.get(path)
        if d is None:
            d = _w_docs[path] = fitz.open(path)
        from PIL import Image
        pix = d[pno].get_pixmap(dpi=dpi, colorspace=fitz.csGRAY, alpha=False)
        img = Image.frombytes("L", (pix.width, pix.height), pix.samples)
        return pno, ocr_image(img), None
    except Exception as e:  # never kill the pool
        return pno, "", str(e)


def to_latex(pil_img):
    global _latex_model
    if _latex_model is False:
        return ""
    try:
        if _latex_model is None:
            from pix2tex.cli import LatexOCR
            _latex_model = LatexOCR()
        return _latex_model(pil_img)
    except Exception:
        _latex_model = False
        return ""


# ----------------------------------------------------------------- detection helpers
def looks_garbled(text):
    if not text.strip():
        return True
    bad = text.count("�") + text.count("(cid:") * 5
    private = sum(1 for c in text if 0xE000 <= ord(c) <= 0xF8FF)
    return (bad + private) / max(len(text), 1) > 0.03


def is_math_line(text):
    t = text.strip()
    if len(t) < 3 or len(t) > 160:
        return False
    if any(c in STRONG_MATH for c in t):
        return True
    if sum(c in GREEK for c in t) >= 1 and any(c in "=+−-" for c in t):
        return True
    return bool(MATH_RE.search(t)) and sum(c.isalpha() for c in t) / len(t) < 0.75


def _overlap(a, b):
    # PyMuPDF's Rect.intersects/union ignore zero-width/height rects (straight lines), so do it by hand.
    return a.x0 <= b.x1 and b.x0 <= a.x1 and a.y0 <= b.y1 and b.y0 <= a.y1


def _union(a, b):
    return fitz.Rect(min(a.x0, b.x0), min(a.y0, b.y0), max(a.x1, b.x1), max(a.y1, b.y1))


def _merge_rects(rects, gap=12, cell=8.0):
    """Group rectangles that touch or lie within `gap` points. Grid + union-find: fast even for
    pages with tens of thousands of drawing commands."""
    parent = {}

    def find(c):
        while parent[c] != c:
            parent[c] = parent[parent[c]]
            c = parent[c]
        return c

    h = gap / 2
    firsts = []
    for r in rects:
        x0, y0, x1, y1 = r.x0 - h, r.y0 - h, r.x1 + h, r.y1 + h
        cells = [(cx, cy) for cx in range(int(x0 // cell), int(x1 // cell) + 1)
                 for cy in range(int(y0 // cell), int(y1 // cell) + 1)]
        for c in cells:
            parent.setdefault(c, c)
        a = find(cells[0])
        for c in cells[1:]:
            b = find(c)
            if a != b:
                parent[b] = a
        firsts.append(cells[0])
    boxes = {}
    for r, c in zip(rects, firsts):
        root = find(c)
        boxes[root] = _union(boxes[root], r) if root in boxes else fitz.Rect(r)
    return list(boxes.values())


def find_vector_graphs(page, drawings, table_rects):
    rects = []
    for d in drawings:
        r = d.get("rect")
        if r is None:
            continue
        if r.width > page.rect.width * 0.98 and r.height > page.rect.height * 0.98:
            continue  # page background
        rects.append(r)
    graphs = []
    for r in _merge_rects(rects):
        if r.width < 60 or r.height < 40:
            continue
        if any(_overlap(r, t) and r.get_area() > 0 and
               (min(r.x1, t.x1) - max(r.x0, t.x0)) * (min(r.y1, t.y1) - max(r.y0, t.y0)) > 0.5 * r.get_area()
               for t in table_rects):
            continue
        graphs.append(r)
    return graphs


def _inside(rect, rects):
    c = fitz.Point((rect.x0 + rect.x1) / 2, (rect.y0 + rect.y1) / 2)
    for i, r in enumerate(rects):
        if r.contains(c):
            return i
    return -1


# ----------------------------------------------------------------- docx helpers
_RPR_ORDER = ["rStyle", "rFonts", "b", "bCs", "i", "iCs", "caps", "smallCaps", "strike", "dstrike", "outline",
              "shadow", "emboss", "imprint", "noProof", "snapToGrid", "vanish", "webHidden", "color", "spacing",
              "w", "kern", "position", "sz", "szCs", "highlight", "u", "effect", "bdr", "shd", "fitText",
              "vertAlign", "rtl", "cs", "em", "lang", "eastAsianLayout", "specVanish", "oMath"]
_BAD_XML = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f￾￿]")


def clean(text):
    return _BAD_XML.sub("", text)


def _set_rpr_extras(rpr, size_pt=None, bold=False, italic=False):
    rf = rpr.find(qn("w:rFonts"))
    if rf is None:
        rf = OxmlElement("w:rFonts")
        rpr.append(rf)
    rf.set(qn("w:ascii"), LATIN_FONT)
    rf.set(qn("w:hAnsi"), LATIN_FONT)
    rf.set(qn("w:cs"), CS_FONT)
    rf.set(qn("w:eastAsia"), CS_FONT)
    for tag, on in (("w:bCs", bold), ("w:iCs", italic)):
        if on and rpr.find(qn(tag)) is None:
            rpr.append(OxmlElement(tag))
    if size_pt and rpr.find(qn("w:szCs")) is None:
        e = OxmlElement("w:szCs")
        e.set(qn("w:val"), str(int(round(size_pt * 2))))
        rpr.append(e)
    kids = list(rpr)  # keep schema order so Word never reports "unreadable content"
    kids.sort(key=lambda k: _RPR_ORDER.index(k.tag.split("}")[1]) if k.tag.split("}")[1] in _RPR_ORDER else 99)
    for k in kids:
        rpr.remove(k)
    for k in kids:
        rpr.append(k)


def style_run(run, size_pt, bold=False, italic=False, color=None, sup=False):
    run.font.size = Pt(size_pt)
    run.bold = bold or None
    run.italic = italic or None
    if sup:
        run.font.superscript = True
    if color:
        run.font.color.rgb = RGBColor(*color)
    _set_rpr_extras(run._element.get_or_add_rPr(), size_pt, bold, italic)


def add_text(par, text, size_pt=11, **kw):
    run = par.add_run(clean(text))
    style_run(run, size_pt, **kw)
    return run


def add_picture_par(doc, img_bytes, width_pt, max_w_in, center=True):
    par = doc.add_paragraph()
    if center:
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_after = Pt(4)
    width = min(Inches(width_pt / 72), Inches(max_w_in))
    par.add_run().add_picture(io.BytesIO(img_bytes), width=width)
    doc._img_bytes = getattr(doc, "_img_bytes", 0) + len(img_bytes)
    return par


def add_caption(doc, text, center=True):
    if not text.strip():
        return
    par = doc.add_paragraph()
    if center:
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    par.paragraph_format.space_after = Pt(6)
    add_text(par, text, 8.5, italic=True, color=(110, 110, 110))


def add_table(doc, rows):
    rows = [[clean((c or "").strip()) for c in r] for r in rows if r]
    if not rows:
        return
    w = max(len(r) for r in rows)
    rows = [r + [""] * (w - len(r)) for r in rows]
    keep = [j for j in range(w) if any(r[j] for r in rows)]  # drop columns empty in every row
    rows = [[r[j] for j in keep] for r in rows]
    w = len(keep)
    if w == 0:
        return
    t = doc.add_table(rows=len(rows), cols=w)
    t.style = "Table Grid"
    for i, r in enumerate(rows):
        for j in range(w):
            add_text(t.cell(i, j).paragraphs[0], r[j], 10)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def render(page, rect, dpi, fmt="png"):
    pix = page.get_pixmap(clip=rect & page.rect, dpi=dpi, alpha=False)
    return pix.tobytes("jpeg", jpg_quality=82) if fmt == "jpeg" else pix.tobytes("png")


def _add_paragraph(doc, lines, brect, page_w, base_x, body):
    """lines: [(spans, rect, text)] -> one Word paragraph with PDF-like emphasis and alignment."""
    par = doc.add_paragraph()
    par.paragraph_format.space_after = Pt(6)
    maxsize = max(s["size"] for ln in lines for s in ln[0])
    is_head = maxsize >= body * 1.3
    cx = (brect.x0 + brect.x1) / 2
    if abs(cx - page_w / 2) < 12 and brect.width < page_w * 0.8:
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif 10 < brect.x0 - base_x < page_w * 0.35:
        par.paragraph_format.left_indent = Pt(brect.x0 - base_x)
    for i, (spans, rect, _) in enumerate(lines):
        for s in spans:
            if not s["text"]:
                continue
            fl, fn = s["flags"], s["font"].lower()
            col = s["color"]
            add_text(par, s["text"], min(max(round(s["size"] * 2) / 2, 6), 40),
                     bold=bool(fl & 16) or "bold" in fn or is_head,
                     italic=bool(fl & 2) or "italic" in fn or "oblique" in fn,
                     sup=bool(fl & 1),
                     color=None if col == 0 else ((col >> 16) & 255, (col >> 8) & 255, col & 255))
        if i < len(lines) - 1:
            full = rect.x1 >= brect.x1 - 0.12 * brect.width and brect.width > page_w * 0.3
            if full:
                add_text(par, " ", body)
            else:
                par.add_run().add_break()


def _image_bytes(pdf, page, xref, rect):
    """Original JPEG/PNG bytes when safe (small, no transparency mask, RGB/gray); else a rendered JPEG."""
    try:
        info = pdf.extract_image(xref)
        if (info and info.get("ext") in ("jpeg", "jpg", "png") and not info.get("smask")
                and info.get("colorspace", 3) in (1, 3) and len(info["image"]) <= 3_000_000):
            return info["image"]
    except Exception:
        pass
    return render(page, rect, 150, "jpeg")


def _body_size(pdf):
    sizes = []
    for pno in range(min(len(pdf), 10)):
        for b in pdf[pno].get_text("dict")["blocks"]:
            if b["type"] == 0:
                for ln in b["lines"]:
                    for s in ln["spans"]:
                        sizes += [round(s["size"])] * max(len(s["text"].strip()), 0)
    return statistics.median(sizes) if sizes else 11


# ----------------------------------------------------------------- one digital page
def _digital_page(doc, pdf, page, table_rects, tables, drawings, body, max_w_in, latex):
    graph_rects = find_vector_graphs(page, drawings, table_rects)
    items = []  # (y0, x0, kind, data)
    blocks = [b for b in page.get_text("dict", sort=True)["blocks"] if b["type"] == 0]
    base_x = min([b["bbox"][0] for b in blocks] or [36])

    # pull short text right next to a chart (axis labels, titles, legends) into the chart picture
    short_lines = [fitz.Rect(ln["bbox"]) for b in blocks for ln in b["lines"]
                   if 0 < len("".join(s["text"] for s in ln["spans"]).strip()) <= 40
                   and _inside(fitz.Rect(ln["bbox"]), table_rects) < 0]
    for gi, g in enumerate(graph_rects):
        changed = True
        while changed:
            changed = False
            zone = fitz.Rect(g.x0 - 20, g.y0 - 20, g.x1 + 20, g.y1 + 20)
            for lr in short_lines:
                if _overlap(zone, lr) and not (g.x0 <= lr.x0 and lr.x1 <= g.x1 and g.y0 <= lr.y0 and lr.y1 <= g.y1):
                    g = _union(g, lr)
                    changed = True
        graph_rects[gi] = g
    fig_text = [[] for _ in graph_rects]

    for b in blocks:
        brect, para = fitz.Rect(b["bbox"]), []

        def flush():
            if para:
                items.append((para[0][1].y0, para[0][1].x0, "para", (list(para), brect)))
                para.clear()

        for ln in b["lines"]:
            r = fitz.Rect(ln["bbox"])
            text = "".join(s["text"] for s in ln["spans"])
            if not text.strip() or _inside(r, table_rects) >= 0:
                continue
            gi = _inside(r, graph_rects)
            if gi >= 0:
                fig_text[gi].append(text.strip())
                continue
            if is_math_line(text):
                flush()
                items.append((r.y0, r.x0, "eq", (text.strip(), r)))
            else:
                para.append((ln["spans"], r, text))
        flush()

    for rect, rows in tables:
        items.append((rect.y0, rect.x0, "table", rows))
    for gi, r in enumerate(graph_rects):
        items.append((r.y0, r.x0, "graph", (r, fig_text[gi])))
    try:
        for info in page.get_image_info(xrefs=True):
            r = fitz.Rect(info["bbox"])
            if r.width < 60 or r.height < 60 or not info.get("xref"):
                continue
            if _inside(r, graph_rects) >= 0:
                continue
            items.append((r.y0, r.x0, "image", (r, info["xref"])))
    except Exception:
        pass

    items.sort(key=lambda it: (round(it[0]), it[1]))
    page_area = page.rect.get_area()
    for _, _, kind, data in items:
        if kind == "para":
            _add_paragraph(doc, data[0], data[1], page.rect.width, base_x, body)
        elif kind == "table":
            add_table(doc, data)
        elif kind == "eq":
            text, r = data
            crop = fitz.Rect(r.x0 - 4, r.y0 - 3, r.x1 + 4, r.y1 + 3)
            png = render(page, crop, 300)
            add_picture_par(doc, png, crop.width, max_w_in)
            cap = text
            if latex:
                from PIL import Image
                tex = to_latex(Image.open(io.BytesIO(png)))
                if tex:
                    cap += f"    LaTeX: {tex}"
            add_caption(doc, cap)
        elif kind == "graph":
            r, texts = data
            r = fitz.Rect(r.x0 - 6, r.y0 - 6, r.x1 + 6, r.y1 + 6)  # pad so edge axis lines aren't clipped
            dpi = 220 if r.get_area() < 0.25 * page_area else 150
            add_picture_par(doc, render(page, r, dpi), r.width, max_w_in)
            add_caption(doc, ("Text in figure: " + " | ".join(texts)) if texts else "")
        elif kind == "image":
            r, xref = data
            add_picture_par(doc, _image_bytes(pdf, page, xref, r), r.width, max_w_in)


# ----------------------------------------------------------------- main
def extract(pdf_path, out_path, force_ocr=False, latex=False, scan_page_image=True,
            pages_per_file=200, workers=None, progress=lambda *_: None, cancel=lambda: False):
    """Convert pdf_path to Word. Returns (list_of_docx_paths, notes)."""
    pdf_path, out_path = Path(pdf_path), Path(out_path)
    if out_path.is_dir():
        out_path = out_path / (pdf_path.stem + ".docx")
    out_path = out_path.with_suffix(".docx")
    out_path.parent.mkdir(parents=True, exist_ok=True)
    stem = out_path.with_suffix("")

    pdf = fitz.open(pdf_path)
    if pdf.needs_pass:
        raise ValueError("This PDF is password-protected. Remove the password first.")
    total = len(pdf)
    body = _body_size(pdf)
    notes, parts = [], []
    ocr_langs = _ocr_available()
    if ocr_langs:
        missing = {"ben", "hin", "eng"} - set(ocr_langs)
        if missing:
            notes.append("OCR language data missing for: " + ", ".join(sorted(missing)) +
                         ". Scanned text in those scripts will not be recognised.")
    workers = workers or max(1, min(8, (os.cpu_count() or 2) - 1))
    first = pdf[0].rect
    margin = 0.8
    max_w_in = first.width / 72 - 2 * margin

    def new_doc():
        d = Document()
        s = d.sections[0]
        s.page_width, s.page_height = Pt(first.width), Pt(first.height)
        s.left_margin = s.right_margin = s.top_margin = s.bottom_margin = Inches(margin)
        n = d.styles["Normal"]
        n.font.name = LATIN_FONT
        n.font.size = Pt(11)
        _set_rpr_extras(n.element.get_or_add_rPr(), 11)
        d._img_bytes = 0
        return d

    def save_part(d, last_page):
        path = Path(f"{stem}_part{len(parts) + 1:03d}.docx")
        d.save(str(path))
        parts.append((path, last_page))

    # which pages need OCR (cheap text-layer check), and a lazily created process pool
    need_ocr = {}
    futures = {}
    pool = None

    def needs_ocr(i):
        if i not in need_ocr:
            need_ocr[i] = force_ocr or looks_garbled(pdf[i].get_text("text"))
        return need_ocr[i]

    def submit_ahead(i):
        nonlocal pool
        if not ocr_langs or workers < 2:
            return
        for j in range(i, min(total, i + workers * 3)):
            if j not in futures and needs_ocr(j):
                if pool is None:
                    pool = ProcessPoolExecutor(max_workers=workers)
                futures[j] = pool.submit(_ocr_worker, (str(pdf_path), j, 300))

    def ocr_text(i):
        if i in futures:
            _, text, err = futures.pop(i).result()
            if err:
                raise RuntimeError(err)
            return text
        _, text, err = _ocr_worker((str(pdf_path), i, 300))
        if err:
            raise RuntimeError(err)
        return text

    doc = new_doc()
    pages_in_part = 0
    cancelled = False
    try:
        for i in range(total):
            if cancel():
                cancelled = True
                notes.append(f"Stopped by user after page {i}.")
                break
            progress(i + 1, total, f"Page {i + 1}/{total}")
            page = pdf[i]
            if pages_in_part:
                doc.add_page_break()
            pg = doc.add_paragraph()
            pg.alignment = WD_ALIGN_PARAGRAPH.CENTER
            add_text(pg, f"— Page {i + 1} —", 8, color=(150, 150, 150))
            mark = len(doc.element.body)  # so a failed page can be rolled back
            try:
                if needs_ocr(i):
                    submit_ahead(i)
                    if not ocr_langs:
                        notes.append(f"Page {i + 1}: needs OCR but Tesseract with ben/hin/eng data was not found.")
                        add_caption(doc, "[OCR unavailable - install Tesseract + Bengali/Hindi data. Original page below.]")
                        text = ""
                    else:
                        text = ocr_text(i).strip()
                    for chunk in re.split(r"\n\s*\n", text):
                        chunk = " ".join(l.strip() for l in chunk.splitlines() if l.strip())
                        if chunk:
                            p = doc.add_paragraph()
                            p.paragraph_format.space_after = Pt(6)
                            add_text(p, chunk, 11)
                    if scan_page_image or not text:
                        add_caption(doc, "Original page (graphs, equations and tables exactly as scanned):")
                        add_picture_par(doc, render(page, page.rect, 150, "jpeg"), page.rect.width, max_w_in)
                else:
                    submit_ahead(i + 1)
                    drawings = page.get_drawings()
                    table_rects, tables = [], []
                    if len(drawings) >= 4:  # tables need ruled lines
                        try:
                            for t in page.find_tables().tables:
                                rows = t.extract()
                                if rows and len(rows) > 1:
                                    table_rects.append(fitz.Rect(t.bbox))
                                    tables.append((fitz.Rect(t.bbox), rows))
                        except Exception:
                            pass
                    _digital_page(doc, pdf, page, table_rects, tables, drawings, body, max_w_in, latex)
            except Exception as e:
                notes.append(f"Page {i + 1}: conversion error ({e}); page inserted as an image instead.")
                body_el = doc.element.body
                for el in list(body_el)[mark:]:
                    if el.tag != qn("w:sectPr"):
                        body_el.remove(el)
                try:
                    add_picture_par(doc, render(page, page.rect, 110, "jpeg"), page.rect.width, max_w_in)
                except Exception:
                    pass
            pages_in_part += 1
            page = None
            if i % 50 == 49:
                fitz.TOOLS.store_shrink(100)  # release PyMuPDF caches on long runs
            if pages_in_part >= pages_per_file or doc._img_bytes > MAX_PART_IMAGE_BYTES:
                save_part(doc, i + 1)
                doc = new_doc()
                pages_in_part = 0
        if pages_in_part:
            save_part(doc, total if not cancelled else i)
    finally:
        if pool is not None:
            pool.shutdown(wait=False, cancel_futures=True)
        pdf.close()

    if len(parts) == 1:  # single file: use the plain name
        final = out_path
        parts[0][0].replace(final)
        return [final], notes
    return [p for p, _ in parts], notes
