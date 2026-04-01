#!/usr/bin/env python3
"""
PDF → PPTX 高精度変換ツール  v3.0
=====================================

アーキテクチャ（3フェーズ）:

  Phase 1 — 解析 (PyMuPDF)
    ├─ テキストブロック: 位置・フォント・色・サイズをスパン単位で抽出
    ├─ 画像: 埋め込みラスタ画像を座標付きで取得
    ├─ 矩形/シェイプ: ベクタ描画を矩形近似、全面背景は背景色として処理
    └─ 背景色: 最初の全面矩形から推定

  Phase 2 — 変換 (python-pptx)
    ├─ 座標変換: PDF pt → EMU (1pt = 914400/72 EMU)、スライドサイズに合わせてスケーリング
    ├─ 境界クランプ: スライド外にはみ出す要素を強制クリップ
    ├─ フォントマッピング: 日本語CJKフォント名を代替マッピング
    └─ テキストボックス: word_wrap=False で元レイアウト再現

  Phase 3 — 最適化・保存
    └─ テキストフォントサイズをスケール比で補正

技術的制限（既知）:
  - 曲線・多角形ベクタ → 矩形バウンディングボックスで近似
  - PDFフォントが未埋め込みの場合 → Arial/ゴシックにフォールバック
  - 透明度・グラデーション・ドロップシャドウ → 非対応（無視）
  - 画像として埋め込まれたテキスト → OCRが別途必要（フラグ付き警告）
  - 表（テーブル）→ 行ごとのテキストボックス（セル分割は非対応）
"""

import sys
import io
from dataclasses import dataclass, field
from typing import Optional

import fitz                          # PyMuPDF
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from PIL import Image

# -------------------------------------------------------------------
#  定数
# -------------------------------------------------------------------
PT_TO_EMU   = 914400 / 72   # 1 pt = 12700 EMU
WIDE_INCHES = 13.333333     # 16:9 ワイドスクリーン幅

# 日本語フォントマッピング（PDFフォント名 → Windowsフォント名）
JP_FONT_MAP = {
    "kozminpr6n":    "游明朝",
    "kozgopr6n":     "游ゴシック",
    "hirakakupro":   "ヒラギノ角ゴ Pro W3",
    "hirakakupron":  "ヒラギノ角ゴ ProN W3",
    "hiraminpro":    "ヒラギノ明朝 Pro W3",
    "msgothic":      "ＭＳ ゴシック",
    "mspgothic":     "ＭＳ Ｐゴシック",
    "msmincho":      "ＭＳ 明朝",
    "yugothic":      "游ゴシック",
    "yumincho":      "游明朝",
    "notosanscjk":   "Noto Sans CJK JP",
    "heiseikakugo":  "ＭＳ ゴシック",
    "heiseimincho":  "ＭＳ 明朝",
    "ipagothic":     "IPAゴシック",
    "ipamincho":     "IPA明朝",
    "meiryo":        "メイリオ",
    "meiryoui":      "Meiryo UI",
}


# -------------------------------------------------------------------
#  データモデル
# -------------------------------------------------------------------

@dataclass
class TextSpan:
    text: str
    font_name: str
    font_size: float
    bold: bool
    italic: bool
    color: tuple


@dataclass
class TextBlock:
    x0: float; y0: float
    x1: float; y1: float
    spans: list
    align: str = "left"


@dataclass
class ImageBlock:
    x0: float; y0: float
    x1: float; y1: float
    image_bytes: bytes
    ext: str = "png"


@dataclass
class RectBlock:
    x0: float; y0: float
    x1: float; y1: float
    fill_color: Optional[tuple]
    line_color: Optional[tuple]
    line_width: float = 1.0


@dataclass
class PageData:
    page_num: int
    width_pt: float
    height_pt: float
    texts: list = field(default_factory=list)
    images: list = field(default_factory=list)
    rects: list = field(default_factory=list)
    bg_color: Optional[tuple] = None


# -------------------------------------------------------------------
#  ユーティリティ
# -------------------------------------------------------------------

def _to_rgb(c):
    if c is None:
        return None
    if isinstance(c, float) and 0.0 <= c <= 1.0:
        g = int(c * 255); return (g, g, g)
    if isinstance(c, int):
        return ((c >> 16) & 0xFF, (c >> 8) & 0xFF, c & 0xFF)
    if isinstance(c, (tuple, list)):
        if len(c) == 1: g = int(c[0]*255); return (g,g,g)
        if len(c) == 3: return tuple(max(0,min(255,int(x*255))) for x in c)
        if len(c) == 4:
            C,M,Y,K = c
            return (int(255*(1-C)*(1-K)), int(255*(1-M)*(1-K)), int(255*(1-Y)*(1-K)))
    return None


def _clamp_rgb(c):
    if c is None: return None
    return tuple(max(0, min(255, int(v))) for v in c)


def _resolve_font(raw_name: str, bold: bool, italic: bool):
    """PDFフォント名 → (フォント名, bold, italic)"""
    name = raw_name
    for kw in ("-BoldItalic", "-BoldOblique", "BoldItalic"):
        if kw in name: name = name.replace(kw,""); bold = italic = True
    for kw in ("-Bold","Bold","-Heavy","Heavy","-Black","Black"):
        if kw in name: name = name.replace(kw,""); bold = True
    for kw in ("-Italic","-Oblique","Italic","Oblique"):
        if kw in name: name = name.replace(kw,""); italic = True
    name = name.strip("-_ ")
    lower = name.lower()
    for key, mapped in JP_FONT_MAP.items():
        if key in lower:
            return (mapped, bold, italic)
    return (name if name else "Arial", bold, italic)


# -------------------------------------------------------------------
#  Phase 1 — PDF 解析
# -------------------------------------------------------------------

def _detect_bg_color(page):
    pw, ph = page.rect.width, page.rect.height
    for path in page.get_drawings():
        r = path.get("rect")
        if r and r.x0 <= 2 and r.y0 <= 2 and r.x1 >= pw-2 and r.y1 >= ph-2:
            fill = _clamp_rgb(_to_rgb(path.get("fill")))
            if fill:
                return fill
    return None


def extract_texts(page):
    from collections import Counter
    blocks = []
    raw = page.get_text("dict", flags=fitz.TEXT_PRESERVE_WHITESPACE)
    for blk in raw.get("blocks", []):
        if blk.get("type") != 0:
            continue
        bbox = blk["bbox"]
        bx_mid = (bbox[0]+bbox[2])/2
        spans_out, aligns = [], []
        for line in blk.get("lines", []):
            lx_mid = (line["bbox"][0]+line["bbox"][2])/2
            diff = lx_mid - bx_mid
            aligns.append("center" if abs(diff)<15 else ("right" if diff>25 else "left"))
            for span in line.get("spans", []):
                txt = span.get("text","")
                if not txt: continue
                flags = span.get("flags",0)
                color = _clamp_rgb(_to_rgb(span.get("color",0)))
                spans_out.append(TextSpan(
                    text=txt,
                    font_name=span.get("font","Arial"),
                    font_size=span.get("size",12),
                    bold=bool(flags & (1<<4)),
                    italic=bool(flags & (1<<1)),
                    color=color if color else (0,0,0),
                ))
        if not spans_out: continue
        align = Counter(aligns).most_common(1)[0][0] if aligns else "left"
        blocks.append(TextBlock(
            x0=bbox[0],y0=bbox[1],x1=bbox[2],y1=bbox[3],
            spans=spans_out, align=align,
        ))
    return blocks


def extract_images(page):
    images = []
    for info in page.get_images(full=True):
        xref = info[0]
        rects = page.get_image_rects(xref)
        if not rects: continue
        r = rects[0]
        if (r.x1-r.x0) < 5 or (r.y1-r.y0) < 5: continue
        try:
            d = page.parent.extract_image(xref)
            raw, ext = d.get("image",b""), d.get("ext","png").lower()
            if not raw: continue
            if ext not in ("png","jpeg","jpg"):
                buf = io.BytesIO()
                Image.open(io.BytesIO(raw)).save(buf,"PNG")
                raw, ext = buf.getvalue(), "png"
            images.append(ImageBlock(x0=r.x0,y0=r.y0,x1=r.x1,y1=r.y1,image_bytes=raw,ext=ext))
        except:
            continue
    return images


def extract_rects(page):
    rects = []
    pw, ph = page.rect.width, page.rect.height
    for path in page.get_drawings():
        r = path.get("rect")
        if not r: continue
        # ページ外をクリップ
        x0 = max(0.0, float(r.x0))
        y0 = max(0.0, float(r.y0))
        x1 = min(pw,  float(r.x1))
        y1 = min(ph,  float(r.y1))
        w, h = x1-x0, y1-y0
        if w < 3 or h < 3: continue
        # 全面背景はスキップ
        if x0 <= 2 and y0 <= 2 and x1 >= pw-2 and y1 >= ph-2: continue
        fill   = _clamp_rgb(_to_rgb(path.get("fill")))
        stroke = _clamp_rgb(_to_rgb(path.get("color")))
        if fill is None and stroke is None: continue
        rects.append(RectBlock(
            x0=x0,y0=y0,x1=x1,y1=y1,
            fill_color=fill, line_color=stroke,
            line_width=float(path.get("width") or 1.0),
        ))
    return rects


def analyze_pdf(pdf_path: str):
    doc = fitz.open(pdf_path)
    pages = []
    n = len(doc)
    warned_pages = []
    for i, page in enumerate(doc):
        print(f"  解析中: {i+1}/{n} ページ", end="\r")
        # 画像テキスト化チェック
        has_imgs = len(page.get_images(full=True)) > 0
        has_text = any(b.get("type")==0 for b in page.get_text("dict")["blocks"])
        if has_imgs and not has_text:
            warned_pages.append(i+1)
        pages.append(PageData(
            page_num=i+1,
            width_pt=page.rect.width, height_pt=page.rect.height,
            texts=extract_texts(page),
            images=extract_images(page),
            rects=extract_rects(page),
            bg_color=_detect_bg_color(page),
        ))
    print()
    doc.close()
    if warned_pages:
        print(f"  ⚠  ページ {warned_pages} : テキストが画像化されている可能性があります")
        print(f"      → OCRが必要な場合は pdf2image + pytesseract を検討してください")
    return pages


# -------------------------------------------------------------------
#  Phase 2 — PPTX 構築
# -------------------------------------------------------------------

ALIGN_MAP = {
    "left":   PP_ALIGN.LEFT,
    "center": PP_ALIGN.CENTER,
    "right":  PP_ALIGN.RIGHT,
}


def _set_slide_bg(slide, rgb: tuple):
    """スライド背景色を XML 直接操作で設定"""
    from lxml import etree
    r, g, b = rgb
    NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main"
    NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
    xml = (
        f'<p:bg xmlns:p="{NS_P}" xmlns:a="{NS_A}">'
        f'<p:bgPr>'
        f'<a:solidFill><a:srgbClr val="{r:02X}{g:02X}{b:02X}"/></a:solidFill>'
        f'<a:effectLst/>'
        f'</p:bgPr>'
        f'</p:bg>'
    )
    slide.shapes._spTree.getparent().insert(2, etree.fromstring(xml))


def build_slide(prs: Presentation, page: PageData):
    slide = prs.slides.add_slide(prs.slide_layouts[6])
    sw = int(prs.slide_width)
    sh = int(prs.slide_height)
    pw = page.width_pt
    ph = page.height_pt

    def _emu(pt): return int(pt * PT_TO_EMU)
    sx = sw / _emu(pw)
    sy = sh / _emu(ph)

    def ex(v): return int(_emu(v) * sx)
    def ey(v): return int(_emu(v) * sy)
    def ew(v): return max(int(_emu(max(v,3)) * sx), 1)
    def eh(v): return max(int(_emu(max(v,3)) * sy), 1)

    def clamp(l, t, w, h):
        l = max(0, min(l, sw-1))
        t = max(0, min(t, sh-1))
        w = max(1, min(w, sw-l))
        h = max(1, min(h, sh-t))
        return l, t, w, h

    # 背景色
    if page.bg_color:
        _set_slide_bg(slide, page.bg_color)

    # 矩形シェイプ
    for rect in page.rects:
        if not rect.fill_color and not rect.line_color: continue
        l, t, w, h = clamp(ex(rect.x0),ey(rect.y0),ew(rect.x1-rect.x0),eh(rect.y1-rect.y0))
        try:
            shp = slide.shapes.add_shape(1, l, t, w, h)
            if rect.fill_color:
                shp.fill.solid()
                shp.fill.fore_color.rgb = RGBColor(*rect.fill_color)
            else:
                shp.fill.background()
            if rect.line_color:
                shp.line.color.rgb = RGBColor(*rect.line_color)
                shp.line.width = Pt(max(rect.line_width, 0.25))
            else:
                shp.line.fill.background()
        except Exception as e:
            print(f"    ⚠ シェイプスキップ: {e}")

    # 画像
    for img in page.images:
        l, t, w, h = clamp(ex(img.x0),ey(img.y0),ew(img.x1-img.x0),eh(img.y1-img.y0))
        try:
            slide.shapes.add_picture(io.BytesIO(img.image_bytes), l, t, w, h)
        except Exception as e:
            print(f"    ⚠ 画像スキップ: {e}")

    # テキストボックス
    font_scale = min(sx, sy)
    for tb in page.texts:
        raw_w = max(tb.x1-tb.x0, 20) + 20   # +20pt 余白で折り返し防止
        raw_h = max(tb.y1-tb.y0, 10) + 8
        l, t, w, h = clamp(ex(tb.x0),ey(tb.y0),ew(raw_w),eh(raw_h))

        txb = slide.shapes.add_textbox(l, t, w, h)
        tf = txb.text_frame
        tf.word_wrap = False
        tf.margin_left = tf.margin_right = tf.margin_top = tf.margin_bottom = Pt(0)

        para = tf.paragraphs[0]
        para.alignment = ALIGN_MAP.get(tb.align, PP_ALIGN.LEFT)

        for span in tb.spans:
            for li, line_text in enumerate(span.text.split("\n")):
                if li > 0:
                    para = tf.add_paragraph()
                    para.alignment = ALIGN_MAP.get(tb.align, PP_ALIGN.LEFT)
                if not line_text: continue
                run = para.add_run()
                run.text = line_text
                fn, bd, it = _resolve_font(span.font_name, span.bold, span.italic)
                f = run.font
                f.name   = fn
                f.size   = Pt(span.font_size * font_scale)
                f.bold   = bd
                f.italic = it
                f.color.rgb = RGBColor(*span.color)
    return slide


# -------------------------------------------------------------------
#  Phase 3 — アセンブリ・保存
# -------------------------------------------------------------------

def convert(pdf_path: str, output_path: str, slide_width_inches: float = WIDE_INCHES):
    """
    PDF → PPTX 変換メイン関数。

    Args:
        pdf_path: 入力 PDF
        output_path: 出力 PPTX
        slide_width_inches: スライド幅インチ（デフォルト 13.33 = 16:9）

    Returns:
        output_path (str)
    """
    print(f"\n{'═'*54}")
    print(f"  PDF → PPTX 高精度変換  v3.0")
    print(f"  入力 : {pdf_path}")
    print(f"  出力 : {output_path}")
    print(f"{'═'*54}")

    print("\n[Phase 1]  PDF 解析中...")
    pages = analyze_pdf(pdf_path)
    print(f"  → {len(pages)} ページ 解析完了")

    first_page = pages[0]
    aspect     = first_page.height_pt / first_page.width_pt
    slide_w    = Inches(slide_width_inches)
    slide_h    = int(slide_w * aspect)

    print("\n[Phase 2]  PPTX 構築中...")
    prs = Presentation()
    prs.slide_width  = slide_w
    prs.slide_height = slide_h
    for i, page in enumerate(pages):
        print(f"  スライド {i+1}/{len(pages)} を構築中 ...", end="\r")
        build_slide(prs, page)
    print()

    print("\n[Phase 3]  保存中...")
    prs.save(output_path)
    print(f"  → 保存完了: {output_path}")

    n_text  = sum(len(p.texts)  for p in pages)
    n_img   = sum(len(p.images) for p in pages)
    n_rect  = sum(len(p.rects)  for p in pages)
    print(f"\n{'─'*44}")
    print(f"  スライド数        :  {len(pages)}")
    print(f"  テキストボックス  :  {n_text}")
    print(f"  画像              :  {n_img}")
    print(f"  シェイプ（矩形）  :  {n_rect}")
    print(f"{'─'*44}\n")
    return output_path


# -------------------------------------------------------------------
#  CLI エントリーポイント
# -------------------------------------------------------------------

if __name__ == "__main__":
    if len(sys.argv) < 3:
        print("使い方: python pdf_to_pptx.py <input.pdf> <output.pptx> [幅インチ]")
        print("例    : python pdf_to_pptx.py slides.pdf output.pptx 13.33")
        sys.exit(1)
    _width = float(sys.argv[3]) if len(sys.argv) > 3 else WIDE_INCHES
    convert(sys.argv[1], sys.argv[2], _width)
