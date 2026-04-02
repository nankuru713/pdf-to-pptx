"""
PDF → PPTX 変換ツール（Gemini AI版 v2）
"""

import io, time, tempfile, os, base64, json, re
import streamlit as st

st.set_page_config(
    page_title="PDF → PPTX コンバーター",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="collapsed",
)

st.markdown("""
<style>
html, body, [data-testid="stAppViewContainer"] { background-color: #0f0f1a; color: #e8e8f0; }
[data-testid="stHeader"] { background: transparent; }
.main .block-container { padding-top: 2rem; padding-bottom: 3rem; max-width: 720px; }
.banner { background: #1a1a2e; border: 1px solid #e9456033; border-radius: 16px; padding: 2rem 2.5rem; margin-bottom: 2rem; border-left: 4px solid #e94560; }
.banner h1 { font-size: 1.8rem; font-weight: 700; color: #fff; margin: 0 0 0.4rem; }
.banner p  { color: #8888aa; margin: 0; font-size: 0.95rem; }
.accent { color: #e94560; }
.stat-grid { display: grid; grid-template-columns: repeat(3,1fr); gap: 10px; margin: 1rem 0; }
.stat-box { background: #1a1a2e; border: 1px solid #252540; border-radius: 10px; padding: 1rem; text-align: center; }
.stat-box .num { font-size: 1.8rem; font-weight: 700; color: #e94560; }
.stat-box .lbl { font-size: 0.75rem; color: #8888aa; margin-top: 2px; }
.info-card { background: #1a1a2e; border: 1px solid #252540; border-radius: 12px; padding: 1.2rem 1.5rem; margin: 0.5rem 0; }
.info-card h4 { color: #e8e8f0; margin: 0 0 0.5rem; font-size: 0.95rem; }
.info-card p  { color: #8888aa; margin: 0; font-size: 0.85rem; line-height: 1.6; }
.stButton > button { background: #e94560 !important; color: white !important; border: none !important; border-radius: 10px !important; padding: 0.65rem 2rem !important; font-size: 1rem !important; font-weight: 600 !important; width: 100% !important; }
[data-testid="stDownloadButton"] > button { background: #1a1a2e !important; color: #e94560 !important; border: 2px solid #e94560 !important; border-radius: 10px !important; width: 100% !important; font-size: 1rem !important; font-weight: 600 !important; padding: 0.65rem 2rem !important; }
[data-testid="stDownloadButton"] > button:hover { background: #e94560 !important; color: white !important; }
hr { border-color: #252540 !important; }
</style>
""", unsafe_allow_html=True)


def analyze_with_gemini(page_img_bytes: bytes, api_key: str) -> dict:
    """Gemini APIでスライドを解析してJSON構造を返す"""
    import google.generativeai as genai
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-2.0-flash")

    prompt = """このPowerPointスライドの画像を解析してください。

以下のJSON形式で返してください。JSONのみ返し、説明文は不要です：

{
  "background_color": "#1a1a2e",
  "elements": [
    {
      "type": "text",
      "text": "テキスト内容",
      "x": 0.05,
      "y": 0.05,
      "w": 0.9,
      "h": 0.1,
      "font_size": 24,
      "bold": true,
      "color": "#ffffff",
      "bg_color": null,
      "border_color": null,
      "align": "center"
    },
    {
      "type": "rect",
      "x": 0.0,
      "y": 0.8,
      "w": 1.0,
      "h": 0.2,
      "fill_color": "#003366",
      "border_color": null
    }
  ]
}

ルール:
- x,y,w,h は画像全体を1.0とした相対値（左上が原点）
- スライド上のすべてのテキストを必ず抽出すること
- 背景の矩形・色付きボックスも含めること
- JSONのみ返すこと（```json などのマークダウン記法も不要）"""

    response = model.generate_content([
        {"mime_type": "image/png", "data": base64.b64encode(page_img_bytes).decode()},
        prompt
    ])

    raw = response.text.strip()
    # マークダウンコードブロックを除去
    raw = re.sub(r"^```[a-zA-Z]*\n?", "", raw)
    raw = re.sub(r"\n?```$", "", raw)
    raw = raw.strip()

    try:
        return json.loads(raw)
    except json.JSONDecodeError:
        # JSON部分を抽出して再試行
        match = re.search(r'\{.*\}', raw, re.DOTALL)
        if match:
            return json.loads(match.group())
        raise


def hex_to_rgb(h):
    if not h or not isinstance(h, str):
        return (0, 0, 0)
    h = h.lstrip("#")
    if len(h) == 6:
        return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))
    return (0, 0, 0)


def build_slide_from_analysis(prs, analysis, page_img_bytes):
    from pptx.util import Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from lxml import etree

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    sw, sh = int(prs.slide_width), int(prs.slide_height)

    def rx(v): return max(0, int(float(v) * sw))
    def ry(v): return max(0, int(float(v) * sh))
    def rw(v): return max(1, int(float(v) * sw))
    def rh(v): return max(1, int(float(v) * sh))

    def clamp_box(l, t, w, h):
        l = max(0, min(int(l), sw - 1))
        t = max(0, min(int(t), sh - 1))
        w = max(1, min(int(w), sw - l))
        h = max(1, min(int(h), sh - t))
        return l, t, w, h

    # 背景色を設定
    bg_color = analysis.get("background_color", "")
    if bg_color and len(bg_color) == 7:
        try:
            r, g, b = hex_to_rgb(bg_color)
            NP = "http://schemas.openxmlformats.org/presentationml/2006/main"
            NA = "http://schemas.openxmlformats.org/drawingml/2006/main"
            xml = (f'<p:bg xmlns:p="{NP}" xmlns:a="{NA}"><p:bgPr>'
                   f'<a:solidFill><a:srgbClr val="{r:02X}{g:02X}{b:02X}"/></a:solidFill>'
                   f'<a:effectLst/></p:bgPr></p:bg>')
            slide.shapes._spTree.getparent().insert(2, etree.fromstring(xml))
        except:
            # 背景設定失敗時はページ画像を背景として配置
            slide.shapes.add_picture(io.BytesIO(page_img_bytes), 0, 0, sw, sh)
    else:
        # 背景色がない場合はページ画像を背景として配置
        slide.shapes.add_picture(io.BytesIO(page_img_bytes), 0, 0, sw, sh)

    ALIGN_MAP = {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
    }

    elements = analysis.get("elements", [])

    for el in elements:
        try:
            et = el.get("type", "")
            x = el.get("x", 0)
            y = el.get("y", 0)
            w = el.get("w", 0.1)
            h = el.get("h", 0.1)
            l, t, bw, bh = clamp_box(rx(x), ry(y), rw(w), rh(h))

            if et == "rect":
                shp = slide.shapes.add_shape(1, l, t, bw, bh)
                fc = el.get("fill_color")
                if fc:
                    shp.fill.solid()
                    shp.fill.fore_color.rgb = RGBColor(*hex_to_rgb(fc))
                else:
                    shp.fill.background()
                bc = el.get("border_color")
                if bc:
                    shp.line.color.rgb = RGBColor(*hex_to_rgb(bc))
                    shp.line.width = Pt(1)
                else:
                    shp.line.fill.background()

            elif et == "text":
                text_content = el.get("text", "")
                if not text_content:
                    continue

                txb = slide.shapes.add_textbox(l, t, bw, bh)
                tf = txb.text_frame
                tf.word_wrap = True
                tf.margin_left = tf.margin_right = Pt(3)
                tf.margin_top = tf.margin_bottom = Pt(2)

                # テキストボックスの背景色
                bg_c = el.get("bg_color")
                if bg_c:
                    txb.fill.solid()
                    txb.fill.fore_color.rgb = RGBColor(*hex_to_rgb(bg_c))

                # 枠線
                border_c = el.get("border_color")
                if border_c:
                    txb.line.color.rgb = RGBColor(*hex_to_rgb(border_c))
                    txb.line.width = Pt(1.5)

                align = ALIGN_MAP.get(el.get("align", "left"), PP_ALIGN.LEFT)
                font_size = max(float(el.get("font_size", 16)) * 0.75, 8)
                is_bold = el.get("bold", False)
                text_color = el.get("color", "#000000")

                lines = text_content.split("\n")
                para = tf.paragraphs[0]
                para.alignment = align

                for li, line in enumerate(lines):
                    if li > 0:
                        para = tf.add_paragraph()
                        para.alignment = align
                    if not line.strip():
                        continue
                    run = para.add_run()
                    run.text = line
                    f = run.font
                    f.size = Pt(font_size)
                    f.bold = is_bold
                    f.color.rgb = RGBColor(*hex_to_rgb(text_color))

        except Exception as e:
            continue  # 1要素の失敗でスキップ

    return slide


def convert_pdf_gemini(pdf_bytes, api_key, slide_width_inches=13.333, progress_cb=None):
    import fitz
    from pptx import Presentation
    from pptx.util import Inches

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name

    try:
        doc = fitz.open(tmp_path)
        n = len(doc)
        first = doc[0]
        aspect = first.rect.height / first.rect.width
        doc.close()

        prs = Presentation()
        prs.slide_width = Inches(slide_width_inches)
        prs.slide_height = int(Inches(slide_width_inches) * aspect)

        stats = {"pages": n, "texts": 0, "elements": 0, "errors": []}

        for i in range(n):
            if progress_cb:
                progress_cb(i, n, f"ページ {i+1}/{n} を画像化中...")

            # ページを高解像度画像に変換
            doc2 = fitz.open(tmp_path)
            pix = doc2[i].get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
            img_bytes = pix.tobytes("png")
            doc2.close()

            if progress_cb:
                progress_cb(i, n, f"ページ {i+1}/{n} をGemini AIで解析中...")

            try:
                analysis = analyze_with_gemini(img_bytes, api_key)
                elements = analysis.get("elements", [])
                n_txt = sum(1 for e in elements if e.get("type") == "text")
                stats["texts"] += n_txt
                stats["elements"] += len(elements)

                if progress_cb:
                    progress_cb(i, n, f"ページ {i+1}/{n}: テキスト{n_txt}個を配置中...")

                build_slide_from_analysis(prs, analysis, img_bytes)

            except Exception as e:
                stats["errors"].append(f"P{i+1}: {str(e)[:50]}")
                if progress_cb:
                    progress_cb(i, n, f"ページ {i+1}/{n}: AI解析失敗→画像として配置")
                # フォールバック: 画像をそのまま配置
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                slide.shapes.add_picture(
                    io.BytesIO(img_bytes), 0, 0,
                    int(prs.slide_width), int(prs.slide_height)
                )

        buf = io.BytesIO()
        prs.save(buf)
        return buf.getvalue(), stats

    finally:
        os.unlink(tmp_path)


# ── UI ──────────────────────────────────────────────────────

gemini_key = ""
if hasattr(st, "secrets"):
    gemini_key = st.secrets.get("GEMINI_API_KEY", "")

st.markdown("""
<div class="banner">
  <h1>📄 PDF <span class="accent">→</span> PPTX コンバーター</h1>
  <p>Gemini AIが画像化されたテキストも認識して、編集可能なPowerPointに変換します（<span class="accent">無料</span>）</p>
</div>
""", unsafe_allow_html=True)

# APIキー入力欄
if not gemini_key:
    st.markdown("### 🔑 Google AI Studio APIキー")
    gemini_key = st.text_input(
        "APIキー（無料で取得できます）",
        type="password",
        placeholder="AIzaSy...",
    )
    if not gemini_key:
        st.info("💡 [aistudio.google.com](https://aistudio.google.com) → 「Get API key」→「Create API key」で取得")
        st.stop()

with st.sidebar:
    st.markdown("### ⚙️ 変換設定")
    slide_width = st.select_slider(
        "スライドサイズ",
        options=[10.0, 12.0, 13.333, 16.0],
        value=13.333,
        format_func=lambda x: {
            10.0: "10\" (4:3)",
            12.0: "12\" (カスタム)",
            13.333: "13.3\" (16:9 標準)",
            16.0: "16\" (ワイド)",
        }[x],
    )
    st.markdown("---")
    st.markdown("### ✅ 変換できる要素")
    for icon, label in [
        ("✅", "画像化テキスト（AI認識）"),
        ("✅", "テキストボックス（編集可能）"),
        ("✅", "背景色"),
        ("✅", "色付きボックス・矩形"),
        ("⚠️", "複雑なグラデーション"),
    ]:
        st.markdown(f"{icon} {label}")

# ファイルアップロード
uploaded = st.file_uploader(
    "PDFファイルをここにドロップ、またはクリックして選択",
    type=["pdf"],
    help="スライド形式のPDF推奨（最大200MB）",
)

def fmt_size(n):
    if n < 1024: return f"{n} B"
    if n < 1024**2: return f"{n/1024:.1f} KB"
    return f"{n/1024**2:.1f} MB"

if uploaded:
    st.markdown("---")
    c1, c2, c3 = st.columns(3)
    with c1: st.metric("ファイル名", uploaded.name[:14] + "…" if len(uploaded.name) > 14 else uploaded.name)
    with c2: st.metric("サイズ", fmt_size(uploaded.size))
    with c3: st.metric("AIモード", "Gemini")
    st.markdown("")

    if st.button("🚀　AI変換を開始する", use_container_width=True):
        pdf_bytes = uploaded.read()

        progress_bar = st.progress(0.0)
        status_text = st.empty()

        import fitz as _fitz
        _doc = _fitz.open(stream=pdf_bytes, filetype="pdf")
        n_pages = len(_doc)
        _doc.close()

        def progress_cb(page_i, total, msg):
            pct = (page_i + 0.5) / total
            progress_bar.progress(min(pct, 0.99))
            status_text.markdown(f"**{msg}**")

        try:
            pptx_bytes, stats = convert_pdf_gemini(
                pdf_bytes, gemini_key, slide_width,
                progress_cb=progress_cb,
            )
            progress_bar.progress(1.0)
            status_text.markdown("**✅ 変換完了！**")

            if stats["errors"]:
                st.warning("⚠️ 一部のページでエラーが発生しました：\n" + "\n".join(stats["errors"]))

            st.success(f"✨ {stats['pages']}スライドを変換！テキスト **{stats['texts']}個** が編集可能です。")

            st.markdown(f"""
<div class="stat-grid">
  <div class="stat-box"><div class="num">{stats['pages']}</div><div class="lbl">スライド</div></div>
  <div class="stat-box"><div class="num">{stats['texts']}</div><div class="lbl">編集可能テキスト</div></div>
  <div class="stat-box"><div class="num">{stats['elements']}</div><div class="lbl">総要素数</div></div>
</div>
""", unsafe_allow_html=True)

            out_name = uploaded.name.replace(".pdf", "_AI変換.pptx").replace(".PDF", "_AI変換.pptx")
            st.download_button(
                label="⬇️　PPTXをダウンロード",
                data=pptx_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )

        except Exception as e:
            import traceback
            progress_bar.empty()
            status_text.empty()
            st.error(f"❌ エラーが発生しました: {e}")
            with st.expander("エラー詳細（開発者向け）"):
                st.code(traceback.format_exc())

else:
    st.markdown("")
    c1, c2 = st.columns(2)
    with c1:
        st.markdown('<div class="info-card"><h4>🎯 こんなPDFに最適</h4><p>・画像化されたスライド<br>・テキスト埋め込みの資料<br>・レイアウト再編集が必要なもの</p></div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="info-card"><h4>⚡ AI変換の特徴</h4><p>・画像内テキストをAIが認識<br>・編集可能なテキストボックスに変換<br>・Gemini Flash（無料）使用</p></div>', unsafe_allow_html=True)

st.markdown("---")
st.markdown("<p style='text-align:center;color:#555577;font-size:0.8rem;'>PDF → PPTX Converter（Gemini AI版 v2）| Powered by Google Gemini + python-pptx</p>", unsafe_allow_html=True)
