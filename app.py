"""
PDF → PPTX 変換ツール（Gemini AI版・無料）
Streamlit アプリ（クラウドデプロイ対応）
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
.banner {
    background: #1a1a2e; border: 1px solid #e9456033;
    border-radius: 16px; padding: 2rem 2.5rem; margin-bottom: 2rem;
    border-left: 4px solid #e94560;
}
.banner h1 { font-size: 1.8rem; font-weight: 700; color: #fff; margin: 0 0 0.4rem; }
.banner p  { color: #8888aa; margin: 0; font-size: 0.95rem; }
.banner .accent { color: #e94560; }
.stat-grid { display: grid; grid-template-columns: repeat(3,1fr); gap: 10px; margin: 1rem 0; }
.stat-box { background: #1a1a2e; border: 1px solid #252540; border-radius: 10px; padding: 1rem; text-align: center; }
.stat-box .num { font-size: 1.8rem; font-weight: 700; color: #e94560; }
.stat-box .lbl { font-size: 0.75rem; color: #8888aa; margin-top: 2px; }
.warn-card { background: #2a1a10; border: 1px solid #e94560aa; border-radius: 10px; padding: 1rem 1.2rem; margin: 1rem 0; font-size: 0.85rem; color: #ffaa80; }
.info-card { background: #1a1a2e; border: 1px solid #252540; border-radius: 12px; padding: 1.2rem 1.5rem; margin: 0.5rem 0; }
.info-card h4 { color: #e8e8f0; margin: 0 0 0.5rem; font-size: 0.95rem; }
.info-card p  { color: #8888aa; margin: 0; font-size: 0.85rem; line-height: 1.6; }
[data-testid="stFileUploader"] { background: #1a1a2e; border: 2px dashed #333355; border-radius: 12px; padding: 1.5rem; }
.stButton > button { background: #e94560 !important; color: white !important; border: none !important; border-radius: 10px !important; padding: 0.65rem 2rem !important; font-size: 1rem !important; font-weight: 600 !important; width: 100% !important; }
.stButton > button:hover { background: #c73550 !important; }
[data-testid="stDownloadButton"] > button { background: #1a1a2e !important; color: #e94560 !important; border: 2px solid #e94560 !important; border-radius: 10px !important; width: 100% !important; font-size: 1rem !important; font-weight: 600 !important; padding: 0.65rem 2rem !important; }
[data-testid="stDownloadButton"] > button:hover { background: #e94560 !important; color: white !important; }
.stProgress > div > div > div { background: linear-gradient(90deg,#e94560,#ff6b8a) !important; border-radius: 4px !important; }
hr { border-color: #252540 !important; }
[data-testid="stExpander"] { background: #1a1a2e; border: 1px solid #252540; border-radius: 10px; }
</style>
""", unsafe_allow_html=True)


# ── Gemini AI でスライド解析 ────────────────────────────────
SYSTEM_PROMPT = """あなたはPowerPointスライドの構造解析専門家です。
スライド画像を詳細に分析し、以下のJSON形式のみで返してください（説明文・マークダウン不要）:

{
  "background": {"type": "color", "color": "#RRGGBB"},
  "elements": [
    {
      "type": "text",
      "x": 0〜1の相対位置,
      "y": 0〜1の相対位置,
      "w": 0〜1の相対幅,
      "h": 0〜1の相対高さ,
      "text": "テキスト内容（改行は\\nで表現）",
      "font_size": ポイント数,
      "bold": true,
      "color": "#RRGGBB",
      "bg_color": "#RRGGBB or null",
      "align": "left|center|right",
      "border_color": "#RRGGBB or null"
    },
    {
      "type": "rect",
      "x": 0〜1, "y": 0〜1, "w": 0〜1, "h": 0〜1,
      "fill_color": "#RRGGBB or null",
      "border_color": "#RRGGBB or null"
    },
    {
      "type": "image_region",
      "x": 0〜1, "y": 0〜1, "w": 0〜1, "h": 0〜1,
      "description": "画像の説明"
    }
  ]
}

重要:
- 座標は画像全体を1×1とした相対値（左上原点）
- すべてのテキストを漏れなく抽出
- 背景や装飾の矩形も含める
- JSONのみ返答"""


def analyze_with_gemini(page_img_bytes: bytes, api_key: str) -> dict:
    import google.generativeai as genai
    genai.configure(api_key=api_key)
    model = genai.GenerativeModel("gemini-1.5-flash")

    img_b64 = base64.b64encode(page_img_bytes).decode()

    response = model.generate_content([
        {"mime_type": "image/png", "data": img_b64},
        SYSTEM_PROMPT + "\nこのスライドを詳細に解析し、すべての要素をJSONで返してください。"
    ])

    raw = response.text.strip()
    raw = re.sub(r"^```[a-z]*\n?", "", raw)
    raw = re.sub(r"\n?```$", "", raw)
    return json.loads(raw.strip())


# ── PPTXスライド構築 ────────────────────────────────────────
def hex_to_rgb(h):
    h = h.lstrip("#")
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4)) if len(h) == 6 else (0,0,0)

def build_slide_from_ai(prs, analysis, page_img_bytes):
    from pptx.util import Pt
    from pptx.dml.color import RGBColor
    from pptx.enum.text import PP_ALIGN
    from lxml import etree

    slide = prs.slides.add_slide(prs.slide_layouts[6])
    sw, sh = int(prs.slide_width), int(prs.slide_height)

    def rx(v): return int(v * sw)
    def ry(v): return int(v * sh)
    def clamp(l,t,w,h):
        l=max(0,min(l,sw-1)); t=max(0,min(t,sh-1))
        return l,t,max(1,min(w,sw-l)),max(1,min(h,sh-t))

    # 背景
    bg = analysis.get("background", {})
    bg_color = bg.get("color","")
    if bg_color and bg.get("type") == "color":
        r,g,b = hex_to_rgb(bg_color)
        NP="http://schemas.openxmlformats.org/presentationml/2006/main"
        NA="http://schemas.openxmlformats.org/drawingml/2006/main"
        xml=(f'<p:bg xmlns:p="{NP}" xmlns:a="{NA}"><p:bgPr>'
             f'<a:solidFill><a:srgbClr val="{r:02X}{g:02X}{b:02X}"/></a:solidFill>'
             f'<a:effectLst/></p:bgPr></p:bg>')
        slide.shapes._spTree.getparent().insert(2, etree.fromstring(xml))
    else:
        # 背景画像として配置
        slide.shapes.add_picture(io.BytesIO(page_img_bytes), 0, 0, sw, sh)

    ALIGN_MAP = {"left":PP_ALIGN.LEFT,"center":PP_ALIGN.CENTER,"right":PP_ALIGN.RIGHT}

    for el in analysis.get("elements", []):
        et = el.get("type")
        l,t,w,h = clamp(rx(el["x"]),ry(el["y"]),rx(el["w"]),ry(el["h"]))

        if et == "rect":
            try:
                shp = slide.shapes.add_shape(1,l,t,w,h)
                fc = el.get("fill_color")
                if fc: shp.fill.solid(); shp.fill.fore_color.rgb=RGBColor(*hex_to_rgb(fc))
                else: shp.fill.background()
                bc = el.get("border_color")
                if bc: shp.line.color.rgb=RGBColor(*hex_to_rgb(bc)); shp.line.width=Pt(1)
                else: shp.line.fill.background()
            except: pass

        elif et == "text":
            txb = slide.shapes.add_textbox(l,t,w,h)
            tf = txb.text_frame
            tf.word_wrap = True
            tf.margin_left=tf.margin_right=tf.margin_top=tf.margin_bottom=Pt(2)
            bg_c = el.get("bg_color")
            if bg_c:
                txb.fill.solid()
                txb.fill.fore_color.rgb=RGBColor(*hex_to_rgb(bg_c))
            bc = el.get("border_color")
            if bc:
                txb.line.color.rgb=RGBColor(*hex_to_rgb(bc))
                txb.line.width=Pt(1.5)
            lines = el.get("text","").split("\n")
            align = ALIGN_MAP.get(el.get("align","left"), PP_ALIGN.LEFT)
            fs = max(el.get("font_size",18)*0.75, 8)
            para = tf.paragraphs[0]
            para.alignment = align
            for li, line in enumerate(lines):
                if li > 0:
                    para = tf.add_paragraph()
                    para.alignment = align
                if not line: continue
                run = para.add_run()
                run.text = line
                f = run.font
                f.size = Pt(fs)
                f.bold = el.get("bold", False)
                tc = el.get("color","#000000")
                f.color.rgb = RGBColor(*hex_to_rgb(tc))


def convert_with_gemini(pdf_bytes, api_key, slide_width_inches, progress_cb=None):
    import fitz
    from pptx import Presentation
    from pptx.util import Inches

    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(pdf_bytes); tmp_path = tmp.name

    try:
        doc = fitz.open(tmp_path)
        n = len(doc)
        first = doc[0]
        aspect = first.rect.height / first.rect.width
        doc.close()

        prs = Presentation()
        prs.slide_width  = Inches(slide_width_inches)
        prs.slide_height = int(Inches(slide_width_inches) * aspect)

        stats = {"pages": n, "texts": 0, "elements": 0}

        for i in range(n):
            if progress_cb: progress_cb(f"📑 [{i+1}/{n}] ページを画像化中...")
            doc2 = fitz.open(tmp_path)
            pix = doc2[i].get_pixmap(matrix=fitz.Matrix(2,2), alpha=False)
            img_bytes = pix.tobytes("png")
            doc2.close()

            if progress_cb: progress_cb(f"🤖 [{i+1}/{n}] Gemini AIが解析中...")
            try:
                analysis = analyze_with_gemini(img_bytes, api_key)
                n_txt = sum(1 for e in analysis.get("elements",[]) if e.get("type")=="text")
                stats["texts"] += n_txt
                stats["elements"] += len(analysis.get("elements",[]))
                if progress_cb: progress_cb(f"✏️  [{i+1}/{n}] {n_txt}個のテキストを配置中...")
                build_slide_from_ai(prs, analysis, img_bytes)
            except Exception as e:
                if progress_cb: progress_cb(f"⚠ [{i+1}/{n}] AI解析失敗→画像として配置: {e}")
                slide = prs.slides.add_slide(prs.slide_layouts[6])
                slide.shapes.add_picture(io.BytesIO(img_bytes),0,0,int(prs.slide_width),int(prs.slide_height))

        buf = io.BytesIO()
        prs.save(buf)
        return buf.getvalue(), stats
    finally:
        os.unlink(tmp_path)


# ── UI ─────────────────────────────────────────────────────

# APIキー取得（Secretsまたは入力欄）
gemini_key = st.secrets.get("GEMINI_API_KEY", "") if hasattr(st, "secrets") else ""

st.markdown("""
<div class="banner">
  <h1>📄 PDF <span class="accent">→</span> PPTX コンバーター</h1>
  <p>Gemini AIが画像化されたテキストも認識して、編集可能なPowerPointに変換します（<span class="accent">無料</span>）</p>
</div>
""", unsafe_allow_html=True)

# APIキー入力欄（Secretsに設定されていない場合のみ表示）
if not gemini_key:
    st.markdown("### 🔑 Google AI Studio APIキーを入力")
    gemini_key = st.text_input(
        "APIキー（無料）",
        type="password",
        placeholder="AIzaSy...",
        help="aistudio.google.com で無料取得できます",
    )
    if not gemini_key:
        st.info("💡 [aistudio.google.com](https://aistudio.google.com) → 「Get API key」で無料取得できます")

with st.sidebar:
    st.markdown("### ⚙️ 変換設定")
    slide_width = st.select_slider(
        "スライド幅",
        options=[10.0, 12.0, 13.333, 16.0],
        value=13.333,
        format_func=lambda x: {10.0:"10\"(4:3)",12.0:"12\"",13.333:"13.3\"(16:9)",16.0:"16\"(ワイド)"}[x],
    )
    st.markdown("---")
    st.markdown("### ✅ 変換できる要素")
    for icon, label in [("✅","テキスト（AI認識・編集可能）"),("✅","背景色"),("✅","図形・矩形"),("✅","埋め込み画像"),("⚠️","複雑なグラデーション")]:
        st.markdown(f"{icon} {label}")
    st.markdown("---")
    st.markdown("### 💰 料金")
    st.markdown("Gemini 1.5 Flash は**無料枠あり**\n\n1分あたり15リクエストまで無料")

uploaded = st.file_uploader(
    "PDFファイルをここにドロップ、またはクリックして選択",
    type=["pdf"],
    help="スライド形式のPDFを推奨（最大200MB）",
)

def format_size(n):
    if n < 1024: return f"{n} B"
    if n < 1024**2: return f"{n/1024:.1f} KB"
    return f"{n/1024**2:.1f} MB"

if uploaded:
    st.markdown("---")
    c1,c2,c3 = st.columns(3)
    with c1: st.metric("ファイル名", uploaded.name[:16]+"…" if len(uploaded.name)>16 else uploaded.name)
    with c2: st.metric("サイズ", format_size(uploaded.size))
    with c3: st.metric("AIモード", "Gemini")
    st.markdown("")

    if st.button("🚀　AI変換を開始する", use_container_width=True):
        if not gemini_key:
            st.error("❌ APIキーを入力してください")
            st.stop()

        pdf_bytes = uploaded.read()
        progress_bar = st.progress(0)
        status_text  = st.empty()
        log_area     = st.empty()
        logs = []

        def progress_cb(msg):
            logs.append(msg)
            status_text.markdown(f"**{msg.strip()}**")
            log_area.markdown("\n".join(f"- {l}" for l in logs[-5:]))

        try:
            import fitz as _fitz
            _doc = _fitz.open(stream=pdf_bytes, filetype="pdf")
            n_pages = len(_doc); _doc.close()

            progress_bar.progress(0.1)
            pptx_bytes, stats = convert_with_gemini(
                pdf_bytes, gemini_key, slide_width,
                progress_cb=progress_cb,
            )
            progress_bar.progress(1.0)
            status_text.markdown("**✅ 変換完了！**")
            log_area.empty()

            st.success(f"✨ {stats['pages']} スライドを変換しました！テキスト {stats['texts']} 個が編集可能です。")

            st.markdown(f"""
<div class="stat-grid">
  <div class="stat-box"><div class="num">{stats['pages']}</div><div class="lbl">スライド</div></div>
  <div class="stat-box"><div class="num">{stats['texts']}</div><div class="lbl">編集可能テキスト</div></div>
  <div class="stat-box"><div class="num">{stats['elements']}</div><div class="lbl">総要素数</div></div>
</div>
""", unsafe_allow_html=True)

            out_name = uploaded.name.replace(".pdf","_AI変換.pptx").replace(".PDF","_AI変換.pptx")
            st.download_button(
                label="⬇️　PPTXをダウンロード",
                data=pptx_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )

        except Exception as e:
            import traceback
            progress_bar.empty(); status_text.empty()
            st.error(f"❌ エラー: {e}")
            with st.expander("詳細"):
                st.code(traceback.format_exc())

else:
    st.markdown("")
    c1,c2 = st.columns(2)
    with c1:
        st.markdown('<div class="info-card"><h4>🎯 こんなPDFに最適</h4><p>・画像化されたスライド<br>・テキスト埋め込みの資料<br>・レイアウト再編集が必要なスライド</p></div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="info-card"><h4>⚡ AI変換の特徴</h4><p>・画像内テキストもAIが認識<br>・編集可能なテキストボックス<br>・Gemini Flash（無料枠あり）</p></div>', unsafe_allow_html=True)

st.markdown("---")
st.markdown("<p style='text-align:center;color:#555577;font-size:0.8rem;'>PDF → PPTX Converter（Gemini AI版）| Powered by Google Gemini + python-pptx</p>", unsafe_allow_html=True)
