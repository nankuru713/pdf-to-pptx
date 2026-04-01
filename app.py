"""
PDF → PPTX 変換ツール
Streamlit アプリ（クラウドデプロイ対応）
"""

import io
import time
import tempfile
import os
import streamlit as st

# ページ設定（最初に呼ぶ必要あり）
st.set_page_config(
    page_title="PDF → PPTX コンバーター",
    page_icon="📄",
    layout="centered",
    initial_sidebar_state="collapsed",
)

# ─── カスタムCSS ───────────────────────────────────────────────
st.markdown("""
<style>
/* ベース */
html, body, [data-testid="stAppViewContainer"] {
    background-color: #0f0f1a;
    color: #e8e8f0;
}
[data-testid="stHeader"] { background: transparent; }

/* メインコンテナ */
.main .block-container {
    padding-top: 2rem;
    padding-bottom: 3rem;
    max-width: 720px;
}

/* タイトルバナー */
.banner {
    background: linear-gradient(135deg, #1a1a2e 0%, #16213e 100%);
    border: 1px solid #e9456033;
    border-radius: 16px;
    padding: 2rem 2.5rem;
    margin-bottom: 2rem;
    position: relative;
    overflow: hidden;
}
.banner::before {
    content: "";
    position: absolute;
    top: 0; left: 0;
    width: 4px; height: 100%;
    background: #e94560;
    border-radius: 4px 0 0 4px;
}
.banner h1 {
    font-size: 1.8rem;
    font-weight: 700;
    color: #ffffff;
    margin: 0 0 0.4rem 0;
    letter-spacing: -0.5px;
}
.banner p {
    color: #8888aa;
    margin: 0;
    font-size: 0.95rem;
}
.banner .accent { color: #e94560; }

/* アップロードエリア */
[data-testid="stFileUploader"] {
    background: #1a1a2e;
    border: 2px dashed #333355;
    border-radius: 12px;
    padding: 1.5rem;
    transition: border-color 0.2s;
}
[data-testid="stFileUploader"]:hover {
    border-color: #e94560;
}
[data-testid="stFileUploaderDropzoneInstructions"] {
    color: #8888aa;
}

/* ボタン */
.stButton > button {
    background: #e94560 !important;
    color: white !important;
    border: none !important;
    border-radius: 10px !important;
    padding: 0.65rem 2rem !important;
    font-size: 1rem !important;
    font-weight: 600 !important;
    width: 100% !important;
    transition: all 0.2s !important;
    letter-spacing: 0.3px;
}
.stButton > button:hover {
    background: #c73550 !important;
    transform: translateY(-1px);
    box-shadow: 0 4px 20px #e9456044 !important;
}
.stButton > button:active { transform: translateY(0); }

/* ダウンロードボタン */
[data-testid="stDownloadButton"] > button {
    background: #1a1a2e !important;
    color: #e94560 !important;
    border: 2px solid #e94560 !important;
    border-radius: 10px !important;
    width: 100% !important;
    font-size: 1rem !important;
    font-weight: 600 !important;
    padding: 0.65rem 2rem !important;
    transition: all 0.2s !important;
}
[data-testid="stDownloadButton"] > button:hover {
    background: #e94560 !important;
    color: white !important;
}

/* 情報カード */
.info-card {
    background: #1a1a2e;
    border: 1px solid #252540;
    border-radius: 12px;
    padding: 1.2rem 1.5rem;
    margin: 0.5rem 0;
}
.info-card h4 { color: #e8e8f0; margin: 0 0 0.5rem 0; font-size: 0.95rem; }
.info-card p  { color: #8888aa; margin: 0; font-size: 0.85rem; line-height: 1.6; }

/* 統計グリッド */
.stat-grid {
    display: grid;
    grid-template-columns: repeat(3, 1fr);
    gap: 10px;
    margin: 1rem 0;
}
.stat-box {
    background: #1a1a2e;
    border: 1px solid #252540;
    border-radius: 10px;
    padding: 1rem;
    text-align: center;
}
.stat-box .num { font-size: 1.8rem; font-weight: 700; color: #e94560; }
.stat-box .lbl { font-size: 0.75rem; color: #8888aa; margin-top: 2px; }

/* 警告カード */
.warn-card {
    background: #2a1a10;
    border: 1px solid #e94560aa;
    border-radius: 10px;
    padding: 1rem 1.2rem;
    margin: 1rem 0;
    font-size: 0.85rem;
    color: #ffaa80;
}

/* プログレス */
.stProgress > div > div > div {
    background: linear-gradient(90deg, #e94560, #ff6b8a) !important;
    border-radius: 4px !important;
}

/* セパレーター */
hr { border-color: #252540 !important; }

/* expander */
[data-testid="stExpander"] {
    background: #1a1a2e;
    border: 1px solid #252540;
    border-radius: 10px;
}
</style>
""", unsafe_allow_html=True)


# ─── ヘルパー ─────────────────────────────────────────────────

def format_size(n_bytes: int) -> str:
    if n_bytes < 1024:       return f"{n_bytes} B"
    if n_bytes < 1024**2:    return f"{n_bytes/1024:.1f} KB"
    return f"{n_bytes/1024**2:.1f} MB"


@st.cache_data(show_spinner=False)
def run_conversion(pdf_bytes: bytes, slide_width: float) -> tuple[bytes, dict]:
    """
    PDF バイト列を受け取り、PPTX バイト列と統計情報を返す。
    Streamlit のキャッシュ対応（同じPDFは再変換しない）。
    """
    from pdf_to_pptx import analyze_pdf, build_slide, WIDE_INCHES
    from pptx import Presentation
    from pptx.util import Inches

    # 一時ファイルに書き出し（PyMuPDF がファイルパスを要求する場合の対策）
    with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as tmp:
        tmp.write(pdf_bytes)
        tmp_path = tmp.name

    try:
        pages = analyze_pdf(tmp_path)

        first  = pages[0]
        aspect = first.height_pt / first.width_pt
        slide_w = Inches(slide_width)
        slide_h = int(slide_w * aspect)

        prs = Presentation()
        prs.slide_width  = slide_w
        prs.slide_height = slide_h

        for page in pages:
            build_slide(prs, page)

        buf = io.BytesIO()
        prs.save(buf)
        pptx_bytes = buf.getvalue()

        stats = {
            "slides":  len(pages),
            "texts":   sum(len(p.texts)  for p in pages),
            "images":  sum(len(p.images) for p in pages),
            "shapes":  sum(len(p.rects)  for p in pages),
            "warnings": [
                p.page_num for p in pages
                if len(p.images) > 0 and len(p.texts) == 0
            ],
        }
        return pptx_bytes, stats

    finally:
        os.unlink(tmp_path)


# ─── UI ───────────────────────────────────────────────────────

# バナー
st.markdown("""
<div class="banner">
  <h1>📄 PDF <span class="accent">→</span> PPTX コンバーター</h1>
  <p>テキスト・図・画像・背景色を維持したまま、編集可能なPowerPointに変換します</p>
</div>
""", unsafe_allow_html=True)


# サイドバー（設定）
with st.sidebar:
    st.markdown("### ⚙️ 変換設定")
    slide_width = st.select_slider(
        "スライド幅（インチ）",
        options=[10.0, 12.0, 13.333, 16.0],
        value=13.333,
        format_func=lambda x: {
            10.0:   "10.0\" (4:3)",
            12.0:   "12.0\" (カスタム)",
            13.333: "13.3\" (16:9 標準)",
            16.0:   "16.0\" (ワイド)",
        }[x],
    )
    st.caption("16:9スライドには13.3\"を推奨")

    st.markdown("---")
    st.markdown("### 📋 変換できる要素")
    items = [
        ("✅", "テキスト（位置・色・サイズ）"),
        ("✅", "背景色"),
        ("✅", "矩形・図形"),
        ("✅", "埋め込み画像"),
        ("⚠️", "表（行単位・列分割なし）"),
        ("❌", "グラデーション・影"),
        ("❌", "画像化テキスト（OCR必要）"),
    ]
    for icon, label in items:
        st.markdown(f"{icon} {label}")

    st.markdown("---")
    st.markdown("### 🔗 デプロイ方法")
    st.markdown("""
1. このフォルダをGitHubにpush
2. [streamlit.io/cloud](https://streamlit.io/cloud) でログイン
3. リポジトリを選択して Deploy
4. URLを共有するだけ！
""")


# ファイルアップロード
uploaded = st.file_uploader(
    "PDFファイルをここにドロップ、またはクリックして選択",
    type=["pdf"],
    help="スライド形式のPDFを推奨（最大200MB）",
)

if uploaded:
    st.markdown("---")

    # ファイル情報
    col1, col2, col3 = st.columns(3)
    with col1:
        st.metric("ファイル名", uploaded.name[:20] + ("…" if len(uploaded.name) > 20 else ""))
    with col2:
        st.metric("ファイルサイズ", format_size(uploaded.size))
    with col3:
        st.metric("形式", "PDF")

    st.markdown("")

    # 変換ボタン
    if st.button("🚀　変換を開始する", use_container_width=True):

        pdf_bytes = uploaded.read()

        # プログレス表示
        progress_bar = st.progress(0)
        status_text  = st.empty()

        steps = [
            (0.1,  "📖 PDFを読み込んでいます..."),
            (0.3,  "🔍 テキスト・画像・図形を解析中..."),
            (0.55, "🎨 スライドレイアウトを再構築中..."),
            (0.75, "✏️  テキストボックスを配置中..."),
            (0.9,  "💾 PPTXファイルを生成中..."),
        ]
        for pct, msg in steps:
            progress_bar.progress(pct)
            status_text.markdown(f"**{msg}**")
            time.sleep(0.3)

        try:
            pptx_bytes, stats = run_conversion(pdf_bytes, slide_width)

            progress_bar.progress(1.0)
            status_text.markdown("**✅ 変換完了！**")
            time.sleep(0.4)
            progress_bar.empty()
            status_text.empty()

            # 完了バナー
            st.success(f"✨ 変換完了！{stats['slides']} スライドをPPTXに変換しました。")

            # 統計グリッド
            st.markdown(f"""
<div class="stat-grid">
  <div class="stat-box">
    <div class="num">{stats['slides']}</div>
    <div class="lbl">スライド</div>
  </div>
  <div class="stat-box">
    <div class="num">{stats['texts']}</div>
    <div class="lbl">テキストボックス</div>
  </div>
  <div class="stat-box">
    <div class="num">{stats['images']}</div>
    <div class="lbl">画像</div>
  </div>
</div>
""", unsafe_allow_html=True)

            # 警告
            if stats["warnings"]:
                st.markdown(f"""
<div class="warn-card">
⚠️ ページ {stats['warnings']} はテキストが画像化されている可能性があります。
そのページのテキストは編集できません（画像として保持）。
</div>
""", unsafe_allow_html=True)

            # ダウンロードボタン
            out_name = uploaded.name.replace(".pdf", ".pptx")
            st.download_button(
                label="⬇️　PPTXをダウンロード",
                data=pptx_bytes,
                file_name=out_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True,
            )

            # 詳細情報（折りたたみ）
            with st.expander("📊 変換詳細を見る"):
                st.markdown(f"- スライド数: **{stats['slides']}**")
                st.markdown(f"- テキストボックス: **{stats['texts']}**")
                st.markdown(f"- 埋め込み画像: **{stats['images']}**")
                st.markdown(f"- シェイプ（矩形）: **{stats['shapes']}**")
                st.markdown(f"- スライドサイズ: **{slide_width}\" 幅**")
                if stats["warnings"]:
                    st.warning(f"画像化テキストの疑いがあるページ: {stats['warnings']}")

        except Exception as e:
            progress_bar.empty()
            status_text.empty()
            st.error(f"❌ 変換に失敗しました: {str(e)}")
            with st.expander("エラー詳細"):
                import traceback
                st.code(traceback.format_exc())

else:
    # 使い方ガイド（アップロード前）
    st.markdown("")
    col1, col2 = st.columns(2)
    with col1:
        st.markdown("""
<div class="info-card">
<h4>🎯 こんなPDFに最適</h4>
<p>
・PowerPointから書き出されたPDF<br>
・テキスト埋め込みのビジネス資料<br>
・レイアウト再編集が必要なスライド
</p>
</div>
""", unsafe_allow_html=True)
    with col2:
        st.markdown("""
<div class="info-card">
<h4>⚡ 変換の特徴</h4>
<p>
・テキストは編集可能なテキストボックス<br>
・背景色・図形の色を忠実に再現<br>
・日本語フォントのマッピング対応
</p>
</div>
""", unsafe_allow_html=True)

# フッター
st.markdown("---")
st.markdown(
    "<p style='text-align:center; color:#555577; font-size:0.8rem;'>"
    "PDF → PPTX Converter v3.0 &nbsp;|&nbsp; "
    "Powered by PyMuPDF + python-pptx"
    "</p>",
    unsafe_allow_html=True,
)
