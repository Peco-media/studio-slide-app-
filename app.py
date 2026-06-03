"""
app.py
撮影スタジオ候補スライド 自動生成アプリ
"""
import streamlit as st
from PIL import Image
import io
import json
import anthropic
from slide_generator import generate_slides

# ── ページ設定 ────────────────────────────────────────────────────────
st.set_page_config(
    page_title="スタジオ候補スライド生成",
    page_icon="📷",
    layout="wide",
)

# ── スタイル ──────────────────────────────────────────────────────────
st.markdown("""
<style>
.main { background: #F8FAFC; }
h1 { color: #2C4A6E !important; }
.stButton > button {
    background: #2C4A6E;
    color: white;
    border: none;
    border-radius: 8px;
    padding: 0.6rem 1.5rem;
    font-size: 15px;
    font-weight: 500;
    width: 100%;
}
.stButton > button:hover { background: #3a5f8a; color: white; }
.info-box {
    background: #EEF6FC;
    border: 1px solid #D6E8F5;
    border-radius: 8px;
    padding: 0.75rem 1rem;
    font-size: 13px;
    color: #2C4A6E;
    margin-bottom: 1rem;
}
</style>
""", unsafe_allow_html=True)


# ── AIでスタジオ情報を抽出 ─────────────────────────────────────────────
def extract_studio_info_with_ai(images: list[Image.Image], studio_name: str) -> dict:
    """
    アップロードされた画像群からAIがスタジオ情報を読み取りJSONで返す
    """
    import base64

    client = anthropic.Anthropic(api_key=st.secrets["ANTHROPIC_API_KEY"])

    # 画像をbase64に変換（最大3枚）
    content = []
    for img in images[:3]:
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        b64 = base64.b64encode(buf.getvalue()).decode()
        content.append({
            "type": "image",
            "source": {"type": "base64", "media_type": "image/png", "data": b64}
        })

    content.append({
        "type": "text",
        "text": f"""この撮影スタジオ「{studio_name}」のスクリーンショット画像から情報を読み取り、
以下のJSON形式のみで返してください。説明文や```は不要です。

{{
  "name": "スタジオ名",
  "area": "エリア（区名など）",
  "address": "住所",
  "access": "最寄り駅・徒歩分数",
  "size": "面積",
  "ceiling": "天井高",
  "rooms": "控室情報",
  "kitchen": "キッチン有無",
  "natural_light": "自然光の有無",
  "electric": "電気容量",
  "wifi": "Wi-Fi速度",
  "min_time": "最低利用時間",
  "price_still": "スチール撮影料金",
  "price_movie": "動画撮影料金",
  "pet": "ペット可否",
  "tags": ["特徴タグ1", "特徴タグ2", ...],
  "note": "スタジオの特徴（2〜3文）",
  "caution": "確認が必要な事項",
  "url": "URL（わかれば）"
}}"""
    })

    try:
        response = client.messages.create(
            model="claude-sonnet-4-20250514",
            max_tokens=1000,
            messages=[{"role": "user", "content": content}]
        )
        text = response.content[0].text.strip()
        # JSON部分だけ抽出
        start = text.find("{")
        end   = text.rfind("}") + 1
        return json.loads(text[start:end])
    except Exception as e:
        # 読み取り失敗時はスタジオ名だけ返す
        return {
            "name": studio_name,
            "area": "—", "address": "—", "access": "—",
            "size": "—", "ceiling": "—", "rooms": "—",
            "kitchen": "—", "natural_light": "—",
            "electric": "—", "wifi": "—", "min_time": "—",
            "price_still": "—", "price_movie": "—", "pet": "—",
            "tags": [], "note": "", "caution": "", "url": ""
        }


# ── メイン UI ─────────────────────────────────────────────────────────
st.title("📷 スタジオ候補スライド 自動生成")
st.markdown(
    '<div class="info-box">スタジオのスクリーンショットをアップロードするだけで、'
    '比較スライド（.pptx）を自動生成します。</div>',
    unsafe_allow_html=True
)

# サイドバー: 基本設定
with st.sidebar:
    st.header("⚙️ 設定")
    client_name = st.text_input("クライアント名（任意）", placeholder="例：ネスレ ピュリナ ワン")
    st.markdown("---")
    st.markdown("**検索条件**")
    cond_pet     = st.checkbox("ペット可", value=True)
    cond_natural = st.checkbox("自然光あり", value=True)
    cond_kitchen = st.checkbox("キッチンあり")
    cond_parking = st.checkbox("駐車場あり")
    st.markdown("---")
    budget = st.selectbox("予算上限（スチール/h）",
                           ["〜15,000円","〜20,000円","〜30,000円","上限なし"],
                           index=2)
    st.markdown("---")
    st.markdown("**使い方**")
    st.markdown("""
1. スタジオ名を入力
2. スクショをアップロード
3. 「スライド生成」ボタンを押す
4. pptxをダウンロード
""")


# スタジオ追加UI
st.markdown("## スタジオを追加")

if "studios" not in st.session_state:
    st.session_state.studios = []

with st.expander("➕ スタジオを追加する", expanded=len(st.session_state.studios) == 0):
    col1, col2 = st.columns([2, 3])
    with col1:
        studio_name = st.text_input("スタジオ名", placeholder="例：PLEASE GREEN")
        use_ai = st.checkbox("AIで自動読み取り（推奨）", value=True)
        if not use_ai:
            st.caption("手動で情報を入力します（スライドには写真のみ使用）")
    with col2:
        uploaded = st.file_uploader(
            "スクリーンショットをアップロード（複数可）",
            type=["png","jpg","jpeg"],
            accept_multiple_files=True,
            key="uploader"
        )
        if uploaded:
            cols = st.columns(min(len(uploaded), 4))
            for i, f in enumerate(uploaded[:4]):
                cols[i].image(f, use_column_width=True)

    if st.button("このスタジオを追加"):
        if not studio_name:
            st.error("スタジオ名を入力してください")
        elif not uploaded:
            st.error("スクリーンショットを1枚以上アップロードしてください")
        else:
            with st.spinner(f"「{studio_name}」の情報を読み取り中..."):
                images = [Image.open(io.BytesIO(f.read())) for f in uploaded]
                for f in uploaded:
                    f.seek(0)

                if use_ai:
                    info = extract_studio_info_with_ai(images, studio_name)
                else:
                    info = {"name": studio_name, "tags": [], "note": "", "caution": "", "url": ""}

                info["photos"] = images
                st.session_state.studios.append(info)
                st.success(f"✅ 「{info.get('name', studio_name)}」を追加しました！")
                st.rerun()


# 追加済みスタジオ一覧
if st.session_state.studios:
    st.markdown("## 追加済みスタジオ")
    nums = "①②③④⑤⑥⑦⑧"
    for i, s in enumerate(st.session_state.studios):
        with st.expander(f"{nums[i]} {s.get('name','（未設定）')}　— {s.get('area','')}", expanded=False):
            col1, col2, col3 = st.columns([1,1,1])
            col1.markdown(f"**エリア** {s.get('area','—')}")
            col1.markdown(f"**広さ** {s.get('size','—')}")
            col1.markdown(f"**控室** {s.get('rooms','—')}")
            col2.markdown(f"**自然光** {s.get('natural_light','—')}")
            col2.markdown(f"**ペット** {s.get('pet','—')}")
            col2.markdown(f"**キッチン** {s.get('kitchen','—')}")
            col3.markdown(f"**料金（スチール）** {s.get('price_still','—')}")
            col3.markdown(f"**最低時間** {s.get('min_time','—')}")
            if s.get("photos"):
                ph_cols = st.columns(min(len(s["photos"]), 4))
                for j, img in enumerate(s["photos"][:4]):
                    ph_cols[j].image(img, use_column_width=True)
            if st.button(f"削除", key=f"del_{i}"):
                st.session_state.studios.pop(i)
                st.rerun()

    st.markdown("---")

    # スライド生成ボタン
    if len(st.session_state.studios) >= 1:
        st.markdown(f"### スライド生成（{len(st.session_state.studios)}スタジオ）")
        if st.button("🎉 スライドを生成する (.pptx)"):
            with st.spinner("スライドを生成中... しばらくお待ちください"):
                try:
                    pptx_bytes = generate_slides(
                        st.session_state.studios,
                        client_name=client_name
                    )
                    fname = (client_name.replace(" ","_") if client_name else "studio") + "_candidates.pptx"
                    st.success("✅ スライドが完成しました！")
                    st.download_button(
                        label="⬇️ pptxをダウンロード",
                        data=pptx_bytes,
                        file_name=fname,
                        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                    )
                    st.balloons()
                except Exception as e:
                    st.error(f"生成中にエラーが発生しました: {e}")

        if st.button("🗑️ リセット（全スタジオ削除）"):
            st.session_state.studios = []
            st.rerun()

else:
    st.info("↑ スタジオを追加してください")

# フッター
st.markdown("---")
st.caption("スタジオ候補スライド自動生成アプリ　|　Powered by Claude AI")
