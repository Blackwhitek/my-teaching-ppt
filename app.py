import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
import google.generativeai as genai
import os
from dotenv import load_dotenv
from datetime import datetime
import requests
from io import BytesIO

load_dotenv()

st.set_page_config(page_title="教學投影片生成器", layout="wide")
st.title("🧑‍🏫 教學投影片生成器")
st.markdown("### AI 自動生成 + 圖片自動插入（Pexels + Canva 建議）")

api_key = os.getenv("GOOGLE_API_KEY")
pexels_key = os.getenv("PEXELS_API_KEY")

if not api_key:
    st.error("請在 Secrets 設定 GOOGLE_API_KEY")
    st.stop()

genai.configure(api_key=api_key)

with st.sidebar:
    st.header("設定")
    model_name = st.selectbox("AI 模型", ["gemini-2.5-flash"])
    num_slides = st.slider("投影片張數", 8, 20, 12)

uploaded_file = st.file_uploader("上傳參考資料", type=["pdf", "docx", "txt"])
topic = st.text_input("專題名稱", placeholder="例如：台灣如何出聲")

if st.button("🚀 生成教學投影片", type="primary", use_container_width=True):
    if not uploaded_file and not topic:
        st.error("請輸入主題或上傳資料")
    else:
        with st.spinner("正在生成投影片並自動插入圖片..."):
            # 讀取檔案...
            file_content = ""
            if uploaded_file:
                # 讀取程式碼保持不變
                pass

            prompt = f"""
主題：{topic}
內容：{file_content[:12000]}

生成 {num_slides} 張教學投影片，遵守6×6規則。
每張必須有清楚的「建議圖片」描述。
"""

            model = genai.GenerativeModel(model_name)
            response = model.generate_content(prompt)
            text = response.text

            prs = Presentation()
            parts = text.split("**第")

            inserted = 0
            for i, part in enumerate(parts):
                if len(part.strip()) < 10: continue

                slide = prs.slides.add_slide(prs.slide_layouts[1])
                title = slide.shapes.title
                title.text = f"第{i}張 " + part.split("\n")[0].replace("：", "").strip()[:55]

                # 文字內容
                try:
                    tf = slide.placeholders[1].text_frame
                    tf.clear()
                    for line in part.split("\n"):
                        if line.strip().startswith("-"):
                            p = tf.add_paragraph()
                            p.text = line.strip("- •").strip()
                            p.font.size = Pt(20)
                except:
                    pass

                # 自動插入圖片 (Pexels)
                if "建議圖片：" in part or "圖片：" in part:
                    try:
                        desc = part.split("建議圖片：")[-1].split("\n")[0].strip()[:100]
                        if pexels_key:
                            r = requests.get(
                                f"https://api.pexels.com/v1/search?query={desc}&per_page=1&orientation=landscape",
                                headers={"Authorization": pexels_key},
                                timeout=8
                            )
                            if r.status_code == 200:
                                photos = r.json().get("photos", [])
                                if photos:
                                    img_url = photos[0]["src"]["large"]
                                    img_resp = requests.get(img_url, timeout=8)
                                    if img_resp.status_code == 200:
                                        slide.shapes.add_picture(BytesIO(img_resp.content), Inches(7), Inches(1.5), Inches(5.8))
                                        inserted += 1
                    except:
                        pass

            filename = f"教學投影片_{topic[:20]}_{datetime.now().strftime('%m%d_%H%M')}.pptx"
            prs.save(filename)

            st.success(f"✅ 生成完成！成功插入 {inserted} 張圖片")
            st.info("💡 如果圖片沒出現，可在 Canva 搜尋 AI 建議的描述快速補圖")

            with open(filename, "rb") as f:
                st.download_button("📥 下載 .pptx", f, filename, use_container_width=True)

            with st.expander("AI 生成完整內容"):
                st.write(text)
