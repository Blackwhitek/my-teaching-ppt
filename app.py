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
st.title("🧑‍🏫 依參考資料自動生成教學投影片")
st.markdown("### 自動分析 + Pexels 高品質圖片自動插入")

api_key = os.getenv("GOOGLE_API_KEY")
pexels_api_key = os.getenv("PEXELS_API_KEY")  # 新增

if not api_key:
    st.error("❌ 未設定 GOOGLE_API_KEY")
    st.stop()

genai.configure(api_key=api_key)

with st.sidebar:
    st.header("生成設定")
    model_name = st.selectbox("AI 模型", ["gemini-2.5-flash", "gemini-2.0-flash"])
    num_slides = st.slider("投影片張數", 8, 25, 12)

# 上傳檔案 + Pexels Key 提示
uploaded_file = st.file_uploader("上傳參考資料（PDF / Word / TXT）", type=["pdf", "docx", "txt"])
topic = st.text_input("專題名稱", placeholder="例如：光合作用原理")

if st.button("🚀 生成含真實圖片的教學投影片", type="primary", use_container_width=True):
    if not uploaded_file and not topic:
        st.error("請上傳資料或輸入主題")
    else:
        with st.spinner("AI 分析中 + 正在尋找高品質教育圖片..."):
            
            # 讀取檔案內容（略）
            file_content = ""
            if uploaded_file:
                # ... (保持原本的讀取程式碼)
                pass

            # Prompt 優化
            prompt = f"""...（保持你原本的 prompt）..."""

            model = genai.GenerativeModel(model_name)
            response = model.generate_content(prompt)
            generated_text = response.text

            # 生成 PPTX + Pexels 圖片
            prs = Presentation()
            parts = generated_text.split("**第")

            for i, part in enumerate(parts):
                if len(part.strip()) < 15: continue

                slide = prs.slides.add_slide(prs.slide_layouts[1])
                title = slide.shapes.title
                title.text = f"第{i}張 " + part.split("\n")[0].replace("：", "").strip()[:60]

                # 插入文字內容
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

                # === Pexels 自動插入圖片 ===
                if "建議圖片：" in part or "圖片：" in part:
                    try:
                        desc = part.split("建議圖片：")[-1].split("\n")[0].strip()[:100]
                        headers = {"Authorization": pexels_api_key}
                        r = requests.get(
                            f"https://api.pexels.com/v1/search?query={desc}&per_page=1&orientation=landscape",
                            headers=headers,
                            timeout=10
                        )
                        if r.status_code == 200:
                            photos = r.json().get("photos", [])
                            if photos:
                                img_url = photos[0]["src"]["large"]
                                img_resp = requests.get(img_url, timeout=10)
                                if img_resp.status_code == 200:
                                    slide.shapes.add_picture(BytesIO(img_resp.content), 
                                                           Inches(6.8), Inches(1.2), Inches(6))
                    except:
                        pass

            filename = f"教學投影片_{topic[:20]}_{datetime.now().strftime('%m%d_%H%M')}.pptx"
            prs.save(filename)

            st.success("✅ 生成完成！已嘗試插入 Pexels 高品質圖片")
            with open(filename, "rb") as f:
                st.download_button("📥 下載 .pptx", f, filename, use_container_width=True)

            with st.expander("AI 生成內容"):
                st.write(generated_text)
