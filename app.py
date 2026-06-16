import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
import google.generativeai as genai
import os
from dotenv import load_dotenv
from datetime import datetime

load_dotenv()

st.set_page_config(page_title="教學投影片生成器", layout="wide")
st.title("🧑‍🏫 依參考資料自動生成教學投影片")
st.markdown("### 自動分析 + 圖片建議")

api_key = os.getenv("GOOGLE_API_KEY")
if not api_key:
    st.error("❌ 未設定 GOOGLE_API_KEY")
    st.stop()

genai.configure(api_key=api_key)

with st.sidebar:
    st.header("生成設定")
    model_name = st.selectbox("AI 模型", ["gemini-2.5-flash", "gemini-2.0-flash"])
    num_slides = st.slider("投影片張數", 8, 20, 12)

uploaded_file = st.file_uploader("上傳參考資料（PDF / Word / TXT）", type=["pdf", "docx", "txt"])
topic = st.text_input("專題名稱", placeholder="例如：台灣如何出聲")

if st.button("🚀 生成教學投影片", type="primary", use_container_width=True):
    if not uploaded_file and not topic:
        st.error("請上傳資料或輸入主題")
    else:
        with st.spinner("AI 正在嚴格按照格式生成投影片..."):
            
            file_content = ""
            if uploaded_file:
                try:
                    if uploaded_file.type == "application/pdf":
                        import PyPDF2
                        pdf = PyPDF2.PdfReader(uploaded_file)
                        file_content = "\n".join([page.extract_text() for page in pdf.pages if page.extract_text()])
                    elif uploaded_file.type.startswith("application/vnd.openxmlformats"):
                        from docx import Document
                        doc = Document(uploaded_file)
                        file_content = "\n".join([para.text for para in doc.paragraphs])
                except:
                    pass

            # 強制格式 Prompt
            prompt = f"""
你是一位嚴格遵守格式的教學簡報專家。

主題：{topic}
參考資料：{file_content[:12000]}

請生成 {num_slides} 張教學投影片，**必須嚴格遵守以下格式**，不要加入任何額外解釋：

**第1張：** [簡潔標題]
**內容：**
- 要點1（最多6個字）
- 要點2
**講者筆記：** [詳細說明]
**建議圖片：** [具體描述]

**第2張：** [簡潔標題]
**內容：**
- 要點1
**講者筆記：** ...
**建議圖片：** ...

請從第1張開始，一直到第{num_slides}張為止，嚴格按照上面格式輸出，不要有其他文字。
"""

            model = genai.GenerativeModel(model_name, generation_config={"temperature": 0.7})
            response = model.generate_content(prompt)
            generated_text = response.text

            # 生成 PPTX
            prs = Presentation()
            parts = generated_text.split("**第")

            for i, part in enumerate(parts):
                if len(part.strip()) < 10:
                    continue

                slide = prs.slides.add_slide(prs.slide_layouts[1])
                title_shape = slide.shapes.title
                title_shape.text = f"第{i}張 " + part.split("\n")[0].replace("：", "").strip()[:50]

                try:
                    tf = slide.placeholders[1].text_frame
                    tf.clear()
                    for line in part.split("\n"):
                        cleaned = line.strip()
                        if cleaned.startswith("-") or cleaned.startswith("•"):
                            p = tf.add_paragraph()
                            p.text = cleaned.strip("- •").strip()
                            p.font.size = Pt(20)
                except:
                    pass

            filename = f"教學投影片_{topic[:15]}_{datetime.now().strftime('%m%d_%H%M')}.pptx"
            prs.save(filename)

            st.success(f"✅ 已生成 {len(prs.slides)} 張投影片！")
            
            with open(filename, "rb") as f:
                st.download_button("📥 下載 .pptx 檔案", f, filename, use_container_width=True)

            with st.expander("查看 AI 原始輸出"):
                st.write(generated_text)
