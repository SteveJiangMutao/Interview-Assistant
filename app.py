import streamlit as st
import google.generativeai as genai
import tempfile
import os
import time
import json
from docx import Document
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn # 用于设置中文字体
import io
import datetime
from google.api_core import retry

# --- 页面配置 / Page Config ---
st.set_page_config(
    page_title="Intelligent Interview System",
    layout="wide",
    page_icon="🧬",
    initial_sidebar_state="expanded"
)

# --- CSS 样式 / CSS Styling ---
st.markdown("""
<style>
    .main-header { font-size: 2.0rem; color: #2c3e50; font-weight: bold; margin-bottom: 10px; }
    .sub-header { font-size: 1.0rem; color: #7f8c8d; margin-bottom: 20px; }
    .developer-credit { font-size: 0.85rem; color: #95a5a6; margin-top: 50px; border-top: 1px solid #bdc3c7; padding-top: 10px; }
    div[data-testid="stFileUploader"] { margin-top: 20px; }
</style>
""", unsafe_allow_html=True)

# --- Session State ---
if 'analysis_result' not in st.session_state:
    st.session_state['analysis_result'] = None

# --- Word 格式化辅助函数 / Word Formatting Helper ---
def set_font_style(run, font_size=10.5, bold=False):
    """
    强制设置中西文混排字体：
    English: Times New Roman
    Chinese: Microsoft YaHei (微软雅黑)
    Color: Black
    """
    run.font.name = 'Times New Roman'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    run.font.size = Pt(font_size)
    run.font.color.rgb = RGBColor(0, 0, 0) # 纯黑
    run.bold = bold

def add_styled_paragraph(doc, text, style='Normal', bold=False, size=10.5, line_spacing=1.5):
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = line_spacing # 行间距
    p.paragraph_format.space_after = Pt(6) # 段后距
    
    # 如果是标题，左对齐；如果是正文，两端对齐(可选，这里保持默认左对齐)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT 
    
    run = p.add_run(str(text))
    set_font_style(run, font_size=size, bold=bold)
    return p

# --- Word 生成逻辑 / Word Generation Logic ---
def generate_word_report(data, company, product, date, mode, logo_file=None):
    doc = Document()
    
    # 0. 页眉 Logo (Header Logo)
    if logo_file is not None:
        section = doc.sections[0]
        header = section.header
        p_header = header.paragraphs[0]
        p_header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        # 调整图片大小，例如宽度 1.5 英寸
        run_header = p_header.add_run()
        run_header.add_picture(logo_file, width=Inches(1.5))

    # 1. 标题 (Title)
    # 朴素大号字体，左对齐
    title_text = f"{company} - {product} Interview Record"
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    run_title = p_title.add_run(title_text)
    set_font_style(run_title, font_size=18, bold=True)
    p_title.paragraph_format.space_after = Pt(24)
    
    # 2. 基础信息 (Meta Info)
    info_text = f"Date: {date} | Type: {'Commercial/Industry' if mode == 'commercial' else 'Clinical/Expert'}"
    add_styled_paragraph(doc, info_text, size=10, bold=False)
    
    doc.add_paragraph("-" * 80)

    # 3. 执行摘要 (Executive Summary)
    add_styled_paragraph(doc, '1. Executive Summary / 执行摘要', size=14, bold=True)
    summary = data.get('executive_summary', 'No content generated.')
    add_styled_paragraph(doc, summary)

    # 4. 结构化维度分析 (Structured Analysis)
    add_styled_paragraph(doc, '2. Detailed Analysis / 详细维度分析', size=14, bold=True)
    
    # 映射表 (保留英文 Key 以匹配 JSON，Value 用于文档标题)
    # 这里的 Value 可以根据 AI 输出的语言动态调整，但为了保险，我们直接用 AI 输出的 Key 
    # 或者我们假设 AI 会根据语种输出对应的 Key，这里我们做通用处理
    
    structured = data.get('structured_analysis', {})
    
    if structured:
        for key, points in structured.items():
            # 标题处理：去掉下划线，首字母大写
            clean_title = key.replace("_", " ").title()
            add_styled_paragraph(doc, clean_title, size=12, bold=True)
            
            if isinstance(points, list):
                for point in points:
                    # 使用特殊符号作为 Bullet
                    p = add_styled_paragraph(doc, f"• {point}")
                    p.paragraph_format.left_indent = Inches(0.25)
            else:
                p = add_styled_paragraph(doc, str(points))

    # 5. 其他维度 (Other Dimensions)
    other_dims = data.get('other_dimensions', {})
    if other_dims:
        add_styled_paragraph(doc, '3. Other Findings / 其他发现', size=14, bold=True)
        for k, v in other_dims.items():
            add_styled_paragraph(doc, str(k), size=12, bold=True)
            if isinstance(v, list):
                for point in v:
                    p = add_styled_paragraph(doc, f"• {point}")
                    p.paragraph_format.left_indent = Inches(0.25)
            else:
                add_styled_paragraph(doc, str(v))

    # 6. Q&A 实录 (Q&A Log)
    add_styled_paragraph(doc, '4. Q&A Transcript / 访谈实录', size=14, bold=True)
    qa_log = data.get('qa_log', [])
    
    if isinstance(qa_log, list):
        for qa in qa_log:
            if isinstance(qa, dict):
                q_text = qa.get('question', 'N/A')
                a_text = qa.get('answer', 'N/A')
                note = qa.get('context_note', None)

                # Q - 加粗
                add_styled_paragraph(doc, f"Q: {q_text}", bold=True)
                
                # A - 正常
                add_styled_paragraph(doc, f"A: {a_text}")
                
                # Note - 斜体 (用辅助函数模拟)
                if note:
                    p_note = doc.add_paragraph()
                    run_note = p_note.add_run(f"[Note: {note}]")
                    set_font_style(run_note, font_size=9)
                    run_note.italic = True
                    p_note.paragraph_format.left_indent = Inches(0.5)
                
                # 增加一点间距
                doc.add_paragraph().paragraph_format.space_after = Pt(2)

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- 核心逻辑类 / Core Logic ---
class InterviewAnalyzer:
    def __init__(self, api_key):
        self.api_key = api_key
        try:
            genai.configure(api_key=self.api_key)
            # 使用 1.5 Flash 保证长文本处理能力
            self.model = genai.GenerativeModel('gemini-3-flash-preview') 
        except Exception as e:
            st.error(f"API Error: {e}")

    def process_audio(self, audio_file_path):
        try:
            myfile = genai.upload_file(audio_file_path)
            with st.spinner("🎧 Uploading & Processing Audio... / 正在上传并解析音频..."):
                while myfile.state.name == "PROCESSING":
                    time.sleep(2)
                    myfile = genai.get_file(myfile.name)
            if myfile.state.name == "FAILED":
                st.error("Audio processing failed.")
                return None
            return myfile
        except Exception as e:
            st.error(f"Upload Error: {e}")
            return None

    def analyze_interview(self, audio_resource, mode):
        # 1. 框架定义 (Framework)
        if mode == "commercial":
            framework_desc = """
            1. **Market Size & Scale (CRITICAL)**: 
               - Extract ALL numbers related to market size, volume, and revenue.
               - **LOGIC FORMULA REQUIREMENT**: Wherever possible, provide the logic used to calculate the size (e.g., "Market Size = 50k procedures * $200 ASP = $10M").
               - TAM/SAM/SOM breakdown.
            2. **Competition Landscape**: Market shares, competitor strengths/weaknesses.
            3. **Sales & Marketing**: Pricing models, sales force structure.
            4. **Channel & Access**: Distribution model, hospital listing (入院) status.
            5. **Industry Trends**: VBP (集采), DRG/DIP impact.
            """
        else: # clinical
            framework_desc = """
            1. **Clinical Value**: Efficacy, safety, comparison with Gold Standard.
            2. **Adoption & Usage**: Procedures per month, indication expansion.
            3. **Competitive Comparison**: Head-to-head comparison in clinical practice.
            4. **Unmet Needs**: Pain points in current surgery/therapy.
            5. **Future Outlook**: Expectations for next-gen technology.
            """

        # 2. Prompt 深度优化
        system_prompt = f"""
        You are a **Senior Medical Device Industry Expert** working for Clearstate.
        Your task is to extract a comprehensive interview record from the audio.

        ### 🌍 LANGUAGE INSTRUCTION:
        - **Auto-Detect**: If the interview is in Chinese, output the report in **Simplified Chinese**.
        - **Auto-Detect**: If the interview is in English, output the report in **English**.

        ### 🧠 CONTEXTUAL CORRECTION (Medical Device Domain):
        - You must intelligently correct ASR errors based on medical context.
        - Examples: 
          - "亚培" -> "雅培 (Abbott)"
          - "强生" -> "强生 (J&J)"
          - "美敦力" -> "美敦力 (Medtronic)"
          - "吻合器" (Stapler), "超声刀" (Ultrasonic Scalpel), etc.

        ### 📊 DATA PRECISION RULES:
        - **Numbers are Sacred**: Do not miss any digits.
        - **Logic Formulas**: For any market sizing data, explicitly state the calculation logic if mentioned (e.g., "Volume x Price").

        ### FRAMEWORK:
        {framework_desc}

        ### OUTPUT JSON FORMAT:
        {{
            "executive_summary": "Summary of key insights.",
            "structured_analysis": {{
                "Dimension_Name": ["Point 1", "Point 2 (Logic: A * B = C)"]
            }},
            "other_dimensions": {{
                "Topic": ["Detail"]
            }},
            "qa_log": [
                {{
                    "question": "Question text",
                    "answer": "Answer text",
                    "context_note": "Correction note or context"
                }}
            ]
        }}
        """
        
        # 安全设置全放开
        safety_settings = [
            {"category": "HARM_CATEGORY_HARASSMENT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_HATE_SPEECH", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_SEXUALLY_EXPLICIT", "threshold": "BLOCK_NONE"},
            {"category": "HARM_CATEGORY_DANGEROUS_CONTENT", "threshold": "BLOCK_NONE"},
        ]
        
        try:
            response = self.model.generate_content(
                [audio_resource, system_prompt],
                safety_settings=safety_settings,
                request_options={"timeout": 600}
            )
            
            try:
                text = response.text
                if "```json" in text:
                    text = text.replace("```json", "").replace("```", "")
                return json.loads(text.strip())
            except ValueError:
                st.error("Error: Model output was not valid JSON.")
                return None

        except Exception as e:
            st.error(f"Analysis Interrupted: {e}")
            return None

# --- UI 主程序 / Main UI ---
with st.sidebar:
    st.title("Clearstate AI")
    st.caption("Intelligent Qualitative Interview System")
    
    # 开发者署名
    st.markdown("""
    <div class='developer-credit'>
    Developed by <b>Steve Jiang</b><br>
    Clearstate Consulting
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    api_key = st.text_input("Gemini API Key", type="password")
    
    st.markdown("### 📝 Project Info / 项目信息")
    company_name = st.text_input("Company / 公司名称", placeholder="e.g. Medtronic / 美敦力")
    product_name = st.text_input("Product / 产品领域", placeholder="e.g. Stapler / 吻合器")
    interview_date = st.date_input("Date / 访谈日期", datetime.date.today())
    
    # Logo 上传
    st.markdown("### 🖼️ Report Logo / 报告Logo")
    uploaded_logo = st.file_uploader("Upload Logo (Optional)", type=['png', 'jpg', 'jpeg'])
    
    st.markdown("### 🛠️ Mode / 模式")
    interview_mode = st.radio(
        "Select Type / 选择类型",
        ("commercial", "clinical"),
        format_func=lambda x: "🏭 Commercial (商业/厂商)" if x == "commercial" else "👨‍⚕️ Clinical (临床/专家)"
    )
    
    if st.button("🗑️ Reset / 重置"):
        st.session_state['analysis_result'] = None
        st.rerun()

# 主标题
st.markdown('<div class="main-header">智能定性访谈报告生成系统</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Intelligent Qualitative Interview Report Generation System</div>', unsafe_allow_html=True)

# --- 上传区域 ---
uploaded_file = st.file_uploader("📂 Upload Audio / 上传录音 (MP3/M4A Recommended)", type=['mp3', 'wav', 'm4a'])

if uploaded_file and st.session_state['analysis_result'] is None:
    if not api_key:
        st.error("Please enter API Key in the sidebar. / 请在侧边栏输入 API Key。")
    elif not company_name or not product_name:
        st.warning("Please fill in Company & Product info. / 请填写公司和产品信息。")
    else:
        st.audio(uploaded_file, format='audio/mp3')
        if st.button("🚀 Start Analysis / 开始分析", type="primary"):
            analyzer = InterviewAnalyzer(api_key)
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=f".{uploaded_file.name.split('.')[-1]}") as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_file_path = tmp_file.name

            with st.status("🤖 AI is processing... / AI 正在处理...", expanded=True) as status:
                st.write("🎧 Uploading audio to Gemini... / 正在上传音频...")
                audio_resource = analyzer.process_audio(tmp_file_path)
                
                if audio_resource:
                    st.write("🧠 Analyzing (Context: Medical Device)... / 正在分析 (医疗器械语境)...")
                    result = analyzer.analyze_interview(audio_resource, interview_mode)
                    
                    if result:
                        st.session_state['analysis_result'] = result
                        status.update(label="✅ Done! / 完成！", state="complete", expanded=False)
                        os.remove(tmp_file_path)
                        st.rerun()

# --- 结果展示与导出 ---
if st.session_state['analysis_result']:
    res = st.session_state['analysis_result']
    
    st.success("✅ Analysis Complete. Please download the report. / 分析完成，请下载报告。")
    
    file_date_str = interview_date.strftime("%Y%m%d")
    file_name = f"Interview_Record_{company_name}_{product_name}_{file_date_str}.docx"
    
    # 传入 Logo 文件对象
    docx_file = generate_word_report(res, company_name, product_name, interview_date, interview_mode, uploaded_logo)
    
    st.download_button(
        label=f"📥 Download Word Report / 下载 Word 报告",
        data=docx_file,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary"
    )

    st.markdown("---")
    st.markdown("### 📊 Preview / 预览")
    st.write(res.get('executive_summary'))

