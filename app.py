import streamlit as st
import google.generativeai as genai
import tempfile
import os
import time
import json
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm  # 引入 Cm 用于精确控制 Logo 高度
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml.ns import qn
import io
import datetime
from google.api_core import retry

# --- 页面配置 / Page Config ---
st.set_page_config(
    page_title="Clearstate Interview System",
    layout="wide",
    page_icon="🧬",
    initial_sidebar_state="expanded"
)

# --- CSS 样式 / CSS Styling ---
st.markdown("""
<style>
    .main-header { font-size: 2.0rem; color: #2c3e50; font-weight: bold; margin-bottom: 5px; }
    .sub-header { font-size: 1.0rem; color: #7f8c8d; margin-bottom: 20px; }
    .developer-credit { font-size: 0.85rem; color: #95a5a6; margin-top: 50px; border-top: 1px solid #bdc3c7; padding-top: 10px; }
    div[data-testid="stFileUploader"] { margin-top: 20px; }
</style>
""", unsafe_allow_html=True)

# --- Session State ---
if 'analysis_result' not in st.session_state:
    st.session_state['analysis_result'] = None

# --- Word 格式化辅助函数 (升级版) ---
def set_font_style(run, font_size=11, bold=False):
    """
    字体设置：
    - English: Times New Roman
    - Chinese: Microsoft YaHei
    - Color: Black (RGB 0,0,0)
    """
    run.font.name = 'Times New Roman'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    run.font.size = Pt(font_size)
    run.font.color.rgb = RGBColor(0, 0, 0)
    run.bold = bold

def add_styled_paragraph(doc, text, bold=False, size=11, level=None):
    """
    段落设置：
    - Line Spacing: 1.0 (Single)
    - Space Before/After: 3 Pt
    """
    p = doc.add_paragraph()
    
    # 间距设置
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT 
    
    run = p.add_run(str(text))
    set_font_style(run, font_size=size, bold=bold)
    return p

# --- Word 生成逻辑 (重构版) ---
def generate_word_report(data, company, product, date, mode, logo_file=None):
    doc = Document()
    
    # 0. 页眉 Logo (Header Logo) - 修正为 1cm 高度
    if logo_file is not None:
        section = doc.sections[0]
        header = section.header
        p_header = header.paragraphs[0]
        p_header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
        run_header = p_header.add_run()
        # 核心修改：高度固定为 1cm，宽度自适应
        run_header.add_picture(logo_file, height=Cm(1.0))

    # 1. 标题 (Title) - 朴素左对齐
    title_text = f"{company} - {product} Interview Record"
    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_title.paragraph_format.space_after = Pt(12)
    run_title = p_title.add_run(title_text)
    set_font_style(run_title, font_size=16, bold=True) # 稍微加大一点总标题
    
    # 2. 基础信息 (Meta Info)
    info_text = f"Date: {date} | Type: {'Commercial/Industry' if mode == 'commercial' else 'Clinical/Expert'}"
    add_styled_paragraph(doc, info_text, size=10.5, bold=False)
    
    doc.add_paragraph("-" * 80)

    # 3. 执行摘要 (Executive Summary) - 一级标题 14 Bold
    add_styled_paragraph(doc, '1. Executive Summary / 执行摘要', size=14, bold=True)
    summary = data.get('executive_summary', 'No content generated.')
    add_styled_paragraph(doc, summary, size=11)

    # 4. 结构化维度分析 (Structured Analysis)
    add_styled_paragraph(doc, '2. Detailed Analysis / 详细维度分析', size=14, bold=True)
    
    structured = data.get('structured_analysis', {})
    
    if structured:
        for key, points in structured.items():
            # 二级标题 12 Bold
            clean_title = key.replace("_", " ").title()
            add_styled_paragraph(doc, clean_title, size=12, bold=True)
            
            if isinstance(points, list):
                for point in points:
                    # 正文 11 Normal
                    p = add_styled_paragraph(doc, f"• {point}", size=11)
                    p.paragraph_format.left_indent = Inches(0.25)
            else:
                add_styled_paragraph(doc, str(points), size=11)

    # 5. 其他维度 (Other Findings) - 仅当 AI 无法整合时才显示
    other_dims = data.get('other_dimensions', {})
    if other_dims:
        add_styled_paragraph(doc, '3. Other Findings / 其他发现', size=14, bold=True)
        for k, v in other_dims.items():
            add_styled_paragraph(doc, str(k), size=12, bold=True)
            if isinstance(v, list):
                for point in v:
                    p = add_styled_paragraph(doc, f"• {point}", size=11)
                    p.paragraph_format.left_indent = Inches(0.25)
            else:
                add_styled_paragraph(doc, str(v), size=11)

    # Q&A 部分已移除

    bio = io.BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --- 核心逻辑类 ---
class InterviewAnalyzer:
    def __init__(self, api_key):
        self.api_key = api_key
        try:
            genai.configure(api_key=self.api_key)
            self.model = genai.GenerativeModel('gemini-3-pro-preview') 
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
        # 1. 框架定义 (Framework) - 强调数据和逻辑
        if mode == "commercial":
            framework_desc = """
            1. **Market Size & Scale (DATA CRITICAL)**: 
               - Extract ALL numerical data about market size, volume, revenue, and growth rates.
               - **LOGIC FORMULA**: You MUST provide the calculation logic if mentioned (e.g., "Total = 50 hospitals * 200 cases/hospital").
            2. **Competition Landscape**: Market shares (%), competitor strengths/weaknesses, sales team sizes.
            3. **Sales & Marketing**: Pricing (ASP), channel margins, promotion strategies.
            4. **Channel & Access**: Distribution structure, admission (入院) barriers.
            5. **Industry Trends**: VBP impact, policy changes.
            """
        else: # clinical
            framework_desc = """
            1. **Clinical Value & Efficacy**: Specific clinical outcomes, comparison with Gold Standard.
            2. **Adoption & Usage**: Monthly procedure volumes, patient selection criteria.
            3. **Competitive Comparison**: Brand A vs Brand B in clinical practice (pros/cons).
            4. **Unmet Needs & Pain Points**: Detailed description of current limitations.
            5. **Future Expectations**: Specific features desired in next-gen products.
            """

        # 2. Prompt 深度优化 - 强调整合和准确性
        system_prompt = f"""
        You are a **Senior Medical Device Consultant** at Clearstate.
        Task: Create a rigorous, data-driven interview report.

        ### 🚨 CRITICAL INSTRUCTIONS:
        1.  **DATA PRECISION**: Capture EVERY number exactly as spoken. Do not round up or summarize vaguely. If the expert says "12.5%", write "12.5%", not "about 12%".
        2.  **LOGIC & INSIGHTS**: Do not just list facts. Explain the **"Why"** and **"How"**. If a competitor is growing, explain the specific reason given (e.g., "aggressive pricing," "better sales coverage").
        3.  **INTEGRATION**: Try to fit ALL information into the main "Structured Analysis" framework. Only use "Other Dimensions" for topics that absolutely do not fit the main categories.
        4.  **NO Q&A**: Do not output a Q&A transcript. Focus on the analysis.
        5.  **CONTEXT CORRECTION**: Correct ASR errors (e.g., "亚培" -> "雅培 Abbott", "强生" -> "强生 J&J").

        ### LANGUAGE:
        - Output in the **same language** as the interview audio (Chinese or English).

        ### FRAMEWORK:
        {framework_desc}

        ### OUTPUT JSON:
        {{
            "executive_summary": "High-level summary of the key takeaways (300 words).",
            "structured_analysis": {{
                "Dimension_Name": [
                    "Point 1: Detailed insight with numbers.", 
                    "Point 2: Logic formula (A * B = C)."
                ]
            }},
            "other_dimensions": {{
                "Topic": ["Detail"]
            }}
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
    if uploaded_logo:
        st.caption("Logo will be resized to 1cm height in Word.")
    
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

