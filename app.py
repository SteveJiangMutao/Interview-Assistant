import streamlit as st
import google.generativeai as genai
import tempfile
import os
import time
import json
from docx import Document
from docx.shared import Pt, RGBColor, Inches, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_ALIGNMENT
from docx.oxml.ns import qn
import io
import datetime
from google.api_core import retry

# --- 🔧 配置项：Logo 文件 ---
# 必须在 GitHub 仓库根目录上传名为 logo.jpg 的文件
LOGO_PATH = "logo.jpg" 

# --- 页面配置 ---
st.set_page_config(
    page_title="Clearstate Interview System",
    layout="wide",
    page_icon="🧬",
    initial_sidebar_state="expanded"
)

# --- CSS 样式 ---
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

# --- 🧹 文本清洗函数 ---
def clean_text(text):
    """去除 Markdown 符号，保持文本纯净"""
    if isinstance(text, str):
        text = text.replace("**", "").replace("__", "")
        text = text.replace("##", "").replace("###", "")
        return text.strip()
    return text

# --- Word 格式化辅助函数 ---
def set_font_style(run, font_size=11, bold=False):
    run.font.name = 'Times New Roman'
    run.element.rPr.rFonts.set(qn('w:eastAsia'), '微软雅黑')
    run.font.size = Pt(font_size)
    run.font.color.rgb = RGBColor(0, 0, 0)
    run.bold = bold

def add_styled_paragraph(doc, text, bold=False, size=11, is_bullet=False):
    clean_content = clean_text(str(text))
    p = doc.add_paragraph()
    p.paragraph_format.line_spacing = 1.0
    p.paragraph_format.space_before = Pt(3)
    p.paragraph_format.space_after = Pt(3)
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT 
    
    # --- 悬挂缩进逻辑 (Strict Hanging Indent) ---
    if is_bullet:
        # 1. 设置缩进距离 (0.25 英寸)
        indent_size = Inches(0.25)
        
        # 2. 整体左缩进 (控制第二行及以后的起始位置)
        p.paragraph_format.left_indent = indent_size
        
        # 3. 首行负缩进 (让圆点回到左边)
        p.paragraph_format.first_line_indent = -indent_size
        
        # 4. 添加制表位 (Tab Stop)
        # 这一步至关重要：它强制第一行在圆点后的文字也跳到 indent_size 的位置开始
        p.paragraph_format.tab_stops.add_tab_stop(indent_size, WD_TAB_ALIGNMENT.LEFT)
        
        # 5. 组合文本：圆点 + Tab + 内容
        final_text = f"•\t{clean_content}"
        run = p.add_run(final_text)
    else:
        run = p.add_run(clean_content)
    
    set_font_style(run, font_size=size, bold=bold)
    return p

# --- 🌍 标题映射字典 ---
SECTION_HEADERS = {
    "commercial": {
        "zh": {
            "market_size": "1. 市场规模与体量 (Market Size)",
            "competition": "2. 竞争格局 (Competition)",
            "sales_marketing": "3. 销售与营销策略 (Sales & Marketing)",
            "channel_access": "4. 渠道与准入 (Channel & Access)",
            "trends": "5. 行业趋势 (Industry Trends)"
        },
        "en": {
            "market_size": "1. Market Size & Scale",
            "competition": "2. Competition Landscape",
            "sales_marketing": "3. Sales & Marketing Strategy",
            "channel_access": "4. Channel & Access Strategy",
            "trends": "5. Industry Trends"
        }
    },
    "clinical": {
        "zh": {
            "clinical_value": "1. 临床价值与疗效 (Clinical Value)",
            "adoption": "2. 临床应用与术式 (Adoption & Usage)",
            "competition": "3. 竞品对比 (Competitive Comparison)",
            "pain_points": "4. 未满足需求与痛点 (Unmet Needs)",
            "expectations": "5. 未来预期 (Future Expectations)"
        },
        "en": {
            "clinical_value": "1. Clinical Value & Efficacy",
            "adoption": "2. Adoption & Usage",
            "competition": "3. Competitive Comparison",
            "pain_points": "4. Unmet Needs & Pain Points",
            "expectations": "5. Future Expectations"
        }
    }
}

# --- Word 生成逻辑 ---
def generate_word_report(data, company, product, date, mode):
    doc = Document()
    
    # 0. Logo (右上角, 高度 0.65cm)
    section = doc.sections[0]
    header = section.header
    p_header = header.paragraphs[0]
    p_header.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    if os.path.exists(LOGO_PATH):
        try:
            run_header = p_header.add_run()
            # 🚨 修改：高度调整为 0.65cm
            run_header.add_picture(LOGO_PATH, height=Cm(0.65))
        except Exception as e:
            print(f"Logo Error: {e}")
    
    # 语言判断
    lang = data.get('language', 'en')
    if 'zh' in lang.lower() or 'chinese' in lang.lower() or 'cn' in lang.lower():
        lang_code = 'zh'
    else:
        lang_code = 'en'

    # 1. 标题与基础信息
    if lang_code == 'zh':
        title_text = f"{company} - {product} 访谈记录"
        # 🚨 修改：Commercial 对应 商业/厂商 (Trade)
        type_text = '商业/厂商' if mode == 'commercial' else '临床/专家'
        date_prefix = "访谈日期"
        type_prefix = "访谈类型"
        exec_title = "1. 执行摘要"
        other_title = "3. 其他发现"
    else:
        title_text = f"{company} - {product} Interview Record"
        # 🚨 修改：Commercial -> Trade
        type_text = 'Trade' if mode == 'commercial' else 'Clinical/Expert'
        date_prefix = "Date"
        type_prefix = "Type"
        exec_title = "1. Executive Summary"
        other_title = "3. Other Findings"

    p_title = doc.add_paragraph()
    p_title.alignment = WD_ALIGN_PARAGRAPH.LEFT
    p_title.paragraph_format.space_after = Pt(12)
    run_title = p_title.add_run(title_text)
    set_font_style(run_title, font_size=16, bold=True)
    
    # Meta Info
    info_text = f"{date_prefix}: {date} | {type_prefix}: {type_text}"
    add_styled_paragraph(doc, info_text, size=10.5, bold=False)
    doc.add_paragraph("-" * 80)

    # 2. Executive Summary
    add_styled_paragraph(doc, exec_title, size=14, bold=True)
    summary = data.get('executive_summary', '')
    add_styled_paragraph(doc, summary, size=11)

    # 3. Structured Analysis
    header_map = SECTION_HEADERS.get(mode, {}).get(lang_code, {})
    structured = data.get('structured_analysis', {})
    
    if structured:
        section_2_title = "2. 详细维度分析" if lang_code == 'zh' else "2. Detailed Analysis"
        add_styled_paragraph(doc, section_2_title, size=14, bold=True)

        # 强制顺序
        key_order = []
        if mode == 'commercial':
            key_order = ['market_size', 'competition', 'sales_marketing', 'channel_access', 'trends']
        else:
            key_order = ['clinical_value', 'adoption', 'competition', 'pain_points', 'expectations']

        for key in key_order:
            if key in structured:
                points = structured[key]
                display_title = header_map.get(key, key.title())
                add_styled_paragraph(doc, display_title, size=12, bold=True)
                
                if isinstance(points, list):
                    for point in points:
                        # 启用严格对齐的 Bullet 模式
                        add_styled_paragraph(doc, point, size=11, is_bullet=True)
                else:
                    add_styled_paragraph(doc, str(points), size=11)

    # 4. Other Findings
    other_dims = data.get('other_dimensions', {})
    if other_dims:
        add_styled_paragraph(doc, other_title, size=14, bold=True)
        for k, v in other_dims.items():
            clean_k = clean_text(k)
            add_styled_paragraph(doc, clean_k, size=12, bold=True)
            if isinstance(v, list):
                for point in v:
                    add_styled_paragraph(doc, point, size=11, is_bullet=True)
            else:
                add_styled_paragraph(doc, str(v), size=11)

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
            # 强制使用 gemini-3-pro-preview
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
        # 框架定义
        if mode == "commercial":
            keys_instruction = """
            Use these EXACT keys for `structured_analysis`:
            - `market_size` (for Market Size & Scale)
            - `competition` (for Competition Landscape)
            - `sales_marketing` (for Sales & Marketing)
            - `channel_access` (for Channel & Access)
            - `trends` (for Industry Trends)
            """
            framework_desc = """
            1. Market Size & Scale: Numbers, volume, revenue. (LOGIC FORMULA REQUIRED).
            2. Competition Landscape: Shares, strengths, weaknesses.
            3. Sales & Marketing: Pricing, promotion.
            4. Channel & Access: Distributors, admission.
            5. Industry Trends: Policy, macro environment.
            """
        else: # clinical
            keys_instruction = """
            Use these EXACT keys for `structured_analysis`:
            - `clinical_value` (for Clinical Value)
            - `adoption` (for Adoption & Usage)
            - `competition` (for Competitive Comparison)
            - `pain_points` (for Unmet Needs)
            - `expectations` (for Future Expectations)
            """
            framework_desc = """
            1. Clinical Value: Efficacy, safety.
            2. Adoption & Usage: Procedure volume, indications.
            3. Competitive Comparison: Brand vs Brand.
            4. Unmet Needs: Pain points.
            5. Future Expectations: Next-gen features.
            """

        system_prompt = f"""
        You are a **Senior Medical Device Consultant** at Clearstate.
        Task: Create a rigorous, data-driven interview report.

        ### 🚨 CRITICAL INSTRUCTIONS:
        1.  **LANGUAGE CONSISTENCY**: Detect the language of the interview. 
            - If Chinese: Output ALL content in Simplified Chinese.
            - If English: Output ALL content in English.
            - **Set the `language` field in JSON to "zh" or "en".**
        2.  **NO MARKDOWN**: Do NOT use bolding marks (like **text**) in the JSON values. Output plain text only.
        3.  **STRICTLY NO TRANSLATION OF NAMES**: 
            - **KEEP IT VERBATIM**. 
            - If the expert says "Medtronic", write "Medtronic". Do NOT write "Medtronic (美敦力)".
            - If the expert says "Stapler", write "Stapler". Do NOT write "Stapler (吻合器)".
            - Do not add any parenthetical translations for company names, products, or technical terms.
        4.  **DATA PRECISION**: Capture EVERY number. Provide logic formulas for calculations.
        5.  **INTEGRATION**: Fit information into the main framework.

        ### FRAMEWORK KEYS:
        {keys_instruction}

        ### FRAMEWORK DETAILS:
        {framework_desc}

        ### OUTPUT JSON:
        {{
            "language": "zh", 
            "executive_summary": "Summary...",
            "structured_analysis": {{
                "market_size": [
                    "Point 1", 
                    "Point 2"
                ]
            }},
            "other_dimensions": {{
                "Topic": ["Detail"]
            }}
        }}
        """
        
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

# --- UI 主程序 ---
with st.sidebar:
    st.title("Clearstate AI")
    st.caption("Intelligent Qualitative Interview System")
    
    st.markdown("""
    <div class='developer-credit'>
    Developed by <b>Steve Jiang</b><br>
    Clearstate Consulting
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("---")
    
    api_key = st.text_input("Gemini API Key", type="password")
    
    st.markdown("### 📝 Project Info / 项目信息")
    company_name = st.text_input("Company / 公司名称", placeholder="e.g. Medtronic")
    product_name = st.text_input("Product / 产品领域", placeholder="e.g. Stapler")
    interview_date = st.date_input("Date / 访谈日期", datetime.date.today())
    
    # 🚨 修改：UI Label 更新
    st.markdown("### 🛠️ Interviewee Type / 访谈对象类型")
    interview_mode = st.radio(
        "Select Type / 选择类型",
        ("commercial", "clinical"),
        # 🚨 修改：Commercial -> Trade
        format_func=lambda x: "🏭 Trade (商业/厂商)" if x == "commercial" else "👨‍⚕️ Clinical (临床/专家)"
    )
    
    if st.button("🗑️ Reset / 重置"):
        st.session_state['analysis_result'] = None
        st.rerun()

st.markdown('<div class="main-header">智能定性访谈报告生成系统</div>', unsafe_allow_html=True)
st.markdown('<div class="sub-header">Intelligent Qualitative Interview Report Generation System</div>', unsafe_allow_html=True)

uploaded_file = st.file_uploader("📂 Upload Audio / 上传录音 (MP3/M4A Recommended)", type=['mp3', 'wav', 'm4a'])

if uploaded_file and st.session_state['analysis_result'] is None:
    if not api_key:
        st.error("Please enter API Key in the sidebar. / 请在侧边栏输入 API Key。")
    elif not company_name or not product_name:
        st.warning("Please fill in Company & Product info. / 请填写公司和产品信息。")
    else:
        st.audio(uploaded_file, format='audio/mp3')
        if st.button("🚀 Start Analysis (Gemini 3 Pro)", type="primary"):
            analyzer = InterviewAnalyzer(api_key)
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=f".{uploaded_file.name.split('.')[-1]}") as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_file_path = tmp_file.name

            with st.status("🤖 AI is processing... / AI 正在处理...", expanded=True) as status:
                st.write("🎧 Uploading audio to Gemini... / 正在上传音频...")
                audio_resource = analyzer.process_audio(tmp_file_path)
                
                if audio_resource:
                    st.write("🧠 Analyzing (Model: gemini-3-pro-preview)... / 正在分析...")
                    result = analyzer.analyze_interview(audio_resource, interview_mode)
                    
                    if result:
                        st.session_state['analysis_result'] = result
                        status.update(label="✅ Done! / 完成！", state="complete", expanded=False)
                        os.remove(tmp_file_path)
                        st.rerun()

if st.session_state['analysis_result']:
    res = st.session_state['analysis_result']
    
    st.success("✅ Analysis Complete. Please download the report. / 分析完成，请下载报告。")
    
    file_date_str = interview_date.strftime("%Y%m%d")
    file_name = f"Interview_Record_{company_name}_{product_name}_{file_date_str}.docx"
    
    docx_file = generate_word_report(res, company_name, product_name, interview_date, interview_mode)
    
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
