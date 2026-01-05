import streamlit as st
import google.generativeai as genai
import tempfile
import os
import time
import json
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import io
import datetime
from google.api_core import retry

# --- 页面配置 ---
st.set_page_config(
    page_title="ConsultAI Pro (Uncensored)",
    layout="wide",
    page_icon="🔓",
    initial_sidebar_state="expanded"
)

# --- CSS 样式 ---
st.markdown("""
<style>
    .main-header { font-size: 2.2rem; color: #003366; font-weight: bold; margin-bottom: 10px; }
    .sub-header { font-size: 1.0rem; color: #666; margin-bottom: 20px; border-left: 4px solid #d93025; padding-left: 10px; }
    div[data-testid="stFileUploader"] { margin-top: 20px; }
</style>
""", unsafe_allow_html=True)

# --- Session State ---
if 'analysis_result' not in st.session_state:
    st.session_state['analysis_result'] = None

# --- Word 生成函数 (保持不变) ---
def generate_word_report(data, company, product, date, mode):
    doc = Document()
    title_text = f"{company} - {product} 访谈记录"
    heading = doc.add_heading(title_text, 0)
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run(f"访谈时间: {date} | 访谈类型: {'商业/厂商' if mode == 'commercial' else '临床/专家'}")
    run.italic = True
    run.font.color.rgb = RGBColor(100, 100, 100)
    doc.add_paragraph("-" * 50).alignment = WD_ALIGN_PARAGRAPH.CENTER

    doc.add_heading('1. 执行摘要 (Executive Summary)', level=1)
    doc.add_paragraph(data.get('executive_summary', '无摘要内容'))

    doc.add_heading('2. 结构化维度分析', level=1)
    comm_map = {
        "Market Size & Scale": "2.1 市场规模与体量",
        "Competition Landscape": "2.2 竞争格局",
        "Sales & Marketing Strategy": "2.3 销售与营销策略",
        "Channel Strategy": "2.4 渠道与准入策略",
        "New Product Development (NPD)": "2.5 新产品开发计划",
        "Industry Trends": "2.6 行业总体趋势"
    }
    clin_map = {
        "Technology Prospects": "2.1 技术市场前景",
        "Hospital Adoption": "2.2 医院落地与使用情况",
        "Competition (Clinical View)": "2.3 竞品竞争情况 (临床视角)",
        "Clinical Pain Points": "2.4 临床痛点与未满足需求",
        "User Experience": "2.5 专家使用体验",
        "Expectations": "2.6 专家预期与展望"
    }
    current_map = comm_map if mode == "commercial" else clin_map
    structured = data.get('structured_analysis', {})
    
    for eng_key, cn_title in current_map.items():
        found_key = None
        for k in structured.keys():
            if eng_key.lower() in k.lower().replace("_", " "):
                found_key = k
                break
        if found_key:
            doc.add_heading(cn_title, level=2)
            for point in structured[found_key]:
                doc.add_paragraph(point, style='List Bullet')

    other_dims = data.get('other_dimensions', {})
    if other_dims:
        doc.add_heading('3. 其他重要维度 (新发现)', level=1)
        for k, v in other_dims.items():
            doc.add_heading(k, level=2)
            for point in v:
                doc.add_paragraph(point, style='List Bullet')

    doc.add_heading('4. 访谈详细实录 (Q&A)', level=1)
    qa_log = data.get('qa_log', [])
    for qa in qa_log:
        p_q = doc.add_paragraph()
        run_q = p_q.add_run(f"Q: {qa['question']}")
        run_q.bold = True
        run_q.font.color.rgb = RGBColor(0, 51, 102)
        p_a = doc.add_paragraph(f"A: {qa['answer']}")
        if qa.get('context_note'):
            p_note = doc.add_paragraph(f"[注: {qa['context_note']}]")
            p_note.style = 'Quote'

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
            # 使用 1.5 Flash (最稳定)
            self.model = genai.GenerativeModel('gemini-2.5-flash') 
        except Exception as e:
            st.error(f"API 配置错误: {e}")

    def process_audio(self, audio_file_path):
        try:
            myfile = genai.upload_file(audio_file_path)
            with st.spinner("🎧 正在上传并解析音频..."):
                while myfile.state.name == "PROCESSING":
                    time.sleep(2)
                    myfile = genai.get_file(myfile.name)
            if myfile.state.name == "FAILED":
                st.error("音频解析失败。")
                return None
            return myfile
        except Exception as e:
            st.error(f"上传错误: {e}")
            return None

    def analyze_interview(self, audio_resource, mode):
        # 1. 定义框架
        if mode == "commercial":
            framework_desc = """
            1. **Market Size & Scale**: Numbers, growth rates, TAM/SAM.
            2. **Competition Landscape**: Competitor names, market shares, strengths/weaknesses.
            3. **Sales & Marketing Strategy**: Pricing, sales team structure, promotion methods.
            4. **Channel Strategy**: Distributors, hospital listing (入院), regional coverage.
            5. **New Product Development (NPD)**: R&D pipeline, launch dates.
            6. **Industry Trends**: Policy impact (VBP/DRG), macro trends.
            """
        else: # clinical
            framework_desc = """
            1. **Technology Prospects**: Clinical value, future potential.
            2. **Hospital Adoption**: Usage rate, department acceptance, billing codes.
            3. **Competition (Clinical View)**: Comparison with other brands/therapies in practice.
            4. **Clinical Pain Points**: Unmet needs, side effects, limitations of current tech.
            5. **User Experience**: Ease of use, learning curve, preference.
            6. **Expectations**: What improvements do they want?
            """

        # 2. 定义 Prompt
        system_prompt = f"""
        You are a Senior Strategy Consultant.
        Task: Extract a **Comprehensive Interview Record** from the audio.

        ### 🚨 STRICT RULES:
        1.  **Source of Truth:** ONLY use info from audio. NO external knowledge.
        2.  **Completeness:** Capture ALL numbers, names, and specific details.
        3.  **Structure:** Follow the framework below strictly.
        4.  **New Dimensions:** Put anything outside the framework into "other_dimensions".

        ### FRAMEWORK:
        {framework_desc}

        ### OUTPUT JSON:
        {{
            "executive_summary": "300 words summary.",
            "structured_analysis": {{
                "dimension_key": ["Detail 1", "Detail 2"]
            }},
            "other_dimensions": {{
                "Topic Name": ["Detail 1"]
            }},
            "qa_log": [
                {{
                    "question": "Consultant question",
                    "answer": "Expert answer",
                    "context_note": "Context if needed"
                }}
            ]
        }}
        **Language:** Simplified Chinese.
        """
        
        # 3. 🚨 核心修复：关闭所有安全过滤器 🚨
        safety_settings = [
            {
                "category": "HARM_CATEGORY_HARASSMENT",
                "threshold": "BLOCK_NONE"
            },
            {
                "category": "HARM_CATEGORY_HATE_SPEECH",
                "threshold": "BLOCK_NONE"
            },
            {
                "category": "HARM_CATEGORY_SEXUALLY_EXPLICIT",
                "threshold": "BLOCK_NONE"
            },
            {
                "category": "HARM_CATEGORY_DANGEROUS_CONTENT",
                "threshold": "BLOCK_NONE"
            },
        ]
        
        try:
            # 调用 API，带上 safety_settings 和 timeout
            response = self.model.generate_content(
                [audio_resource, system_prompt],
                safety_settings=safety_settings, # 允许所有内容通过
                request_options={"timeout": 600} # 允许 10 分钟超时
            )
            
            # 检查是否因为其他原因被拦截
            if response.prompt_feedback:
                 if response.prompt_feedback.block_reason:
                     st.warning(f"⚠️ 警告: 输入内容可能触发生存策略: {response.prompt_feedback.block_reason}")

            # 尝试获取文本
            try:
                text = response.text
                if "```json" in text:
                    text = text.replace("```json", "").replace("```", "")
                return json.loads(text.strip())
            except ValueError:
                # 如果 response.text 依然报错，打印详细的 candidate 信息以便调试
                st.error("❌ 模型生成被中断，未返回有效文本。")
                st.write("Debug Info (Finish Reason):", response.candidates[0].finish_reason)
                st.write("Debug Info (Safety Ratings):", response.candidates[0].safety_ratings)
                return None

        except Exception as e:
            st.error(f"分析过程中断: {e}")
            return None

# --- UI 主程序 ---
with st.sidebar:
    st.title("🔓 ConsultAI Pro")
    st.caption("Uncensored Version")
    api_key = st.text_input("Gemini API Key", type="password")
    
    st.markdown("### 📝 报告基础信息")
    company_name = st.text_input("公司名称", placeholder="例如：美敦力")
    product_name = st.text_input("产品/领域", placeholder="例如：吻合器")
    interview_date = st.date_input("访谈时间", datetime.date.today())
    
    st.markdown("---")
    st.markdown("### 🛠️ 访谈场景")
    interview_mode = st.radio(
        "选择类型：",
        ("commercial", "clinical"),
        format_func=lambda x: "🏭 厂商/商业" if x == "commercial" else "👨‍⚕️ 临床/专家"
    )
    
    if st.button("🗑️ 清空当前记录"):
        st.session_state['analysis_result'] = None
        st.rerun()

st.markdown(f'<div class="main-header">{company_name if company_name else "未命名公司"} - 访谈智能梳理系统</div>', unsafe_allow_html=True)

# --- 上传区域 ---
uploaded_file = st.file_uploader("📂 上传录音文件 (建议 MP3/M4A)", type=['mp3', 'wav', 'm4a'])

if uploaded_file and st.session_state['analysis_result'] is None:
    if not api_key:
        st.error("请先在左侧输入 API Key")
    elif not company_name or not product_name:
        st.warning("⚠️ 请先在左侧侧边栏填写【公司名称】和【产品/领域】。")
    else:
        st.audio(uploaded_file, format='audio/mp3')
        if st.button("🚀 开始分析 (无限制模式)", type="primary"):
            analyzer = InterviewAnalyzer(api_key)
            
            with tempfile.NamedTemporaryFile(delete=False, suffix=f".{uploaded_file.name.split('.')[-1]}") as tmp_file:
                tmp_file.write(uploaded_file.getvalue())
                tmp_file_path = tmp_file.name

            with st.status("🤖 AI 正在处理...", expanded=True) as status:
                st.write("🎧 正在上传音频...")
                audio_resource = analyzer.process_audio(tmp_file_path)
                
                if audio_resource:
                    st.write("🧠 正在提取结构化数据 (已关闭安全拦截)...")
                    result = analyzer.analyze_interview(audio_resource, interview_mode)
                    
                    if result:
                        st.session_state['analysis_result'] = result
                        status.update(label="✅ 整理完成！", state="complete", expanded=False)
                        os.remove(tmp_file_path)
                        st.rerun()

# --- 结果展示与导出 ---
if st.session_state['analysis_result']:
    res = st.session_state['analysis_result']
    
    st.success("✅ 分析完成，请下载 Word 报告")
    
    file_date_str = interview_date.strftime("%Y%m%d")
    file_name = f"{company_name}_{product_name}_访谈记录_{file_date_str}.docx"
    
    docx_file = generate_word_report(res, company_name, product_name, interview_date, interview_mode)
    
    st.download_button(
        label=f"📥 下载 Word 报告: {file_name}",
        data=docx_file,
        file_name=file_name,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary"
    )

    st.markdown("---")
    st.markdown("### 📊 网页版预览")
    st.write(res.get('executive_summary'))

