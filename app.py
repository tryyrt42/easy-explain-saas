"""
쉬운 문서 해석기 — 사이드바 복구 + PPTX 지원 추가 버전
- 🔧 수정 1: 랜딩 페이지의 min-width 1200px 누수로 사이드바가 화면 밖으로 밀려나던 버그 제거
- 🔧 수정 2: 사이드바 강제 노출 CSS 추가 (구버전/신버전 Streamlit 모두 대응)
- 🆕 추가: PPTX 파일 지원 (슬라이드 단위로 해석)
"""
import docx  
import io    
import streamlit as st
import fitz  
import google.generativeai as genai
from pptx import Presentation  # 🆕 PPTX 지원
from supabase import create_client, Client

# ============================================================
# ⚙️ 1. 페이지 설정
# ============================================================
st.set_page_config(
    page_title="쉬운 문서 해석기",
    page_icon="📄",
    layout="wide",
    initial_sidebar_state="expanded" 
)

# 안전한 세션 초기화
if "user" not in st.session_state:
    st.session_state["user"] = None
if "interpret_cache" not in st.session_state:
    st.session_state["interpret_cache"] = {}

# ============================================================
# 💡 전역 CSS — 사이드바를 명시적으로 강제 노출
# ============================================================
st.markdown("""
<style>
    /* === 배경 === */
    .stApp { background-color: #0f172a; }
    
    /* === 🔥 메인 컨텐츠 위로 끌어올리기 (기본 6rem → 1rem) === */
    .main .block-container,
    [data-testid="stMainBlockContainer"],
    [data-testid="stAppViewBlockContainer"] {
        padding-top: 1rem !important;
    }
    
    /* === 상단 헤더는 투명하게만 === */
    header[data-testid="stHeader"] { 
        background: transparent !important;
        height: 0 !important;
    }
    
    /* === 🚨 우측 상단 잔존 요소들 (toolbarMode='minimal'이 못 잡는 것들만 백업으로) === */
    /* ⚠️ stToolbar 자체는 절대 건드리지 않음! 사이드바 토글이 그 안에 있어서 같이 사라짐 */
    .stDeployButton,
    .stAppDeployButton,
    [data-testid="stMainMenu"],
    [class*="viewerBadge_container"],
    [class*="viewerBadge_link"] {
        display: none !important;
    }
    
    /* === 🚨 사이드바 무조건 보이게 (toolbar 제거해도 사이드바는 살아있음) === */
    [data-testid="stSidebar"] {
        display: block !important;
        visibility: visible !important;
        background-color: #1e293b !important;
        border-right: 1px solid rgba(255,255,255,0.1) !important;
    }
    
    /* === 사이드바 토글 버튼은 stToolbar와 별개 셀렉터라 안전 === */
    [data-testid="collapsedControl"],
    [data-testid="stSidebarCollapsedControl"],
    [data-testid="stSidebarCollapseButton"],
    [data-testid="stSidebarHeader"] button {
        display: flex !important;
        visibility: visible !important;
        opacity: 1 !important;
        z-index: 999999 !important;
    }
    
    /* === 헤드라인 그라데이션 === */
    h1 { 
        background: linear-gradient(90deg, #d8b4fe, #818cf8); 
        -webkit-background-clip: text; 
        -webkit-text-fill-color: transparent; 
        font-weight: 800 !important; 
    }
    
    /* === Primary 버튼 === */
    button[kind="primary"] { 
        background: linear-gradient(135deg, #6366f1 0%, #a855f7 100%) !important; 
        border: none !important; 
        color: white !important; 
        font-weight: 600 !important; 
        border-radius: 8px !important; 
        box-shadow: 0 4px 15px rgba(168, 85, 247, 0.4) !important; 
        transition: all 0.3s ease !important; 
    }
    button[kind="primary"]:hover { 
        transform: translateY(-2px); 
        box-shadow: 0 6px 20px rgba(168, 85, 247, 0.6) !important; 
    }
    
    /* === 컨테이너 === */
    [data-testid="stVerticalBlock"] > div > div { border-radius: 12px; }
    div[data-testid="stContainer"] { 
        border: 1px solid rgba(255, 255, 255, 0.1) !important; 
        background-color: rgba(30, 41, 59, 0.4) !important; 
        backdrop-filter: blur(10px); 
    }
    [data-testid="stFileUploadDropzone"] { 
        border: 2px dashed rgba(129, 140, 248, 0.5) !important; 
        background-color: rgba(15, 23, 42, 0.3) !important; 
        border-radius: 12px !important; 
    }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 🔒 2. Supabase 연결 및 F5 새로고침 방어 로직
# ============================================================
SUPABASE_URL = "https://nufvazmyuvhqkeysfwla.supabase.co"
SUPABASE_KEY = st.secrets["SUPABASE_KEY"]
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
MODEL_NAME = "gemini-3.1-flash-lite" 

# F5 방어 로직
if st.session_state.get("user") is None and "logged_in_email" in st.query_params:
    saved_email = st.query_params["logged_in_email"]
    response = supabase.table("users").select("*").eq("email", saved_email).execute()
    if len(response.data) > 0:
        st.session_state["user"] = response.data[0]
    else:
        st.query_params.clear()

# ============================================================
# 💎 3. 요금제 팝업(모달) 창
# ============================================================
@st.dialog("💎 플랜 업그레이드 안내")
def show_pricing_modal():
    st.write("서비스의 무제한 기능을 경험해 보세요.")
    col_free, col_pro = st.columns(2)

    user_info = st.session_state.get("user", {})

    with col_free:
        with st.container(height=420, border=True):
            st.subheader("FREE")
            st.markdown("## ₩ 0 / 월")
            st.markdown("""<div style='min-height: 180px; color: #94a3b8;'>✔️ <b>매월 3장</b> 해석 제공<br>✔️ 기본 문서 텍스트 추출<br>✔️ 일반 속도 처리</div>""", unsafe_allow_html=True)
            if user_info.get('plan_type') == 'FREE':
                st.button("현재 이용 중", disabled=True, key="modal_free_btn", use_container_width=True)
            else:
                st.button("FREE 플랜", disabled=True, key="modal_free_btn_dis", use_container_width=True)

    with col_pro:
        with st.container(height=420, border=True):
            st.subheader("PRO (인기)")
            st.markdown("## ₩ 9,900 / 월")
            st.markdown("""<div style='min-height: 180px; color: #94a3b8;'>✔️ <b>월 1,000장 해석 제공</b><br>✔️ 1타 강사 / 비유 모드 완벽 지원<br>✔️ 한도 초과 스트레스 없는 쾌적함</div>""", unsafe_allow_html=True)
            BASE_CHECKOUT_LINK = "https://easy-explain-saas.lemonsqueezy.com/checkout/buy/7a87b27c-335a-42c9-9995-54eb03fb49a3"
            current_user_email = user_info.get('email', '')
            final_checkout_link = f"{BASE_CHECKOUT_LINK}?checkout[email]={current_user_email}"
            
            if user_info.get('plan_type') == 'PRO':
                st.button("현재 이용 중 (PRO)", disabled=True, key="modal_pro_btn", use_container_width=True)
            elif user_info.get('plan_type') == 'ADMIN':
                st.button("👑 마스터 계정 사용 중", disabled=True, key="modal_admin_btn", use_container_width=True)
            else:
                st.link_button("Pro 구독하기", final_checkout_link, type="primary", use_container_width=True)

# ============================================================
# 🚪 4. 랜딩 페이지 — 문제의 min-width 1200px 제거!
# ============================================================
if st.session_state.get("user") is None:
    st.markdown("""
    <style>
        /* 🔧 수정: min-width 1200px 제거 (이게 사이드바를 밀어내던 원인!) */
        div[data-testid="stHorizontalBlock"] {
            flex-wrap: nowrap !important;
            align-items: center !important;
        }
        /* 좌측 컬럼 550px 고정 */
        div[data-testid="column"]:first-child {
            flex: 0 0 550px !important; 
            border-right: 1px solid rgba(255, 255, 255, 0.2);
            padding-right: 3rem !important;
        }
        div[data-testid="column"]:nth-child(2) {
            flex: 1 1 auto !important;
            padding-left: 3rem !important;
        }
        [data-testid="stImage"] img {
            border-radius: 12px;
            border: 1px solid rgba(255, 255, 255, 0.2);
            box-shadow: 0px 4px 40px rgba(129, 140, 248, 0.35);
        }
    </style>
    """, unsafe_allow_html=True)

    col_left, col_right = st.columns([1, 2])
    
    with col_left:
        st.markdown("<h1 style='font-size: 3.2rem; line-height: 1.2;'>어려운 기술 문서,<br>이제 가장 쉽게 읽으세요.</h1>", unsafe_allow_html=True)
        st.markdown("<p style='color: #f8fafc; font-size: 1.1rem; margin-top: 1.5rem; margin-bottom: 2.5rem;'>복잡한 영문 매뉴얼, 번역기 돌리며 고생하지 마세요. AI가 핵심만 짚어 가장 이해하기 쉬운 한글로 설명해 드립니다.</p>", unsafe_allow_html=True)
        
        with st.container(border=True):
            st.markdown("<h3 style='text-align: center; font-weight: 700;'>문서 해석 시작하기</h3>", unsafe_allow_html=True)
            st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
            email_input = st.text_input("이메일 주소", placeholder="example@email.com", label_visibility="collapsed")
            login_btn = st.button("✨ 이메일로 간편하게 시작하기", type="primary", use_container_width=True)
            
            if login_btn and email_input:
                response = supabase.table("users").select("*").eq("email", email_input).execute()
                if len(response.data) > 0:
                    st.session_state["user"] = response.data[0]
                else:
                    new_user = {"email": email_input, "plan_type": "FREE", "used_pages": 0}
                    insert_res = supabase.table("users").insert(new_user).execute()
                    st.session_state["user"] = insert_res.data[0]
                
                st.query_params["logged_in_email"] = email_input
                st.rerun()  

    with col_right:
        try:
            st.image("result_preview.png", use_container_width=True, output_format="PNG")
        except:
            st.info("💡 여기에 결과물 스샷(result_preview.png)이 큼직하게 표시됩니다.")
            
    st.stop() 

# ============================================================
# 👤 5. 유저 사이드바
# ============================================================
user_data = st.session_state.get("user", {})

with st.sidebar:
    st.markdown(f"**👤 계정**: {user_data.get('email', '')}")
    st.markdown(f"**💳 플랜**: {user_data.get('plan_type', '')}")

    if user_data.get('plan_type') == 'ADMIN':
        st.markdown(f"**📄 사용량**: {user_data.get('used_pages', 0)} 장 (👑 무제한)")
    else:
        st.markdown(f"**📄 사용량**: {user_data.get('used_pages', 0)} 장")

    st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)

    if st.button("💎 플랜 업그레이드", use_container_width=True):
        show_pricing_modal()

    if st.button("로그아웃", use_container_width=True):
        st.session_state["user"] = None
        st.query_params.clear() 
        st.rerun()

# ============================================================
# 🔥 6. 프롬프트 세팅
# ============================================================
PROMPT_TEMPLATES = {
    "👨‍🏫 1타 강사 해설 모드": """너는 반도체/EDA 업계를 주름잡는 1타 강사입니다. 
- 톤앤매너: 수강생이 절대 졸 수 없게 만드는 리듬감 있고 흡입력 있는 존댓말.
- 제약사항: "자, 여러분", "집중하십시오" 같은 뻔한 서론 절대 금지.
- 특징: 실무 맥락을 짚어주고 강조할 부분은 굵은 글씨 처리.""",

    "💡 비유 모드": """너는 복잡한 기술을 일상생활에 빗대어 설명하는 천재적인 블로거입니다.
- 톤앤매너: 정중하지만 무릎을 탁 치게 만드는 센스 있는 존댓말.
- 제약사항: 인사말 서론 금지. 바로 비유 진입.
- 특징: 어려운 개념을 직관적으로 이해되게 찰떡 비유.""",

    "😎 촌철살인 동네 형 모드": """너는 산전수전 다 겪은 실무 에이스 친한 동네 형입니다.
- 톤앤매너: 핵심만 짚어주는 거침없고 직관적인 반말. 
- 어투 제약사항(매우 중요): 명령조(~해라, ~한다) 절대 금지. 친근한 구어체(~해, ~야, ~거야, ~거든) 사용.
- 특징: 복잡한 이론 걷어내고 팩트 폭격 뼈대만 꽂아주기."""
}

def build_prompt(text: str, mode: str) -> str:
    return f"""{PROMPT_TEMPLATES[mode]}

== 구조 지침 ==
1. 핵심 타격: "> [가장 뼈때리는 한 줄 요약]" 
2. 찰진 해설: 명령어, 파라미터 컨셉에 맞게.
3. 실무 한 줄 팁: 실무자 조언으로 마무리.

== 규칙 ==
- 기술 용어 영문 유지
- 마크다운 표 활용
- 볼드체 뒤에 조사 띄어쓰기

== 해석할 문서 ==
{text}"""

# ============================================================
# ⚙️ 7. 메인 화면: 파일 업로드 (PPTX 추가!)
# ============================================================
top_left, top_right = st.columns(2, gap="large")

with top_left:
    st.title("📄 쉬운 문서 해석기")
    st.caption("어려운 기술 문서, 불필요한 사설 없이 핵심만 명확하게 짚어드립니다.")
    # 🆕 PPTX 추가
    uploaded_file = st.file_uploader(
        "문서 파일 업로드 (PDF, TXT, DOCX, PPTX)", 
        type=["pdf", "txt", "docx", "pptx"]
    )

with top_right:
    # 💡 margin-top 128px로 왼쪽 업로드 박스 밑바닥과 정확히 정렬
    st.markdown("<div style='margin-top: 128px;'></div>", unsafe_allow_html=True) 
    with st.container(border=True):
        st.markdown("### 해석 컨트롤러")
        selected_mode = st.selectbox(
            "해석 스타일 선택",
            list(PROMPT_TEMPLATES.keys()),
            index=2, 
            label_visibility="collapsed" 
        )

if uploaded_file is None:
    st.info("👆 좌측에 문서를 업로드하면 툴이 시작됩니다.")
    st.stop()

# ============================================================
# ⚙️ 8. 파일 파싱 — PPTX 처리 추가
# ============================================================
file_id = f"{uploaded_file.name}_{uploaded_file.size}"
file_ext = uploaded_file.name.split('.')[-1].lower()

if st.session_state.get("file_id") != file_id:
    page_images, page_texts = [], []
    with st.spinner("📖 문서 읽는 중..."):
        if file_ext == "pdf":
            pdf_bytes = uploaded_file.read()
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            for i in range(len(doc)):
                page = doc[i]
                pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
                page_images.append(pix.tobytes("png"))
                page_texts.append(page.get_text())
            doc.close()
        
        # 🆕 PPTX 처리 — 슬라이드 = 페이지
        elif file_ext == "pptx":
            prs = Presentation(io.BytesIO(uploaded_file.read()))
            for slide_idx, slide in enumerate(prs.slides, start=1):
                slide_parts = []
                
                # 1) 일반 텍스트 박스에서 텍스트 추출
                for shape in slide.shapes:
                    if shape.has_text_frame:
                        for para in shape.text_frame.paragraphs:
                            para_text = "".join(run.text for run in para.runs)
                            if para_text.strip():
                                slide_parts.append(para_text)
                
                # 2) 표 안에 있는 텍스트도 추출
                for shape in slide.shapes:
                    if shape.has_table:
                        for row in shape.table.rows:
                            row_text = " | ".join(
                                cell.text.strip() 
                                for cell in row.cells 
                                if cell.text.strip()
                            )
                            if row_text:
                                slide_parts.append(row_text)
                
                # 3) 슬라이드 노트(발표자 메모)도 함께 해석에 활용
                if slide.has_notes_slide:
                    notes = slide.notes_slide.notes_text_frame.text.strip()
                    if notes:
                        slide_parts.append(f"\n[발표자 노트]\n{notes}")
                
                slide_full = "\n".join(slide_parts) if slide_parts else "(빈 슬라이드)"
                page_texts.append(f"[슬라이드 {slide_idx} / {len(prs.slides)}]\n\n{slide_full}")
            
            page_images = [None] * len(page_texts)
        
        else:
            # TXT / DOCX 처리
            raw_text = ""
            if file_ext == "txt":
                raw_bytes = uploaded_file.read()
                try: raw_text = raw_bytes.decode('utf-8')
                except: raw_text = raw_bytes.decode('cp949', errors='ignore')
            elif file_ext == "docx":
                doc_file = docx.Document(io.BytesIO(uploaded_file.read()))
                raw_text = "\n".join([para.text for para in doc_file.paragraphs])
            chunk_size = 1500
            if not raw_text.strip(): 
                page_texts = ["(내용이 없습니다)"]
            else: 
                page_texts = [raw_text[i:i+chunk_size] for i in range(0, len(raw_text), chunk_size)]
            page_images = [None] * len(page_texts)

    st.session_state["file_id"] = file_id
    st.session_state["page_images"] = page_images
    st.session_state["page_texts"] = page_texts
    st.session_state["total_pages"] = len(page_texts)

total_pages = st.session_state.get("total_pages", 1)
page_images = st.session_state.get("page_images", [])
page_texts = st.session_state.get("page_texts", [])

# 🆕 PPTX는 '슬라이드'로 표기
unit_label = "슬라이드" if file_ext == "pptx" else "페이지"
st.success(f"✅ 총 {total_pages} {unit_label} 로드 완료")
st.divider()

col_pdf, col_result = st.columns([1, 1], gap="large")

with col_pdf:
    st.markdown(f"### {file_ext.upper()} 원본")
    view_page = st.number_input(
        f"📄 이동할 {unit_label} 번호 입력", 
        min_value=1, max_value=total_pages, value=1, step=1
    )
    
    with st.container(height=800, border=True):
        if page_images and page_images[view_page - 1] is not None:
            st.image(
                page_images[view_page - 1], 
                caption=f"━━━ {unit_label} {view_page} / {total_pages} ━━━", 
                use_container_width=True
            )
        elif page_texts:
            st.text_area(
                "문서 내용", 
                page_texts[view_page - 1], 
                height=700, 
                disabled=True, 
                label_visibility="collapsed"
            )

cache_key = f"{file_id}_{view_page}_{selected_mode}"
is_cached = cache_key in st.session_state.get("interpret_cache", {})

with col_result:
    st.markdown(f"### {selected_mode.split()[1]} {selected_mode.split()[2]} 답변")
    
    status_col, btn_col = st.columns([3, 2])
    with status_col:
        st.text_input(
            "✨ 현재 상태", 
            value="🟢 메모리에서 불러옴" if is_cached else "⏳ 해석 대기 중", 
            disabled=True
        )
    with btn_col:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        interpret_btn = st.button(
            f"✨ 현재 {unit_label} 해석", 
            type="primary" if not is_cached else "secondary", 
            use_container_width=True
        )
    
    with st.container(height=800, border=True):
        if interpret_btn and not is_cached:
            if GEMINI_API_KEY == "":
                st.error("🔑 Secrets 세팅에 GEMINI_API_KEY를 정상 등록해 주세요.")
            else:
                text = page_texts[view_page - 1].strip() if page_texts else ""
                if not text:
                    st.warning("⚠️ 추출 가능한 텍스트가 없습니다.")
                else:
                    try:
                        genai.configure(api_key=GEMINI_API_KEY)
                        model = genai.GenerativeModel(
                            MODEL_NAME, 
                            generation_config=genai.types.GenerationConfig(max_output_tokens=8192)
                        )
                        
                        current_user = st.session_state.get("user", {})
                        is_admin = (current_user.get('plan_type') == 'ADMIN')
                        
                        if not is_admin and current_user.get('plan_type') == 'FREE' and current_user.get('used_pages', 0) >= 3:
                            st.error("🚫 무료 제공량(3장)을 모두 소진했습니다.")
                        else:
                            with st.spinner(f"🧠 [ADMIN] 분석 중..." if is_admin else f"🧠 분석 중..."):
                                response = model.generate_content(build_prompt(text, selected_mode))
                            
                            if not is_admin:
                                new_used = current_user.get('used_pages', 0) + 1
                                supabase.table("users").update({"used_pages": new_used}).eq("email", current_user.get('email')).execute()
                                st.session_state["user"]['used_pages'] = new_used
                            
                            st.session_state["interpret_cache"][cache_key] = response.text
                            st.rerun()
                    except Exception as e:
                        st.error(f"❌ 오류: {e}")

        if is_cached:
            st.markdown(st.session_state["interpret_cache"][cache_key])
        elif not interpret_btn:
            st.info(f"👆 상단의 **[✨ 현재 {unit_label} 해석]** 버튼을 눌러주세요.")
