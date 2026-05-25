"""
쉬운 문서 해석기 — PRO 최종 통합본 (진짜 절대 고정 레이아웃)
- UX 개선: 스트림릿의 억지 반응형(자동 리사이징)을 flex 속성으로 강제 무력화
- 로그인 화면: 좌측 420px, 우측 850px 완전 고정 (브라우저를 줄이면 가로 스크롤만 생김)
"""
import docx  
import io    
import streamlit as st
import fitz  
import google.generativeai as genai
from supabase import create_client, Client

# ============================================================
# ⚙️ 1. 페이지 설정 및 전역 고급 다크 스타일 주입
# ============================================================
st.set_page_config(
    page_title="쉬운 문서 해석기",
    page_icon="📄",
    layout="wide",
)

# 세션 초기화
if "user" not in st.session_state:
    st.session_state["user"] = None
if "interpret_cache" not in st.session_state:
    st.session_state["interpret_cache"] = {}

st.markdown("""
<style>
    .stApp { background-color: #0f172a; overflow-x: auto !important; }
    [data-testid="stToolbar"] { visibility: hidden !important; }
    [data-testid="stSidebar"] { min-width: 300px !important; }
    h1 { background: linear-gradient(90deg, #d8b4fe, #818cf8); -webkit-background-clip: text; -webkit-text-fill-color: transparent; font-weight: 800 !important; }
    button[kind="primary"] { background: linear-gradient(135deg, #6366f1 0%, #a855f7 100%) !important; border: none !important; color: white !important; font-weight: 600 !important; border-radius: 8px !important; box-shadow: 0 4px 15px rgba(168, 85, 247, 0.4) !important; transition: all 0.3s ease !important; }
    button[kind="primary"]:hover { transform: translateY(-2px); box-shadow: 0 6px 20px rgba(168, 85, 247, 0.6) !important; }
    [data-testid="stVerticalBlock"] > div > div { border-radius: 12px; }
    div[data-testid="stContainer"] { border: 1px solid rgba(255, 255, 255, 0.1) !important; background-color: rgba(30, 41, 59, 0.4) !important; backdrop-filter: blur(10px); }
    [data-testid="stFileUploadDropzone"] { border: 2px dashed rgba(129, 140, 248, 0.5) !important; background-color: rgba(15, 23, 42, 0.3) !important; border-radius: 12px !important; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# 💎 2. 요금제 팝업(모달) 창 기능정의
# ============================================================
@st.dialog("💎 플랜 업그레이드 안내")
def show_pricing_modal():
    st.write("서비스의 무제한 기능을 경험해 보세요.")
    col_free, col_pro = st.columns(2)

    with col_free:
        with st.container(height=420, border=True):
            st.subheader("FREE")
            st.markdown("## ₩ 0 / 월")
            st.markdown(
                """<div style='min-height: 180px; color: #94a3b8;'>
                ✔️ <b>매월 3장</b> 해석 제공<br>
                ✔️ 기본 문서 텍스트 추출<br>
                ✔️ 일반 속도 처리
                </div>""", 
                unsafe_allow_html=True
            )
            if st.session_state["user"]['plan_type'] == 'FREE':
                st.button("현재 이용 중", disabled=True, key="modal_free_btn", use_container_width=True)
            else:
                st.button("FREE 플랜", disabled=True, key="modal_free_btn_dis", use_container_width=True)

    with col_pro:
        with st.container(height=420, border=True):
            st.subheader("PRO (인기)")
            st.markdown("## ₩ 9,900 / 월")
            st.markdown(
                """<div style='min-height: 180px; color: #94a3b8;'>
                ✔️ <b>월 1,000장 해석 제공</b><br>
                ✔️ 1타 강사 / 비유 모드 완벽 지원<br>
                ✔️ 한도 초과 스트레스 없는 쾌적함
                </div>""", 
                unsafe_allow_html=True
            )
            BASE_CHECKOUT_LINK = "https://easy-explain-saas.lemonsqueezy.com/checkout/buy/7a87b27c-335a-42c9-9995-54eb03fb49a3"
            current_user_email = st.session_state["user"]['email']
            final_checkout_link = f"{BASE_CHECKOUT_LINK}?checkout[email]={current_user_email}"
            
            if st.session_state["user"]['plan_type'] == 'PRO':
                st.button("현재 이용 중 (PRO)", disabled=True, key="modal_pro_btn", use_container_width=True)
            elif st.session_state["user"]['plan_type'] == 'ADMIN':
                st.button("👑 마스터 계정 사용 중", disabled=True, key="modal_admin_btn", use_container_width=True)
            else:
                st.link_button("Pro 구독하기", final_checkout_link, type="primary", use_container_width=True)

# ============================================================
# 🔒 3. API 키 설정 및 로그인 시스템 (절대 크기 강제 고정)
# ============================================================
SUPABASE_URL = "https://nufvazmyuvhqkeysfwla.supabase.co"
SUPABASE_KEY = st.secrets["SUPABASE_KEY"]
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
GEMINI_API_KEY = st.secrets["GEMINI_API_KEY"]
MODEL_NAME = "gemini-3.1-flash-lite" 

if st.session_state["user"] is None:
    # 💡 [핵심] 로그인 화면에서만 작동하는 '절대 크기 강제 고정' 핵(Hack)
    st.markdown("""
    <style>
        /* 1. 전체 블록 컨테이너 최소 너비 1300px 강제 고정 */
        [data-testid="block-container"] {
            min-width: 1300px !important;
            max-width: 1300px !important;
            padding-top: 5vh !important;
        }
        
        /* 2. 컬럼 자동 줄바꿈 완전 차단 */
        [data-testid="stHorizontalBlock"] {
            flex-wrap: nowrap !important;
            gap: 3rem !important;
        }
        
        /* 3. 좌측 구역(텍스트/로그인): 무조건 420px 고정 (스트림릿 반응형 무시) */
        [data-testid="column"]:nth-child(1) {
            width: 420px !important;
            min-width: 420px !important;
            max-width: 420px !important;
            flex: 0 0 420px !important; 
            border-right: 1px solid rgba(255, 255, 255, 0.15);
            padding-right: 3rem !important;
        }
        
        /* 4. 우측 구역(사진): 무조건 850px 고정 (스트림릿 반응형 무시) */
        [data-testid="column"]:nth-child(2) {
            width: 850px !important;
            min-width: 850px !important;
            max-width: 850px !important;
            flex: 0 0 850px !important;
        }
        
        /* 스크린샷 테두리 마감 */
        [data-testid="stImage"] img {
            border-radius: 12px;
            border: 1px solid rgba(255, 255, 255, 0.2);
            box-shadow: 0px 4px 40px rgba(129, 140, 248, 0.35);
        }
    </style>
    """, unsafe_allow_html=True)

    col_left, col_right = st.columns(2)
    
    with col_left:
        st.markdown("<div style='margin-top: 5vh;'></div>", unsafe_allow_html=True)
        st.markdown("""
        <h1 style='font-size: 3rem; line-height: 1.3;'>어려운 기술 문서,<br>이제 가장 쉽게 읽으세요.</h1>
        <p style='color: #f8fafc; font-size: 1.05rem; margin-top: 1.5rem; margin-bottom: 2.5rem;'>복잡한 영문 매뉴얼, 번역기 돌리며 고생하지 마세요. AI가 핵심만 짚어 가장 이해하기 쉬운 한글로 설명해 드립니다.</p>
        """, unsafe_allow_html=True)
        
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
                st.rerun()  

    with col_right:
        st.markdown("<div style='margin-top: 5vh;'></div>", unsafe_allow_html=True)
        try:
            st.image("result_preview.png", use_container_width=True, output_format="PNG")
        except:
            st.info("💡 여기에 결과물 스샷(result_preview.png)이 큼직하게 표시됩니다.")
            
    st.stop()

# ============================================================
# 👤 4. 유저 상태창 및 워크스페이스 제어 (사이드바)
# ============================================================
st.sidebar.markdown(f"**👤 계정**: {st.session_state['user']['email']}")
st.sidebar.markdown(f"**💳 플랜**: {st.session_state['user']['plan_type']}")

if st.session_state["user"]['plan_type'] == 'ADMIN':
    st.sidebar.markdown(f"**📄 사용량**: {st.session_state['user']['used_pages']} 장 (👑 무제한)")
else:
    st.sidebar.markdown(f"**📄 사용량**: {st.session_state['user']['used_pages']} 장")

st.sidebar.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)

if st.sidebar.button("💎 플랜 업그레이드", use_container_width=True):
    show_pricing_modal()

if st.sidebar.button("로그아웃", use_container_width=True):
    st.session_state["user"] = None
    st.rerun()

# ============================================================
# 🔥 5. 3종 핵심 모드 프롬프트
# ============================================================
PROMPT_TEMPLATES = {
    "👨‍🏫 1타 강사 해설 모드": """너는 반도체/EDA 업계를 주름잡는 1타 강사입니다. 
- 톤앤매너: 수강생이 절대 졸 수 없게 만드는 리듬감 있고 흡입력 있는 존댓말.
- 제약사항: "자, 여러분", "집중하십시오", "오늘 다룰 내용은" 같은 뻔한 서론이나 인사말은 절대 금지합니다.
- 특징: 실무 맥락을 짚어주고, 강조할 부분은 굵은 글씨로 처리하여 가독성을 높이세요.""",

    "💡 비유 모드": """너는 복잡한 기술을 일상생활에 빗대어 설명하는 천재적인 블로거입니다.
- 톤앤매너: 정중하지만 무릎을 탁 치게 만드는 센스 있는 존댓말.
- 제약사항: 인사말, 서론, 뻔한 추임새 절대 금지. 바로 본론의 비유로 진입하세요.
- 특징: 아무리 어려운 개념도 일상 개념에 찰떡같이 비유해서 직관적으로 이해되게 만들어 주세요.""",

    "😎 촌철살인 동네 형 모드": """너는 산전수전 다 겪은 실무 에이스 친한 동네 형입니다.
- 톤앤매너: 아끼는 동생에게 알려주듯 핵심만 짚어주는 거침없고 직관적인 반말. 
- 어투 제약사항(매우 중요): 딱딱한 명령조나 권위적인 말투(~해라, ~마라, ~한다)는 절대 금지합니다. 반드시 친근하고 자연스러운 구어체 말투(~해, ~야, ~하지 마, ~거야, ~거든)로 끝맺으세요.
- 일반 제약사항: 억지스러운 유행어나 반복적인 추임새, 인사말은 절대 쓰지 마세요.
- 특징: 복잡한 이론은 과감히 걷어내고, 실무에서 이 개념이 어떻게 작동하는지 뼈대만 날카롭게 팩트 폭격으로 꽂아주세요."""
}

def build_prompt(text: str, mode: str) -> str:
    system_base = PROMPT_TEMPLATES[mode]
    return f"""{system_base}

== 구조 지침 ==
1. 핵심 타격: "> [가장 뼈때리는 한 줄 요약]" 형태로 문서 맨 처음에 인용문으로 강렬하게 배치하세요. 서론 없이 바로 요약부터 나와야 합니다.
2. 찰진 해설: 본문의 핵심 단락, 명령어, 파라미터를 쪼개서 선택된 모드의 컨셉에 맞게 풀어주세요.
3. 실무 한 줄 팁: 마지막에 실무자를 위한 조언으로 마무리해 주세요.

== 🚨 출력 필수 규칙 ==
- 기술 용어 영문 유지(중요): 업계 표준 기술 용어, 고유 명사, 아키텍처 이름, 파라미터, 약어 등은 억지로 한글로 번역하지 말고 반드시 영문 그대로 사용하십시오. (예: '응용 프로그램 인터페이스' 대신 'API', '플로어플랜' 대신 'Floorplan' 등 원문 뉘앙스 유지)
- 데이터 정상화: 원문에 있는 수치나 표 데이터는 절대 뭉개거나 생략하지 마십시오. 반드시 마크다운 표(|---|---|) 문법을 사용하여 깔끔한 그리드 형태로 재구성하십시오.
- 볼드체 한국어 버그 방지(매우 중요): 마크다운 볼드체(**강조**) 뒤에 한국어 조사(입니다, 은, 는 등)가 정규 표현식 오류가 나지 않도록 반드시 `**강조 단어** 입니다` 처럼 뒤를 한 칸 띄우거나, `**강조 단어입니다**` 처럼 조사를 포함하세요.

== 해석할 문서 ==
{text}

위 텍스트를 [{mode}] 컨셉에 맞춰 설명해 주세요."""

# ============================================================
# ⚙️ 6. 상단 파일 로더 및 컨트롤러 조립 
# ============================================================
top_left, top_right = st.columns(2, gap="large")

with top_left:
    st.title("📄 쉬운 문서 해석기")
    st.caption("어려운 기술 문서, 불필요한 사설 없이 핵심만 명확하게 짚어드립니다.")
    uploaded_file = st.file_uploader("문서 파일 업로드 (PDF, TXT, DOCX)", type=["pdf", "txt", "docx"])

with top_right:
    st.markdown("<div style='margin-top: 50px;'></div>", unsafe_allow_html=True) 
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
# ⚙️ 7. 파일 파싱 및 텍스트 추출 로직 
# ============================================================
file_id = f"{uploaded_file.name}_{uploaded_file.size}"
file_ext = uploaded_file.name.split('.')[-1].lower()

if st.session_state.get("file_id") != file_id:
    page_images = []
    page_texts = []
    
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
        else:
            raw_text = ""
            if file_ext == "txt":
                raw_bytes = uploaded_file.read()
                try:
                    raw_text = raw_bytes.decode('utf-8')
                except UnicodeDecodeError:
                    raw_text = raw_bytes.decode('cp949', errors='ignore')
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

total_pages = st.session_state["total_pages"]
page_images = st.session_state["page_images"]
page_texts = st.session_state["page_texts"]

st.success(f"✅ 총 {total_pages} 페이지(구간) 로드 완료")
st.divider()

# ============================================================
# 🔥 8. 하단 메인 레이아웃 (좌우 대칭 매칭)
# ============================================================
col_pdf, col_result = st.columns([1, 1], gap="large")

with col_pdf:
    st.markdown(f"### {file_ext.upper()} 원본")
    
    view_page = st.number_input(
        "📄 이동할 페이지 번호 입력", 
        min_value=1, 
        max_value=total_pages, 
        value=1, 
        step=1
    )
    
    with st.container(height=800, border=True):
        if page_images[view_page - 1] is not None:
            st.image(
                page_images[view_page - 1],
                caption=f"━━━ 페이지 {view_page} / {total_pages} ━━━",
                use_container_width=True,
            )
        else:
            st.markdown(f"**━━━ 가상 페이지 {view_page} / {total_pages} ━━━**")
            st.text_area(
                "문서 내용", 
                page_texts[view_page - 1], 
                height=700, 
                disabled=True, 
                label_visibility="collapsed"
            )

cache_key = f"{file_id}_{view_page}_{selected_mode}"
is_cached = cache_key in st.session_state["interpret_cache"]

with col_result:
    st.markdown(f"### {selected_mode.split()[1]} {selected_mode.split()[2]}의 실시간 답변")
    
    status_col, btn_col = st.columns([3, 2])
    with status_col:
        status_text = "🟢 메모리에서 불러옴" if is_cached else "⏳ 해석 대기 중"
        st.text_input("✨ 현재 상태", value=status_text, disabled=True)
        
    with btn_col:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        interpret_btn = st.button(
            "✨ 현재 페이지 해석",
            type="primary" if not is_cached else "secondary",
            use_container_width=True,
        )
    
    with st.container(height=800, border=True):
        if interpret_btn and not is_cached:
            if GEMINI_API_KEY == "":
                st.error("🔑 Secrets 세팅에 GEMINI_API_KEY를 정상 등록해 주세요.")
            else:
                text = page_texts[view_page - 1].strip()

                if not text:
                    st.warning("⚠️ 해당 페이지에 추출 가능한 텍스트가 없습니다.")
                else:
                    try:
                        genai.configure(api_key=GEMINI_API_KEY)
                        generation_config = genai.types.GenerationConfig(max_output_tokens=8192)
                        model = genai.GenerativeModel(MODEL_NAME, generation_config=generation_config)
                        prompt = build_prompt(text, selected_mode)

                        user_info = st.session_state["user"]
                        is_admin = (user_info['plan_type'] == 'ADMIN')
                        
                        if not is_admin and user_info['plan_type'] == 'FREE' and user_info['used_pages'] >= 3:
                            st.error("🚫 무료 제공량(3장)을 모두 소진했습니다.")
                            st.info("사이드바의 [💎 플랜 업그레이드] 버튼을 눌러 결제 후 무제한으로 사용해 보세요!")
                        else:
                            spinner_msg = f"🧠 [ADMIN] 페이지 {view_page}를 분석 중 (무제한)..." if is_admin else f"🧠 페이지 {view_page}를 분석 중입니다..."
                            with st.spinner(spinner_msg):
                                response = model.generate_content(prompt)
                            
                            if not is_admin:
                                new_used_pages = user_info['used_pages'] + 1
                                supabase.table("users").update({"used_pages": new_used_pages}).eq("email", user_info['email']).execute()
                                st.session_state["user"]['used_pages'] = new_used_pages
                            
                            st.session_state["interpret_cache"][cache_key] = response.text
                            st.rerun()

                    except Exception as e:
                        st.error(f"❌ 오류 발생: {e}")

        if is_cached:
            st.markdown(st.session_state["interpret_cache"][cache_key])
        elif not interpret_btn:
            st.info(f"👆 왼쪽에서 원하는 페이지로 이동한 뒤, 상단의 **[✨ 현재 페이지 해석]** 버튼을 눌러주세요.")
