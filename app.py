"""
쉬운 문서 해석기 — MVP v2.2 (TXT/DOCX 지원 추가 및 텍스트 뷰어 통합)
- 기능 추가: .txt, .docx 확장자 업로드 지원 및 1500자 단위 가상 페이지 분할
- UI 업데이트: 문서 타입에 따라 원본 뷰어를 이미지/텍스트 모드로 자동 전환
"""
import docx  # 최상단에 추가 (pip install python-docx)
import io    # 최상단에 추가
import streamlit as st
import fitz  # PyMuPDF
import google.generativeai as genai
from supabase import create_client, Client

# Supabase DB 연결 세팅
SUPABASE_URL = "https://nufvazmyuvhqkeysfwla.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im51ZnZhem15dXZocWtleXNmd2xhIiwicm9sZSI6ImFub24iLCJpYXQiOjE3Nzk2NDIxODcsImV4cCI6MjA5NTIxODE4N30.6MNFMzQjEkeItf7OTMr82gg5pn66s3ksxwnrIff14po"

# 클라이언트 생성
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)

# ============================================================
# 🔒 [5단계] 로그인 시스템
# ============================================================
if "user" not in st.session_state:
    st.session_state.user = None

# 로그인되어 있지 않다면 로그인 화면만 보여주고 멈춤
if st.session_state.user is None:
    st.title("🔒 쉬운 문서 해석기 로그인")
    st.write("서비스를 이용하려면 이메일을 입력해주세요.")
    
    email_input = st.text_input("이메일 주소")
    login_btn = st.button("시작하기")
    
    if login_btn and email_input:
        # 1. DB에 유저가 있는지 확인
        response = supabase.table("users").select("*").eq("email", email_input).execute()
        
        if len(response.data) > 0:
            # 2. 기존 가입자면 정보 가져오기
            st.session_state.user = response.data[0]
        else:
            # 3. 처음 온 사람이면 DB에 새로 등록 (기본 FREE 플랜)
            new_user = {"email": email_input, "plan_type": "FREE", "used_pages": 0}
            insert_res = supabase.table("users").insert(new_user).execute()
            st.session_state.user = insert_res.data[0]
            
        st.success("✅ 로그인 성공!")
        st.rerun()  # 화면을 새로고침해서 메인 앱으로 진입
        
    st.stop() # 로그인을 안 했으면 이 아래 코드(메인 앱)는 절대 실행되지 않음

# ============================================================
# 👤 유저 상태창 (사이드바)
# ============================================================
st.sidebar.markdown(f"**👤 계정**: {st.session_state.user['email']}")
st.sidebar.markdown(f"**💳 플랜**: {st.session_state.user['plan_type']}")

# 💡 ADMIN 계정일 경우 사용량 표시에 (무제한) 안내 추가
if st.session_state.user['plan_type'] == 'ADMIN':
    st.sidebar.markdown(f"**📄 사용량**: {st.session_state.user['used_pages']} 장 (👑 무제한)")
else:
    st.sidebar.markdown(f"**📄 사용량**: {st.session_state.user['used_pages']} 장")

if st.sidebar.button("로그아웃"):
    st.session_state.user = None
    st.rerun()

# ============================================================
# 페이지 설정
# ============================================================
st.set_page_config(
    page_title="쉬운 문서 해석기",
    page_icon="📄",
    layout="wide",
)

# 👇👇 여기서부터 추가! 고급스러운 UI를 위한 커스텀 CSS 주입 👇👇
st.markdown("""
<style>
    /* 1. 메인 타이틀 그라데이션 텍스트 효과 */
    h1 {
        background: linear-gradient(90deg, #d8b4fe, #818cf8);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        font-weight: 800 !important;
    }
    
    /* 2. Primary 버튼 (결제, 해석하기) 고급 그라데이션 & 네온 그림자 효과 */
    button[kind="primary"] {
        background: linear-gradient(135deg, #6366f1 0%, #a855f7 100%) !important;
        border: none !important;
        color: white !important;
        font-weight: 600 !important;
        border-radius: 8px !important;
        box-shadow: 0 4px 15px rgba(168, 85, 247, 0.4) !important;
        transition: all 0.3s ease !important;
    }
    
    /* Primary 버튼에 마우스 올렸을 때 살짝 떠오르는 애니메이션 */
    button[kind="primary"]:hover {
        transform: translateY(-2px);
        box-shadow: 0 6px 20px rgba(168, 85, 247, 0.6) !important;
    }
    
    /* 3. 컨테이너(박스) 테두리를 부드럽게 만들고 아주 얇은 형광 라인 추가 */
    [data-testid="stVerticalBlock"] > div > div {
        border-radius: 12px;
    }
    div[data-testid="stContainer"] {
        border: 1px solid rgba(255, 255, 255, 0.1) !important;
        background-color: rgba(30, 41, 59, 0.5) !important;
        backdrop-filter: blur(10px); /* 반투명 유리 효과 */
    }
    
    /* 4. 파일 업로더 점선 테두리 고급화 */
    [data-testid="stFileUploadDropzone"] {
        border: 2px dashed rgba(129, 140, 248, 0.5) !important;
        background-color: rgba(15, 23, 42, 0.3) !important;
        border-radius: 12px !important;
    }
</style>
""", unsafe_allow_html=True)
# 👆👆 커스텀 CSS 주입 끝 👆👆
# ============================================================
# 💎 요금제 및 결제 안내 UI
# ============================================================
st.divider()
st.markdown("### 💎 플랜 업그레이드")

# 2칸으로 깔끔하게 나눕니다
col_free, col_pro = st.columns(2)

with col_free:
    with st.container(border=True):
        st.subheader("FREE")
        st.markdown("## ₩ 0 / 월")
        
        # 텍스트 박스 높이를 강제로 맞춰서 아래 버튼을 밀어내어 정렬합니다.
        st.markdown(
            """<div style='min-height: 100px;'>
            ✔️ <b>매월 3장</b> 해석 제공<br>
            ✔️ 기본 문서 텍스트 추출<br>
            ✔️ 일반 속도 처리
            </div>""", 
            unsafe_allow_html=True
        )
        
        if st.session_state.user['plan_type'] == 'FREE':
            st.button("현재 이용 중", disabled=True, use_container_width=True)
        else:
            st.button("FREE 플랜", disabled=True, use_container_width=True)

with col_pro:
    with st.container(border=True):
        st.subheader("PRO (인기)")
        st.markdown("## ₩ 9,900 / 월")
        
        # 동일한 높이를 주어 버튼 위치를 완벽하게 동일 선상에 맞춥니다.
        st.markdown(
            """<div style='min-height: 100px;'>
            ✔️ <b>월 1,000장 해석 제공</b><br>
            ✔️ 1타 강사 / 비유 모드 완벽 지원<br>
            ✔️ 한도 초과 스트레스 없는 쾌적함
            </div>""", 
            unsafe_allow_html=True
        )
        
        BASE_CHECKOUT_LINK = "https://easy-explain-saas.lemonsqueezy.com/checkout/buy/7a87b27c-335a-42c9-9995-54eb03fb49a3"
        current_user_email = st.session_state.user['email']
        final_checkout_link = f"{BASE_CHECKOUT_LINK}?checkout[email]={current_user_email}"
        
        if st.session_state.user['plan_type'] == 'PRO':
            st.button("현재 이용 중 (PRO)", disabled=True, use_container_width=True)
        elif st.session_state.user['plan_type'] == 'ADMIN':
            st.button("👑 마스터 계정 사용 중", disabled=True, use_container_width=True)
        else:
            st.link_button("Pro 구독하기", final_checkout_link, type="primary", use_container_width=True)
# ============================================================
# ⚙️ 환경 설정
# ============================================================
GEMINI_API_KEY = ""
MODEL_NAME = "gemini-3.1-flash-lite" 

if "interpret_cache" not in st.session_state:
    st.session_state.interpret_cache = {}

# ============================================================
# 🔥 3종 핵심 모드 프롬프트 (동네 형 모드 추가)
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

    "😎 촌철살인 동네 형 모드": """너는 산전수전 다 겪은 실무 에이스 선배입니다.
- 톤앤매너: 친한 후배에게 핵심만 짚어주는 거침없고 직관적인 반말.
- 제약사항: "야, 복잡하게 생각할 거 없고", "정신 똑바로 차려", "이게 왜 필요하냐면" 같은 억지스러운 유행어나 반복적인 추임새, 인사말은 절대 쓰지 마세요.
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
- 볼드체 한국어 버그 방지(매우 중요): 마크다운 볼드체(**강조**) 뒤에 한국어 조사(입니다, 은, 는 등)가 띄어쓰기 없이 붙으면 기호가 노출되는 오류가 납니다. 반드시 `**강조 단어** 입니다` 처럼 뒤를 한 칸 띄우거나, `**강조 단어입니다**` 처럼 조사를 볼드체 안에 포함시키세요. 따옴표와 볼드체는 섞어 쓰지 마십시오.
- 서론 금지: "설명해 드리겠습니다", "자 시작합니다" 등 요약 앞뒤에 붙는 모든 쓸데없는 인사말과 추임새를 엄격히 금지합니다.

== 해석할 문서 ==
{text}

위 텍스트를 [{mode}] 컨셉에 맞춰 설명해 주세요."""


# ============================================================
# 메인 화면 구성
# ============================================================
# ============================================================
# 상단 레이아웃 (좌: 파일 업로드 / 우: 컨트롤러)
# ============================================================
top_left, top_right = st.columns(2, gap="large") 
with top_right:
    st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
    st.markdown("### 해석 컨트롤러")
    selected_mode = st.selectbox(
        "해석 스타일 선택",
        list(PROMPT_TEMPLATES.keys()),
        index=2, # 기본값을 동네 형 모드로 세팅
    )
    st.caption("☝️ 원하는 스타일을 미리 선택해 두세요.")

with top_left:
    st.title("📄 쉬운 문서 해석기")
    st.caption("어려운 기술 문서, 불필요한 사설 없이 핵심만 명확하게 짚어드립니다.")
    uploaded_file = st.file_uploader("문서 파일 업로드 (PDF, TXT, DOCX)", type=["pdf", "txt", "docx"])

if uploaded_file is None:
    st.info("👆 좌측에 문서를 업로드하면 툴이 시작됩니다.")
    st.stop()

# 문서 로드 + 렌더링
file_id = f"{uploaded_file.name}_{uploaded_file.size}"
file_ext = uploaded_file.name.split('.')[-1].lower()

if st.session_state.get("file_id") != file_id:
    page_images = []
    page_texts = []
    
    with st.spinner("📖 문서 읽는 중..."):
        if file_ext == "pdf":
            # PDF 처리 로직
            pdf_bytes = uploaded_file.read()
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
            for i in range(len(doc)):
                page = doc[i]
                pix = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5))
                page_images.append(pix.tobytes("png"))
                page_texts.append(page.get_text())
            doc.close()
        else:
            # TXT, DOCX 텍스트 추출 로직
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
            
            # 텍스트가 너무 길면 1500자씩 잘라서 가상 페이지로 분할
            chunk_size = 1500
            if not raw_text.strip():
                page_texts = ["(내용이 없습니다)"]
            else:
                page_texts = [raw_text[i:i+chunk_size] for i in range(0, len(raw_text), chunk_size)]
            
            # 이미지가 없으므로 None으로 채움
            page_images = [None] * len(page_texts)

    st.session_state.file_id = file_id
    st.session_state.page_images = page_images
    st.session_state.page_texts = page_texts
    st.session_state.total_pages = len(page_texts)

total_pages = st.session_state.total_pages
page_images = st.session_state.page_images
page_texts = st.session_state.page_texts

st.success(f"✅ 총 {total_pages} 페이지(구간) 로드 완료")
st.divider()

# ============================================================
# 1. 상단 컨트롤러 (UI 정렬 유지)
# ============================================================
st.markdown("### 해석 컨트롤러")

selected_mode = st.selectbox(
    "해석 스타일 선택",
    list(PROMPT_TEMPLATES.keys()),
    index=1,  
)
st.caption("☝️ 원하는 스타일을 선택하고, 아래 원본 뷰어에서 이동한 페이지의 해석 버튼을 눌러주세요.")

st.divider()

# ============================================================
# 2. 하단: 좌/우 분할 레이아웃
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
            # PDF인 경우 이미지 출력
            st.image(
                page_images[view_page - 1],
                caption=f"━━━ 페이지 {view_page} / {total_pages} ━━━",
                use_container_width=True,
            )
        else:
            # TXT, DOCX인 경우 텍스트 뷰어 출력
            st.markdown(f"**━━━ 가상 페이지 {view_page} / {total_pages} ━━━**")
            st.text_area(
                "문서 내용", 
                page_texts[view_page - 1], 
                height=700, 
                disabled=True, 
                label_visibility="collapsed"
            )

# 캐시 확인
cache_key = f"{file_id}_{view_page}_{selected_mode}"
is_cached = cache_key in st.session_state.interpret_cache

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
                st.error("🔑 코드 상단의 GEMINI_API_KEY 변수에 실제 API 키를 입력해주세요.")
            else:
                text = page_texts[view_page - 1].strip()

                if not text:
                    st.warning("⚠️ 해당 페이지에 추출 가능한 텍스트가 없습니다. (그림표 전용 페이지 등)")
                else:
                    try:
                        genai.configure(api_key=GEMINI_API_KEY)
                        generation_config = genai.types.GenerationConfig(max_output_tokens=8192)
                        model = genai.GenerativeModel(MODEL_NAME, generation_config=generation_config)
                        prompt = build_prompt(text, selected_mode)

                        # === [요금제 문지기 로직 시작] ===
                        user_info = st.session_state.user
                        is_admin = (user_info['plan_type'] == 'ADMIN')  # 💡 ADMIN 여부 확인 변수 생성
                        
                        # 무료 유저이면서 3장 이상 썼다면 차단 (단, ADMIN 계정은 이 조건문을 무조건 통과)
                        if not is_admin and user_info['plan_type'] == 'FREE' and user_info['used_pages'] >= 3:
                            st.error("🚫 무료 제공량(3장)을 모두 소진했습니다.")
                            st.info("무제한 해석 기능을 이용하려면 PRO 플랜으로 업그레이드해 주세요!")
                        else:
                            # 번역 실행 (FREE 한도 내, PRO, 그리고 ADMIN 계정 모두 정상 작동)
                            spinner_msg = f"🧠 [ADMIN] 페이지 {view_page}를 분석 중입니다 (무제한 패스)..." if is_admin else f"🧠 페이지 {view_page}를 분석 중입니다..."
                            with st.spinner(spinner_msg):
                                response = model.generate_content(prompt)
                            
                            # 💡 ADMIN 계정이 아닐 때만 DB 및 세션에 실질적인 사용량을 누적시킵니다.
                            if not is_admin:
                                new_used_pages = user_info['used_pages'] + 1
                                supabase.table("users").update({"used_pages": new_used_pages}).eq("email", user_info['email']).execute()
                                st.session_state.user['used_pages'] = new_used_pages
                            
                            st.session_state.interpret_cache[cache_key] = response.text
                            st.rerun()
                        # === [요금제 문지기 로직 끝] ===

                    except Exception as e:
                        st.error(f"❌ 오류 발생: {e}")

        if is_cached:
            st.markdown(st.session_state.interpret_cache[cache_key])
        elif not interpret_btn:
            st.info(f"👆 ¼쪽에서 원하는 페이지로 이동한 뒤, 상단의 **[✨ 현재 페이지 해석]** 버튼을 눌러주세요.")
