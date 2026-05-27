"""
쉬운 문서 해석기 — Easy-Easy 브랜딩 + 랜딩 페이지 개선 버전
- 🎨 유지: Easy-Easy 브랜드 헤더, 좌우 레이아웃 고정, 은은한 수직선
- 🎨 유지: 좌우 패널 세로 길이 대폭 확장 (스크롤 최소화 1200px)
- 🚀 극대화: '📖 원서 독해 & 영단어 학습 모드' 프롬프트 엔진 풀업그레이드
    (구동사/숙어 무제한 대량 추출, 소문자 강제, 불필요한 평 삭제, 복잡한 구문 분석 추가)
"""
import docx  
import io    
import streamlit as st
import fitz  
import google.generativeai as genai
from supabase import create_client, Client

# ============================================================
# ⚙️ 1. 페이지 설정
# ============================================================
st.set_page_config(
    page_title="Easy-Easy | 쉬운 문서 해석기",
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
    
    /* === 🚨 우측 상단 Streamlit Cloud 버튼들 정밀 타격 === */
    [data-testid="stToolbarActions"] {
        display: none !important;
    }
    [class*="viewerBadge_container"],
    [class*="viewerBadge_link"],
    [class*="ViewerBadge"] {
        display: none !important;
    }
    .stDeployButton,
    .stAppDeployButton,
    [data-testid="stMainMenu"] {
        display: none !important;
    }
    
    /* === 🚨 우측 하단 'Manage app' 버튼 (Streamlit Cloud 운영 메뉴) 완전 제거 === */
    [data-testid="manage-app-button"],
    [data-testid="manage-app-button-container"],
    [data-testid*="ManageApp"],
    [data-testid*="manageApp"],
    [class*="manage-app"],
    [class*="ManageApp"],
    .stStatusWidget,
    [data-testid="stStatusWidget"],
    [data-testid="stBottomBlockContainer"],
    iframe[title*="Manage"],
    iframe[title*="manage"],
    [aria-label*="Manage app"] {
        display: none !important;
        visibility: hidden !important;
    }
    
    /* === 🚨 사이드바 무조건 보이게 === */
    [data-testid="stSidebar"] {
        display: block !important;
        visibility: visible !important;
        background-color: #1e293b !important;
        border-right: 1px solid rgba(255,255,255,0.1) !important;
    }
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
    
    /* === 🎯 업로더 ↔ 컨트롤러 하단 자동 정렬 (반응형) === */
    div[data-testid="stHorizontalBlock"]:has([data-testid="stFileUploader"]) {
        align-items: flex-end !important;
    }
    div[data-testid="stHorizontalBlock"]:has([data-testid="stFileUploader"]) 
    > div[data-testid="column"],
    div[data-testid="stHorizontalBlock"]:has([data-testid="stFileUploader"]) 
    > div[data-testid="stColumn"] {
        align-self: flex-end !important;
    }
    
    /* === 🖱 Expander 헤더 클릭 영역 === */
    [data-testid="stExpander"] summary,
    [data-testid="stExpander"] details > summary {
        cursor: pointer !important;
        width: fit-content !important;
        display: inline-flex !important;
        align-items: center !important;
        gap: 8px !important;
        padding: 0.6rem 1.1rem !important;
        min-height: 52px !important;
        border-radius: 8px !important;
        transition: background-color 0.15s ease !important;
    }
    [data-testid="stExpander"] summary * {
        cursor: pointer !important;
    }
    [data-testid="stExpander"] summary:hover {
        background-color: rgba(168, 85, 247, 0.1) !important;
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

if st.session_state.get("user") is None and "logged_in_email" in st.query_params:
    saved_email = st.query_params["logged_in_email"]
    response = supabase.table("users").select("*").eq("email", saved_email).execute()
    if len(response.data) > 0:
        st.session_state["user"] = response.data[0]
    else:
        st.query_params.clear()

# ============================================================
# 💎 3. 요금제 팝업
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
# 🚪 4. 랜딩 페이지 — Easy-Easy 브랜딩
# ============================================================
if st.session_state.get("user") is None:
    st.markdown("""
    <style>
        .brand-header {
            display: flex;
            align-items: center;
            gap: 14px;
            margin: 0 0 3.5rem 0;
            padding-top: 0.5rem;
        }
        .brand-name {
            font-size: 1.7rem;
            font-weight: 800;
            background: linear-gradient(90deg, #d8b4fe 0%, #818cf8 100%);
            -webkit-background-clip: text;
            -webkit-text-fill-color: transparent;
            letter-spacing: -0.03em;
            line-height: 1;
        }
        
        div[data-testid="stHorizontalBlock"] {
            flex-wrap: nowrap !important;
            align-items: flex-start !important;
            gap: 5rem !important;
        }
        
        div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:first-child,
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:first-child {
            flex: 0 0 500px !important; 
            width: 500px !important;
            min-width: 500px !important;
            max-width: 500px !important;
            padding-right: 3rem !important;
            position: relative;
        }
        
        /* 은은한 수직 디바이더 */
        div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:first-child::after,
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:first-child::after {
            content: '';
            position: absolute;
            top: 0;
            bottom: 0;
            right: 0;
            width: 1px;
            background: linear-gradient(180deg, 
                transparent 0%, 
                rgba(168, 85, 247, 0.3) 4%, 
                rgba(168, 85, 247, 0.3) 96%, 
                transparent 100%
            );
        }
        
        div[data-testid="stHorizontalBlock"] > div[data-testid="stColumn"]:nth-child(2),
        div[data-testid="stHorizontalBlock"] > div[data-testid="column"]:nth-child(2) {
            flex: 1 1 auto !important;
            width: auto !important;
            min-width: 0 !important;
            padding-left: 3rem !important;
        }
        [data-testid="stImage"] img {
            border-radius: 12px;
            border: 1px solid rgba(255, 255, 255, 0.2);
            box-shadow: 0px 4px 40px rgba(129, 140, 248, 0.35);
        }
    </style>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    <div class="brand-header">
        <svg width="42" height="42" viewBox="0 0 42 42" fill="none" xmlns="http://www.w3.org/2000/svg">
            <defs>
                <linearGradient id="ee-grad-back" x1="0%" y1="0%" x2="100%" y2="100%">
                    <stop offset="0%" stop-color="#c4b5fd"/>
                    <stop offset="100%" stop-color="#8b5cf6"/>
                </linearGradient>
                <linearGradient id="ee-grad-front" x1="0%" y1="0%" x2="100%" y2="100%">
                    <stop offset="0%" stop-color="#818cf8"/>
                    <stop offset="100%" stop-color="#4f46e5"/>
                </linearGradient>
            </defs>
            <rect x="3" y="3" width="22" height="22" rx="7" fill="url(#ee-grad-back)" opacity="0.6"/>
            <rect x="17" y="17" width="22" height="22" rx="7" fill="url(#ee-grad-front)"/>
        </svg>
        <span class="brand-name">Easy-Easy</span>
    </div>
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
# 🔥 6. 프롬프트 세팅 (모드별 완벽 분리)
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

    "😎 촌철살인 동네형 모드": """너는 산전수전 다 겪은 실무 에이스 친한 동네 형입니다.
- 톤앤매너: 핵심만 짚어주는 거침없고 직관적인 반말. 
- 어투 제약사항(매우 중요): 명령조(~해라, ~한다) 절대 금지. 친근한 구어체(~해, ~야, ~거야, ~거든) 사용.
- 특징: 복잡한 이론 걷어내고 팩트 폭격 뼈대만 꽂아주기.""",
    
    # 🔥 원서 독해 특화 프롬프트
    "📖 원서 독해 & 영단어 학습 모드": """너는 영어 소설, 기사, 원서를 깊이 있게 분석하고 한국인 학습자의 가려운 곳을 완벽하게 긁어주는 1타 영어 독해 강사입니다.
- 톤앤매너: 학생이 사전을 켤 필요 없이 이 글 하나만으로 완벽히 학습할 수 있도록 짚어주는 친절하고 밀도 높은 존댓말.
- 제약사항: 뻔한 인사말이나 서론 절대 금지.
- 특징: 독해의 흐름이 끊기지 않도록 **구동사, 전치사구, 관용표현**을 집요하고 무제한으로 파헤칩니다."""
}

def build_prompt(text: str, mode: str) -> str:
    # 🚀 원서 독해 모드 - 무제한 추출 및 소문자 강제 로직 적용
    if mode == "📖 원서 독해 & 영단어 학습 모드":
        return f"""{PROMPT_TEMPLATES[mode]}

== 구조 지침 (반드시 따를 것) ==
### 1️⃣ 원문 뉘앙스 요약
"[원문의 전체적인 분위기, 상황, 그리고 행간에 숨겨진 의미를 한국어로 매끄럽고 흡입력 있게 요약]"

### 2️⃣ 📖 독해 필수 영단어, 숙어 & 구동사 총정리 (표)
- 원문에 등장하는 독해 필수 영단어, 숙어(Idioms), 구동사(Phrasal Verbs), 전치사구, 연어(Collocations) 등을 **사전이 필요 없을 정도로 최대한 많이, 빠짐없이, 무제한으로** 추출해. 
- 단순히 쉬운 단어보다는 한국인 학습자가 헷갈려하는 '전치사의 뉘앙스'나 '다의어의 문맥상 쓰임' 등 가려운 곳을 시원하게 긁어줘야 해. (원문 길이에 비례해 최소 수십 개 이상 아낌없이 추출할 것)
- ⚠️ **[절대 규칙]** 표 안의 영단어/숙어는 고유명사가 아닌 이상 **절대 첫 글자를 대문자로 시작하지 마** (전부 무조건 소문자로만 작성할 것. 예: Never mind -> never mind).
- 표 컬럼: | 표현 (소문자) | 품사 | 문맥상 의미 | 뉘앙스 설명 및 원문 활용 |

### 3️⃣ 🧠 구문 분석 (Syntax) & 원어민의 표현법
- 한국인들이 해석하다가 구조가 꼬이기 쉬운 복잡한 긴 문장, 도치구문, 혹은 생략구문 등을 1~3개 골라서, 문장 뼈대(주어/동사/수식어구)를 낱낱이 해부하고 직독직해 하는 방법을 가르쳐 줘.
- 단순히 뜻만 알려주지 말고 "왜 원어민은 이런 전치사를 썼는지, 왜 이런 구조로 말하는지" 원어민의 사고방식을 바탕으로 깊이 있게 설명해.

== 해석할 문서 ==
{text}"""
    else:
        # 기술 문서 모드용 지침
        return f"""{PROMPT_TEMPLATES[mode]}

== 구조 지침 (반드시 따를 것) ==
### 1️⃣ 핵심 한 줄
"[페이지 전체를 한 문장으로 꿰뚫는 요약]"

### 2️⃣ 찰진 해설 (짧게 끝내지 말 것!)
- 페이지에 나오는 **모든 중요 개념을 빠짐없이** 다룰 것
- 단순 번역이 아니라 **왜 필요한지**, **어떻게 동작하는지**, **무엇과 연결되는지** 맥락까지
- 명령어 / 파라미터 / 옵션이 등장하면 각각의 역할을 **표**로 재구성
- 어려운 개념은 일상 비유나 구체적 예시로 풀기

### 3️⃣ 실무 인사이트
- 실무에서 자주 마주치는 함정·실수·오해
- 왜 이게 중요한가 (성능·비용·QoR·수율 등 실제 영향)

### 4️⃣ 한 줄 정리
가장 마지막에 페이지 핵심을 한 문장으로 압축.

== 절대 원칙 ==
- 기술 용어는 영문 그대로 유지 (Fusion Compiler, LVT 등).
- 볼드체 뒤에는 조사 띄어쓰기.
- 첫 줄에 인사말 금지. 바로 본론 진입.

== 해석할 문서 ==
{text}"""

# ============================================================
# ⚙️ 7. 메인 화면 — 접이식 상단 + 전체화면 토글
# ============================================================

if "fullscreen_result" not in st.session_state:
    st.session_state["fullscreen_result"] = False
if "selected_mode" not in st.session_state:
    st.session_state["selected_mode"] = list(PROMPT_TEMPLATES.keys())[2]  
if "include_next_page" not in st.session_state:
    st.session_state["include_next_page"] = False

mode_keys = list(PROMPT_TEMPLATES.keys())

with st.expander("문서 & 해석 설정", expanded=True):
    top_left, top_right = st.columns(2, gap="large")
    
    with top_left:
        st.markdown("""
        <div style='display: flex; align-items: center; gap: 14px; margin: 0.25rem 0 0.5rem 0;'>
            <svg width="40" height="40" viewBox="0 0 40 40" fill="none" xmlns="http://www.w3.org/2000/svg">
                <defs>
                    <linearGradient id="title-emblem-back-app" x1="0%" y1="0%" x2="100%" y2="100%">
                        <stop offset="0%" stop-color="#c4b5fd"/>
                        <stop offset="100%" stop-color="#8b5cf6"/>
                    </linearGradient>
                    <linearGradient id="title-emblem-front-app" x1="0%" y1="0%" x2="100%" y2="100%">
                        <stop offset="0%" stop-color="#818cf8"/>
                        <stop offset="100%" stop-color="#4f46e5"/>
                    </linearGradient>
                </defs>
                <rect x="2" y="2" width="22" height="22" rx="7" fill="url(#title-emblem-back-app)" opacity="0.6"/>
                <rect x="16" y="16" width="22" height="22" rx="7" fill="url(#title-emblem-front-app)"/>
            </svg>
            <h1 style='margin: 0; padding: 0; line-height: 1.1;'>쉬운 문서 해석기</h1>
        </div>
        """, unsafe_allow_html=True)
        st.caption("어려운 기술 문서, 불필요한 사설 없이 핵심만 명확하게 짚어드립니다.")
        uploaded_file = st.file_uploader(
            "문서 파일 업로드 (PDF, TXT, DOCX)", 
            type=["pdf", "txt", "docx"],
            key="file_uploader_main"
        )
    
    with top_right:
        with st.container(border=True):
            st.markdown("### 해석 컨트롤러")
            selected_mode = st.selectbox(
                "해석 스타일 선택",
                mode_keys,
                index=mode_keys.index(st.session_state["selected_mode"]),
                label_visibility="collapsed",
                key="mode_selector_main"
            )
            st.session_state["selected_mode"] = selected_mode
    
    if uploaded_file is not None:
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
                else:
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
            st.session_state["file_ext"] = file_ext
            st.session_state["page_images"] = page_images
            st.session_state["page_texts"] = page_texts
            st.session_state["total_pages"] = len(page_texts)
        
        total_pages_show = st.session_state.get("total_pages", 1)
        st.success(f"✅ 총 {total_pages_show} 페이지 로드 완료")
    else:
        st.info("👆 좌측에 문서를 업로드하면 툴이 시작됩니다.")

if uploaded_file is None and "file_id" not in st.session_state:
    st.stop()

total_pages = st.session_state.get("total_pages", 1)
page_images = st.session_state.get("page_images", [])
page_texts = st.session_state.get("page_texts", [])
file_id = st.session_state.get("file_id", "")
file_ext = st.session_state.get("file_ext", "pdf")

def run_interpretation(text, mode, cache_key, pages_used=1):
    if GEMINI_API_KEY == "":
        st.error("🔑 Secrets 세팅에 GEMINI_API_KEY를 정상 등록해 주세요.")
        return False
    if not text:
        st.warning("⚠️ 추출 가능한 텍스트가 없습니다.")
        return False
    
    try:
        genai.configure(api_key=GEMINI_API_KEY)
        model = genai.GenerativeModel(
            MODEL_NAME, 
            generation_config=genai.types.GenerationConfig(max_output_tokens=8192)
        )
        current_user = st.session_state.get("user", {})
        is_admin = (current_user.get('plan_type') == 'ADMIN')
        
        if not is_admin and current_user.get('plan_type') == 'FREE':
            current_used = current_user.get('used_pages', 0)
            if current_used + pages_used > 3:
                st.error(f"🚫 무료 한도 초과: 현재 {current_used}장 사용 + 이번 {pages_used}장 = 한도 3장 초과")
                return False
        
        spinner_msg = f"🧠 [ADMIN] {pages_used}페이지 분석 중..." if is_admin else f"🧠 {pages_used}페이지 분석 중..."
        with st.spinner(spinner_msg):
            response = model.generate_content(build_prompt(text, mode))
        
        if not is_admin:
            new_used = current_user.get('used_pages', 0) + pages_used
            supabase.table("users").update({"used_pages": new_used}).eq("email", current_user.get('email')).execute()
            st.session_state["user"]['used_pages'] = new_used
        
        st.session_state["interpret_cache"][cache_key] = response.text
        return True
    except Exception as e:
        st.error(f"❌ 오류: {e}")
        return False

# ============================================================
# 🖥 전체화면 모드 OR 분할 보기 모드
# ============================================================
if st.session_state["fullscreen_result"]:
    fs_top = st.columns([2, 2, 4, 2])
    
    with fs_top[0]:
        view_page = st.number_input(
            f"📄 시작 페이지 (총 {total_pages})", 
            min_value=1, max_value=total_pages, value=1, step=1,
            key="view_page_input"
        )
    
    with fs_top[1]:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        include_next_raw_fs = st.checkbox(
            "📚 다음 페이지 포함",
            key="include_next_page",
            disabled=(view_page >= total_pages),
            help=f"체크 시 {view_page}~{min(view_page+1, total_pages)}페이지 해석"
        )
    include_next = include_next_raw_fs and (view_page < total_pages)
    
    with fs_top[2]:
        st.markdown(
            f"<div style='margin-top: 32px; color: #94a3b8;'>"
            f"🎭 <b>{selected_mode}</b>"
            f"</div>", 
            unsafe_allow_html=True
        )
    
    with fs_top[3]:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        if st.button("◀ 분할 보기로", use_container_width=True, key="exit_fullscreen"):
            st.session_state["fullscreen_result"] = False
            st.rerun()
    
    pages_suffix = "_2pg" if include_next else ""
    cache_key = f"{file_id}_{view_page}{pages_suffix}_{selected_mode}"
    is_cached = cache_key in st.session_state.get("interpret_cache", {})
    
    if not is_cached:
        num_pages_label = 2 if include_next else 1
        run_col = st.columns([8, 2])
        with run_col[1]:
            if st.button(f"✨ {num_pages_label}페이지 해석", type="primary", use_container_width=True, key="run_fs"):
                if include_next:
                    text = (
                        page_texts[view_page - 1].strip() 
                        + "\n\n--- 다음 페이지 ---\n\n" 
                        + page_texts[view_page].strip()
                    )
                    pages_used = 2
                else:
                    text = page_texts[view_page - 1].strip() if page_texts else ""
                    pages_used = 1
                
                if run_interpretation(text, selected_mode, cache_key, pages_used=pages_used):
                    st.rerun()
    
    with st.container(height=1200, border=True):
        if is_cached:
            st.markdown(st.session_state["interpret_cache"][cache_key])
        else:
            num_pages_label = 2 if include_next else 1
            st.info(f"👆 위의 **[✨ {num_pages_label}페이지 해석]** 버튼을 눌러주세요.")

else:
    st.divider()
    col_pdf, col_result = st.columns([1, 1], gap="large")
    
    with col_pdf:
        st.markdown(f"### {file_ext.upper()} 원본")
        
        page_row = st.columns([3, 2])
        with page_row[0]:
            view_page = st.number_input(
                "📄 시작 페이지", 
                min_value=1, max_value=total_pages, value=1, step=1,
                key="view_page_input"
            )
        with page_row[1]:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            include_next_raw = st.checkbox(
                "📚 다음 페이지 포함",
                key="include_next_page",
                disabled=(view_page >= total_pages),
                help=f"체크 시 {view_page}~{min(view_page+1, total_pages)}페이지를 함께 해석합니다"
            )
        include_next = include_next_raw and (view_page < total_pages)
        
        with st.container(height=1200, border=True):
            pages_to_show = [view_page]
            if include_next:
                pages_to_show.append(view_page + 1)
            
            for idx, pg in enumerate(pages_to_show):
                if page_images and page_images[pg - 1] is not None:
                    st.image(
                        page_images[pg - 1], 
                        caption=f"━━━ 페이지 {pg} / {total_pages} ━━━", 
                        use_container_width=True
                    )
                elif page_texts:
                    st.text_area(
                        f"페이지 {pg}", 
                        page_texts[pg - 1], 
                        height=580 if include_next else 1100, 
                        disabled=True, 
                        label_visibility="visible" if include_next else "collapsed",
                        key=f"page_text_{pg}_{idx}"
                    )
                if idx < len(pages_to_show) - 1:
                    st.markdown("<hr style='border-color: rgba(168,85,247,0.2);'>", unsafe_allow_html=True)
    
    pages_suffix = "_2pg" if include_next else ""
    cache_key = f"{file_id}_{view_page}{pages_suffix}_{selected_mode}"
    is_cached = cache_key in st.session_state.get("interpret_cache", {})
    
    with col_result:
        header_col, fs_btn_col = st.columns([4, 2])
        with header_col:
            mode_parts = selected_mode.split()
            display_title = " ".join(mode_parts[1:])
            st.markdown(f"### {display_title} 답변")
        with fs_btn_col:
            st.markdown("<div style='margin-top: 4px;'></div>", unsafe_allow_html=True)
            if st.button("🔍 전체화면", use_container_width=True, key="enter_fullscreen"):
                st.session_state["fullscreen_result"] = True
                st.rerun()
        
        status_col, btn_col = st.columns([3, 2])
        with status_col:
            st.text_input(
                "✨ 현재 상태", 
                value="🟢 메모리에서 불러옴" if is_cached else "⏳ 해석 대기 중", 
                disabled=True
            )
        with btn_col:
            st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
            btn_label = f"✨ {len(pages_to_show)}페이지 해석"
            interpret_btn = st.button(
                btn_label, 
                type="primary" if not is_cached else "secondary", 
                use_container_width=True,
                key="interpret_btn_split"
            )
        
        with st.container(height=1200, border=True):
            if interpret_btn and not is_cached:
                if include_next:
                    text = (
                        page_texts[view_page - 1].strip() 
                        + "\n\n--- 다음 페이지 ---\n\n" 
                        + page_texts[view_page].strip()
                    )
                    pages_used = 2
                else:
                    text = page_texts[view_page - 1].strip() if page_texts else ""
                    pages_used = 1
                
                if run_interpretation(text, selected_mode, cache_key, pages_used=pages_used):
                    st.rerun()
            
            if is_cached:
                st.markdown(st.session_state["interpret_cache"][cache_key])
            elif not interpret_btn:
                st.info(f"👆 상단의 **[✨ {len(pages_to_show)}페이지 해석]** 버튼을 눌러주세요.")
