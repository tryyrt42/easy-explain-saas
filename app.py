"""
쉬운 문서 해석기 — Easy-Easㅇy 브랜딩 + 랜딩 페이지 개선 버전
- 🎨 유지: Easy-Easy 브랜드 헤더, 좌우 레이아웃 고정, 은은한 수직선
- 🎨 유지: 좌우 패널 세로 길이 대폭 확장 (스크롤 최소화 1200px)
- 🚀 극대화: '📖 원서 독해 & 영단어 학습 모드' 프롬프트 엔진 풀업그레이드
    (구동사/숙어 40~100개 무제한 대량 추출, 소문자 강제, 불필요 평 삭제, 복잡한 구문 분석 추가)
"""
import docx  
import io    
import os
import re
import gc
import tempfile
import hashlib
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
    [data-testid*="ManageMenu"],
    [data-testid*="manageMenu"],
    [class*="manage-app"],
    [class*="ManageApp"],
    [class*="manageMenu"],
    [class*="ManageMenu"],
    [class*="appActions"],
    [class*="AppActions"],
    [class*="floatingMenu"],
    .stStatusWidget,
    [data-testid="stStatusWidget"],
    [data-testid="stBottomBlockContainer"],
    [data-testid="stAppFooter"],
    iframe[title*="Manage"],
    iframe[title*="manage"],
    iframe[src*="share.streamlit.io"],
    button[title*="Manage"],
    button[aria-label*="Manage"],
    a[href*="share.streamlit.io"],
    [aria-label*="Manage app"] {
        display: none !important;
        visibility: hidden !important;
        opacity: 0 !important;
        pointer-events: none !important;
    }
    
    /* JS 주입용 보이지 않는 iframe 자체도 숨김 */
    iframe[srcdoc*="hideManageApp"] {
        display: none !important;
        height: 0 !important;
        width: 0 !important;
    }
    
    /* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
    /* 🛡 1차 방어: 최대 specificity로 정확히 타격                   */
    /* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
    html body button[data-testid="manage-app-button"],
    html body button[class^="_terminalButton_"],
    html body button[class*="_terminalButton_"] {
        display: none !important;
        visibility: hidden !important;
        opacity: 0 !important;
        width: 0 !important;
        height: 0 !important;
        padding: 0 !important;
        margin: 0 !important;
        border: 0 !important;
        position: absolute !important;
        left: -99999px !important;
        pointer-events: none !important;
    }
    
    /* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
    /* ⚠️ Manage app 버튼은 Streamlit Cloud 외부 iframe 소속이라
       앱 코드로 못 막음. 호스팅 이전 시 자동 해결.                */
    /* ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━ */
    
    
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
    
    /* === 📖 TXT/DOCX 원본 가독성 향상 === */
    .reading-area {
        font-family: 'Pretendard', -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;
        font-size: 17px;
        line-height: 1.85;
        color: #e2e8f0;
        letter-spacing: 0.005em;
        padding: 1rem 1.5rem;
    }
    .reading-area p {
        margin: 0 0 1.4em 0;
        text-indent: 0;
    }
    .reading-area p:last-child {
        margin-bottom: 0;
    }
    .reading-area .page-heading {
        font-size: 14px;
        font-weight: 600;
        color: #a78bfa;
        letter-spacing: 0.1em;
        text-transform: uppercase;
        margin: 0 0 1rem 0;
        padding-bottom: 0.5rem;
        border-bottom: 1px solid rgba(168, 85, 247, 0.2);
    }
    .reading-area .page-divider {
        margin: 2.5rem 0;
        border: none;
        border-top: 1px dashed rgba(168, 85, 247, 0.3);
    }
    
    /* === 📚 영단어 학습 모드 — 호버 시 원본 하이라이트 === */
    .vocab-mark {
        transition: background-color 0.18s, color 0.18s, box-shadow 0.18s;
        border-radius: 3px;
        padding: 0 2px;
    }
    .vocab-mark.highlighted {
        background-color: #fde047 !important;
        color: #422006 !important;
        font-weight: 600;
        box-shadow: 0 0 0 1px #fbbf24;
    }
    .vocab-source {
        cursor: default !important;
        transition: background-color 0.15s !important;
        position: relative;
    }
    .vocab-source:hover {
        background-color: rgba(74, 222, 128, 0.18) !important;
    }

    /* === 붙여넣기 textarea: 크기 조절 핸들 제거 (요청) === */
    .stTextArea textarea {
        resize: none !important;
    }

    /* === 알림바(준비완료/안내) 두께 줄이기 === */
    [data-testid="stAlert"] {
        padding-top: 0.55rem !important;
        padding-bottom: 0.55rem !important;
    }

    /* === 해석 모드 카드: 투명 버튼을 카드 위에 겹쳐서 클릭 영역으로 === */
    div[class*="st-key-modebtn"] {
        margin-top: -54px !important;   /* 위 카드와 겹치도록 끌어올림 */
        margin-bottom: 10px !important; /* 다음 카드와 간격 */
    }
    div[class*="st-key-modebtn"] button {
        height: 50px !important;
        background: transparent !important;
        border: none !important;
        color: transparent !important;  /* 버튼 글자 숨김 (카드가 보임) */
        box-shadow: none !important;
    }
    div[class*="st-key-modebtn"] button:hover {
        background: transparent !important;
        color: transparent !important;
    }
    div[class*="st-key-modebtn"] button:focus,
    div[class*="st-key-modebtn"] button:active {
        color: transparent !important;
        box-shadow: none !important;
    }
</style>
""", unsafe_allow_html=True)

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# ⚠️ Manage app 버튼: Streamlit Cloud의 외부 iframe 영역이라 
# 앱 코드(CSS/JS)로 막을 수 없음. Render/Railway 등 다른 호스팅으로
# 이전 시 자동 해결됨. 우리 CSS는 일부 변형 케이스만 보호.
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

# ============================================================
# 🔒 2. Supabase 연결 및 F5 새로고침 방어 로직
# ============================================================
SUPABASE_URL = "https://nufvazmyuvhqkeysfwla.supabase.co"

# 🌐 Streamlit Cloud / Render 양쪽 지원 헬퍼
# - Streamlit Cloud: .streamlit/secrets.toml 에서 자동 로드
# - Render: 환경변수에서 자동 로드
def get_secret(key, default=None):
    import os
    val = os.environ.get(key)
    if val:
        return val
    try:
        return st.secrets[key]
    except Exception:
        return default

SUPABASE_KEY = get_secret("SUPABASE_KEY")
supabase: Client = create_client(SUPABASE_URL, SUPABASE_KEY)
GEMINI_API_KEY = get_secret("GEMINI_API_KEY")
MODEL_NAME = "gemini-3.1-flash-lite" 

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 🔑 Gemini API 키 풀 (라운드로빈 + cooldown 자동 fallback)
# - 단일 키: GEMINI_API_KEY (기존 호환)
# - 다중 키: GEMINI_API_KEYS = "key1,key2,key3" (콤마 구분)
# - 둘 다 있으면 다중 키 우선
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
import threading
import time as _time_mod

def _init_gemini_keys():
    """Secrets에서 키들을 읽어서 풀 구성"""
    multi = get_secret("GEMINI_API_KEYS", "")
    if multi:
        keys = [k.strip() for k in multi.split(",") if k.strip()]
    elif GEMINI_API_KEY:
        keys = [GEMINI_API_KEY]
    else:
        keys = []
    return keys

_GEMINI_KEYS = _init_gemini_keys()
_GEMINI_KEY_IDX = 0                  # round-robin 인덱스
_GEMINI_COOLDOWN = {}                # {key: cooldown_until_timestamp}
_GEMINI_LOCK = threading.Lock()      # 키 선택 + Gemini 호출 보호 (configure는 global state)


def _pick_available_key():
    """다음 사용 가능한 키 선택 (cooldown 회피). 호출자가 lock 보유 상태여야 함.
    반환: 키 문자열 or None (전부 cooldown 중)"""
    global _GEMINI_KEY_IDX
    if not _GEMINI_KEYS:
        return None
    now = _time_mod.time()
    n = len(_GEMINI_KEYS)
    for _ in range(n):
        key = _GEMINI_KEYS[_GEMINI_KEY_IDX]
        _GEMINI_KEY_IDX = (_GEMINI_KEY_IDX + 1) % n
        if _GEMINI_COOLDOWN.get(key, 0) <= now:
            return key
    return None


def _mark_key_cooldown(key, seconds=300):
    """키를 cooldown 풀에 넣음. 기본 5분 (분당 한도 회복 시간 + 여유)"""
    _GEMINI_COOLDOWN[key] = _time_mod.time() + seconds


def call_gemini_with_pool(prompt, model_name=MODEL_NAME, max_output_tokens=16384):
    """키 풀에서 자동 선택해서 Gemini 호출.
    429/quota 에러면 그 키를 cooldown에 넣고 다음 키로 자동 재시도.
    전부 실패하면 RuntimeError."""
    if not _GEMINI_KEYS:
        raise RuntimeError("Gemini API 키가 설정되지 않았습니다. Secrets에 GEMINI_API_KEY 또는 GEMINI_API_KEYS를 등록해주세요.")
    
    n = len(_GEMINI_KEYS)
    last_error = None
    for attempt in range(n):
        # 키 선택 + Gemini 호출을 하나의 critical section으로 묶음
        # (google-generativeai SDK는 genai.configure가 global state라
        #  여러 키 동시 사용 시 race condition 위험. 직렬화로 안전 확보)
        with _GEMINI_LOCK:
            key = _pick_available_key()
            if key is None:
                break  # 전부 cooldown
            try:
                genai.configure(api_key=key)
                model = genai.GenerativeModel(
                    model_name,
                    generation_config=genai.types.GenerationConfig(
                        max_output_tokens=max_output_tokens
                    )
                )
                return model.generate_content(prompt)
            except Exception as e:
                err_str = str(e).lower()
                last_error = e
                # quota / rate limit 류면 cooldown 후 다음 키 시도
                rate_limit_signals = (
                    '429', 'quota', 'rate limit', 'rate-limit',
                    'resource_exhausted', 'resourceexhausted', 'too many requests'
                )
                if any(s in err_str for s in rate_limit_signals):
                    _mark_key_cooldown(key, seconds=300)
                    continue
                # 기타 에러 — 일시적일 수 있으니 다음 키 시도하되, 짧은 cooldown
                _mark_key_cooldown(key, seconds=30)
                continue
    
    if last_error:
        raise last_error
    raise RuntimeError(
        "모든 Gemini 키가 일시적으로 사용 불가합니다 (cooldown 중). 1~5분 후 다시 시도해주세요."
    )


def get_gemini_pool_status():
    """디버그용: 현재 풀 상태 반환 (관리자 페이지에서 보여줄 수 있음)"""
    now = _time_mod.time()
    return [
        {
            "key_preview": k[:10] + "..." + k[-4:] if len(k) > 14 else "***",
            "cooldown_remaining_sec": max(0, int(_GEMINI_COOLDOWN.get(k, 0) - now)),
        }
        for k in _GEMINI_KEYS
    ]

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 💾 외부 영속 캐시 (F5 새로고침에도 살아남음)
# - @st.cache_resource는 컨테이너 라이프타임 동안 같은 dict 인스턴스 반환
# - 세션 state가 날아가도 이 dict는 유지됨
# - 로그아웃 시에만 명시적으로 비움
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
@st.cache_resource
def get_persistent_file_cache():
    """이메일별 파일 데이터 영속 저장소. 컨테이너 재시작 시까지 유지."""
    return {}  # {email: {file_id, file_ext, pdf_path, page_texts, total_pages, interpret_cache}}

def cache_file_for_user(email, **data):
    """파일 데이터를 외부 캐시에 저장"""
    if not email:
        return
    cache = get_persistent_file_cache()
    cache[email] = data

def restore_file_for_user(email):
    """외부 캐시에서 파일 데이터 복구. 성공 시 True"""
    if not email:
        return False
    cache = get_persistent_file_cache()
    if email not in cache:
        return False
    data = cache[email]
    for key, value in data.items():
        st.session_state[key] = value
    return True

def clear_file_for_user(email):
    """외부 캐시에서 특정 사용자 데이터 삭제 (로그아웃 시)"""
    if not email:
        return
    cache = get_persistent_file_cache()
    cache.pop(email, None)


# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 📖 파일 파싱 함수 — Lazy PDF 렌더링 (대용량 PDF 메모리 안전)
# - parse_pdf_lazy: PDF를 디스크 임시파일로 저장 + 텍스트만 추출 (빠름)
# - render_pdf_page_image: 보는 페이지만 즉석 렌더링 (메모리 안전)
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def get_pdf_temp_path(file_id):
    """file_id 기반 결정적 경로 → 같은 파일은 같은 경로 사용"""
    safe_id = hashlib.md5(file_id.encode()).hexdigest()[:16]
    return os.path.join(tempfile.gettempdir(), f"easyeasy_{safe_id}.pdf")


def parse_pdf_lazy(uploaded_file, file_id):
    """업로드 PDF를 메모리 복사 없이 디스크에 저장 + 페이지 수만 확인.
    핵심: uploaded_file.read()(40MB 복사본 생성) 대신
    getbuffer()(복사 없는 memoryview)로 디스크에 바로 씀 → 메모리 절약"""
    pdf_path = get_pdf_temp_path(file_id)
    # getbuffer() = zero-copy. .read()처럼 새 40MB 복사본을 안 만듦
    with open(pdf_path, 'wb') as f:
        f.write(uploaded_file.getbuffer())
    # 새 파일이므로 이전에 캐시된 doc 비우기 (stale 방지)
    get_cached_pdf_doc.clear()
    # 페이지 수 확인 (캐시에 등록됨 → 이후 페이지 넘길 때 재사용)
    doc = get_cached_pdf_doc(pdf_path)
    total_pages = len(doc)
    gc.collect()
    return pdf_path, total_pages


@st.cache_resource(max_entries=1, show_spinner=False)
def get_cached_pdf_doc(pdf_path):
    """열린 PDF 문서를 캐싱 — 페이지 넘길 때마다 재오픈/재인덱싱 방지.
    max_entries=1: 메모리 절약 (한 번에 한 문서만 캐시)"""
    return fitz.open(pdf_path)


def get_pdf_page_text(pdf_path, page_index):
    """PDF 한 페이지의 텍스트만 즉석 추출 (AI 해석용). 캐시된 doc 재사용."""
    if not pdf_path or not os.path.exists(pdf_path):
        return ""
    try:
        doc = get_cached_pdf_doc(pdf_path)  # 재사용 (재오픈 X)
        if 0 <= page_index < len(doc):
            return doc[page_index].get_text()
        return ""
    except Exception:
        return ""


def get_page_text(file_ext, pdf_path, page_texts, page_index):
    """페이지 텍스트 통합 헬퍼 — PDF는 즉석 추출, TXT/DOCX는 메모리에서"""
    if file_ext == "pdf":
        return get_pdf_page_text(pdf_path, page_index)
    if page_texts and 0 <= page_index < len(page_texts):
        return page_texts[page_index] or ""
    return ""


def render_pdf_page_image(pdf_path, page_index, dpi=2.0):
    """PDF 한 페이지만 즉석 렌더링. 캐시된 doc 재사용 + 적응형 해상도.
    화질 우선: 가장 긴 변 ~2200px까지 허용 (선명하게)."""
    if not pdf_path or not os.path.exists(pdf_path):
        return None
    pix = None
    try:
        doc = get_cached_pdf_doc(pdf_path)  # 재사용 (재오픈 X — 인덱싱 안 함)
        if page_index < 0 or page_index >= len(doc):
            return None
        page = doc[page_index]
        # 적응형 해상도: 가장 긴 변이 ~2200px 넘지 않게 (화질 좋게, 메모리 과하지 않게)
        rect = page.rect
        max_dim = max(rect.width, rect.height)
        target_px = 2200
        scale = min(dpi, target_px / max_dim) if max_dim > 0 else dpi
        pix = page.get_pixmap(matrix=fitz.Matrix(scale, scale))
        # JPEG quality 92 — 선명한 화질
        img_bytes = pix.tobytes("jpeg", jpg_quality=92)
        return img_bytes
    except Exception:
        return None
    finally:
        pix = None
        gc.collect()


def get_or_render_page(pdf_path, page_index, dpi=2.0):
    """세션 캐시 활용 (최근 본 4페이지 유지) — 2페이지 모드에서도 재렌더링 최소화"""
    cache_key = "_pdf_page_cache"
    cache = st.session_state.setdefault(cache_key, {})
    full_key = f"{pdf_path}::{page_index}"
    
    if full_key in cache:
        return cache[full_key]
    
    img = render_pdf_page_image(pdf_path, page_index, dpi)
    if img is None:
        return None
    
    # LRU: 4페이지 초과 시 가장 오래된 것 제거
    # (2페이지 모드 = 현재 쌍 2개 + 직전 쌍 일부 캐시 가능)
    if len(cache) >= 4:
        oldest_key = next(iter(cache))
        del cache[oldest_key]
    cache[full_key] = img
    return img


def parse_txt(txt_bytes):
    """TXT 파싱 → page_texts 리스트 반환"""
    try:
        raw_text = txt_bytes.decode('utf-8')
    except Exception:
        raw_text = txt_bytes.decode('cp949', errors='ignore')
    chunk_size = 1500
    if not raw_text.strip():
        return ["(내용이 없습니다)"]
    return [raw_text[i:i+chunk_size] for i in range(0, len(raw_text), chunk_size)]


def parse_docx(docx_bytes):
    """DOCX 파싱 → page_texts 리스트 반환"""
    doc_file = docx.Document(io.BytesIO(docx_bytes))
    raw_text = "\n".join([para.text for para in doc_file.paragraphs])
    chunk_size = 1500
    if not raw_text.strip():
        return ["(내용이 없습니다)"]
    return [raw_text[i:i+chunk_size] for i in range(0, len(raw_text), chunk_size)]


# F5 새로고침 방어: 사용자 + 파일 데이터 모두 복구
if st.session_state.get("user") is None and "logged_in_email" in st.query_params:
    saved_email = st.query_params["logged_in_email"]
    response = supabase.table("users").select("*").eq("email", saved_email).execute()
    if len(response.data) > 0:
        st.session_state["user"] = response.data[0]
        # 🆕 외부 캐시에서 파일 데이터 복구 (있으면)
        restore_file_for_user(saved_email)
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
        <span class="brand-name">Easy-Easy (쉽다쉬워)</span>
    </div>
    """, unsafe_allow_html=True)

    col_left, col_right = st.columns([1, 2])
    
    with col_left:
        st.markdown("<h1 style='font-size: 3.2rem; line-height: 1.2;'>어려운 문서<br>이제 쉽게 읽으세요</h1>", unsafe_allow_html=True)
        st.markdown("<p style='color: #f8fafc; font-size: 1.1rem; margin-top: 1.5rem; margin-bottom: 1.2rem;'>모든 분야의 복잡한 문서, 번역기 돌리며 고생하지 마세요. AI가 핵심만 짚어 가장 이해하기 쉬운 한글로 설명해 드립니다.</p>", unsafe_allow_html=True)
        
        # 🆕 4가지 모드 뱃지 — SVG 엠블럼
        st.markdown("""
        <p style='color: #a78bfa; font-size: 0.85rem; font-weight: 600; letter-spacing: 0.05em; margin-bottom: 0.7rem;'>
            ✨ 4가지 해석 모드 지원
        </p>
        <div style='display: flex; flex-wrap: wrap; gap: 8px; margin-bottom: 2rem;'>
            <span style='background: rgba(99, 102, 241, 0.15); border: 1px solid rgba(99, 102, 241, 0.4); color: #c7d2fe; padding: 6px 12px; border-radius: 20px; font-size: 0.85rem; font-weight: 500; display: inline-flex; align-items: center; gap: 6px;'>
                <svg width="16" height="16" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                    <rect x="3" y="4" width="18" height="16" rx="1.5" fill="none" stroke="#818cf8" stroke-width="1.5"/>
                    <line x1="6.5" y1="9" x2="14" y2="9" stroke="#818cf8" stroke-width="1.5" opacity="0.45" stroke-linecap="round"/>
                    <line x1="6.5" y1="13" x2="17.5" y2="13" stroke="#818cf8" stroke-width="2.5" stroke-linecap="round"/>
                    <line x1="6.5" y1="17" x2="11" y2="17" stroke="#818cf8" stroke-width="1.5" opacity="0.45" stroke-linecap="round"/>
                </svg>
                1타 강사 해설
            </span>
            <span style='background: rgba(168, 85, 247, 0.15); border: 1px solid rgba(168, 85, 247, 0.4); color: #e9d5ff; padding: 6px 12px; border-radius: 20px; font-size: 0.85rem; font-weight: 500; display: inline-flex; align-items: center; gap: 6px;'>
                <svg width="16" height="16" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                    <path d="M3 9 L7 5 L11 9 L7 13 Z" fill="#c084fc" opacity="0.75"/>
                    <circle cx="17" cy="14" r="4" fill="#c084fc"/>
                    <path d="M10.5 11 L13.5 11 M10.5 13 L13.5 13" stroke="#c084fc" stroke-width="1.5" stroke-linecap="round"/>
                </svg>
                비유 모드
            </span>
            <span style='background: rgba(236, 72, 153, 0.15); border: 1px solid rgba(236, 72, 153, 0.4); color: #fbcfe8; padding: 6px 12px; border-radius: 20px; font-size: 0.85rem; font-weight: 500; display: inline-flex; align-items: center; gap: 6px;'>
                <svg width="16" height="16" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                    <circle cx="12" cy="12" r="9" fill="none" stroke="#f472b6" stroke-width="1.5" opacity="0.4"/>
                    <circle cx="12" cy="12" r="6" fill="none" stroke="#f472b6" stroke-width="1.5" opacity="0.65"/>
                    <circle cx="12" cy="12" r="3" fill="none" stroke="#f472b6" stroke-width="1.5"/>
                    <circle cx="12" cy="12" r="1.2" fill="#f472b6"/>
                </svg>
                촌철살인 동네형
            </span>
            <span style='background: rgba(34, 197, 94, 0.15); border: 1px solid rgba(34, 197, 94, 0.4); color: #bbf7d0; padding: 6px 12px; border-radius: 20px; font-size: 0.85rem; font-weight: 500; display: inline-flex; align-items: center; gap: 6px;'>
                <svg width="16" height="16" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg">
                    <path d="M3 5.5 C 3 5.5, 7 4, 12 5.5 C 17 4, 21 5.5, 21 5.5 L 21 19 C 21 19, 17 17.5, 12 19 C 7 17.5, 3 19, 3 19 Z" fill="none" stroke="#4ade80" stroke-width="1.5" stroke-linejoin="round"/>
                    <line x1="12" y1="5.5" x2="12" y2="19" stroke="#4ade80" stroke-width="1.5"/>
                    <text x="7.5" y="14" font-family="Georgia, serif" font-size="6.5" font-weight="600" fill="#4ade80" text-anchor="middle">A</text>
                    <text x="16.5" y="14" font-family="-apple-system, sans-serif" font-size="6.5" font-weight="600" fill="#4ade80" text-anchor="middle">가</text>
                </svg>
                원서 독해 & 영단어
            </span>
        </div>
        """, unsafe_allow_html=True)
        
        with st.container(border=True):
            st.markdown("<h3 style='text-align: center; font-weight: 700;'>문서 해석 시작하기</h3>", unsafe_allow_html=True)
            st.markdown("<div style='margin-top: 10px;'></div>", unsafe_allow_html=True)
            
            # OTP 흐름 상태 관리
            if "otp_sent" not in st.session_state:
                st.session_state["otp_sent"] = False
            if "pending_email" not in st.session_state:
                st.session_state["pending_email"] = ""
            
            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
            # 1단계: 이메일 입력 → 인증 코드 전송
            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
            if not st.session_state["otp_sent"]:
                email_input = st.text_input(
                    "이메일 주소", 
                    placeholder="example@email.com", 
                    label_visibility="collapsed",
                    key="login_email"
                )
                send_btn = st.button(
                    "✨ 인증 코드 받기", 
                    type="primary", 
                    use_container_width=True
                )
                st.markdown(
                    "<p style='font-size: 0.8rem; color: #94a3b8; text-align: center; margin-top: 8px;'>"
                    "입력한 이메일로 인증 코드가 발송됩니다"
                    "</p>", 
                    unsafe_allow_html=True
                )
                
                if send_btn:
                    # 간단한 이메일 형식 체크
                    import re
                    email_pattern = r"^[\w\.\-+]+@[\w\.\-]+\.\w+$"
                    if not email_input or not re.match(email_pattern, email_input.strip()):
                        st.error("⚠️ 올바른 이메일 형식이 아닙니다.")
                    else:
                        try:
                            with st.spinner("📨 인증 코드 발송 중..."):
                                # Supabase Auth로 OTP 발송 (기본 이메일 서비스 사용)
                                supabase.auth.sign_in_with_otp({
                                    "email": email_input.strip(),
                                    "options": {"should_create_user": True}
                                })
                            st.session_state["otp_sent"] = True
                            st.session_state["pending_email"] = email_input.strip()
                            st.rerun()
                        except Exception as e:
                            err_msg = str(e)
                            if "rate" in err_msg.lower() or "limit" in err_msg.lower():
                                st.error("⏱️ 메일 발송 한도에 도달했습니다. **5~10분 후** 다시 시도해주세요.")
                            else:
                                st.error(f"⚠️ 전송 실패: {err_msg}")
            
            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
            # 2단계: OTP 코드 입력 → 검증 → 로그인
            # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
            else:
                st.info(
                    f"**{st.session_state['pending_email']}** 으로 "
                    f"인증 코드를 보냈습니다.\n\n"
                    f"⚠️ 메일 안의 **링크는 누르지 마시고**, "
                    f"**숫자 인증 코드**만 아래에 입력해주세요. *(스팸함도 확인)*"
                )
                otp_input = st.text_input(
                    "인증 코드", 
                    placeholder="00000000",
                    max_chars=8,
                    key="otp_input"
                )
                verify_btn = st.button(
                    "✅ 인증하고 시작하기", 
                    type="primary", 
                    use_container_width=True
                )
                
                # 🔄 재전송 (전체 너비, 보조 액션)
                if st.button("인증 코드 재전송", use_container_width=True, key="resend_otp"):
                    try:
                        with st.spinner("재전송 중..."):
                            supabase.auth.sign_in_with_otp({
                                "email": st.session_state["pending_email"],
                                "options": {"should_create_user": True}
                            })
                        st.success("✅ 재전송 완료. 메일을 다시 확인해주세요.")
                    except Exception as e:
                        err_msg = str(e)
                        if "rate" in err_msg.lower() or "limit" in err_msg.lower():
                            st.error("⏱️ 메일 발송 한도에 도달했습니다. **5~10분 후** 다시 시도해주세요.")
                        else:
                            st.error(f"⚠️ {err_msg}")
                
                # ⬅️ 이메일 다시 입력 (전체 너비, escape 액션)
                if st.button("다른 이메일로 다시 시작", use_container_width=True, key="back_to_email"):
                    st.session_state["otp_sent"] = False
                    st.session_state["pending_email"] = ""
                    st.rerun()
                
                if verify_btn:
                    if not otp_input or not (6 <= len(otp_input.strip()) <= 8):
                        st.error("⚠️ 인증 코드를 정확히 입력해주세요.")
                    else:
                        try:
                            with st.spinner("🔐 인증 중..."):
                                # Supabase Auth OTP 검증
                                auth_response = supabase.auth.verify_otp({
                                    "email": st.session_state["pending_email"],
                                    "token": otp_input.strip(),
                                    "type": "email"
                                })
                            
                            if auth_response and auth_response.user:
                                # 인증 성공 → 우리 users 테이블과 동기화
                                verified_email = st.session_state["pending_email"]
                                response = supabase.table("users").select("*").eq("email", verified_email).execute()
                                
                                if len(response.data) > 0:
                                    st.session_state["user"] = response.data[0]
                                else:
                                    new_user = {"email": verified_email, "plan_type": "FREE", "used_pages": 0}
                                    insert_res = supabase.table("users").insert(new_user).execute()
                                    st.session_state["user"] = insert_res.data[0]
                                
                                # OTP 상태 정리
                                st.session_state["otp_sent"] = False
                                st.session_state["pending_email"] = ""
                                st.query_params["logged_in_email"] = verified_email
                                st.rerun()
                            else:
                                st.error("⚠️ 인증 실패. 코드를 다시 확인해주세요.")
                        except Exception as e:
                            err_msg = str(e)
                            if "expired" in err_msg.lower():
                                st.error("⏱️ 인증 코드가 만료되었습니다. 재전송 해주세요.")
                            elif "invalid" in err_msg.lower():
                                st.error("⚠️ 잘못된 인증 코드입니다.")
                            else:
                                st.error(f"⚠️ 인증 실패: {err_msg}")

    with col_right:
        # 🆕 4가지 모드 스크린샷 수직 배치 (각 섹션마다 SVG 엠블럼)
        modes_preview = [
            (
                "1타 강사 해설 모드", 
                "assets/mode_instructor.webp", 
                "강의실에서 듣는 듯 명쾌한 해설. 개념과 원리를 차근차근 풀어줍니다.",
                """<svg width="28" height="28" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" style="flex-shrink:0;">
                    <rect x="3" y="4" width="18" height="16" rx="1.5" fill="none" stroke="#818cf8" stroke-width="1.5"/>
                    <line x1="6.5" y1="9" x2="14" y2="9" stroke="#818cf8" stroke-width="1.5" opacity="0.45" stroke-linecap="round"/>
                    <line x1="6.5" y1="13" x2="17.5" y2="13" stroke="#818cf8" stroke-width="2.5" stroke-linecap="round"/>
                    <line x1="6.5" y1="17" x2="11" y2="17" stroke="#818cf8" stroke-width="1.5" opacity="0.45" stroke-linecap="round"/>
                </svg>"""
            ),
            (
                "비유 모드", 
                "assets/mode_analogy.webp", 
                "어려운 개념을 일상 비유로 직관적으로 이해. 의사·요리사·운전 등 친숙한 비유 활용.",
                """<svg width="28" height="28" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" style="flex-shrink:0;">
                    <path d="M3 9 L7 5 L11 9 L7 13 Z" fill="#c084fc" opacity="0.75"/>
                    <circle cx="17" cy="14" r="4" fill="#c084fc"/>
                    <path d="M10.5 11 L13.5 11 M10.5 13 L13.5 13" stroke="#c084fc" stroke-width="1.5" stroke-linecap="round"/>
                </svg>"""
            ),
            (
                "촌철살인 동네형 모드", 
                "assets/mode_brother.webp", 
                "동네 형이 풀어주듯 핵심만 콕콕. 군더더기 없이 본질만 짚어줍니다.",
                """<svg width="28" height="28" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" style="flex-shrink:0;">
                    <circle cx="12" cy="12" r="9" fill="none" stroke="#f472b6" stroke-width="1.5" opacity="0.4"/>
                    <circle cx="12" cy="12" r="6" fill="none" stroke="#f472b6" stroke-width="1.5" opacity="0.65"/>
                    <circle cx="12" cy="12" r="3" fill="none" stroke="#f472b6" stroke-width="1.5"/>
                    <circle cx="12" cy="12" r="1.2" fill="#f472b6"/>
                </svg>"""
            ),
            (
                "원서 독해 & 영단어 학습 모드", 
                "assets/mode_english.webp", 
                "원서·논문을 읽으며 핵심 영단어/숙어/구문까지 한 번에 학습.",
                """<svg width="28" height="28" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" style="flex-shrink:0;">
                    <path d="M3 5.5 C 3 5.5, 7 4, 12 5.5 C 17 4, 21 5.5, 21 5.5 L 21 19 C 21 19, 17 17.5, 12 19 C 7 17.5, 3 19, 3 19 Z" fill="none" stroke="#4ade80" stroke-width="1.5" stroke-linejoin="round"/>
                    <line x1="12" y1="5.5" x2="12" y2="19" stroke="#4ade80" stroke-width="1.5"/>
                    <text x="7.5" y="14" font-family="Georgia, serif" font-size="6.5" font-weight="600" fill="#4ade80" text-anchor="middle">A</text>
                    <text x="16.5" y="14" font-family="-apple-system, sans-serif" font-size="6.5" font-weight="600" fill="#4ade80" text-anchor="middle">가</text>
                </svg>"""
            ),
        ]
        
        for idx, (mode_name, img_path, mode_desc, emblem_svg) in enumerate(modes_preview):
            # 🆕 첫 번째 섹션은 좌측 h1과 상단 정렬되도록 margin-top 0
            margin_top = "0" if idx == 0 else "2rem"
            st.markdown(
                f"<div style='display: flex; align-items: center; gap: 10px; margin-top: {margin_top}; margin-bottom: 0.3rem;'>"
                f"{emblem_svg}"
                f"<h3 style='color: #e2e8f0; font-size: 1.3rem; margin: 0;'>{mode_name}</h3>"
                f"</div>"
                f"<p style='color: #94a3b8; font-size: 0.95rem; margin-bottom: 1rem;'>{mode_desc}</p>",
                unsafe_allow_html=True
            )
            try:
                st.image(img_path, use_container_width=True)
            except Exception:
                st.info(f"💡 `{img_path}` 파일이 필요합니다.")
            st.markdown("<div style='margin-bottom: 1.5rem;'></div>", unsafe_allow_html=True)
            
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
        # 🧹 로그아웃 시 모든 세션 데이터 + 외부 캐시 완전 정리
        user_email = st.session_state.get("user", {}).get("email", "")
        
        # 🆕 외부 영속 캐시에서도 삭제 (F5와 구분되는 핵심!)
        clear_file_for_user(user_email)
        
        keys_to_clear = [
            # 사용자
            "user",
            # 업로드된 파일 관련 데이터
            "file_id", "file_ext", "pdf_path", "page_texts", "total_pages",
            "_pdf_page_cache",
            # 해석 캐시
            "interpret_cache",
            # UI 토글 상태
            "include_next_page", "fullscreen_result", "fullscreen_pdf",
            # 모드 선택
            "selected_mode",
            # OTP 인증 상태
            "otp_sent", "pending_email", "otp_input",
            # Streamlit 위젯 키들
            "file_uploader_main", "mode_selector_main", "view_page_input",
            "login_email", "input_mode_main", "pasted_text_main", "paste_committed",
        ]
        for key in keys_to_clear:
            if key in st.session_state:
                del st.session_state[key]
        
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
- ⚠️ **[경고: 임의 요약 절대 금지]** 원문의 길이에 비례하여 한국인 학습자가 모를 수 있는 모든 단어, 구동사(Phrasal Verbs), 전치사구, 고어(Archaic), 비유적 표현, 연어(Collocations)를 **최소 40개에서 100개 사이로 무자비하게 전부 추출**할 것. 단어 추출을 중간에 멈추거나 적당히 요약하면 실패한 결과로 간주함. 사전을 통째로 옮기듯 영혼까지 끌어모아 다 발라낼 것.
- 단순히 쉬운 단어보다는 한국인 학습자가 헷갈려하는 '전치사의 뉘앙스'나 '다의어의 문맥상 쓰임' 등 가려운 곳을 시원하게 긁어줘야 해.
- ⚠️ **[표제어 작성 절대 규칙 — 하이라이트 매칭에 직결되니 반드시 지킬 것]**
  
  🔥 **핵심 원칙**: 표제어는 원문에서 **그대로 복사해서 붙여넣은 것**처럼 적어. 한 글자라도 바꾸면 안 돼. 사전형 변환·표준화·재구성·재배열 일체 금지. 원문에 실제로 등장한 그 글자 시퀀스 그대로.
  
  📋 **체크리스트 (표제어 적기 전에 매번 확인할 것)**:
  1. 이 표제어 그대로 원문에서 Ctrl+F로 검색했을 때 반드시 찾을 수 있어야 한다
  2. 동사·형용사·부사 형태 변환 금지 (carelessly를 careless로 바꾸지 마, smoky를 smoke로 바꾸지 마)
  3. 단수·복수 변환 금지 (couriers를 courier로 바꾸지 마, skirmishes를 skirmish로 바꾸지 마)
  4. 시제 변환 금지 (tucked를 tuck으로 바꾸지 마, closed를 close로 바꾸지 마)
  5. 어순 변환 금지 (원문이 "set the table for"면 그대로, "set for the table"로 재배열 금지)
  6. 양 끝의 마침표·쉼표·콜론만 제외, 단어 내부 하이픈·아포스트로피는 유지
  
  ✅ **올바른 예시 (원문 그대로)**:
  - 원문 "couriers" → 표제어 "couriers" (절대 "courier" X)
  - 원문 "closing in on" → 표제어 "closing in on" (절대 "close in on" X)
  - 원문 "skirmishes" → 표제어 "skirmishes" (절대 "skirmish" X)
  - 원문 "a careless look" → 표제어 "careless look" 또는 "a careless look" (절대 "carelessly look" X — 다른 단어를 섞어넣지 마)
  - 원문 "a smoky light" → 표제어 "smoky light" 또는 "smoky" (절대 "smoke light" X)
  - 원문 "fragile cease-fire" → 표제어 "fragile cease-fire"
  - 원문이 문장 첫 단어 "Quitting" → 표제어 "Quitting"
  
  🆕 **사이에 단어가 끼는 구동사·표현은 `...` 으로 줄여 적기 (권장)**:
  긴 구절을 통째로 적기 부담스러우면 사이를 `...`로 줄임. 단, `...` 양쪽 단어는 원문 그대로:
  - 원문 "stuffed a shirt or two into" → "stuffed... into" (또는 통째로)
  - 원문 "tucked it under" → "tucked... under" (또는 "tucked it under")
  - 원문 "picked the book up" → "picked... up" (또는 통째로)
  
  ⚠️ **금지 패턴 (이렇게 적으면 매칭 실패함)**:
  - ❌ "carelessly look" (원문엔 "careless look" 인데 임의로 부사 만듦)
  - ❌ "smoke light" (원문엔 "smoky light" 인데 명사로 바꿈)
  - ❌ "tuck under" (원문엔 "tucked it under" 인데 원형으로 바꿈)
  - ❌ "stuff a shirt into" (원문엔 "stuffed a shirt or two into" 인데 단축·원형화)
  - ❌ `**following before me**` (표제어를 마크다운 볼드/이탤릭 등으로 감싸는 것 금지: `**`, `*`, `__`, `_`, `` ` ``, `~~` 등 일체 사용 X. 강조하려 하지 말 것. 표제어는 plain text 그대로)

🆕 **[자기 검증 단계 — 표 작성 후 반드시 실행]**
표 작성을 완료했으면, 표를 다시 한 번 훑으면서 각 표제어가 **원문에 토씨 하나 안 틀리고 정확히 존재하는지** 검증해. (`...` 줄임 표제어는 `...` 양쪽 단어가 원문에 있는지 확인.) 만약 원문에서 못 찾는 표제어가 있다면, 그 행을 **원문 그대로의 형태로 수정**해서 다시 작성할 것. 검증을 건너뛰면 사용자가 호버해도 하이라이트가 안 떠서 학습 경험이 망가져.

- 학습자 친화 안내: 첫 컬럼이 변형된 형태라 어색할 수 있는데, 그건 의도된 거야. 호버 시 원문에서 정확히 매칭되도록 한 거니까 그대로 둘 것.
- 표 컬럼: | 표현 (원문 형태) | 품사 | 문맥상 의미 | 뉘앙스 설명 및 원문 활용 |

### 3️⃣ 🧠 구문 분석 & 원어민의 표현법

영어 보는 눈을 키워주는 게 이 섹션 목표. 분량보다 *선별*과 *깊이*가 핵심.

**🎯 문장 선정 (이 섹션의 생명)**
원문에서 학습 가치가 가장 높은 문장 **2~3개만** 신중하게 골라. 그냥 길거나 복잡한 문장은 제외하고, 다음 중 하나에 해당하는 것만 선정:
- 한국인이 직역하면 놓치는 영어 특유의 발상이 담긴 문장
- 같은 의미를 영어와 한국어가 다르게 생각하는 게 드러나는 문장
- "원어민은 이렇게 쓰는구나" 깨달음을 주는 표현이 있는 문장
- 이 한 문장만 익혀도 다른 글 보는 눈이 달라지는 문장

**🚫 절대 금지**
`[S: I] [V: knew] [O: no one] [수식: in the place]` 같은 태그 분해, 5형식·8품사 분류 — 가독성 빵점. 사용 금지.

**📝 각 선정 문장마다 다음 형식 (모든 라벨은 콜론으로 구분)**

**원문:** 따옴표로 문장 그대로 인용 (가공/축약 금지)

**왜 이 문장을 골랐는가:** 한국인이 이 문장에서 흔히 놓치는 영어의 핵심 감각을 1~2줄로 미리 짚어주기.

**친절한 해설:** 1타 강사가 옆에서 풀어주듯 자연스러운 한국어 산문으로. 한국어식 직역과 영어식 발상의 결정적 차이, 원어민이 왜 이 단어·어순·시제·태를 골랐는지, 비슷한 표현으로 바꾸면 어감이 어떻게 달라지는지, 문장이 풍기는 분위기·뉘앙스를 자연스럽게 녹여 써.

> 톤 예시: "한국어로는 '아는 사람이 없었다'고 *없음*을 말하지만, 영어는 'I knew no one'으로 *내가 알았다, 0명을* 이라고 합니다. 동사가 아니라 목적어를 부정하는 거죠. 영어를 영어답게 쓰는 핵심 감각이에요."

**응용 예문:** 같은 패턴을 활용한 새 영어 문장 1~2개 + 한글 해석.

**핵심 통찰:** 이 문장에서 평생 가져갈 영어의 작동 원리를 한 줄로 굵게.

== 해석할 문서 ==
{text}"""
    else:
        # 기술 문서 모드용 지침
        return f"""{PROMPT_TEMPLATES[mode]}

== 구조 지침 (반드시 따를 것) ==

### 1️⃣ 핵심 한 줄
"[페이지 전체를 한 문장으로 꿰뚫는 요약]"

### 2️⃣ 본격 해설 (이 섹션이 핵심 — 절대 짧게 끝내지 말 것)

⚠️ **[분량 경고]** 이 섹션의 목표는 원문 내용을 빠짐없이 풀어내는 거야. 원문에 개념이 5개 나오면 5개 다, 10개 나오면 10개 다 다뤄야 해. 적당히 요약하거나 중요한 것만 골라 쓰면 실패한 결과로 간주함. **원문에 담긴 정보 밀도만큼 해설도 풍부해야 하고, 원문이 길면(예: 2페이지) 해설도 그만큼 길어져야 해.**

원문에 나오는 **각 개념/주제마다** 다음을 빠짐없이 풀어쓸 것:
- **무엇인지**: 정의와 역할
- **왜 필요한지**: 이게 없으면 어떤 문제가 생기는지
- **어떻게 동작하는지**: 메커니즘·원리를, 단계가 있으면 단계별로 차근차근
- **무엇과 연결되는지**: 앞뒤 개념, 다른 도구·단계·flow와의 관계
- **구체화**: 어려운 개념은 일상 비유나 구체적 예시로 한 번 더 풀기

추가로 반드시 챙길 것:
- 명령어 / 파라미터 / 옵션 / 설정값이 등장하면 각각의 역할을 **표**로 정리
- 숫자·수식·조건·임계값이 나오면 그게 *의미하는 바*까지 해석 (단순 나열 금지)
- 그림·다이어그램·그래프가 언급되면 그게 보여주는 핵심을 글로 설명
- 단순 번역 절대 금지 — 항상 "그래서 이게 무슨 뜻이냐면" 수준으로 한 단계 더 풀어쓰기

### 3️⃣ 인사이트
- 이 내용과 관련해 현업에서 자주 마주치는 함정·실수·오해를 **구체적으로**
- 왜 중요한가 — 성능·비용·QoR·수율·타이밍 등 실제 영향과 연결지어 설명
- 이 지식이 실무 어느 국면에서 어떻게 쓰이는지 맥락

### 4️⃣ 한 줄 정리
가장 마지막에 페이지 핵심을 한 문장으로 압축.

== 절대 원칙 ==
- 기술 용어는 영문 그대로 유지 (Fusion Compiler, LVT 등).
- 볼드체 뒤에는 조사 띄어쓰기.
- 첫 줄에 인사말 금지. 바로 본론 진입.
- **내용 대비 해설이 빈약하면 안 됨. 원문의 모든 정보를 독자가 완벽히 이해하도록 깊고 풍부하게 풀어낼 것.**

== 해석할 문서 ==
{text}"""

# ============================================================
# ⚙️ 7. 메인 화면 — 접이식 상단 + 전체화면 토글
# ============================================================

if "fullscreen_result" not in st.session_state:
    st.session_state["fullscreen_result"] = False
if "fullscreen_pdf" not in st.session_state:
    st.session_state["fullscreen_pdf"] = False
if "selected_mode" not in st.session_state:
    st.session_state["selected_mode"] = list(PROMPT_TEMPLATES.keys())[2]  
if "include_next_page" not in st.session_state:
    st.session_state["include_next_page"] = False

mode_keys = list(PROMPT_TEMPLATES.keys())

with st.expander("문서 & 해석 설정", expanded=True):
    top_left, top_right = st.columns(2, gap="large")
    
    with top_left:
        with st.container(border=True, height=290):
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
            
            # 🆕 입력 방식 토글 — 파일 업로드 / 텍스트 붙여넣기
            input_mode = st.radio(
                "입력 방식",
                ["파일 업로드", "텍스트 붙여넣기"],
                horizontal=True,
                label_visibility="collapsed",
                key="input_mode_main"
            )
            
            uploaded_file = None
            pasted_text = None
            
            if input_mode == "파일 업로드":
                uploaded_file = st.file_uploader(
                    "문서 파일 업로드 (PDF, TXT, DOCX)", 
                    type=["pdf", "txt", "docx"],
                    key="file_uploader_main"
                )
            else:
                # text_area — 제출(Ctrl+Enter/포커스아웃) 시 콜백으로 내용 확정 + 입력칸 비우기
                def _commit_paste():
                    txt = st.session_state.get("pasted_text_main", "")
                    if txt and txt.strip() and len(txt.strip()) >= 10:
                        st.session_state["paste_committed"] = txt
                    st.session_state["pasted_text_main"] = ""  # 입력칸 비우기
                
                st.text_area(
                    "문서 텍스트 붙여넣기 (Ctrl+Enter)",
                    height=68,
                    placeholder="여기에 글을 붙여넣고 Ctrl+Enter",
                    key="pasted_text_main",
                    on_change=_commit_paste,
                )
                pasted_text = st.session_state.get("paste_committed", "")
    
    with top_right:
        with st.container(border=True, height=290):
            st.markdown("### 해석 컨트롤러")
            
            # 모드별 엠블럼(SVG) + 색상 + 설명
            mode_meta = {
                mode_keys[0]: (  # 1타 강사
                    "#818cf8",
                    """<svg width="26" height="26" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" style="flex-shrink:0;">
                        <rect x="3" y="4" width="18" height="16" rx="1.5" fill="none" stroke="#818cf8" stroke-width="1.5"/>
                        <line x1="6.5" y1="9" x2="14" y2="9" stroke="#818cf8" stroke-width="1.5" opacity="0.45" stroke-linecap="round"/>
                        <line x1="6.5" y1="13" x2="17.5" y2="13" stroke="#818cf8" stroke-width="2.5" stroke-linecap="round"/>
                        <line x1="6.5" y1="17" x2="11" y2="17" stroke="#818cf8" stroke-width="1.5" opacity="0.45" stroke-linecap="round"/>
                    </svg>""",
                ),
                mode_keys[1]: (  # 비유
                    "#c084fc",
                    """<svg width="26" height="26" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" style="flex-shrink:0;">
                        <path d="M3 9 L7 5 L11 9 L7 13 Z" fill="#c084fc" opacity="0.75"/>
                        <circle cx="17" cy="14" r="4" fill="#c084fc"/>
                        <path d="M10.5 11 L13.5 11 M10.5 13 L13.5 13" stroke="#c084fc" stroke-width="1.5" stroke-linecap="round"/>
                    </svg>""",
                ),
                mode_keys[2]: (  # 촌철살인 동네형
                    "#f472b6",
                    """<svg width="26" height="26" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" style="flex-shrink:0;">
                        <circle cx="12" cy="12" r="9" fill="none" stroke="#f472b6" stroke-width="1.5" opacity="0.4"/>
                        <circle cx="12" cy="12" r="6" fill="none" stroke="#f472b6" stroke-width="1.5" opacity="0.65"/>
                        <circle cx="12" cy="12" r="3" fill="none" stroke="#f472b6" stroke-width="1.5"/>
                        <circle cx="12" cy="12" r="1.2" fill="#f472b6"/>
                    </svg>""",
                ),
                mode_keys[3]: (  # 원서 독해
                    "#4ade80",
                    """<svg width="26" height="26" viewBox="0 0 24 24" xmlns="http://www.w3.org/2000/svg" style="flex-shrink:0;">
                        <path d="M3 5.5 C 3 5.5, 7 4, 12 5.5 C 17 4, 21 5.5, 21 5.5 L 21 19 C 21 19, 17 17.5, 12 19 C 7 17.5, 3 19, 3 19 Z" fill="none" stroke="#4ade80" stroke-width="1.5" stroke-linejoin="round"/>
                        <line x1="12" y1="5.5" x2="12" y2="19" stroke="#4ade80" stroke-width="1.5"/>
                        <text x="7.5" y="14" font-family="Georgia, serif" font-size="6.5" font-weight="600" fill="#4ade80" text-anchor="middle">A</text>
                        <text x="16.5" y="14" font-family="-apple-system, sans-serif" font-size="6.5" font-weight="600" fill="#4ade80" text-anchor="middle">가</text>
                    </svg>""",
                ),
            }
            
            for idx_m, mode in enumerate(mode_keys):
                color, emblem = mode_meta[mode]
                is_sel = (mode == st.session_state["selected_mode"])
                # 이모지 제거한 순수 모드 이름 (엠블럼이 이모지 대체)
                clean_name = mode.split(" ", 1)[1] if " " in mode else mode
                bg = f"{color}22" if is_sel else "transparent"
                border = f"1px solid {color}" if is_sel else "1px solid rgba(255,255,255,0.1)"
                weight = "700" if is_sel else "500"
                # 카드 비주얼
                st.markdown(
                    f"<div style='display:flex; align-items:center; gap:12px; "
                    f"height:26px; padding:12px 16px; border-radius:10px; "
                    f"background:{bg}; border:{border};'>"
                    f"{emblem}"
                    f"<span style='font-size:17px; font-weight:{weight}; color:#f1f5f9;'>{clean_name}</span>"
                    f"</div>",
                    unsafe_allow_html=True
                )
                # 카드 위에 얹는 투명 클릭 버튼 (CSS로 카드와 겹침)
                if st.button(clean_name, key=f"modebtn{idx_m}", use_container_width=True):
                    st.session_state["selected_mode"] = mode
                    st.rerun()
            
            selected_mode = st.session_state["selected_mode"]
    
    # 🆕 입력 소스 통합 — 파일 업로드 또는 붙여넣기 텍스트
    content_id = None
    content_ext = None
    if uploaded_file is not None:
        content_id = f"{uploaded_file.name}_{uploaded_file.size}"
        content_ext = uploaded_file.name.split('.')[-1].lower()
    elif pasted_text and pasted_text.strip() and len(pasted_text.strip()) >= 10:
        # 붙여넣은 텍스트 → 해시로 고유 ID 생성, TXT로 취급
        text_hash = hashlib.md5(pasted_text.encode('utf-8')).hexdigest()[:12]
        content_id = f"paste_{text_hash}_{len(pasted_text)}"
        content_ext = "txt"
    
    if content_id is not None:
        file_id = content_id
        file_ext = content_ext
        
        if st.session_state.get("file_id") != file_id:
            with st.spinner("📖 문서 읽는 중..."):
                if file_ext == "pdf":
                    # 메모리 복사 없이 디스크로 직접 저장 (uploaded_file 그대로 전달)
                    pdf_path, total_pages_loaded = parse_pdf_lazy(uploaded_file, file_id)
                    page_texts = []  # PDF는 텍스트도 lazy — 보는 페이지만 즉석 추출
                elif file_ext == "txt":
                    # 파일 업로드면 read(), 붙여넣기면 인코딩
                    if uploaded_file is not None:
                        txt_bytes = uploaded_file.read()
                    else:
                        txt_bytes = pasted_text.encode('utf-8')
                    page_texts = parse_txt(txt_bytes)
                    pdf_path = None
                    total_pages_loaded = len(page_texts)
                elif file_ext == "docx":
                    docx_bytes = uploaded_file.read()
                    page_texts = parse_docx(docx_bytes)
                    pdf_path = None
                    total_pages_loaded = len(page_texts)
                else:
                    page_texts = ["(지원하지 않는 파일 형식)"]
                    pdf_path = None
                    total_pages_loaded = 1
            
            st.session_state["file_id"] = file_id
            st.session_state["file_ext"] = file_ext
            st.session_state["pdf_path"] = pdf_path
            st.session_state["page_texts"] = page_texts
            st.session_state["total_pages"] = total_pages_loaded
            # 새 파일 업로드 시 이전 페이지 캐시 초기화
            st.session_state["_pdf_page_cache"] = {}
            
            # 🆕 외부 영속 캐시에도 저장 → F5 새로고침 시 복구 가능
            # ⚠️ pdf_path만 저장 (페이지 이미지는 저장 안 함 — 메모리 절약)
            user_email = st.session_state.get("user", {}).get("email", "")
            cache_file_for_user(
                user_email,
                file_id=file_id,
                file_ext=file_ext,
                pdf_path=pdf_path,
                page_texts=page_texts,
                total_pages=total_pages_loaded,
                interpret_cache=st.session_state["interpret_cache"],
            )
        
        total_pages_show = st.session_state.get("total_pages", 1)
        source_label = "텍스트" if content_id.startswith("paste_") else "문서"
        st.success(f"✅ {source_label} 총 {total_pages_show} 페이지 · 준비 완료 (보는 페이지만 즉석 표시)")
        # 🆕 초대형 문서 부드러운 경고 (막지 않고 안내만)
        if total_pages_show > 800:
            st.warning(
                f"⚠️ {total_pages_show}페이지는 매우 큰 문서예요. 페이지 이동이 느리거나 "
                "가끔 불안정할 수 있어요. 특정 챕터만 분리해서 올리면 훨씬 빠르고 정확합니다."
            )
    elif "file_id" in st.session_state:
        # 🆕 F5 후: 업로더는 비어있지만 캐시에 파일 데이터가 있는 경우
        total_pages_show = st.session_state.get("total_pages", 1)
        st.success(f"✅ 총 {total_pages_show} 페이지 · 준비 완료 (이전 세션 복구)")
    else:
        st.info("👆 좌측에 문서를 업로드하거나 텍스트를 붙여넣으면 툴이 시작됩니다.")

if content_id is None and "file_id" not in st.session_state:
    st.stop()

total_pages = st.session_state.get("total_pages", 1)
pdf_path = st.session_state.get("pdf_path")
page_texts = st.session_state.get("page_texts", [])
file_id = st.session_state.get("file_id", "")
file_ext = st.session_state.get("file_ext", "pdf")

# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
# 📚 영단어 학습 모드 — 호버 하이라이트 (서버사이드 처리 + CSS :has())
# ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

def _word_inflect_pattern(word, suffix):
    """단어 1개에 대한 인플렉션 패턴 — silent-e와 y→ies 변환까지 처리.
    예: close → close/closes/closed + closing/closer (e 빠진 형태)
        try → try/tries/tried (y → ies/ied 변환)
    타이포그래픽 변형(따옴표·대시)도 동등하게 매칭."""
    base_pat = _re_escape_with_typo(word) + suffix
    alternatives = [base_pat]
    # silent-e: 자음 + e로 끝나면 → e 빼고 vowel-suffix 시도
    # 🆕 y 추가: smoke→smoky, shine→shiny, juice→juicy 같은 형용사 파생까지 커버
    if len(word) > 2 and word.endswith('e') and word[-2].lower() not in 'aeiou':
        no_e = _re_escape_with_typo(word[:-1])
        alternatives.append(no_e + r'(?:ing|ed|er|est|ies|ied|y)')
    # y → ies/ied: 자음 + y로 끝나면 → y 빼고 ies/ied/ier/iest 시도
    # 예: try → tries/tried, study → studies/studied
    if len(word) > 2 and word.endswith('y') and word[-2].lower() not in 'aeiou':
        no_y = _re_escape_with_typo(word[:-1])
        alternatives.append(no_y + r'(?:ies|ied|ier|iest|ying)')
    if len(alternatives) > 1:
        return '(?:' + '|'.join(alternatives) + ')'
    return alternatives[0]


def build_vocab_search_pattern(term):
    """영단어 매칭용 정규식
    - 일반 접미사: s, es, ed, ing, er, est, ly, ies, ied, n
    - 자음 중복: stop→stopped, run→running 등
    - 🆕 silent-e: close→closing, write→writing
    - 🆕 y→ies: try→tries, study→studied
    - 🆕 분리형 구동사: "pick up" 패턴이 "pick the book up"도 매칭
    - tilde: 'tuck ~ under' 같은 분리 표기 지원"""
    # 분리 가능한 영어 구동사 파티클
    PARTICLES = {
        'up', 'down', 'in', 'out', 'on', 'off', 'away', 'back', 'over',
        'through', 'around', 'apart', 'together', 'forward', 'across',
        'along', 'about', 'by', 'ahead', 'aside', 'behind', 'under',
    }
    suffix = (
        r'(?:s|es|ed|ing|er|est|d|ly|ies|ied|n'
        r'|[bcdfghjklmnpqrstvwxyz](?:ing|ed|er)'
        r')?'
    )
    if '~' in term:
        parts = [p.strip() for p in term.split('~') if p.strip()]
        sub_patterns = []
        for part in parts:
            words = part.split()
            sub_patterns.append(r'\s+'.join(_word_inflect_pattern(w, suffix) for w in words))
        pattern = r'[^\n.!?]{1,50}?'.join(sub_patterns)
    else:
        words = term.split()
        # 🆕 2단어 구동사 (동사 + 파티클) → 분리형도 매칭
        if len(words) == 2 and words[1].lower() in PARTICLES:
            sep = r'(?:\s+[^\s\n.!?]+){0,4}?\s+'
            pattern = (
                _word_inflect_pattern(words[0], suffix) + sep
                + _word_inflect_pattern(words[1], suffix)
            )
        else:
            pattern = r'\s+'.join(_word_inflect_pattern(w, suffix) for w in words)
    return r'\b' + pattern + r'\b'


def _strip_md_emphasis(s):
    """양 끝의 마크다운 강조 마크업 제거 (**bold**, *italic*, __bold__, _italic_, `code`, ~~s~~).
    AI가 표제어를 자기 멋대로 강조 마크업으로 감싸는 케이스 방어용."""
    s = s.strip()
    # 양 끝 동시 마크업을 반복 제거 (***word*** → **word** → word)
    while True:
        stripped = False
        # 긴 마크업 먼저
        for marker in ('**', '__', '~~', '*', '_', '`'):
            if (s.startswith(marker) and s.endswith(marker)
                    and len(s) > 2 * len(marker)):
                s = s[len(marker):-len(marker)].strip()
                stripped = True
                break
        if not stripped:
            break
    return s


def process_response_for_vocab(response_text, source_text=None):
    """마크다운 응답에서 영단어 테이블 찾아 HTML로 변환 + 표제어 리스트 반환.
    🆕 표제어가 곧 원문에 등장한 형태 그대로 → 추측 없이 정확 매칭 가능
    🆕 source_text 주어지면, 원문에서 매칭 안 되는 표제어 자동 제외 (AI 환각 방어)"""
    terms = []
    
    table_pattern = re.compile(
        r'(\|[^\n]+\|\n)'             # 헤더
        r'(\|[\s\-:|]+\|\n)'           # 구분선
        r'((?:\|[^\n]+\|(?:\n|$))+)',  # 본문 행들
        re.MULTILINE
    )
    
    def replace_table(match):
        header = match.group(1)
        body = match.group(3)
        
        header_cells = [c.strip() for c in header.strip().strip('|').split('|')]
        
        body_rows = []
        for line in body.strip().split('\n'):
            if line.strip():
                cells = [c.strip() for c in line.strip().strip('|').split('|')]
                body_rows.append(cells)
        
        if not body_rows:
            return match.group(0)
        
        # 영단어 테이블인지 확인 (첫 칸이 영어인 행이 50% 이상)
        english_first = sum(
            1 for r in body_rows 
            if r and re.search(r'[a-zA-Z]{2,}', r[0])
        )
        if english_first < len(body_rows) * 0.5:
            return match.group(0)
        
        # 단어 추출 — 표제어 자체가 원문 형태
        # 🆕 마크다운 강조 마크업(**..**, *..*, _, ` 등) 자동 제거
        candidates = []  # [(clean_term, row_id), ...]
        for row in body_rows:
            if row and re.search(r'[a-zA-Z]{2,}', row[0]):
                clean = _strip_md_emphasis(row[0])
                if clean:
                    candidates.append((clean, id(row)))
        
        # 🆕 원문에서 매칭 안 되는 표제어 제외 (AI 환각 필터)
        if source_text and candidates:
            candidate_terms = [c[0] for c in candidates]
            annotated = annotate_text_with_vocab(source_text, candidate_terms)
            matched_in_source = set(re.findall(r'data-word="([^"]+)"', annotated))
            matched_in_source = {m.replace('&quot;', '"') for m in matched_in_source}
            survivors = [(t, rid) for t, rid in candidates if t in matched_in_source]
        else:
            survivors = candidates
        
        survivor_row_ids = {rid for _, rid in survivors}
        local_terms = [t for t, _ in survivors]
        terms.extend(local_terms)
        
        # HTML 테이블 빌드 (필터링된 행은 제외)
        html = ['<table>', '<thead><tr>']
        for h in header_cells:
            html.append(f'<th>{h}</th>')
        html.append('</tr></thead><tbody>')
        
        for row in body_rows:
            is_vocab_row = bool(row) and bool(re.search(r'[a-zA-Z]{2,}', row[0]))
            # 영단어 행인데 필터링됐으면 스킵
            if is_vocab_row and id(row) not in survivor_row_ids:
                continue
            html.append('<tr>')
            for i, cell in enumerate(row):
                if i == 0 and is_vocab_row:
                    clean = _strip_md_emphasis(cell)
                    safe = clean.replace('"', '&quot;')
                    html.append(f'<td class="vocab-source" data-word="{safe}">{clean}</td>')
                else:
                    html.append(f'<td>{cell}</td>')
            html.append('</tr>')
        html.append('</tbody></table>')
        
        return '\n\n' + ''.join(html) + '\n\n'
    
    new_text = table_pattern.sub(replace_table, response_text)
    return new_text, terms


def _re_escape_with_typo(s):
    """re.escape의 변형 — 따옴표·아포스트로피·대시의 다양한 형태를 동등하게 취급.
    🆕 하이픈은 다른 하이픈류 / 공백 / 아예 없음으로도 매칭
       (everlasting ↔ ever-lasting ↔ ever lasting, ceasefire ↔ cease-fire 등)
    PDF/ebook에서 흔히 발생: don't (U+2019) vs don't (U+0027) 같은 미스매치 해결."""
    result = []
    for char in s:
        if char in "'\u2019\u2018\u02BC":   # 단일 따옴표 / 아포스트로피 변형 전부
            result.append("['\u2019\u2018\u02BC]")
        elif char in '"\u201C\u201D':         # 이중 따옴표 변형
            result.append('["\u201C\u201D]')
        elif char in '-\u2013\u2014\u2212':   # 하이픈 / en-dash / em-dash / 마이너스
            # 🆕 하이픈이 있어도 / 없어도 / 공백이어도 다 매칭
            result.append('[\\s\\-\u2013\u2014\u2212]*')
        else:
            result.append(re.escape(char))
    return ''.join(result)


def _build_compound_fuzzy_pattern(term):
    """합성어 변형 매칭용 패턴 (3차 fallback).
    표제어에 하이픈/공백이 없을 때 원문에는 있는 경우를 위해, 내부 위치에
    옵셔널 하이픈/공백을 끼울 수 있는 변형들을 OR로 묶음.
    예: 'everlasting' 표제어 → 원문 'ever-lasting' or 'ever lasting' 매칭
    안전장치: 6자+ 단어만, 양쪽 분할이 3자+, 이미 분리 마커 있으면 적용 안 함."""
    # 이미 분리 마커가 있거나 너무 짧으면 스킵
    if any(c in term for c in ' -\u2013\u2014\u2212\u2026~'):
        return None
    if '...' in term or len(term) < 6:
        return None
    variants = [_re_escape_with_typo(term)]
    for i in range(3, len(term) - 2):
        if term[i-1].isalpha() and term[i].isalpha():
            variants.append(
                _re_escape_with_typo(term[:i])
                + '[\\s\\-\u2013\u2014\u2212]+'
                + _re_escape_with_typo(term[i:])
            )
    if len(variants) <= 1:
        return None
    return r'(?<![A-Za-z0-9])(?:' + '|'.join(variants) + r')(?![A-Za-z0-9])'


def annotate_text_with_vocab(text, terms):
    """원본 텍스트에서 영단어를 <span class='vocab-mark'>로 감싸기.
    🆕 표제어 = 원문 형태 그대로 → 단순 literal 매칭 (대소문자 무시)
    🆕 표제어에 '...' / '…' / '~' 가 있으면 그 자리에 임의 단어들 허용
    🆕 literal로 못 잡으면 인플렉션 패턴으로 fallback (AI 오타·사전형 보정)
       예: 표제어 "smoke" + 원문 "smoky" 매칭됨"""
    if not terms:
        return text
    
    gap_re = re.compile(r'\s*(?:\.{3}|…|~)\s*')
    # 인플렉션 fallback용 suffix
    inflect_suffix = (
        r'(?:s|es|ed|ing|er|est|d|ly|ies|ied|n'
        r'|[bcdfghjklmnpqrstvwxyz](?:ing|ed|er)'
        r')?'
    )
    
    def _build_pattern(parts, use_inflection=False):
        """parts: 단어 묶음 리스트 (...로 나눠진 조각들). use_inflection=True면 단어별 인플렉션 적용"""
        if use_inflection:
            sub_patterns = []
            for part in parts:
                words = part.split()
                sub_patterns.append(r'\s+'.join(_word_inflect_pattern(w, inflect_suffix) for w in words))
        else:
            # 타이포그래픽 변형(따옴표·대시) 동등 매칭
            sub_patterns = [_re_escape_with_typo(p) for p in parts]
        if len(parts) >= 2:
            between = r'\s+[^\n.!?]+?\s+'
            inner = between.join(sub_patterns)
        else:
            inner = sub_patterns[0]
        return r'(?<![A-Za-z0-9])' + inner + r'(?![A-Za-z0-9])'
    
    matches = []
    for term in terms:
        if not term:
            continue
        parts = [p.strip() for p in gap_re.split(term) if p.strip()]
        if not parts:
            continue
        try:
            # 1차: literal 정확 매칭
            pat = _build_pattern(parts, use_inflection=False)
            found = list(re.finditer(pat, text, re.IGNORECASE))
            # 2차: 못 찾으면 인플렉션 fallback (AI 오타·사전형 보정)
            if not found:
                pat = _build_pattern(parts, use_inflection=True)
                found = list(re.finditer(pat, text, re.IGNORECASE))
            # 3차: 합성어 fuzzy (everlasting ↔ ever-lasting 같은 변형)
            if not found and len(parts) == 1:
                fuzzy = _build_compound_fuzzy_pattern(parts[0])
                if fuzzy:
                    found = list(re.finditer(fuzzy, text, re.IGNORECASE))
            for m in found:
                matches.append((m.start(), m.end(), term, m.group(0)))
        except re.error:
            continue
    
    if not matches:
        return text
    
    # 겹치는 매치 제거 (앞쪽 + 더 긴 것 우선)
    matches.sort(key=lambda x: (x[0], -(x[1] - x[0])))
    non_overlapping = []
    for m in matches:
        if not non_overlapping or m[0] >= non_overlapping[-1][1]:
            non_overlapping.append(m)
    
    result = []
    last_end = 0
    for start, end, term, matched_text in non_overlapping:
        result.append(text[last_end:start])
        safe_term = term.replace('"', '&quot;')
        result.append(f'<span class="vocab-mark" data-word="{safe_term}">{matched_text}</span>')
        last_end = end
    result.append(text[last_end:])
    return ''.join(result)


def build_vocab_hover_css(terms):
    """CSS :has() 선택자로 양방향 호버 하이라이트 (JS 불필요)"""
    if not terms:
        return ''
    rules = []
    seen = set()
    for term in terms:
        if isinstance(term, dict):
            term = term.get('term', '')
        if not term or term in seen:
            continue
        seen.add(term)
        safe = term.replace('\\', '\\\\').replace('"', '\\"')
        rules.append(
            f'body:has([data-word="{safe}"]:hover) .vocab-mark[data-word="{safe}"] {{ '
            f'background-color: #fde047 !important; color: #422006 !important; '
            f'font-weight: 600; box-shadow: 0 0 0 1px #fbbf24; }}'
        )
        rules.append(
            f'body:has([data-word="{safe}"]:hover) .vocab-source[data-word="{safe}"] {{ '
            f'background-color: rgba(74, 222, 128, 0.22) !important; }}'
        )
    return '<style>' + '\n'.join(rules) + '</style>'


def render_pages_in_container(pages_to_show, pdf_path, page_texts, total_pages, vocab_terms=None, container_height=1200):
    """페이지(들)를 컨테이너 안에 표시.
    PDF=pdf_path에서 즉석 렌더링 (lazy), TXT/DOCX=텍스트 HTML
    vocab_terms: 영단어 학습 모드 시 원문에서 하이라이트할 단어 리스트"""
    with st.container(height=container_height, border=True):
        for idx, pg in enumerate(pages_to_show):
            # 🆕 PDF 모드: 디스크에서 해당 페이지만 즉석 렌더링
            if pdf_path and os.path.exists(pdf_path):
                img_bytes = get_or_render_page(pdf_path, pg - 1)
                if img_bytes:
                    st.image(
                        img_bytes,
                        caption=f"━━━ 페이지 {pg} / {total_pages} ━━━",
                        use_container_width=True
                    )
                else:
                    st.warning(f"⚠️ 페이지 {pg} 렌더링 실패")
            # TXT/DOCX (텍스트만) — HTML로 가독성 향상
            elif page_texts:
                text_content = page_texts[pg - 1] or ""
                # 단락 분리 (\n\n 또는 \n 1개도 단락으로 처리)
                paragraphs = [p.strip() for p in text_content.split('\n\n') if p.strip()]
                if not paragraphs:
                    # \n\n 없으면 \n으로 분리
                    paragraphs = [p.strip() for p in text_content.split('\n') if p.strip()]
                
                # HTML 빌드
                html = f'<div class="reading-area">'
                html += f'<div class="page-heading">페이지 {pg} / {total_pages}</div>'
                for para in paragraphs:
                    # 🆕 영단어 모드: 단어 하이라이트 마킹 먼저
                    if vocab_terms:
                        para = annotate_text_with_vocab(para, vocab_terms)
                    # 단락 안의 줄바꿈은 <br>로
                    para_html = para.replace('\n', '<br>')
                    # HTML 이스케이프 안 하면 < > 가 문제될 수 있지만
                    # 사용자 문서 내용이므로 신뢰 가능
                    html += f'<p>{para_html}</p>'
                html += '</div>'
                
                st.markdown(html, unsafe_allow_html=True)
            
            # 페이지 사이 구분선
            if idx < len(pages_to_show) - 1:
                st.markdown('<hr class="page-divider">', unsafe_allow_html=True)


def run_interpretation(text, mode, cache_key, pages_used=1):
    if not _GEMINI_KEYS:
        st.error("🔑 Secrets 세팅에 GEMINI_API_KEY 또는 GEMINI_API_KEYS를 등록해 주세요.")
        return False
    if not text:
        st.warning("⚠️ 추출 가능한 텍스트가 없습니다.")
        return False
    
    try:
        current_user = st.session_state.get("user", {})
        is_admin = (current_user.get('plan_type') == 'ADMIN')
        
        if not is_admin and current_user.get('plan_type') == 'FREE':
            current_used = current_user.get('used_pages', 0)
            if current_used + pages_used > 3:
                st.error(f"🚫 무료 한도 초과: 현재 {current_used}장 사용 + 이번 {pages_used}장 = 한도 3장 초과")
                return False
        
        spinner_msg = f"🧠 [ADMIN] {pages_used}페이지 분석 중..." if is_admin else f"🧠 {pages_used}페이지 분석 중..."
        with st.spinner(spinner_msg):
            gc.collect()  # 🆕 호출 직전 메모리 정리 (대용량 PDF 캐시와 겹칠 때 OOM 방지)
            # 🆕 키 풀 사용 (자동 fallback)
            response = call_gemini_with_pool(build_prompt(text, mode))
        
        if not is_admin:
            new_used = current_user.get('used_pages', 0) + pages_used
            supabase.table("users").update({"used_pages": new_used}).eq("email", current_user.get('email')).execute()
            st.session_state["user"]['used_pages'] = new_used
        
        st.session_state["interpret_cache"][cache_key] = response.text
        return True
    except Exception as e:
        err_msg = str(e).lower()
        # 사용자 친화적 메시지
        if any(s in err_msg for s in ['quota', 'rate limit', '429', 'resource_exhausted', 'cooldown']):
            st.error("⏰ 일시적으로 사용량이 많습니다. 1~5분 후 다시 시도해주세요.")
        else:
            st.error(f"❌ 오류: {e}")
        return False

# ============================================================
# 🖥 전체화면 모드 OR 분할 보기 모드
# ============================================================
if st.session_state["fullscreen_pdf"]:
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    # 🔍 원본 전체화면 모드 (PDF/TXT/DOCX 풀와이드)
    # ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    fs_pdf_top = st.columns([2, 2, 4, 2])
    
    with fs_pdf_top[0]:
        view_page = st.number_input(
            f"📄 시작 페이지 (총 {total_pages})", 
            min_value=1, max_value=total_pages, value=1, step=1,
            key="view_page_input"
        )
    
    with fs_pdf_top[1]:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        include_next_raw_pdf = st.checkbox(
            "📚 다음 페이지 포함",
            key="include_next_page",
            disabled=(view_page >= total_pages),
        )
    include_next = include_next_raw_pdf and (view_page < total_pages)
    
    with fs_pdf_top[2]:
        st.markdown(
            f"<div style='margin-top: 32px; color: #94a3b8;'>"
            f"📄 <b>{file_ext.upper()} 원본 보기</b>"
            f"</div>", 
            unsafe_allow_html=True
        )
    
    with fs_pdf_top[3]:
        st.markdown("<div style='margin-top: 28px;'></div>", unsafe_allow_html=True)
        if st.button("◀ 분할 보기로", use_container_width=True, key="exit_fullscreen_pdf"):
            st.session_state["fullscreen_pdf"] = False
            st.rerun()
    
    # 풀와이드 원본 렌더링 (가독성 좋게)
    pages_to_show = [view_page]
    if include_next:
        pages_to_show.append(view_page + 1)
    render_pages_in_container(pages_to_show, pdf_path, page_texts, total_pages, container_height=1200)

elif st.session_state["fullscreen_result"]:
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
                        get_page_text(file_ext, pdf_path, page_texts, view_page - 1).strip()
                        + "\n\n--- 다음 페이지 ---\n\n"
                        + get_page_text(file_ext, pdf_path, page_texts, view_page).strip()
                    )
                    pages_used = 2
                else:
                    text = get_page_text(file_ext, pdf_path, page_texts, view_page - 1).strip()
                    pages_used = 1
                
                if run_interpretation(text, selected_mode, cache_key, pages_used=pages_used):
                    st.rerun()
    
    with st.container(height=1200, border=True):
        if is_cached:
            # 🆕 영단어 모드: 테이블 HTML 변환 + 호버 CSS (전체화면에서도 일관된 렌더링)
            response_raw_fs = st.session_state["interpret_cache"][cache_key]
            is_english_mode_fs = "원서 독해" in selected_mode or "영단어" in selected_mode
            if is_english_mode_fs:
                # 🆕 원문 재구성 → AI 환각 표제어 필터링용
                if include_next:
                    src_text_fs = (
                        get_page_text(file_ext, pdf_path, page_texts, view_page - 1).strip()
                        + "\n\n--- 다음 페이지 ---\n\n"
                        + get_page_text(file_ext, pdf_path, page_texts, view_page).strip()
                    )
                else:
                    src_text_fs = get_page_text(file_ext, pdf_path, page_texts, view_page - 1).strip()
                response_processed_fs, vocab_terms_fs = process_response_for_vocab(response_raw_fs, source_text=src_text_fs)
                if vocab_terms_fs:
                    st.markdown(build_vocab_hover_css(vocab_terms_fs), unsafe_allow_html=True)
                st.markdown(response_processed_fs, unsafe_allow_html=True)
            else:
                st.markdown(response_raw_fs)
        else:
            num_pages_label = 2 if include_next else 1
            st.info(f"👆 위의 **[✨ {num_pages_label}페이지 해석]** 버튼을 눌러주세요.")

else:
    # 🟠 주황색 발광 포인트 라인 (기존 회색 divider 대체)
    st.markdown(
        "<div style='height:2px; margin:1.6rem 0; border-radius:2px; "
        "background:linear-gradient(90deg, transparent, #f59e0b 20%, #fb923c 50%, #f59e0b 80%, transparent); "
        "box-shadow:0 0 8px rgba(245,158,11,0.8), 0 0 18px rgba(251,146,60,0.5);'></div>",
        unsafe_allow_html=True
    )
    col_pdf, col_result = st.columns([1, 1], gap="large")
    
    with col_pdf:
        # 🆕 좌측 헤더 + 전체화면 버튼 (우측과 동일 구조로 높이 정렬)
        l_header_col, l_fs_btn_col = st.columns([4, 2])
        with l_header_col:
            st.markdown(f"### {file_ext.upper()} 원본")
        with l_fs_btn_col:
            st.markdown("<div style='margin-top: 4px;'></div>", unsafe_allow_html=True)
            if st.button("🔍 전체화면", use_container_width=True, key="enter_fullscreen_pdf"):
                st.session_state["fullscreen_pdf"] = True
                st.rerun()
        
        # 페이지 선택 + 다음 페이지 포함 체크박스
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
        
        # 페이지 렌더링 (PDF=이미지, TXT/DOCX=가독성 좋은 HTML)
        pages_to_show = [view_page]
        if include_next:
            pages_to_show.append(view_page + 1)
        
        # 🆕 영단어 학습 모드: 캐시된 응답에서 단어 추출 → 원문에 마킹
        pages_suffix = "_2pg" if include_next else ""
        cache_key = f"{file_id}_{view_page}{pages_suffix}_{selected_mode}"
        is_cached = cache_key in st.session_state.get("interpret_cache", {})
        
        vocab_terms = []
        response_processed = ""
        if is_cached:
            response_raw = st.session_state["interpret_cache"][cache_key]
            is_english_mode = "원서 독해" in selected_mode or "영단어" in selected_mode
            if is_english_mode:
                # 🆕 원문 재구성 → AI 환각 표제어 필터링용
                if include_next:
                    src_text = (
                        get_page_text(file_ext, pdf_path, page_texts, view_page - 1).strip()
                        + "\n\n--- 다음 페이지 ---\n\n"
                        + get_page_text(file_ext, pdf_path, page_texts, view_page).strip()
                    )
                else:
                    src_text = get_page_text(file_ext, pdf_path, page_texts, view_page - 1).strip()
                response_processed, vocab_terms = process_response_for_vocab(response_raw, source_text=src_text)
            else:
                response_processed = response_raw
        
        render_pages_in_container(
            pages_to_show, pdf_path, page_texts, total_pages, 
            vocab_terms=vocab_terms, container_height=1200
        )
    
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
                        get_page_text(file_ext, pdf_path, page_texts, view_page - 1).strip()
                        + "\n\n--- 다음 페이지 ---\n\n"
                        + get_page_text(file_ext, pdf_path, page_texts, view_page).strip()
                    )
                    pages_used = 2
                else:
                    text = get_page_text(file_ext, pdf_path, page_texts, view_page - 1).strip()
                    pages_used = 1
                
                if run_interpretation(text, selected_mode, cache_key, pages_used=pages_used):
                    st.rerun()
            
            if is_cached:
                # 🆕 영단어 모드: 동적 CSS로 양방향 호버 (JS 불필요, CSS :has() 사용)
                if vocab_terms:
                    st.markdown(build_vocab_hover_css(vocab_terms), unsafe_allow_html=True)
                st.markdown(response_processed if response_processed else st.session_state["interpret_cache"][cache_key], unsafe_allow_html=True)
            elif not interpret_btn:
                st.info(f"👆 상단의 **[✨ {len(pages_to_show)}페이지 해석]** 버튼을 눌러주세요.")
