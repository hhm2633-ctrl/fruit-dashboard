"""
🍎 과일 위탁판매 발주 대시보드
쿠팡 API 기반 실시간 도매 발주 관리 시스템
"""

import streamlit as st
import hashlib
import hmac
import datetime
import json
import requests
import random
from urllib.parse import urlencode

# ══════════════════════════════════════════
#  페이지 설정 (가장 먼저 실행)
# ══════════════════════════════════════════
st.set_page_config(
    page_title="🍎 과일 발주 대시보드",
    page_icon="🍎",
    layout="wide",
    initial_sidebar_state="expanded",
)

# ══════════════════════════════════════════
#  전역 CSS (모바일 최적화)
# ══════════════════════════════════════════
st.markdown(
    """
<style>
/* ── 모바일 반응형 ── */
@media (max-width: 768px) {
    .main .block-container { padding: 0.8rem 0.4rem; }
    .stButton > button { font-size: 13px; padding: 6px 10px; }
    h1 { font-size: 22px !important; }
}

/* ── 공통 카드 ── */
.dash-header {
    background: linear-gradient(135deg, #FF6B35 0%, #F7931E 55%, #FFCD3C 100%);
    padding: 20px 24px;
    border-radius: 16px;
    margin-bottom: 20px;
    color: white;
}
.dash-header h1 { margin: 0; font-size: 26px; }
.dash-header p  { margin: 4px 0 0; opacity: .88; font-size: 14px; }

/* ── 도매사 탭 헤더 ── */
.tab-header {
    background: #f8fafc;
    border-left: 5px solid #FF6B35;
    border-radius: 10px;
    padding: 12px 16px;
    margin-bottom: 14px;
}
.tab-header h4 { margin: 0; color: #FF6B35; font-size: 17px; }
.tab-header span { color: #64748b; font-size: 13px; }

/* ── 이익 색상 ── */
.profit-pos { color: #16a34a; font-weight: 700; }
.profit-neg { color: #dc2626; font-weight: 700; }

/* ── 로그인 박스 ── */
.login-hint {
    text-align: center;
    color: #94a3b8;
    font-size: 12px;
    margin-top: 10px;
}

/* ── 복사 코드블록 ── */
.copy-block {
    background: #f1f5f9;
    border: 1px solid #cbd5e1;
    border-radius: 8px;
    padding: 14px 16px;
    font-family: 'Courier New', monospace;
    font-size: 14px;
    line-height: 2;
    white-space: pre-wrap;
    word-break: break-all;
}

/* ── 사이드바 헤더 ── */
.sb-header {
    background: linear-gradient(135deg, #FF6B35, #FFCD3C);
    padding: 10px 14px;
    border-radius: 10px;
    color: white;
    text-align: center;
    margin-bottom: 14px;
}
.sb-header h3 { margin: 0; font-size: 16px; }

/* Streamlit 기본 여백 숨김 */
#MainMenu { visibility: hidden; }
footer    { visibility: hidden; }
</style>
""",
    unsafe_allow_html=True,
)


# ══════════════════════════════════════════
#  세션 상태 초기화
# ══════════════════════════════════════════
def _init():
    defaults = {
        "logged_in": False,
        "product_master": [],       # 상품 마스터 리스트
        "orders": [],               # 수집된 주문 리스트
        "api_access_key": "",
        "api_secret_key": "",
        "api_vendor_id": "",
        "dialog_order": None,       # 팝업에 표시할 주문
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v

_init()


# ══════════════════════════════════════════
#  로그인 시스템
# ══════════════════════════════════════════
def _get_password() -> str:
    """secrets.toml 우선, 없으면 기본값."""
    try:
        return st.secrets["app_password"]
    except Exception:
        return "fruit2024!"


def login_page():
    """로그인 UI만 렌더링하고 False를 반환."""
    st.markdown(
        """
        <div style="text-align:center; padding: 48px 16px 24px;">
            <div style="font-size:64px;">🍎</div>
            <h2 style="margin:8px 0 4px;">과일 발주 대시보드</h2>
            <p style="color:#64748b;">쿠팡 위탁판매 실시간 발주 관리 시스템</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    c1, c2, c3 = st.columns([1, 1.2, 1])
    with c2:
        with st.container(border=True):
            st.markdown("#### 🔐 로그인")
            pw = st.text_input("비밀번호", type="password", placeholder="비밀번호 입력")
            if st.button("로그인", use_container_width=True, type="primary"):
                if pw == _get_password():
                    st.session_state.logged_in = True
                    st.rerun()
                else:
                    st.error("❌ 비밀번호가 올바르지 않습니다.")
            st.markdown(
                '<p class="login-hint">기본 비밀번호: <code>fruit2024!</code><br>'
                '<small>.streamlit/secrets.toml의 app_password로 변경 가능</small></p>',
                unsafe_allow_html=True,
            )
    return False


# ══════════════════════════════════════════
#  쿠팡 API 유틸
# ══════════════════════════════════════════
def _make_query_string(params: dict) -> str:
    """
    쿠팡 API 서명용 쿼리 스트링 생성.
    - 키 알파벳 오름차순 정렬
    - RFC 3986 percent-encoding (quote, safe='')
    - urlencode 미사용 → '+' 대신 '%20' 보장, 쿠팡 서버와 동일한 인코딩
    """
    from urllib.parse import quote as _quote
    pairs = sorted(params.items(), key=lambda x: x[0])
    return "&".join(f"{_quote(str(k), safe='')}={_quote(str(v), safe='')}" for k, v in pairs)


def _generate_auth_header(method: str, path: str, query_string: str,
                           access_key: str, secret_key: str) -> tuple[str, str]:
    """
    쿠팡 HMAC-SHA256 Authorization 헤더 생성.
    반환: (auth_header, signed_datetime)
    서명 메시지 형식: {datetime}\\n{METHOD}\\n{path}\\n{queryString}
    """
    dt = datetime.datetime.utcnow().strftime("%y%m%dT%H%M%SZ")
    message = f"{dt}\n{method}\n{path}\n{query_string}"
    signature = hmac.new(
        secret_key.encode("utf-8"),
        message.encode("utf-8"),
        hashlib.sha256,
    ).hexdigest()
    auth = (
        f"CEA algorithm=HmacSHA256, access-key={access_key}, "
        f"signed-date={dt}, signature={signature}"
    )
    return auth, dt


def _official_hmac(method: str, full_url: str, access_key: str, secret_key: str) -> tuple[str, str, str]:
    """
    쿠팡 공식 SDK 패턴을 그대로 따른 HMAC 생성.
    full_url 에서 path 와 query_string 을 추출해 서명.
    반환: (auth_header, datetime_str, debug_message)
    """
    stripped = full_url.replace("https://api-gateway.coupang.com", "")
    dt = datetime.datetime.utcnow().strftime("%y%m%dT%H%M%SZ")

    if "?" in stripped:
        path_part, qs = stripped.split("?", 1)
    else:
        path_part, qs = stripped, ""

    message = f"{dt}\n{method}\n{path_part}\n{qs}"
    signature = hmac.new(
        secret_key.encode("utf-8"),
        message.encode("utf-8"),
        hashlib.sha256,
    ).hexdigest()

    auth = (
        f"CEA algorithm=HmacSHA256, access-key={access_key}, "
        f"signed-date={dt}, signature={signature}"
    )
    return auth, dt, message   # message 는 디버그용


def fetch_coupang_orders(access_key: str, secret_key: str, vendor_id: str,
                         debug: bool = False):
    """
    쿠팡 오픈 API – 발주서 리스트 조회 (오늘 날짜 기준).
    debug=True 이면 서명 세부 정보를 st.expander 에 표시.
    """
    BASE   = "https://api-gateway.coupang.com"
    path   = f"/v2/providers/openapi/apis/api/v4/vendors/{vendor_id}/ordersheets"
    method = "GET"

    today = datetime.date.today().isoformat()
    # 쿠팡 API: 쿼리 파라미터 알파벳 오름차순 정렬 필수
    # createdAtFrom < createdAtTo < pageSize < status
    qs = (
        f"createdAtFrom={today}T00:00:00"
        f"&createdAtTo={today}T23:59:59"
        f"&pageSize=50"
        f"&status=ACCEPT"
    )
    full_url = f"{BASE}{path}?{qs}"

    auth, dt, dbg_msg = _official_hmac(method, full_url, access_key, secret_key)

    if debug:
        with st.expander("🔍 API 서명 디버그 정보 (401 발생 시 확인용)", expanded=True):
            st.code(f"[Datetime]  {dt}\n[Method]    {method}\n[Path]      {path}\n[Query]     {qs}", language=None)
            st.markdown("**서명 메시지 (4줄 연결):**")
            st.code(dbg_msg, language=None)
            st.markdown("**요청 URL:**")
            st.code(full_url, language=None)
            st.markdown("**Authorization 헤더:**")
            st.code(auth, language=None)

    headers = {
        "Authorization": auth,
        "Content-Type":  "application/json;charset=UTF-8",
    }

    resp = requests.get(full_url, headers=headers, timeout=10)
    resp.raise_for_status()

    raw_sheets = resp.json().get("data", {}).get("orderSheets", [])
    return _parse_sheets(raw_sheets)


def _parse_sheets(sheets: list) -> list:
    """쿠팡 API 응답 → 내부 주문 딕셔너리 리스트."""
    orders = []
    for sheet in sheets:
        receiver = sheet.get("receiver", {})
        orderer  = sheet.get("orderer", {})
        addr = (receiver.get("addr1", "") + " " + receiver.get("addr2", "")).strip()
        for item in sheet.get("items", []):
            orders.append({
                "order_id":        str(sheet.get("orderId", "")),
                "ordered_at":      sheet.get("orderedAt", ""),
                "option_id":       str(item.get("vendorItemId", "")),   # 쿠팡등록번호
                "product_name":    item.get("vendorItemName", ""),
                "quantity":        int(item.get("shippingCount", 1)),
                "sale_price":      int(item.get("salePrice", 0)),
                "orderer_name":    orderer.get("name", ""),
                "phone":           orderer.get("safeNumber", ""),
                "address":         addr,
                "delivery_message": receiver.get("parcelPrintMessage", ""),
                "status":          sheet.get("status", "ACCEPT"),
            })
    return orders


# ══════════════════════════════════════════
#  테스트 가상 주문 생성
# ══════════════════════════════════════════
_NAMES    = ["김철수", "이영희", "박민준", "최수연", "정도현", "강지은", "윤성호", "임나영"]
_PHONES   = ["010-1234-5678", "010-9876-5432", "010-5555-7777", "010-3333-4444"]
_ADDRS    = [
    "서울특별시 강남구 테헤란로 123 삼성아파트 101동 202호",
    "경기도 성남시 분당구 판교로 456 카카오빌딩 7층",
    "부산광역시 해운대구 해운대로 789 마린시티 30층",
    "인천광역시 연수구 송도과학로 90 아이파크 201동 1502호",
    "대구광역시 수성구 동대구로 100 범어아이파크 B동 304호",
]
_MSGS     = ["부재시 경비실에 맡겨주세요", "문 앞에 놓아주세요", "오전 배송 부탁드립니다", "조심히 다뤄주세요", ""]

_DEFAULT_PRODUCTS = [
    {"option_id": "TEST001", "product_name": "제주 천혜향 5kg",    "cost_price": 18000, "sale_price": 32000, "fee_rate": 10.8, "wholesale_name": "제주과일도매",   "wholesale_url": "https://naver.com"},
    {"option_id": "TEST002", "product_name": "충주 사과 10kg",     "cost_price": 22000, "sale_price": 38000, "fee_rate": 10.8, "wholesale_name": "충주사과마트",   "wholesale_url": ""},
    {"option_id": "TEST003", "product_name": "나주 배 선물세트 3개", "cost_price": 15000, "sale_price": 28000, "fee_rate": 10.8, "wholesale_name": "나주배농원",     "wholesale_url": "https://example.com"},
    {"option_id": "TEST004", "product_name": "샤인머스캣 2kg",      "cost_price": 25000, "sale_price": 45000, "fee_rate": 12.5, "wholesale_name": "제주과일도매",   "wholesale_url": "https://naver.com"},
]


def generate_test_orders() -> list:
    """상품 마스터(또는 기본 상품)를 바탕으로 가상 주문 10건 생성."""
    pool = st.session_state.product_master if st.session_state.product_master else _DEFAULT_PRODUCTS
    orders = []
    for i in range(10):
        prod = random.choice(pool)
        orders.append({
            "order_id":         f"TEST-{i+1:04d}",
            "ordered_at":       datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
            "option_id":        prod["option_id"],
            "product_name":     prod["product_name"],
            "quantity":         random.randint(1, 3),
            "sale_price":       prod.get("sale_price", 30000),
            "orderer_name":     random.choice(_NAMES),
            "phone":            random.choice(_PHONES),
            "address":          random.choice(_ADDRS),
            "delivery_message": random.choice(_MSGS),
            "status":           "ACCEPT",
        })
    return orders


# ══════════════════════════════════════════
#  수익 계산
# ══════════════════════════════════════════
def calc_profit(order: dict, master: dict) -> dict:
    qty        = order["quantity"]
    sale_price = order["sale_price"]
    cost_price = master.get("cost_price", 0)
    fee_rate   = master.get("fee_rate", 10.8) / 100

    revenue = sale_price * qty
    cost    = cost_price * qty
    fee     = revenue * fee_rate
    profit  = revenue - cost - fee
    margin  = (profit / revenue * 100) if revenue > 0 else 0

    return {"revenue": revenue, "cost": cost, "fee": fee, "profit": profit, "margin": margin}


# ══════════════════════════════════════════
#  발주 정보 팝업 (st.dialog)
# ══════════════════════════════════════════
@st.dialog("📋 발주 정보", width="large")
def order_dialog(order: dict, master: dict):
    product_name   = master.get("product_name", order["product_name"])
    wholesale_name = master.get("wholesale_name", "-")
    wholesale_url  = master.get("wholesale_url", "")

    st.markdown(f"### 🛒 {product_name}")
    st.caption(f"도매사: **{wholesale_name}** | 주문 ID: `{order['order_id']}`")

    if wholesale_url:
        st.link_button("🚀 도매 사이트 바로가기", wholesale_url, type="primary")

    st.divider()

    # 필드별 코드블록
    fields = [
        ("👤 주문자 이름",           order["orderer_name"]),
        ("📞 연락처",                order["phone"]),
        ("📍 배송 주소",             order["address"]),
        ("💬 배송 메시지(요청사항)", order["delivery_message"] or "(없음)"),
        ("📦 수량",                  f"{order['quantity']}개"),
    ]
    for label, val in fields:
        st.markdown(f"**{label}**")
        st.code(val, language=None)

    st.divider()
    st.markdown("#### 📋 도매 발주용 통합 복사")
    copy_text = (
        f"주문자: {order['orderer_name']}\n"
        f"연락처: {order['phone']}\n"
        f"주소: {order['address']}\n"
        f"배송메시지: {order['delivery_message'] or '없음'}\n"
        f"상품: {product_name} × {order['quantity']}개"
    )
    st.code(copy_text, language=None)


# ══════════════════════════════════════════
#  사이드바 – 상품 마스터 관리
# ══════════════════════════════════════════
def render_sidebar():
    with st.sidebar:
        st.markdown(
            '<div class="sb-header"><h3>🍊 상품 마스터 관리</h3>'
            "<small>도매사 & 상품 정보</small></div>",
            unsafe_allow_html=True,
        )

        # API 키 설정
        with st.expander("⚙️ 쿠팡 API 설정", expanded=False):
            st.session_state.api_access_key = st.text_input(
                "Access Key", value=st.session_state.api_access_key,
                type="password", key="inp_ak",
            )
            st.session_state.api_secret_key = st.text_input(
                "Secret Key", value=st.session_state.api_secret_key,
                type="password", key="inp_sk",
            )
            st.session_state.api_vendor_id = st.text_input(
                "Vendor ID", value=st.session_state.api_vendor_id, key="inp_vid",
            )

        # 상품 추가 폼
        with st.expander("➕ 상품 추가", expanded=True):
            with st.form("product_form", clear_on_submit=True):
                option_id    = st.text_input("🔑 쿠팡등록번호(옵션ID)*", placeholder="예: 7654321098")
                product_name = st.text_input("📦 제품이름*", placeholder="예: 제주 천혜향 5kg")
                cost_price   = st.number_input("💰 원가(도매가, 원)", min_value=0, value=0, step=500)
                sale_price   = st.number_input("🏷️ 쿠팡 판매가(원)", min_value=0, value=0, step=500)
                fee_rate     = st.number_input("📊 판매 수수료율(%)", min_value=0.0, max_value=100.0, value=10.8, step=0.1)
                wholesale_nm = st.text_input("🏪 도매사 이름*", placeholder="예: 제주과일도매")
                wholesale_url= st.text_input("🔗 도매 주문 페이지 URL", placeholder="https://...")

                if st.form_submit_button("✅ 상품 등록", use_container_width=True, type="primary"):
                    if not option_id or not product_name or not wholesale_nm:
                        st.error("* 표시 필드는 필수입니다.")
                    elif option_id in [p["option_id"] for p in st.session_state.product_master]:
                        st.warning(f"옵션ID '{option_id}'는 이미 등록되어 있습니다.")
                    else:
                        st.session_state.product_master.append({
                            "option_id":     option_id,
                            "product_name":  product_name,
                            "cost_price":    cost_price,
                            "sale_price":    sale_price,
                            "fee_rate":      fee_rate,
                            "wholesale_name":wholesale_nm,
                            "wholesale_url": wholesale_url,
                        })
                        st.success(f"✅ '{product_name}' 등록 완료!")
                        st.rerun()

        # 등록 상품 목록
        count = len(st.session_state.product_master)
        st.markdown(f"#### 📋 등록 상품 ({count}개)")

        if st.session_state.product_master:
            for idx, p in enumerate(st.session_state.product_master):
                with st.container(border=True):
                    c1, c2 = st.columns([5, 1])
                    with c1:
                        st.markdown(f"**{p['product_name']}**")
                        st.caption(f"ID: `{p['option_id']}` | {p['wholesale_name']}")
                        if p["sale_price"] > 0:
                            raw_margin = (p["sale_price"] - p["cost_price"]) / p["sale_price"] * 100 - p["fee_rate"]
                            color = "#16a34a" if raw_margin >= 0 else "#dc2626"
                            st.markdown(
                                f"<small>판매가: ₩{p['sale_price']:,} | "
                                f"마진: <span style='color:{color};font-weight:700'>{raw_margin:.1f}%</span></small>",
                                unsafe_allow_html=True,
                            )
                    with c2:
                        if st.button("🗑️", key=f"del_{idx}", help="삭제"):
                            st.session_state.product_master.pop(idx)
                            st.rerun()
        else:
            st.info("등록된 상품이 없습니다.\n위 폼에서 추가해주세요.")

        # 로그아웃
        st.divider()
        if st.button("🚪 로그아웃", use_container_width=True):
            st.session_state.logged_in = False
            st.session_state.orders = []
            st.rerun()


# ══════════════════════════════════════════
#  메인 대시보드
# ══════════════════════════════════════════
def render_dashboard():
    # ── 헤더 ──
    st.markdown(
        '<div class="dash-header">'
        "<h1>🍎 과일 발주 대시보드</h1>"
        "<p>쿠팡 위탁판매 · 실시간 주문 수집 → 도매사별 발주 관리</p>"
        "</div>",
        unsafe_allow_html=True,
    )

    # ── 액션 버튼 ──
    col_a, col_b, col_c = st.columns([5, 5, 2])

    with col_a:
        if st.button("🔄 쿠팡 API 실시간 동기화", use_container_width=True, type="primary"):
            ak = st.session_state.api_access_key
            sk = st.session_state.api_secret_key
            vid = st.session_state.api_vendor_id
            if not ak or not sk or not vid:
                st.error("❌ 사이드바 → ⚙️ 쿠팡 API 설정에서 Access Key / Secret Key / Vendor ID를 입력하세요.")
            else:
                with st.spinner("쿠팡 API 연결 중…"):
                    try:
                        orders = fetch_coupang_orders(ak, sk, vid, debug=True)
                        st.session_state.orders = orders
                        st.success(f"✅ {len(orders)}건 주문 동기화 완료!")
                        st.rerun()
                    except requests.HTTPError as e:
                        st.error(f"❌ HTTP 오류 {e.response.status_code}: {e.response.text[:300]}")
                    except Exception as e:
                        st.error(f"❌ 오류 발생: {e}")

    with col_b:
        if st.button("🧪 가상 주문 데이터 테스트", use_container_width=True):
            st.session_state.orders = generate_test_orders()
            st.success("✅ 테스트 주문 10건 생성 완료!")
            st.rerun()

    with col_c:
        if st.button("🗑️ 초기화", use_container_width=True):
            st.session_state.orders = []
            st.rerun()

    # ── 주문 없음 안내 ──
    if not st.session_state.orders:
        st.markdown(
            "<br><div style='text-align:center;color:#94a3b8;padding:60px 0'>"
            "<div style='font-size:56px'>📭</div>"
            "<h3>수집된 주문이 없습니다</h3>"
            "<p>위 버튼으로 쿠팡 API를 동기화하거나 테스트 데이터를 생성해보세요.</p>"
            "</div>",
            unsafe_allow_html=True,
        )
        return

    # ── VLOOKUP: 주문 ↔ 상품 마스터 매핑 ──
    master_map = {p["option_id"]: p for p in st.session_state.product_master}

    matched   = []
    unmatched = []
    for order in st.session_state.orders:
        master = master_map.get(order["option_id"])
        if master:
            matched.append({**order, "_master": master})
        else:
            unmatched.append(order)

    # ── 전체 요약 메트릭 ──
    total_orders  = len(st.session_state.orders)
    total_revenue = sum(o["sale_price"] * o["quantity"] for o in st.session_state.orders)
    total_profit  = sum(calc_profit(o, o["_master"])["profit"] for o in matched)
    total_fee     = sum(calc_profit(o, o["_master"])["fee"]    for o in matched)

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("📦 총 주문",   f"{total_orders}건")
    m2.metric("💰 총 매출",   f"₩{total_revenue:,.0f}")
    m3.metric("🏦 수수료 합계", f"₩{total_fee:,.0f}")
    profit_pct = (total_profit / total_revenue * 100) if total_revenue > 0 else 0
    m4.metric("📈 예상 순이익", f"₩{total_profit:,.0f}",
              delta=f"{profit_pct:.1f}%",
              delta_color="normal" if total_profit >= 0 else "inverse")

    st.divider()

    # ── 도매사별 그룹핑 ──
    wholesale_groups: dict[str, dict] = {}
    for order in matched:
        name = order["_master"]["wholesale_name"]
        if name not in wholesale_groups:
            wholesale_groups[name] = {"orders": [], "url": order["_master"]["wholesale_url"]}
        wholesale_groups[name]["orders"].append(order)

    # 탭 이름 목록
    tab_labels = [f"🏪 {n}" for n in wholesale_groups]
    if unmatched:
        tab_labels.append("❓ 미매핑 주문")

    if not tab_labels:
        st.warning(
            "⚠️ 주문이 상품 마스터와 매핑되지 않았습니다.\n"
            "사이드바에서 쿠팡등록번호(옵션ID)를 등록해주세요."
        )
        return

    tabs = st.tabs(tab_labels)

    # ── 도매사 탭 렌더링 ──
    wholesale_names = list(wholesale_groups.keys())

    for tab_obj, label in zip(tabs, tab_labels):
        with tab_obj:
            # 미매핑 탭
            if label.startswith("❓"):
                st.warning(f"⚠️ {len(unmatched)}건의 주문이 상품 마스터와 매핑되지 않았습니다.")
                for o in unmatched:
                    with st.container(border=True):
                        st.markdown(
                            f"**주문 ID:** `{o['order_id']}` | **옵션 ID:** `{o['option_id']}`"
                        )
                        st.markdown(f"**상품명:** {o['product_name']} | 수량: {o['quantity']}개")
                        st.caption("💡 사이드바에서 해당 옵션ID를 등록하면 자동 분류됩니다.")
                continue

            # 일반 도매사 탭
            wholesale_name = label.replace("🏪 ", "")
            group = wholesale_groups[wholesale_name]
            grp_orders = group["orders"]
            grp_url    = group["url"]

            grp_revenue = sum(o["sale_price"] * o["quantity"] for o in grp_orders)
            grp_profit  = sum(calc_profit(o, o["_master"])["profit"] for o in grp_orders)
            pct = (grp_profit / grp_revenue * 100) if grp_revenue > 0 else 0
            profit_cls = "profit-pos" if grp_profit >= 0 else "profit-neg"

            # 탭 헤더 + 도매사 바로가기
            hc1, hc2 = st.columns([3, 1])
            with hc1:
                st.markdown(
                    f'<div class="tab-header">'
                    f"<h4>🏪 {wholesale_name}</h4>"
                    f"<span>주문 {len(grp_orders)}건 &nbsp;|&nbsp; "
                    f"매출 ₩{grp_revenue:,.0f} &nbsp;|&nbsp; "
                    f"예상순이익 <span class='{profit_cls}'>₩{grp_profit:,.0f} ({pct:.1f}%)</span>"
                    f"</span></div>",
                    unsafe_allow_html=True,
                )
            with hc2:
                if grp_url:
                    st.link_button(
                        "🚀 도매 사이트 바로가기",
                        grp_url,
                        use_container_width=True,
                        type="primary",
                    )
                else:
                    st.button("🔗 URL 미등록", disabled=True,
                              use_container_width=True, key=f"no_url_{wholesale_name}")

            # ── 주문 카드 ──
            for o in grp_orders:
                master = o["_master"]
                pf = calc_profit(o, master)

                with st.container(border=True):
                    c1, c2, c3 = st.columns([3, 3, 1])

                    with c1:
                        st.markdown(f"**{master['product_name']}**")
                        st.caption(f"주문 ID: `{o['order_id']}` | 옵션 ID: `{o['option_id']}`")
                        st.caption(f"🕒 {o['ordered_at']}")
                        st.caption(
                            f"👤 {o['orderer_name']} &nbsp; 📞 {o['phone']}"
                        )

                    with c2:
                        mc1, mc2, mc3, mc4 = st.columns(4)
                        mc1.metric("수량",   f"{o['quantity']}개")
                        mc2.metric("매출",   f"₩{pf['revenue']:,.0f}")
                        mc3.metric("원가",   f"₩{pf['cost']:,.0f}")
                        delta_color = "normal" if pf["profit"] >= 0 else "inverse"
                        mc4.metric(
                            "순이익",
                            f"₩{pf['profit']:,.0f}",
                            delta=f"{pf['margin']:.1f}%",
                            delta_color=delta_color,
                        )

                    with c3:
                        st.write("")   # 수직 정렬 보정
                        st.write("")
                        dialog_key = f"dlg_{o['order_id']}_{wholesale_name}"
                        if st.button(
                            "📋 발주 정보\n새창 열기",
                            key=dialog_key,
                            use_container_width=True,
                        ):
                            order_dialog(o, master)


# ══════════════════════════════════════════
#  진입점
# ══════════════════════════════════════════
def main():
    # 로그인 상태가 아니면 로그인 페이지만 표시
    if not st.session_state.logged_in:
        login_page()
        return

    # 로그인 후에만 사이드바와 대시보드 렌더링
    render_sidebar()
    render_dashboard()


if __name__ == "__main__":
    main()
