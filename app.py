# -*- coding: utf-8 -*-
"""
=======================================================================================
B2G 영업 리드 자동화 대시보드 (app.py) - 첨부파일 파싱 기반 이메일 추출 기능 탑재
=======================================================================================
"""

import io
import os
import re
import time
import tempfile
import zipfile
import urllib.parse
from datetime import datetime, timedelta

import pandas as pd
import requests
import streamlit as st
import urllib3

# HTTPS 보안 경고창 숨김
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ---------------------------------------------------------------------------
# 문서 파싱 라이브러리 로드 (설치 안 되어 있어도 앱이 죽지 않도록 방어 코드)
# ---------------------------------------------------------------------------
try: import pdfplumber
except ImportError: pdfplumber = None
try: import docx
except ImportError: docx = None
try: import olefile
except ImportError: olefile = None

# ---------------------------------------------------------------------------
# 0. 기본 설정 및 상수
# ---------------------------------------------------------------------------
st.set_page_config(page_title="B2G 영업 리드 자동화 대시보드", layout="wide")

SERVC_URL = "https://apis.data.go.kr/1230000/ad/BidPublicInfoService/getBidPblancListInfoServcPPSSrch"
THNG_URL = "https://apis.data.go.kr/1230000/ad/BidPublicInfoService/getBidPblancListInfoThngPPSSrch"
ATTACH_BASE_URL = "https://www.g2b.go.kr:8081"

OFFICER_NAME_KEYS = ["ntceInsttOfclNm", "ofclNm", "chrgOfclNm", "ntceInsttOfclEmplyNm"]
OFFICER_TEL_KEYS = ["ntceInsttOfclTelNo", "ofclTelNo", "chrgOfclTelNo", "ntceInsttOfclTelno"]
OFFICER_EMAIL_KEYS = ["ntceInsttOfclEmailAdres", "ofclEmailAdres", "chrgOfclEmail"] 
DMINSTT_NM_KEYS = ["dminsttNm"]          # <--- ➕ 추가!
NTCE_INSTT_NM_KEYS = ["ntceInsttNm"]     # <--- ➕ 추가!
ATTACH_URL_FIELDS = [f"ntceSpecDocUrl{i}" for i in range(1, 11)]

EMAIL_PATTERN = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")

# ---------------------------------------------------------------------------
# 1. 유틸 함수 (이메일 파싱 포함)
# ---------------------------------------------------------------------------
def pick_field(row: dict, candidates: list, default: str = "") -> str:
    for key in candidates:
        val = row.get(key)
        if val not in (None, "", "null", "NaN"):
            return str(val).strip()
    return default

def build_full_url(url: str) -> str:
    if not url: return ""
    url = str(url).strip()
    if url in ("", "none", "null"): return ""
    if url.startswith("http://") or url.startswith("https://"): return url
    if not url.startswith("/"): url = "/" + url
    return ATTACH_BASE_URL + url

def collect_attachment_urls(row: dict) -> list:
    return [build_full_url(row.get(f)) for f in ATTACH_URL_FIELDS if build_full_url(row.get(f))]

def extract_email_from_url(url: str) -> str:
    """첨부파일 URL을 다운로드하여 텍스트를 읽고 이메일 주소를 정규식으로 추출"""
    if not url: return ""
    try:
        resp = requests.get(url, timeout=10, verify=False)
        if resp.status_code != 200: return ""
        
        ext = url.split("?")[0].split(".")[-1].lower()
        content = resp.content
        text = ""
        
        # 1. PDF 분석 (맨 앞 3장만)
        if ext == "pdf" and pdfplumber:
            with pdfplumber.open(io.BytesIO(content)) as pdf:
                for page in pdf.pages[:3]:
                    t = page.extract_text()
                    if t: text += t + "\n"
                    
        # 2. DOCX 워드 분석
        elif ext == "docx" and docx:
            doc = docx.Document(io.BytesIO(content))
            text = "\n".join([p.text for p in doc.paragraphs[:30]])
            
        # 3. HWP 구형 한글 분석
        elif ext == "hwp" and olefile:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".hwp") as tmp:
                tmp.write(content)
                tmp_path = tmp.name
            if olefile.isOleFile(tmp_path):
                ole = olefile.OleFileIO(tmp_path)
                if ole.exists("PrvText"):
                    text = ole.openstream("PrvText").read().decode("utf-16", errors="ignore")
                ole.close()
            os.remove(tmp_path)
            
        # 4. HWPX 신형 한글 분석
        elif ext == "hwpx":
            with zipfile.ZipFile(io.BytesIO(content)) as zf:
                for name in zf.namelist():
                    if name.startswith("Contents/section") and name.endswith(".xml"):
                        xml_content = zf.read(name).decode("utf-8", errors="ignore")
                        text += re.sub("<[^>]+>", " ", xml_content) + "\n"
                        
        emails = EMAIL_PATTERN.findall(text)
        return ", ".join(sorted(set(emails))) if emails else ""
    except Exception:
        return ""

@st.cache_data(show_spinner=False)
def load_master_csv(file_bytes: bytes) -> pd.DataFrame:
    for enc in ("utf-8-sig", "cp949", "euc-kr", "utf-8"):
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding=enc, dtype=str)
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except: continue
    raise ValueError("CSV 인코딩 판별 불가")

# ---------------------------------------------------------------------------
# 2. 나라장터 API 호출
# ---------------------------------------------------------------------------
def _fetch_one_service(base_url: str, service_key: str, begin_dt: str, end_dt: str,
                        max_pages: int, num_of_rows: int, category_label: str,
                        progress_cb=None) -> list:
    results = []
    page_no = 1
    total_count = None
    raw_key = service_key.strip()

    while page_no <= max_pages:
        req_url = f"{base_url}?serviceKey={raw_key}&pageNo={page_no}&numOfRows={num_of_rows}&inqryDiv=1&type=json&inqryBgnDt={begin_dt}&inqryEndDt={end_dt}"
        try:
            resp = requests.get(req_url, timeout=15, verify=False)
            raw_text = resp.text.strip()
            
            if not raw_text: break
            if raw_text.startswith("<") or "Unexpected errors" in raw_text:
                st.error(f"🚨 서버 인증 거부 ({category_label})")
                break
                
            try: data = resp.json()
            except: break

            body = data.get("response", {}).get("body", {})
            items = body.get("items", [])
            if total_count is None: total_count = int(body.get("totalCount", 0) or 0)
                
        except requests.exceptions.RequestException: break

        if not items: break
        if isinstance(items, dict): items = [items]
        
        for it in items: it["_공고구분"] = category_label
        results.extend(items)

        if progress_cb: progress_cb(category_label, page_no, total_count or 0, len(results))
        if total_count is not None and len(results) >= total_count: break
        if len(items) < num_of_rows: break

        page_no += 1
        time.sleep(0.15)
    return results

def fetch_all_bids(service_key: str, days: int, max_pages: int, num_of_rows: int = 999, progress_cb=None) -> pd.DataFrame:
    end_dt = datetime.now()
    begin_dt = end_dt - timedelta(days=days)
    b_str, e_str = begin_dt.strftime("%Y%m%d") + "0000", end_dt.strftime("%Y%m%d") + "2359"

    all_items = []
    all_items.extend(_fetch_one_service(SERVC_URL, service_key, b_str, e_str, max_pages, num_of_rows, "용역", progress_cb))
    all_items.extend(_fetch_one_service(THNG_URL, service_key, b_str, e_str, max_pages, num_of_rows, "물품", progress_cb))
    return pd.DataFrame(all_items)

# ---------------------------------------------------------------------------
# 3. 데이터프레임 가공 (이메일 심층 추출 포함)
# ---------------------------------------------------------------------------
def build_bid_list_df(raw_df: pd.DataFrame, use_email_parsing: bool, progress_bar) -> pd.DataFrame:
    if raw_df.empty: return pd.DataFrame()
    
    rows = []
    total = len(raw_df)
    
    for i, r in raw_df.iterrows():
        row = r.to_dict()
        dminsttNm = pick_field(row, DMINSTT_NM_KEYS)
        ntceInsttNm = pick_field(row, NTCE_INSTT_NM_KEYS)
        attach_urls = collect_attachment_urls(row)

        # 1차 시도: API가 직접 이메일을 준 경우
        final_email = pick_field(row, OFFICER_EMAIL_KEYS, default="")
        
        # 2차 시도: 스위치를 켰고, API 이메일이 없고, 첨부파일이 있는 경우 (파일 열기!)
        if use_email_parsing and not final_email and attach_urls:
            if progress_bar:
                progress_bar.progress(int((i / total) * 100), text=f"🔍 첨부파일 속 이메일 추출 중... ({i+1}/{total}건)")
            final_email = extract_email_from_url(attach_urls[0])

        rows.append({
            "공고구분": row.get("_공고구분", "-"),
            "수요기관명": dminsttNm or "-",
            "공고기관명": ntceInsttNm or "-",
            "사업명": row.get("bidNtceNm", "-"),
            "공고번호": row.get("bidNtceNo", "-"),
            "공고일자": row.get("bidNtceDt", row.get("bidNtceDate", "-")),
            "담당자명": pick_field(row, OFFICER_NAME_KEYS, default="-"),
            "담당자연락처": pick_field(row, OFFICER_TEL_KEYS, default="-"),
            "담당자이메일": final_email if final_email else "-",
            "첨부파일1": attach_urls[0] if len(attach_urls) > 0 else "",
            "기관명": dminsttNm or ntceInsttNm or "-",
        })
        
    df = pd.DataFrame(rows)
    if "공고번호" in df.columns:
        df = df.drop_duplicates(subset=["공고번호", "공고구분"], keep="first").reset_index(drop=True)
    return df

def to_excel_bytes(sheets: dict) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df in sheets.items():
            df.to_excel(writer, sheet_name=sheet_name[:31], index=False)
    return output.getvalue()

# ---------------------------------------------------------------------------
# 4. Streamlit UI 구성
# ---------------------------------------------------------------------------
st.sidebar.header("⚙️ 기본 설정")
api_key = st.sidebar.text_input("나라장터 API 인증키", type="password")
uploaded_csv = st.sidebar.file_uploader("공공기관 마스터 CSV 업로드", type=["csv"])
days = st.sidebar.slider("최근 N일 조회", min_value=1, max_value=30, value=7)

st.sidebar.markdown("---")
st.sidebar.subheader("🔍 심층 분석 옵션")
use_email_parsing = st.sidebar.checkbox(
    "📧 첨부파일을 열어 이메일 자동 추출", 
    value=False, 
    help="스위치를 켜면 공고에 첨부된 문서를 파이썬이 열어 이메일을 긁어옵니다. (공고당 1~2초가량 수집 시간이 추가됩니다)"
)
if use_email_parsing:
    st.sidebar.warning("⚠️ **주의:** 수집된 공고가 100건이면 약 1~2분이 더 소요됩니다. 여유를 가지고 기다려주세요!")

run_btn = st.sidebar.button("🚀 나라장터 전체 입찰공고 수집 시작", use_container_width=True)

master_df = pd.DataFrame()
if uploaded_csv is not None:
    master_df = load_master_csv(uploaded_csv.getvalue())
    st.session_state["master_df_fallback"] = master_df
elif "master_df_fallback" in st.session_state:
    master_df = st.session_state["master_df_fallback"]

st.title("📡 B2G 영업 리드 자동화 대시보드")
if "bid_list_df" not in st.session_state: st.session_state["bid_list_df"] = pd.DataFrame()

if run_btn:
    if not api_key: st.error("사이드바에 나라장터 API 인증키를 입력해주세요.")
    else:
        progress_bar = st.progress(0, text="1단계: 전국 입찰공고 리스트 수집 중...")
        def progress_cb(category_label, page_no, total_count, collected):
            pct = min(90, int((collected / max(total_count, 1)) * 90)) if total_count else 10
            progress_bar.progress(pct, text=f"[{category_label}] 누적 {collected:,}건 수집 중...")

        # 1. 공고 리스트부터 긁어오기
        raw_df = fetch_all_bids(api_key, days, max_pages=15, num_of_rows=999, progress_cb=progress_cb)
        
        # 2. 이메일 추출 및 표 만들기 (스위치 켜져있으면 여기서 파일 분석 시작)
        progress_bar.progress(90, text="2단계: 데이터 정리 및 첨부파일 이메일 추출 중...")
        bid_list_df = build_bid_list_df(raw_df, use_email_parsing, progress_bar)
        
        progress_bar.progress(100, text="🎉 수집 완벽하게 종료되었습니다!")
        st.session_state["bid_list_df"] = bid_list_df
        time.sleep(1)
        progress_bar.empty()

bid_list_df = st.session_state.get("bid_list_df", pd.DataFrame())

kpi1, kpi2, kpi3 = st.columns(3)
kpi1.metric("공공기관 마스터 기관 수", f"{len(master_df):,}" if not master_df.empty else "0")
kpi2.metric("최근 N일 수집 입찰공고 수", f"{len(bid_list_df):,}")
email_cnt = int((bid_list_df["담당자이메일"] != "-").sum()) if not bid_list_df.empty else 0
kpi3.metric("이메일 정보 확보 건수", f"{email_cnt:,}")

st.markdown("---")
tab1, tab2 = st.tabs(["📋 1. 나라장터 최근 입찰공고 및 이메일 리스트", "🏢 2. 내 공공기관 마스터 현황 전체"])

with tab1:
    if bid_list_df.empty:
        st.info("👈 좌측 스위치와 버튼을 눌러 데이터를 수집해주세요.")
    else:
        c1, c2 = st.columns(2)
        kw = c1.text_input("기관명 검색", key="t1")
        has_email = c2.checkbox("💡 이메일 주소가 추출된 공고만 보기")
        
        view_df = bid_list_df.copy()
        if kw: view_df = view_df[view_df["수요기관명"].str.contains(kw, na=False) | view_df["공고기관명"].str.contains(kw, na=False)]
        if has_email: view_df = view_df[view_df["담당자이메일"] != "-"]
        
        # 보기 좋은 순서로 컬럼 정리
        view_cols = ["공고구분", "수요기관명", "사업명", "공고일자", "담당자명", "담당자연락처", "담당자이메일", "첨부파일1"]
        view_cols = [c for c in view_cols if c in view_df.columns]
        
        st.dataframe(
            view_df[view_cols], use_container_width=True, hide_index=True,
            column_config={"첨부파일1": st.column_config.LinkColumn("제안요청서", display_text="📥 다운로드")}
        )
        st.download_button("⬇️ 엑셀 다운로드 (입찰공고 리스트)", to_excel_bytes({"공고리스트": view_df}), file_name="입찰공고_이메일포함.xlsx")

with tab2:
    if master_df.empty: st.warning("사이드바에 마스터 CSV를 업로드해주세요.")
    else:
        st.dataframe(master_df, use_container_width=True, hide_index=True)
        st.download_button("⬇️ 엑셀 다운로드 (마스터 전체)", to_excel_bytes({"마스터": master_df}), file_name="마스터현황.xlsx")