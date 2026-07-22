# -*- coding: utf-8 -*-
"""
=======================================================================================
B2G 영업 리드 자동화 대시보드 (app.py) - 부서명 & 사업자등록번호 투트랙 스캔 추가
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

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# 문서 파싱 및 OCR 라이브러리 로드
try: import pdfplumber
except ImportError: pdfplumber = None
try: import docx
except ImportError: docx = None
try: import olefile
except ImportError: olefile = None
try: import pytesseract
except ImportError: pytesseract = None

st.set_page_config(page_title="B2G 영업 리드 자동화 대시보드", layout="wide")

SERVC_URL = "https://apis.data.go.kr/1230000/ad/BidPublicInfoService/getBidPblancListInfoServcPPSSrch"
THNG_URL = "https://apis.data.go.kr/1230000/ad/BidPublicInfoService/getBidPblancListInfoThngPPSSrch"
ATTACH_BASE_URL = "https://www.g2b.go.kr:8081"

# API 탐색 키 세팅
OFFICER_NAME_KEYS = ["ntceInsttOfclNm", "ofclNm", "chrgOfclNm", "ntceInsttOfclEmplyNm"]
OFFICER_TEL_KEYS = ["ntceInsttOfclTelNo", "ofclTelNo", "chrgOfclTelNo", "ntceInsttOfclTelno"]
OFFICER_EMAIL_KEYS = ["ntceInsttOfclEmailAdrs", "ntceInsttOfclEmailAdres", "ofclEmailAdres", "chrgOfclEmail"] 
OFFICER_DEPT_KEYS = ["ntceInsttOfclDeptNm", "ofclDeptNm", "chrgDptNm", "chrgOfclDptNm", "chrgOfclDeptNm", "ntceInsttDptNm"]

DMINSTT_NM_KEYS = ["dminsttNm"]
NTCE_INSTT_NM_KEYS = ["ntceInsttNm"]
ATTACH_URL_FIELDS = ["bidSpecificationUrl"] + [f"ntceSpecDocUrl{i}" for i in range(1, 11)]

# 정규식 패턴 (이메일, 연락처, 사업자등록번호)
EMAIL_PATTERN = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
PHONE_PATTERN = re.compile(r"0\d{1,2}-\d{3,4}-\d{4}")
BRN_PATTERN = re.compile(r"\b\d{3}-\d{2}-\d{5}\b")  # 사업자번호 000-00-00000 패턴

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

def extract_contacts_from_urls(urls: list) -> tuple:
    """첨부파일에서 이메일, 전화번호, 사업자번호 3가지를 분리해서 추출"""
    if not urls: return "", "", ""
    emails, phones, brns = set(), set(), set()
    
    for url in urls[:2]:
        try:
            resp = requests.get(url, timeout=15, verify=False)
            if resp.status_code != 200: continue
            
            ext = url.split("?")[0].split(".")[-1].lower()
            content = resp.content
            text = ""
            
            # 1. PDF
            if ext == "pdf" and pdfplumber:
                with pdfplumber.open(io.BytesIO(content)) as pdf:
                    pages_to_scan = pdf.pages[:2] + pdf.pages[-2:] if len(pdf.pages) > 4 else pdf.pages
                    for page in pages_to_scan:
                        t = page.extract_text()
                        if t: text += t + "\n"
                        if pytesseract and (not t or len(t.strip()) < 30):
                            try:
                                img = page.to_image(resolution=150).original
                                text += pytesseract.image_to_string(img, lang='eng+kor') + "\n"
                            except: pass
                            
            # 2. DOCX
            elif ext == "docx" and docx:
                doc = docx.Document(io.BytesIO(content))
                text = "\n".join([p.text for p in doc.paragraphs])
                
            # 3. HWP
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
                
            # 4. HWPX
            elif ext == "hwpx":
                with zipfile.ZipFile(io.BytesIO(content)) as zf:
                    for name in zf.namelist():
                        if name.startswith("Contents/section") and name.endswith(".xml"):
                            xml_content = zf.read(name).decode("utf-8", errors="ignore")
                            text += re.sub("<[^>]+>", " ", xml_content) + "\n"
                            
            emails.update(EMAIL_PATTERN.findall(text))
            phones.update(PHONE_PATTERN.findall(text))
            brns.update(BRN_PATTERN.findall(text))
        except Exception:
            continue
            
    return ", ".join(sorted(emails)), ", ".join(sorted(phones)), ", ".join(sorted(brns))

@st.cache_data(show_spinner=False)
def load_master_csv(file_bytes: bytes) -> pd.DataFrame:
    for enc in ("utf-8-sig", "cp949", "euc-kr", "utf-8"):
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding=enc, dtype=str)
            df.columns = [str(c).strip() for c in df.columns]
            return df
        except: continue
    raise ValueError("CSV 인코딩 판별 불가")

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

def build_bid_list_df(raw_df: pd.DataFrame, use_email_parsing: bool, progress_bar) -> pd.DataFrame:
    if raw_df.empty: return pd.DataFrame()
    rows = []
    total = len(raw_df)
    
    for i, r in raw_df.iterrows():
        row = r.to_dict()
        dminsttNm = pick_field(row, DMINSTT_NM_KEYS)
        ntceInsttNm = pick_field(row, NTCE_INSTT_NM_KEYS)
        attach_urls = collect_attachment_urls(row)
        
        api_tel = pick_field(row, OFFICER_TEL_KEYS, default="-")
        api_email = pick_field(row, OFFICER_EMAIL_KEYS, default="-")
        api_dept = pick_field(row, OFFICER_DEPT_KEYS, default="-")
        
        doc_email, doc_tel, doc_brn = "-", "-", "-"
        if use_email_parsing and attach_urls:
            if progress_bar:
                progress_bar.progress(int((i / total) * 100), text=f"🔍 첨부파일 실무부서 연락처 및 사업자번호 스캔 중... ({i+1}/{total}건)")
            extracted_emails, extracted_phones, extracted_brns = extract_contacts_from_urls(attach_urls)
            if extracted_emails: doc_email = extracted_emails
            if extracted_phones: doc_tel = extracted_phones
            if extracted_brns: doc_brn = extracted_brns

        rows.append({
            "공고구분": row.get("_공고구분", "-"),
            "수요기관명": dminsttNm or "-",
            "공고기관명": ntceInsttNm or "-",
            "사업명": row.get("bidNtceNm", "-"),
            "공고번호": row.get("bidNtceNo", "-"),
            "공고일자": row.get("bidNtceDt", row.get("bidNtceDate", "-")),
            "[API] 담당자명": pick_field(row, OFFICER_NAME_KEYS, default="-"),
            "[API] 부서명": api_dept,
            "[API] 연락처": api_tel,
            "[API] 이메일": api_email,
            "[문서] 추출_연락처": doc_tel,
            "[문서] 추출_이메일": doc_email,
            "[문서] 사업자번호": doc_brn,
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

st.sidebar.header("⚙️ 기본 설정")
api_key = st.sidebar.text_input("나라장터 API 인증키", type="password")
uploaded_csv = st.sidebar.file_uploader("공공기관 마스터 CSV 업로드", type=["csv"])
days = st.sidebar.slider("최근 N일 조회", min_value=1, max_value=30, value=7)

st.sidebar.markdown("---")
st.sidebar.subheader("🔍 심층 분석 옵션")
use_email_parsing = st.sidebar.checkbox(
    "📧 제안요청서(RFP) 실무부서 연락처 & 사업자번호 강제 추출", 
    value=False, 
    help="스위치를 켜면 첨부문서를 열어 실제 사업부서의 연락처/이메일과 기관의 사업자등록번호를 모두 수집합니다."
)

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
        progress_bar = st.progress(0, text="1단계: 입찰공고 리스트 수집 중...")
        def progress_cb(category_label, page_no, total_count, collected):
            pct = min(90, int((collected / max(total_count, 1)) * 90)) if total_count else 10
            progress_bar.progress(pct, text=f"[{category_label}] 누적 {collected:,}건 수집 중...")

        raw_df = fetch_all_bids(api_key, days, max_pages=15, num_of_rows=999, progress_cb=progress_cb)
        
        progress_bar.progress(90, text="2단계: 부서명/연락처/사업자번호 투트랙 추출 중...")
        bid_list_df = build_bid_list_df(raw_df, use_email_parsing, progress_bar)
        
        progress_bar.progress(100, text="🎉 수집 및 분석 완료!")
        st.session_state["bid_list_df"] = bid_list_df
        time.sleep(1)
        progress_bar.empty()

bid_list_df = st.session_state.get("bid_list_df", pd.DataFrame())

kpi1, kpi2, kpi3 = st.columns(3)
kpi1.metric("공공기관 마스터 기관 수", f"{len(master_df):,}" if not master_df.empty else "0")
kpi2.metric("최근 N일 수집 입찰공고 수", f"{len(bid_list_df):,}")
email_cnt = int(((bid_list_df.get("[API] 이메일", pd.Series()) != "-") | (bid_list_df.get("[문서] 추출_이메일", pd.Series()) != "-")).sum()) if not bid_list_df.empty else 0
kpi3.metric("이메일 정보 확보 건수", f"{email_cnt:,}")

st.markdown("---")
tab1, tab2 = st.tabs(["📋 1. 나라장터 최근 입찰공고 및 추출 데이터", "🏢 2. 내 공공기관 마스터 현황 전체"])

with tab1:
    if bid_list_df.empty:
        st.info("👈 좌측 스위치와 버튼을 눌러 데이터를 수집해주세요.")
    else:
        c1, c2 = st.columns(2)
        kw = c1.text_input("기관명 검색", key="t1")
        has_email = c2.checkbox("💡 문서 혹은 API에서 이메일이 하나라도 추출된 공고만 보기")
        
        view_df = bid_list_df.copy()
        if kw: view_df = view_df[view_df["수요기관명"].str.contains(kw, na=False) | view_df["공고기관명"].str.contains(kw, na=False)]
        if has_email: 
            view_df = view_df[(view_df["[API] 이메일"] != "-") | (view_df["[문서] 추출_이메일"] != "-")]
        
        # 보기 좋은 순서로 컬럼 정리 (부서명과 사업자번호 추가)
        view_cols = [
            "공고구분", "수요기관명", "사업명", "공고일자", 
            "[API] 담당자명", "[API] 부서명", "[API] 연락처", "[API] 이메일", 
            "[문서] 추출_연락처", "[문서] 추출_이메일", "[문서] 사업자번호", "첨부파일1"
        ]
        view_cols = [c for c in view_cols if c in view_df.columns]
        
        st.dataframe(
            view_df[view_cols], use_container_width=True, hide_index=True,
            column_config={"첨부파일1": st.column_config.LinkColumn("제안요청서", display_text="📥 다운로드")}
        )
        st.download_button("⬇️ 엑셀 다운로드 (입찰공고 리스트)", to_excel_bytes({"공고리스트": view_df}), file_name="입찰공고_데이터확장판.xlsx")

with tab2:
    if master_df.empty: st.warning("사이드바에 마스터 CSV를 업로드해주세요.")
    else:
        view_master = master_df.copy()
        
        if not bid_list_df.empty and "수요기관명" in bid_list_df.columns:
            # 1. 문서에서 추출한 이메일 매핑
            valid_emails = bid_list_df[bid_list_df["[문서] 추출_이메일"] != "-"]
            if not valid_emails.empty:
                email_map = valid_emails.groupby("수요기관명")["[문서] 추출_이메일"].apply(lambda x: ", ".join(set(x))).to_dict()
                if "기관명" in view_master.columns:
                    view_master["최근수집_담당자이메일(문서)"] = view_master["기관명"].map(email_map).fillna("-")
                    cols = view_master.columns.tolist()
                    cols.insert(cols.index("기관명") + 1, cols.pop(cols.index("최근수집_담당자이메일(문서)")))
                    view_master = view_master[cols]
                    
            # 2. 문서에서 추출한 사업자번호 매핑
            valid_brns = bid_list_df[bid_list_df["[문서] 사업자번호"] != "-"]
            if not valid_brns.empty:
                brn_map = valid_brns.groupby("수요기관명")["[문서] 사업자번호"].apply(lambda x: ", ".join(set(x))).to_dict()
                if "기관명" in view_master.columns:
                    view_master["최근수집_사업자번호(문서)"] = view_master["기관명"].map(brn_map).fillna("-")
                    cols = view_master.columns.tolist()
                    # 이메일 뒤에 사업자번호 열 배치
                    insert_idx = cols.index("최근수집_담당자이메일(문서)") + 1 if "최근수집_담당자이메일(문서)" in cols else cols.index("기관명") + 1
                    cols.insert(insert_idx, cols.pop(cols.index("최근수집_사업자번호(문서)")))
                    view_master = view_master[cols]

        st.dataframe(view_master, use_container_width=True, hide_index=True)
        st.download_button("⬇️ 엑셀 다운로드 (마스터 전체)", to_excel_bytes({"마스터": view_master}), file_name="마스터현황_사업자번호매칭.xlsx")