# -*- coding: utf-8 -*-
"""
=======================================================================================
B2G 영업 리드 자동화 대시보드 (app.py) - 깃허브 클라우드 배포 완결판 
(파일 DNA 판독, OCR 하이브리드, 원클릭 타겟팅, 마스터 풀매칭 탑재)
=======================================================================================
"""

import io
import os
import re
import time
import tempfile
import zipfile
from datetime import datetime, timedelta

import pandas as pd
import requests
import streamlit as st
import urllib3

# HTTPS 보안 경고창 숨김
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# ---------------------------------------------------------------------------
# 문서 파싱 및 OCR 라이브러리 로드 (진단용 상태 변수 추가)
# ---------------------------------------------------------------------------
lib_status = {"pdf": False, "docx": False, "olefile": False, "ocr": False}
try: 
    import pdfplumber
    lib_status["pdf"] = True
except ImportError: pdfplumber = None
try: 
    import docx
    lib_status["docx"] = True
except ImportError: docx = None
try: 
    import olefile
    lib_status["olefile"] = True
except ImportError: olefile = None
try: 
    import pytesseract
    # 깃허브 클라우드(리눅스) 환경에 맞춰 기본 설정 사용
    lib_status["ocr"] = True
except ImportError: pytesseract = None

# ---------------------------------------------------------------------------
# 0. 기본 설정 및 상수
# ---------------------------------------------------------------------------
st.set_page_config(page_title="B2G 영업 리드 자동화 대시보드", layout="wide")

SERVC_URL = "https://apis.data.go.kr/1230000/ad/BidPublicInfoService/getBidPblancListInfoServcPPSSrch"
THNG_URL = "https://apis.data.go.kr/1230000/ad/BidPublicInfoService/getBidPblancListInfoThngPPSSrch"
ATTACH_BASE_URL = "https://www.g2b.go.kr:8081"

OFFICER_NAME_KEYS = ["ntceInsttOfclNm", "ofclNm", "chrgOfclNm", "ntceInsttOfclEmplyNm"]
OFFICER_TEL_KEYS = ["ntceInsttOfclTelNo", "ofclTelNo", "chrgOfclTelNo", "ntceInsttOfclTelno"]
OFFICER_EMAIL_KEYS = ["ntceInsttOfclEmailAdrs", "ntceInsttOfclEmailAdres", "ofclEmailAdres", "chrgOfclEmail"] 
OFFICER_DEPT_KEYS = ["ntceInsttOfclDeptNm", "ofclDeptNm", "chrgDptNm", "chrgOfclDptNm", "chrgOfclDeptNm", "ntceInsttDptNm"]

DMINSTT_NM_KEYS = ["dminsttNm"]
NTCE_INSTT_NM_KEYS = ["ntceInsttNm"]
ATTACH_URL_FIELDS = ["bidSpecificationUrl"] + [f"ntceSpecDocUrl{i}" for i in range(1, 11)]

# 정규식 패턴 세팅
EMAIL_PATTERN = re.compile(r"[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}")
PHONE_PATTERN = re.compile(r"0\d{1,2}-\d{3,4}-\d{4}")
BRN_PATTERN = re.compile(r"\b\d{3}-\d{2}-\d{5}\b")

# 원클릭 제안 타겟팅 판독 키워드
DOC_KEYWORDS = ["사업자등록증", "재무제표", "인감증명서", "주주명부", "법인등기부등본", "신용평가등급", "신용평가서", "기업정보"]
SUBMIT_KEYWORDS = ["우편", "이메일", "e-mail", "메일", "방문", "직접", "사본", "스캔본"]

# ---------------------------------------------------------------------------
# 1. 유틸 함수 (파일 DNA 판독 기반 데이터 추출)
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

def extract_docs_and_contacts(urls: list) -> tuple:
    if not urls: return "", "", "", "-", "첨부파일 없음"
    emails, phones, brns = set(), set(), set()
    needs_oneclick = False
    logs = []
    
    for idx, url in enumerate(urls[:2]):
        try:
            # 대용량 파일 대비 타임아웃 40초 연장
            resp = requests.get(url, timeout=40, verify=False)
            if resp.status_code != 200:
                logs.append(f"파일{idx+1} 접근실패({resp.status_code})")
                continue
            
            content = resp.content
            text = ""
            ext_log = "알수없음"
            
            # 파일 고유의 DNA(Magic Bytes)로 강제 판독 (.do 우회)
            if content.startswith(b'%PDF'):
                ext_log = "PDF"
                if not pdfplumber:
                    logs.append(f"파일{idx+1}(PDF) 미설치 패스")
                    continue
                with pdfplumber.open(io.BytesIO(content)) as pdf:
                    pages_to_scan = pdf.pages[:3] + pdf.pages[-3:] if len(pdf.pages) > 6 else pdf.pages
                    for page in pages_to_scan:
                        t = page.extract_text()
                        if t: text += t + "\n"
                        # 텍스트가 없으면 OCR(한국어+영어) 스캔 가동
                        if pytesseract and (not t or len(t.strip()) < 30):
                            try:
                                img = page.to_image(resolution=150).original
                                text += pytesseract.image_to_string(img, lang='eng+kor') + "\n"
                            except Exception: pass
                                
            elif content.startswith(b'PK'): 
                ext_log = "DOCX/HWPX"
                parsed = False
                if docx:
                    try:
                        doc = docx.Document(io.BytesIO(content))
                        text = "\n".join([p.text for p in doc.paragraphs])
                        parsed = True
                    except Exception: pass
                
                if not parsed:
                    try:
                        with zipfile.ZipFile(io.BytesIO(content)) as zf:
                            for name in zf.namelist():
                                if name.startswith("Contents/section") and name.endswith(".xml"):
                                    xml_content = zf.read(name).decode("utf-8", errors="ignore")
                                    text += re.sub("<[^>]+>", " ", xml_content) + "\n"
                    except Exception: pass

            elif content.startswith(b'\xd0\xcf\x11\xe0'):
                ext_log = "HWP"
                if not olefile:
                    logs.append(f"파일{idx+1}(HWP) 미설치 패스")
                    continue
                with tempfile.NamedTemporaryFile(delete=False, suffix=".hwp") as tmp:
                    tmp.write(content)
                    tmp_path = tmp.name
                if olefile.isOleFile(tmp_path):
                    ole = olefile.OleFileIO(tmp_path)
                    if ole.exists("PrvText"):
                        text = ole.openstream("PrvText").read().decode("utf-16", errors="ignore")
                    ole.close()
                os.remove(tmp_path)
            else:
                logs.append(f"파일{idx+1} 미지원포맷")
                continue
            
            if len(text.strip()) < 10: logs.append(f"파일{idx+1}({ext_log}) 텍스트없음(스캔본)")
            else: logs.append(f"파일{idx+1}({ext_log}) 판독성공")
                
            emails.update(EMAIL_PATTERN.findall(text))
            phones.update(PHONE_PATTERN.findall(text))
            brns.update(BRN_PATTERN.findall(text))
            
            if not needs_oneclick:
                has_doc = any(kw in text for kw in DOC_KEYWORDS)
                has_sub = any(kw in text for kw in SUBMIT_KEYWORDS)
                if has_doc and has_sub: needs_oneclick = True
                    
        except requests.exceptions.Timeout:
            logs.append(f"파일{idx+1} 시간초과(40초)")
        except Exception as e:
            logs.append(f"파일{idx+1} 에러({str(e)[:10]})")
            
    oneclick_flag = "💡 원클릭 제안 필요" if needs_oneclick else "-"
    log_msg = " / ".join(logs) if logs else "스캔 안함"
    return ", ".join(sorted(emails)), ", ".join(sorted(phones)), ", ".join(sorted(brns)), oneclick_flag, log_msg

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
# 2. 나라장터 API 호출 영역
# ---------------------------------------------------------------------------
def _fetch_one_service(base_url: str, service_key: str, begin_dt: str, end_dt: str, max_pages: int, num_of_rows: int, category_label: str, progress_cb=None) -> list:
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
            if raw_text.startswith("<") or "Unexpected errors" in raw_text: break
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
        
        doc_email, doc_tel, doc_brn, doc_oneclick, doc_log = "-", "-", "-", "-", "-"
        if use_email_parsing and attach_urls:
            if progress_bar:
                progress_bar.progress(int((i / total) * 100), text=f"🔍 RFP 문서 DNA 판독 및 스캔 중... ({i+1}/{total}건)")
            extracted_emails, extracted_phones, extracted_brns, oneclick_flag, log_msg = extract_docs_and_contacts(attach_urls)
            if extracted_emails: doc_email = extracted_emails
            if extracted_phones: doc_tel = extracted_phones
            if extracted_brns: doc_brn = extracted_brns
            doc_oneclick = oneclick_flag
            doc_log = log_msg
        elif use_email_parsing and not attach_urls:
            doc_log = "첨부파일 없음"

        rows.append({
            "공고구분": row.get("_공고구분", "-"),
            "수요기관명": dminsttNm or "-",
            "공고기관명": ntceInsttNm or "-",
            "사업명": row.get("bidNtceNm", "-"),
            "원클릭 제안여부": doc_oneclick,
            "공고번호": row.get("bidNtceNo", "-"),
            "공고일자": row.get("bidNtceDt", row.get("bidNtceDate", "-")),
            "[API] 담당자명": pick_field(row, OFFICER_NAME_KEYS, default="-"),
            "[API] 부서명": api_dept,
            "[API] 연락처": api_tel,
            "[API] 이메일": api_email,
            "[문서] 추출_연락처": doc_tel,
            "[문서] 추출_이메일": doc_email,
            "[문서] 사업자번호": doc_brn,
            "[문서] 파싱 로그": doc_log,
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
# 3. Streamlit UI 구성
# ---------------------------------------------------------------------------
st.sidebar.header("⚙️ 기본 설정")
api_key = st.sidebar.text_input("나라장터 API 인증키", type="password")
uploaded_csv = st.sidebar.file_uploader("공공기관 마스터 CSV 업로드", type=["csv"])
days = st.sidebar.slider("최근 N일 조회", min_value=1, max_value=30, value=7)

st.sidebar.markdown("---")
st.sidebar.subheader("🔍 심층 분석 옵션")
use_email_parsing = st.sidebar.checkbox("📧 RFP 실무부서 정보 딥 스캔 (원클릭 제안 분석 포함)", value=False)

st.sidebar.markdown("---")
st.sidebar.subheader("🚨 클라우드 시스템 진단")
st.sidebar.write(f"PDF 분석기 (`pdfplumber`): {'✅ 작동중' if lib_status['pdf'] else '❌ 미설치'}")
st.sidebar.write(f"Word 분석기 (`docx`): {'✅ 작동중' if lib_status['docx'] else '❌ 미설치'}")
st.sidebar.write(f"HWP 분석기 (`olefile`): {'✅ 작동중' if lib_status['olefile'] else '❌ 미설치'}")
st.sidebar.write(f"OCR 스캐너 (`tesseract`): {'✅ 작동중' if lib_status['ocr'] else '❌ 미설치'}")

run_btn = st.sidebar.button("🚀 전국 입찰공고 및 제안 타겟팅 수집 시작", use_container_width=True)

master_df = pd.DataFrame()
if uploaded_csv is not None:
    master_df = load_master_csv(uploaded_csv.getvalue())
    st.session_state["master_df_fallback"] = master_df
elif "master_df_fallback" in st.session_state:
    master_df = st.session_state["master_df_fallback"]

st.title("📡 B2G 영업 리드 자동화 대시보드 (클라우드 최적화)")
if "bid_list_df" not in st.session_state: st.session_state["bid_list_df"] = pd.DataFrame()

if run_btn:
    if not api_key: st.error("사이드바에 나라장터 API 인증키를 입력해주세요.")
    else:
        progress_bar = st.progress(0, text="1단계: 입찰공고 리스트 수집 중...")
        raw_df = fetch_all_bids(api_key, days, max_pages=15, num_of_rows=999, progress_cb=lambda c,p,t,col: progress_bar.progress(min(90, int((col/max(t,1))*90)), text=f"[{c}] 누적 {col:,}건 수집 중..."))
        progress_bar.progress(90, text="2단계: 제안 타겟팅 분석 및 담당자 마스터 매핑 중...")
        st.session_state["bid_list_df"] = build_bid_list_df(raw_df, use_email_parsing, progress_bar)
        progress_bar.progress(100, text="🎉 수집 및 분석 완료!")
        time.sleep(1)
        progress_bar.empty()

bid_list_df = st.session_state.get("bid_list_df", pd.DataFrame())
kpi1, kpi2, kpi3 = st.columns(3)
kpi1.metric("공공기관 마스터 기관 수", f"{len(master_df):,}" if not master_df.empty else "0")
kpi2.metric("최근 N일 수집 입찰공고 수", f"{len(bid_list_df):,}")
oneclick_cnt = int((bid_list_df["원클릭 제안여부"] != "-").sum()) if not bid_list_df.empty and "원클릭 제안여부" in bid_list_df.columns else 0
kpi3.metric("원클릭 제안 타겟 공고 수", f"{oneclick_cnt:,}")

st.markdown("---")
tab1, tab2 = st.tabs(["📋 1. 나라장터 입찰공고 및 RFP 추출 데이터", "🏢 2. 내 공공기관 마스터 담당자 매핑 현황"])

with tab1:
    if bid_list_df.empty: st.info("👈 좌측 스위치와 버튼을 눌러 데이터를 수집해주세요.")
    else:
        c1, c2 = st.columns(2)
        kw = c1.text_input("기관명 검색", key="t1")
        has_oneclick = c2.checkbox("💡 '원클릭 제안 필요' 대상 공고만 보기")
        view_df = bid_list_df.copy()
        if kw: view_df = view_df[view_df["수요기관명"].str.contains(kw, na=False) | view_df["공고기관명"].str.contains(kw, na=False)]
        if has_oneclick: view_df = view_df[view_df["원클릭 제안여부"] != "-"]
        
        view_cols = ["공고구분", "수요기관명", "사업명", "원클릭 제안여부", "공고일자", "[API] 담당자명", "[API] 부서명", "[API] 연락처", "[API] 이메일", "[문서] 추출_연락처", "[문서] 추출_이메일", "[문서] 사업자번호", "[문서] 파싱 로그", "첨부파일1"]
        view_cols = [c for c in view_cols if c in view_df.columns]
        st.dataframe(view_df[view_cols], use_container_width=True, hide_index=True, column_config={"첨부파일1": st.column_config.LinkColumn("제안요청서", display_text="📥 다운로드")})
        st.download_button("⬇️ 엑셀 다운로드 (입찰공고 리스트)", to_excel_bytes({"공고리스트": view_df}), file_name="입찰공고_원클릭타겟팅.xlsx")

with tab2:
    if master_df.empty: st.warning("사이드바에 마스터 CSV를 업로드해주세요.")
    else:
        view_master = master_df.copy()
        if not bid_list_df.empty and "수요기관명" in bid_list_df.columns:
            mapping_cols = {"[API] 담당자명": "최근수집_API담당자", "[API] 부서명": "최근수집_API부서", "[API] 연락처": "최근수집_API연락처", "[API] 이메일": "최근수집_API이메일", "[문서] 추출_연락처": "최근수집_문서연락처", "[문서] 추출_이메일": "최근수집_문서이메일", "[문서] 사업자번호": "최근수집_문서사업자번호"}
            for src_col, target_col in mapping_cols.items():
                if src_col in bid_list_df.columns:
                    valid_data = bid_list_df[bid_list_df[src_col] != "-"]
                    if not valid_data.empty:
                        data_map = valid_data.groupby("수요기관명")[src_col].apply(lambda x: " | ".join(sorted(set(str(v) for v in x if str(v) != "-")))).to_dict()
                        if "기관명" in view_master.columns: view_master[target_col] = view_master["기관명"].map(data_map).fillna("-")
            cols = view_master.columns.tolist()
            if "기관명" in cols:
                base_idx = cols.index("기관명") + 1
                for target_col in reversed(list(mapping_cols.values())):
                    if target_col in cols: cols.insert(base_idx, cols.pop(cols.index(target_col)))
                view_master = view_master[cols]
        st.dataframe(view_master, use_container_width=True, hide_index=True)
        st.download_button("⬇️ 엑셀 다운로드 (마스터 전체)", to_excel_bytes({"마스터 매핑현황": view_master}), file_name="마스터현황_전체담당자매핑.xlsx")
