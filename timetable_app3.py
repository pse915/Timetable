# -*- coding: utf-8 -*-
"""
서라벌여중 시간표·결보강 관리 프로그램
2026 최종 완성판
- 결보강 계획서: 보강수업 칸에 '보강 배정된 교사'가 나오도록 수정
- 모든 기존 기능 유지
"""

import io
from datetime import date, datetime, timedelta
from collections import defaultdict
import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import gspread
from gspread.exceptions import APIError, WorksheetNotFound
from google.oauth2.service_account import Credentials

# ==========================================================================================
# 0. 기본 설정
# ==========================================================================================
st.set_page_config(page_title="시간표·결보강 관리", page_icon="📘", layout="wide")

st.markdown("""
<style>
    [data-testid="stDataFrame"] { border: 1px solid #94a3b8 !important; border-radius: 6px; }
    [data-testid="stDataFrame"] [role="gridcell"], 
    [data-testid="stDataFrame"] [role="columnheader"] {
        border-right: 1px solid #cbd5e1 !important;
        border-bottom: 1px solid #cbd5e1 !important;
    }
    .login-box { max-width: 420px; margin: 80px auto; padding: 30px; border: 1px solid #cbd5e1; border-radius: 12px; background: #f8fafc; }
</style>
""", unsafe_allow_html=True)

SCHOOL_NAME = "서라벌여자중학교"
SCHOOL_YEAR = "2026"
DAYS = ["월", "화", "수", "목", "금"]
PERIODS_PER_DAY = {"월": 6, "화": 7, "수": 7, "목": 7, "금": 6}
MAX_PERIOD = 7
WEEKDAY_KR = {0: "월", 1: "화", 2: "수", 3: "목", 4: "금", 5: "토", 6: "일"}

TIMETABLE_SHEET_ID = "1jZhTHyJ8vKXn6tkoFXfY_f52-pj6eQTdVvRCo3cCmBA"
WORK_SHEET_ID = "1g1B1cyZG_tfRn3AD1NZzr30YxYNYFewJeZYdos2obpU"

MAX_HISTORY = 6
ABSENCE_REASONS = ["병가", "연가", "출장", "공가", "조퇴", "외출", "연수", "특별휴가", "기타"]
MAX_LOGIN_ATTEMPTS = 5
SUB_COST = 10000

SUBJECT_GROUP = {
    "국어1": "국어", "국어2": "국어", "사회": "사회", "사회1": "사회", "사회2": "사회", "사회3": "사회",
    "역사": "역사", "도덕1": "도덕", "도덕2": "도덕", "수학": "수학", "수학1": "수학", "수학2": "수학",
    "과학": "과학", "과학1": "과학", "과학2": "과학", "기가": "기술가정",
    "체육1": "체육", "체육2": "체육", "체육3": "체육", "스포": "스포츠",
    "음악": "음악", "음악1": "음악", "음악2": "음악", "미술": "미술",
    "영어": "영어", "영어1": "영어", "영어2": "영어", "영회": "영어",
    "한문": "한문", "일본어": "일본어", "정보": "정보", "진동": "진로활동",
}

ROLE_MASTER = "마스터"
ROLE_EDU = "교육과정부"
ROLE_TEACHER = "일반교사"
ROLE_GUEST = "게스트"
MASTER_ID = "pse915"

ALL_TABS = [
    "시간표 조회",
    "시간강사 관리",
    "결강·보강",
    "시간표 맞교환 & 변경 추천",
    "통계",
    "시간표 변경 테스트용",
    "변경된 교사 주간표",
    "📋 복무 관리 & 판단",
    "🛠️ 다중 출장·전체 조정 추천",
    "🔑 아이디·권한 관리",
    "📑 회원별 탭 권한 관리"
]

DEFAULT_TABS = {
    ROLE_MASTER: ALL_TABS,
    ROLE_EDU: ALL_TABS,
    ROLE_TEACHER: [
        "시간표 조회", "시간강사 관리", "결강·보강",
        "시간표 맞교환 & 변경 추천", "통계",
        "시간표 변경 테스트용", "변경된 교사 주간표", "📋 복무 관리 & 판단"
    ],
    ROLE_GUEST: []
}

# ==========================================================================================
# 유틸
# ==========================================================================================
def safe_int(val, default=0):
    try:
        if pd.isna(val) or val is None or str(val).strip() in ("", "nan", "None"):
            return default
        return int(float(str(val).strip()))
    except Exception:
        return default

def normalize_date_str(d_str):
    if pd.isna(d_str) or not d_str or str(d_str).strip() in ("", "nan", "None"):
        return ""
    try:
        return pd.to_datetime(str(d_str).strip(), errors="coerce").strftime("%Y-%m-%d")
    except Exception:
        return str(d_str).strip()

def subject_group(subject: str) -> str:
    if not isinstance(subject, str) or not subject.strip():
        return ""
    s = subject.strip()
    return SUBJECT_GROUP.get(s, s.rstrip("0123456789"))

def grade_of(class_name: str) -> str:
    if isinstance(class_name, str) and "-" in class_name:
        return class_name.split("-")[0]
    return ""

def format_periods(periods):
    periods = sorted(set(safe_int(p) for p in periods if safe_int(p) >= 0))
    if not periods:
        return ""
    if 0 in periods:
        return "전체"
    if len(periods) == 1:
        return f"{periods[0]}교시"
    ranges = []
    start = prev = periods[0]
    for p in periods[1:]:
        if p == prev + 1:
            prev = p
        else:
            ranges.append(f"{start}~{prev}교시" if start != prev else f"{start}교시")
            start = prev = p
    ranges.append(f"{start}~{prev}교시" if start != prev else f"{start}교시")
    return ", ".join(ranges)

def get_all_teacher_names():
    ts = []
    if "teachers" in st.session_state and not st.session_state.teachers.empty and "교사명" in st.session_state.teachers.columns:
        ts = st.session_state.teachers["교사명"].dropna().astype(str).str.strip().tolist()
    pts = []
    pt = st.session_state.get("part_time", pd.DataFrame())
    if not pt.empty and "시간강사명" in pt.columns:
        pts = pt["시간강사명"].dropna().astype(str).str.strip().tolist()
    return sorted(set([t for t in ts + pts if t]))

def current_user():
    return st.session_state.get("user_id", "")

def current_name():
    return st.session_state.get("user_name", "")

def current_role():
    return st.session_state.get("user_role", ROLE_GUEST)

def is_master():
    return current_role() == ROLE_MASTER

def is_edu_or_master():
    return current_role() in (ROLE_MASTER, ROLE_EDU)

def is_teacher():
    return current_role() == ROLE_TEACHER

def can_manage_ids():
    return is_edu_or_master()

def can_full_data():
    return is_edu_or_master()

def get_user_allowed_tabs():
    role = current_role()
    allowed_str = st.session_state.get("user_allowed_tabs", "")
    if allowed_str and isinstance(allowed_str, str) and allowed_str.strip():
        tabs = [t.strip() for t in allowed_str.split(",") if t.strip()]
        if not can_manage_ids():
            tabs = [t for t in tabs if t not in ("🔑 아이디·권한 관리", "📑 회원별 탭 권한 관리", "🛠️ 다중 출장·전체 조정 추천")]
        return tabs if tabs else DEFAULT_TABS.get(role, [])
    return DEFAULT_TABS.get(role, [])

def get_teacher_subject(teacher_name: str) -> str:
    teachers = st.session_state.get("teachers", pd.DataFrame())
    if not teachers.empty and "교사명" in teachers.columns:
        row = teachers[teachers["교사명"] == teacher_name]
        if not row.empty:
            for col in ["담당과목", "과목", "교과", "전공"]:
                if col in row.columns:
                    val = str(row.iloc[0][col]).strip()
                    if val and val not in ("nan", "None", ""):
                        return val
    tt = st.session_state.get("timetable", pd.DataFrame())
    if not tt.empty:
        sub = tt[tt["교사명"] == teacher_name]
        if not sub.empty and "과목군" in sub.columns:
            most = sub["과목군"].value_counts()
            if not most.empty:
                return most.index[0]
        if not sub.empty and "과목" in sub.columns:
            most = sub["과목"].value_counts()
            if not most.empty:
                return subject_group(most.index[0])
    return ""

# ==========================================================================================
# Google Sheets
# ==========================================================================================
@st.cache_resource(show_spinner=False)
def get_gspread_client():
    scopes = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    if "gcp_service_account" not in st.secrets:
        st.error("GCP Secrets 인증 오류")
        st.stop()
    creds = Credentials.from_service_account_info(st.secrets["gcp_service_account"], scopes=scopes)
    return gspread.authorize(creds)

@st.cache_resource(show_spinner=False)
def get_spreadsheet(spreadsheet_id: str):
    return get_gspread_client().open_by_key(spreadsheet_id)

def get_worksheet(spreadsheet_id: str, sheet_name: str):
    try:
        return get_spreadsheet(spreadsheet_id).worksheet(sheet_name)
    except WorksheetNotFound:
        try:
            return get_spreadsheet(spreadsheet_id).add_worksheet(title=sheet_name, rows=2000, cols=40)
        except Exception:
            return None
    except Exception:
        return None

def df_from_worksheet(ws):
    if ws is None:
        return pd.DataFrame()
    try:
        data = ws.get_all_values()
        if not data or len(data) < 2:
            return pd.DataFrame()
        headers = [str(h).strip() for h in data[0]]
        rows = []
        for row in data[1:]:
            row = list(row) + [""] * max(0, len(headers) - len(row))
            rows.append(["" if c is None else str(c).strip() for c in row[:len(headers)]])
        return pd.DataFrame(rows, columns=headers).replace({"nan": "", "None": "", "NaN": ""})
    except Exception:
        return pd.DataFrame()

def df_to_worksheet(ws, df):
    if ws is None:
        return
    try:
        ws.clear()
        if df is None or df.empty:
            return
        values = [df.fillna("").astype(str).columns.tolist()] + df.fillna("").astype(str).values.tolist()
        ws.update("A1", values, value_input_option="USER_ENTERED")
    except Exception:
        pass

# ==========================================================================================
# 아이디 / 권한 / 예산 / 수업교체신청
# ==========================================================================================
@st.cache_data(ttl=30, show_spinner=False)
def load_id_sheet():
    ws = get_worksheet(WORK_SHEET_ID, "아이디저장함")
    df = df_from_worksheet(ws)
    if df.empty or "아이디" not in df.columns:
        df = pd.DataFrame([{
            "아이디": MASTER_ID, "이름": "관리자", "권한": ROLE_MASTER, "허용탭": ",".join(ALL_TABS)
        }])
        df_to_worksheet(ws, df)
    for c in ["아이디", "이름", "권한", "허용탭"]:
        if c not in df.columns:
            df[c] = ""
    df["권한"] = df["권한"].replace("", ROLE_TEACHER)
    df["이름"] = df["이름"].fillna("").astype(str)
    df["허용탭"] = df["허용탭"].fillna("").astype(str)
    return df

def save_id_sheet(df):
    ws = get_worksheet(WORK_SHEET_ID, "아이디저장함")
    df_to_worksheet(ws, df)
    load_id_sheet.clear()

def save_id_request(name, email, desired_id, memo):
    ws = get_worksheet(WORK_SHEET_ID, "아이디추가요청")
    df = df_from_worksheet(ws)
    new = pd.DataFrame([{
        "이름": name, "이메일": email, "추가아이디": desired_id, "메모": memo,
        "요청시각": datetime.now().strftime("%Y-%m-%d %H:%M:%S"), "처리상태": "대기"
    }])
    if df.empty:
        df = new
    else:
        for c in new.columns:
            if c not in df.columns:
                df[c] = ""
        df = pd.concat([df, new], ignore_index=True)
    df_to_worksheet(ws, df)

@st.cache_data(ttl=15, show_spinner=False)
def load_budget_df():
    ws = get_worksheet(WORK_SHEET_ID, "예산")
    df = df_from_worksheet(ws)
    expected_cols = ["일시", "내용", "변동금액", "잔액"]
    if df.empty or not any(c in df.columns for c in expected_cols + ["보강예산 현황"]):
        init_df = pd.DataFrame([{
            "일시": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "내용": "초기 예산 설정",
            "변동금액": 2200000,
            "잔액": 2200000
        }])
        df_to_worksheet(ws, init_df)
        return init_df
    if "보강예산 현황" in df.columns and "잔액" not in df.columns:
        try:
            val = safe_int(df.iloc[0, 0]) if not df.empty else 2200000
        except Exception:
            val = 2200000
        df = pd.DataFrame([{
            "일시": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "내용": "기존 예산 불러오기",
            "변동금액": 0,
            "잔액": val
        }])
    for c in expected_cols:
        if c not in df.columns:
            df[c] = 0 if c in ("변동금액", "잔액") else ""
    return df

def get_current_budget():
    try:
        df = load_budget_df()
        if df.empty:
            return 2200000
        last_val = safe_int(df.iloc[-1].get("잔액", 2200000))
        return max(0, last_val)
    except Exception:
        return 2200000

def update_budget(change_amount: int, reason: str = "보강"):
    try:
        df = load_budget_df()
        current = get_current_budget()
        new_balance = max(0, current + change_amount)
        new_row = pd.DataFrame([{
            "일시": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "내용": reason,
            "변동금액": change_amount,
            "잔액": new_balance
        }])
        df = pd.concat([df, new_row], ignore_index=True)
        ws = get_worksheet(WORK_SHEET_ID, "예산")
        df_to_worksheet(ws, df)
        load_budget_df.clear()
        return new_balance
    except Exception as e:
        st.warning(f"예산 업데이트 실패: {e}")
        return get_current_budget()

SWAP_REQUEST_COLS = [
    "신청ID", "신청자", "신청자이름", "원본일자", "교사A", "요일A", "교시A", "학급A", "과목A",
    "목표일자", "교사B", "요일B", "교시B", "학급B", "과목B", "유형", "신청시각", "상태"
]

@st.cache_data(ttl=30, show_spinner=False)
def load_swap_requests():
    ws = get_worksheet(WORK_SHEET_ID, "수업교체신청")
    df = df_from_worksheet(ws)
    if df.empty:
        return pd.DataFrame(columns=SWAP_REQUEST_COLS)
    for c in SWAP_REQUEST_COLS:
        if c not in df.columns:
            df[c] = ""
    return df

def save_swap_request(rec: dict):
    df = load_swap_requests()
    new = pd.DataFrame([rec])
    df = pd.concat([df, new], ignore_index=True)
    ws = get_worksheet(WORK_SHEET_ID, "수업교체신청")
    df_to_worksheet(ws, df)
    load_swap_requests.clear()

# ==========================================================================================
# 히스토리
# ==========================================================================================
def push_history(action_name="작업"):
    if "history" not in st.session_state:
        st.session_state.history = []
        st.session_state.history_index = -1
    st.session_state.history = st.session_state.history[:st.session_state.history_index + 1]
    snap = {
        "action": action_name,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "absences": st.session_state.get("absences", pd.DataFrame()).copy(deep=False),
        "subs": st.session_state.get("subs", pd.DataFrame()).copy(deep=False),
        "swaps": st.session_state.get("swaps", pd.DataFrame()).copy(deep=False),
        "part_time": st.session_state.get("part_time", pd.DataFrame()).copy(deep=False),
        "duties": st.session_state.get("duties", pd.DataFrame()).copy(deep=False),
    }
    st.session_state.history.append(snap)
    if len(st.session_state.history) > MAX_HISTORY:
        st.session_state.history.pop(0)
    else:
        st.session_state.history_index += 1

def undo():
    if st.session_state.get("history_index", 0) <= 0:
        return False
    st.session_state.history_index -= 1
    snap = st.session_state.history[st.session_state.history_index]
    for k in ["absences", "subs", "swaps", "part_time", "duties"]:
        st.session_state[k] = snap[k].copy(deep=False)
    _invalidate_all_caches()
    return True

def redo():
    if st.session_state.get("history_index", -1) >= len(st.session_state.get("history", [])) - 1:
        return False
    st.session_state.history_index += 1
    snap = st.session_state.history[st.session_state.history_index]
    for k in ["absences", "subs", "swaps", "part_time", "duties"]:
        st.session_state[k] = snap[k].copy(deep=False)
    _invalidate_all_caches()
    return True

def _invalidate_all_caches():
    st.session_state._data_version = st.session_state.get("_data_version", 0) + 1
    get_effective_timetable_for_date.clear()
    teacher_matrix.clear()
    class_matrix.clear()
    cumulative_sub_count.clear()
    weekly_load.clear()
    load_budget_df.clear()
    load_swap_requests.clear()

# ==========================================================================================
# 데이터 로드
# ==========================================================================================
DUTY_COLS = ["교사명", "일자", "교시", "사유", "상세사유", "등록시각", "입력자"]
PART_TIME_EXTRA_COLS = ["시작일", "종료일", "대체교사"]

def ensure_duty_columns(df):
    if df is None or not isinstance(df, pd.DataFrame):
        return pd.DataFrame(columns=DUTY_COLS)
    df = df.copy()
    for c in DUTY_COLS:
        if c not in df.columns:
            df[c] = 0 if c == "교시" else ""
    df["교시"] = df["교시"].apply(safe_int)
    return df

def ensure_part_time_columns(df):
    base_cols = ["번호", "시간강사명", "담당과목", "과목군", "비고"] + PART_TIME_EXTRA_COLS + [f"{d}{p}" for d in DAYS for p in range(1, 8)]
    if df is None or not isinstance(df, pd.DataFrame):
        return pd.DataFrame(columns=base_cols)
    df = df.copy()
    for c in base_cols:
        if c not in df.columns:
            df[c] = ""
    for c in ["시작일", "종료일"]:
        if c in df.columns:
            df[c] = df[c].apply(normalize_date_str)
    return df

def ensure_input_user(df, default=""):
    if df is None or not isinstance(df, pd.DataFrame):
        return df
    if "입력자" not in df.columns:
        df = df.copy()
        df["입력자"] = default
    return df

@st.cache_data(ttl=300, show_spinner="시간표 로딩...")
def load_timetable_from_gsheet():
    try:
        ti = df_from_worksheet(get_worksheet(TIMETABLE_SHEET_ID, "교사정보"))
        tt = df_from_worksheet(get_worksheet(TIMETABLE_SHEET_ID, "시간표"))
        if tt.empty:
            return ti, pd.DataFrame(columns=["교사명", "요일", "교시", "과목", "학급", "과목군"])
        tt["교시"] = tt["교시"].apply(safe_int)
        for c in ["교사명", "요일", "과목", "학급"]:
            tt[c] = tt[c].astype(str).str.strip()
        if "과목군" not in tt.columns or tt["과목군"].eq("").all():
            tt["과목군"] = tt["과목"].map(subject_group)
        tt = tt[tt["요일"].isin(DAYS) & (tt["교시"] >= 1) & (tt["교시"] <= MAX_PERIOD)]
        tt = tt.drop_duplicates(subset=["교사명", "요일", "교시"]).reset_index(drop=True)
        return ti, tt
    except Exception as e:
        st.error(f"시간표 로드 실패: {e}")
        return pd.DataFrame(), pd.DataFrame()

@st.cache_data(ttl=60, show_spinner="작업 데이터 로딩...")
def load_work_data_from_gsheet():
    try:
        absences = ensure_input_user(df_from_worksheet(get_worksheet(WORK_SHEET_ID, "결강")))
        subs = ensure_input_user(df_from_worksheet(get_worksheet(WORK_SHEET_ID, "보강")))
        swaps = ensure_input_user(df_from_worksheet(get_worksheet(WORK_SHEET_ID, "맞교환")))
        part_time = ensure_part_time_columns(df_from_worksheet(get_worksheet(WORK_SHEET_ID, "시간강사")))
        cumulative = df_from_worksheet(get_worksheet(WORK_SHEET_ID, "누적보강"))
        duties = ensure_duty_columns(df_from_worksheet(get_worksheet(WORK_SHEET_ID, "복무")))

        for df in [absences, subs]:
            if not df.empty:
                if "교시" in df.columns:
                    df["교시"] = df["교시"].apply(safe_int)
                if "일자" in df.columns:
                    df["일자"] = df["일자"].apply(normalize_date_str)
        if not swaps.empty:
            for c in ["원본일자", "목표일자"]:
                if c in swaps.columns:
                    swaps[c] = swaps[c].apply(normalize_date_str)
            for c in ["교시A", "교시B"]:
                if c in swaps.columns:
                    swaps[c] = swaps[c].apply(safe_int)
        if not duties.empty and "일자" in duties.columns:
            duties["일자"] = duties["일자"].apply(normalize_date_str)
        return absences, subs, swaps, part_time, cumulative, duties
    except Exception as e:
        st.error(f"작업 데이터 로드 실패: {e}")
        return (pd.DataFrame(),) * 5 + (pd.DataFrame(columns=DUTY_COLS),)

def save_work_data_to_gsheet():
    if not can_full_data() and not is_teacher():
        st.warning("저장 권한이 없습니다.")
        return False
    try:
        df_to_worksheet(get_worksheet(WORK_SHEET_ID, "결강"), st.session_state.absences)
        df_to_worksheet(get_worksheet(WORK_SHEET_ID, "보강"), st.session_state.subs)
        df_to_worksheet(get_worksheet(WORK_SHEET_ID, "맞교환"), st.session_state.swaps)
        st.session_state.part_time = ensure_part_time_columns(st.session_state.part_time)
        df_to_worksheet(get_worksheet(WORK_SHEET_ID, "시간강사"), st.session_state.part_time)
        st.session_state.duties = ensure_duty_columns(st.session_state.duties)
        df_to_worksheet(get_worksheet(WORK_SHEET_ID, "복무"), st.session_state.duties)
        _invalidate_all_caches()
        return True
    except Exception as e:
        st.error(f"저장 실패: {e}")
        return False

def init_state():
    if "teachers" in st.session_state:
        return
    ti, tt = load_timetable_from_gsheet()
    st.session_state.teachers = ti
    st.session_state.timetable = tt

    absences, subs, swaps, part_time, cumulative, duties = load_work_data_from_gsheet()

    if absences.empty:
        absences = pd.DataFrame(columns=["결강ID", "일자", "요일", "교사명", "사유", "상세사유", "교시", "학급", "과목", "등록시각", "입력자"])
    if subs.empty:
        subs = pd.DataFrame(columns=["결강ID", "일자", "요일", "교시", "학급", "과목", "결강교사", "보강교사", "배정방식", "우선순위", "비고", "등록시각", "입력자"])
    if swaps.empty:
        swaps = pd.DataFrame(columns=["원본일자", "교사A", "요일A", "교시A", "학급A", "과목A",
                                      "목표일자", "교사B", "요일B", "교시B", "학급B", "과목B", "유형", "시간강사구인", "등록시각", "입력자"])
    if part_time.empty:
        part_time = pd.DataFrame(columns=["번호", "시간강사명", "담당과목", "과목군", "비고", "시작일", "종료일", "대체교사"] + [f"{d}{p}" for d in DAYS for p in range(1, 8)])

    st.session_state.absences = ensure_input_user(absences)
    st.session_state.subs = ensure_input_user(subs)
    st.session_state.swaps = ensure_input_user(swaps)
    st.session_state.part_time = ensure_part_time_columns(part_time)
    st.session_state.cumulative = cumulative
    st.session_state.duties = ensure_duty_columns(duties)
    st.session_state._data_version = 0
    st.session_state.history = []
    st.session_state.history_index = -1
    st.session_state.test_swaps = pd.DataFrame()
    push_history("초기 상태")

# ==========================================================================================
# 핵심 로직 (기존과 동일)
# ==========================================================================================
@st.cache_data(show_spinner=False, ttl=180)
def get_effective_timetable_for_date(on_date: str, version: int = 0, use_test: bool = False) -> pd.DataFrame:
    norm = normalize_date_str(on_date)
    if not norm:
        return st.session_state.timetable.copy()

    try:
        day = WEEKDAY_KR[datetime.strptime(norm, "%Y-%m-%d").weekday()]
    except Exception:
        return st.session_state.timetable.copy()

    tt = st.session_state.timetable
    if tt.empty:
        return pd.DataFrame(columns=["교사명", "요일", "교시", "과목", "학급", "과목군", "원본교사"])

    base = tt[tt["요일"] == day]
    current = {}
    for r in base.itertuples(index=False):
        p = safe_int(r.교시)
        t = str(r.교사명).strip()
        current[(t, p)] = {
            "교사명": t, "요일": day, "교시": p,
            "과목": str(r.과목).strip(), "학급": str(r.학급).strip(),
            "과목군": str(getattr(r, "과목군", subject_group(r.과목))).strip(),
            "원본교사": ""
        }

    swaps = st.session_state.swaps
    if not swaps.empty:
        mask = (swaps["원본일자"] == norm) | (swaps["목표일자"] == norm)
        for sw in swaps[mask].itertuples(index=False):
            t_a, p_a = str(sw.교사A).strip(), safe_int(sw.교시A)
            t_b, p_b = str(sw.교사B).strip(), safe_int(sw.교시B)
            typ = str(getattr(sw, "유형", "")).strip()
            s_a = str(getattr(sw, "과목A", "")).strip()
            c_a = str(getattr(sw, "학급A", "")).strip()
            s_b = str(getattr(sw, "과목B", "")).strip()
            c_b = str(getattr(sw, "학급B", "")).strip()

            if sw.원본일자 == norm:
                current.pop((t_a, p_a), None)
                if typ in ["1:1 맞교환", "1:1맞교환", "직접1:1"] and t_b:
                    current[(t_b, p_a)] = {"교사명": t_b, "요일": day, "교시": p_a,
                                           "과목": s_b or s_a, "학급": c_b or c_a,
                                           "과목군": subject_group(s_b or s_a), "원본교사": ""}
            if sw.목표일자 == norm:
                if typ in ["1:1 맞교환", "1:1맞교환", "직접1:1"]:
                    current.pop((t_b, p_b), None)
                    if t_a and p_b:
                        current[(t_a, p_b)] = {"교사명": t_a, "요일": day, "교시": p_b,
                                               "과목": s_a, "학급": c_a, "과목군": subject_group(s_a), "원본교사": ""}
                elif "연계" in typ and t_a and p_b:
                    current[(t_a, p_b)] = {"교사명": t_a, "요일": day, "교시": p_b,
                                           "과목": s_a, "학급": c_a, "과목군": subject_group(s_a), "원본교사": ""}

    if use_test:
        test_swaps = st.session_state.get("test_swaps", pd.DataFrame())
        if not test_swaps.empty:
            mask = (test_swaps["원본일자"] == norm) | (test_swaps["목표일자"] == norm)
            for sw in test_swaps[mask].itertuples(index=False):
                t_a, p_a = str(sw.교사A).strip(), safe_int(sw.교시A)
                t_b, p_b = str(sw.교사B).strip(), safe_int(sw.교시B)
                s_a = str(getattr(sw, "과목A", "")).strip()
                c_a = str(getattr(sw, "학급A", "")).strip()
                if sw.원본일자 == norm:
                    current.pop((t_a, p_a), None)
                    if t_b:
                        current[(t_b, p_a)] = {"교사명": t_b, "요일": day, "교시": p_a,
                                               "과목": s_a, "학급": c_a, "과목군": subject_group(s_a), "원본교사": ""}
                if sw.목표일자 == norm and t_a and p_b:
                    current.pop((t_b, p_b), None)
                    current[(t_a, p_b)] = {"교사명": t_a, "요일": day, "교시": p_b,
                                           "과목": s_a, "학급": c_a, "과목군": subject_group(s_a), "원본교사": ""}

    subs = st.session_state.subs
    if not subs.empty:
        day_subs = subs[subs["일자"] == norm]
        for r in day_subs.itertuples(index=False):
            p = safe_int(r.교시)
            abs_t = str(r.결강교사).strip()
            sub_t = str(r.보강교사).strip()
            if sub_t and p > 0:
                current.pop((abs_t, p), None)
                current[(sub_t, p)] = {
                    "교사명": sub_t, "요일": day, "교시": p,
                    "과목": str(r.과목).strip(), "학급": str(r.학급).strip(),
                    "과목군": subject_group(str(r.과목)), "원본교사": ""
                }

    pt_df = st.session_state.get("part_time", pd.DataFrame())
    if not pt_df.empty and "시작일" in pt_df.columns:
        for _, prow in pt_df.iterrows():
            start = normalize_date_str(prow.get("시작일", ""))
            end = normalize_date_str(prow.get("종료일", ""))
            if not (start and end and start <= norm <= end):
                continue
            orig = str(prow.get("대체교사", "")).strip()
            pt_name = str(prow.get("시간강사명", "")).strip()
            if not orig or not pt_name or orig == pt_name:
                continue
            keys_to_move = [k for k in list(current.keys()) if k[0] == orig]
            for k in keys_to_move:
                lesson = current.pop(k)
                lesson["교사명"] = pt_name
                lesson["원본교사"] = orig
                current[(pt_name, k[1])] = lesson

    df = pd.DataFrame(list(current.values()))
    if df.empty:
        df = pd.DataFrame(columns=["교사명", "요일", "교시", "과목", "학급", "과목군", "원본교사"])
    return df

def get_swap_origin_info(teacher: str, on_date: str, period: int) -> str:
    norm_date = normalize_date_str(on_date)
    if not norm_date:
        return ""
    swaps = st.session_state.get("swaps", pd.DataFrame())
    if swaps.empty:
        return ""
    p = safe_int(period)
    mask1 = (swaps["목표일자"] == norm_date) & (swaps["교사A"] == teacher) & (swaps["교시B"] == p)
    if mask1.any():
        row = swaps[mask1].iloc[0]
        return f"{row.get('요일A','')}{safe_int(row.get('교시A',0))}({row.get('교사B','')})"
    mask2 = (swaps["원본일자"] == norm_date) & (swaps["교사B"] == teacher) & (swaps["교시A"] == p)
    if mask2.any():
        row = swaps[mask2].iloc[0]
        return f"{row.get('요일B','')}{safe_int(row.get('교시B',0))}({row.get('교사A','')})"
    return ""

def has_duty(teacher: str, on_date: str, period: int = None) -> bool:
    duties = st.session_state.duties
    if duties.empty:
        return False
    norm = normalize_date_str(on_date)
    mask = (duties["교사명"] == teacher) & (duties["일자"] == norm)
    if not mask.any():
        return False
    if period is None:
        return True
    periods = duties.loc[mask, "교시"].tolist()
    return 0 in periods or safe_int(period) in periods

def is_free(teacher: str, day: str, period: int, on_date: str = None, e_tt=None) -> bool:
    p = safe_int(period)
    norm = normalize_date_str(on_date)
    if has_duty(teacher, norm, p):
        return False
    if e_tt is None:
        e_tt = get_effective_timetable_for_date(norm, st.session_state.get("_data_version", 0))
    if not e_tt.empty and ((e_tt["교사명"] == teacher) & (e_tt["교시"] == p)).any():
        return False
    return True

@st.cache_data(show_spinner=False)
def cumulative_sub_count(start_date=None, end_date=None, version=0):
    s = st.session_state.subs
    base = {t: 0 for t in st.session_state.teachers["교사명"].tolist()} if not st.session_state.teachers.empty else {}
    if s.empty or "보강교사" not in s.columns:
        return base
    if start_date and end_date:
        s = s[(s["일자"] >= normalize_date_str(start_date)) & (s["일자"] <= normalize_date_str(end_date))]
    counts = s["보강교사"].value_counts()
    for k, v in counts.items():
        if k in base:
            base[k] = int(v)
    return base

@st.cache_data(show_spinner=False)
def weekly_load(version=0):
    tt = st.session_state.timetable
    return tt["교사명"].value_counts().to_dict() if not tt.empty else {}

def recommend_substitutes(day, period, subject, class_name, absent_teacher, on_date, top_n=20, include_part_time=False, e_tt=None):
    """해당 시간대에 비어있는 모든 교사를 검색 (우선순위 유지)"""
    teachers = st.session_state.teachers
    if teachers.empty:
        return pd.DataFrame()
    norm = normalize_date_str(on_date)
    if e_tt is None:
        e_tt = get_effective_timetable_for_date(norm, st.session_state.get("_data_version", 0))
    grp = subject_group(subject)
    grade = grade_of(class_name)
    cum = cumulative_sub_count(version=st.session_state.get("_data_version", 0))
    load = weekly_load(version=st.session_state.get("_data_version", 0))
    max_cum = max(cum.values()) if cum else 0

    free = [t for t in teachers["교사명"].tolist()
            if t != absent_teacher and not has_duty(t, norm) and is_free(t, day, period, norm, e_tt)]

    rows = []
    for t in free:
        my = e_tt[e_tt["교사명"] == t] if not e_tt.empty else pd.DataFrame()
        my_groups = set(my["과목군"]) if not my.empty else set()
        my_grades = {grade_of(c) for c in my["학급"]} if not my.empty else set()

        if grp in my_groups and grade in my_grades:
            prio, label, score = 1, "1순위 · 동일 과목 & 동일 학년", 100
        elif grp in my_groups:
            prio, label, score = 2, "2순위 · 동일 과목", 70
        elif grade in my_grades:
            prio, label, score = 3, "3순위 · 동일 학년", 45
        else:
            prio, label, score = 4, "4순위 · 전체 공강", 20

        score += (max_cum - cum.get(t, 0)) * 2 + max(0, 22 - load.get(t, 0)) * 0.3

        t_row = teachers[teachers["교사명"] == t]
        rows.append({
            "보강교사": t,
            "유형": "정규교사",
            "우선순위": label,
            "_prio": prio,
            "담당과목": t_row["담당과목"].iloc[0] if not t_row.empty and "담당과목" in t_row.columns else "",
            "주당시수": load.get(t, 0),
            "누적보강": cum.get(t, 0),
            "추천점수": round(score, 1)
        })

    if not rows:
        return pd.DataFrame()

    df = pd.DataFrame(rows)
    return df.sort_values(["_prio", "추천점수"], ascending=[True, False]).drop(columns=["_prio"]).head(top_n).reset_index(drop=True)

def add_substitute(cid, on_date, day, period, class_name, subject, absent_teacher, sub_teacher, method, priority, memo):
    push_history(f"보강 배정 ({sub_teacher})")
    s = st.session_state.subs
    p = safe_int(period)
    norm = normalize_date_str(on_date)
    if not s.empty:
        s = s[~((s["결강ID"] == cid) & (s["교시"] == p))]
    new = pd.DataFrame([{
        "결강ID": cid, "일자": norm, "요일": day, "교시": p,
        "학급": class_name, "과목": subject, "결강교사": absent_teacher,
        "보강교사": sub_teacher, "배정방식": method, "우선순위": priority,
        "비고": memo, "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "입력자": current_user()
    }])
    st.session_state.subs = pd.concat([s, new], ignore_index=True)
    save_work_data_to_gsheet()
    update_budget(-SUB_COST, f"보강 1건 ({sub_teacher} ← {absent_teacher})")

def cancel_substitute(cid, period):
    if not can_full_data():
        s = st.session_state.subs
        p = safe_int(period)
        m = s[(s["결강ID"] == cid) & (s["교시"] == p)]
        if not m.empty and m.iloc[0].get("입력자") != current_user():
            st.warning("본인이 입력한 데이터만 취소할 수 있습니다.")
            return
    push_history(f"보강 취소 ({period}교시)")
    s = st.session_state.subs
    p = safe_int(period)
    st.session_state.subs = s[~((s["결강ID"] == cid) & (s["교시"] == p))].reset_index(drop=True)
    save_work_data_to_gsheet()
    update_budget(+SUB_COST, f"보강 취소 복구 ({period}교시)")

def do_swap(a, b, date_a, date_b, is_part_time_purpose=False, is_test=False):
    rec = {
        "원본일자": normalize_date_str(date_a), "교사A": a["교사명"], "요일A": a["요일"], "교시A": safe_int(a["교시"]),
        "학급A": a.get("학급", ""), "과목A": a.get("과목", ""),
        "목표일자": normalize_date_str(date_b), "교사B": b["교사명"], "요일B": b["요일"], "교시B": safe_int(b["교시"]),
        "학급B": b.get("학급", ""), "과목B": b.get("과목", ""),
        "유형": "1:1 맞교환", "시간강사구인": "Y" if is_part_time_purpose else "N",
        "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "입력자": current_user()
    }
    if is_test:
        st.session_state.test_swaps = pd.concat([st.session_state.get("test_swaps", pd.DataFrame()), pd.DataFrame([rec])], ignore_index=True)
        return True
    push_history(f"맞교환 ({a['교사명']} ↔ {b['교사명']})")
    st.session_state.swaps = pd.concat([st.session_state.swaps, pd.DataFrame([rec])], ignore_index=True)
    save_work_data_to_gsheet()
    return True

def do_linked_swap(a, teacher_b, date_a, date_b, day_b, period_b, is_part_time_purpose=False, is_test=False, subject_b=None):
    """subject_b를 명시적으로 받을 수 있도록 수정 (과목 버그 해결)"""
    rec = {
        "원본일자": normalize_date_str(date_a), "교사A": a["교사명"], "요일A": a["요일"], "교시A": safe_int(a["교시"]),
        "학급A": a.get("학급", ""), "과목A": a.get("과목", ""),
        "목표일자": normalize_date_str(date_b), "교사B": teacher_b, "요일B": day_b, "교시B": safe_int(period_b),
        "학급B": a.get("학급", ""), "과목B": subject_b if subject_b is not None else a.get("과목", ""),
        "유형": "연계 공강 교환", "시간강사구인": "Y" if is_part_time_purpose else "N",
        "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "입력자": current_user()
    }
    if is_test:
        st.session_state.test_swaps = pd.concat([st.session_state.get("test_swaps", pd.DataFrame()), pd.DataFrame([rec])], ignore_index=True)
        return True
    push_history(f"연계교환 ({a['교사명']} → {teacher_b})")
    st.session_state.swaps = pd.concat([st.session_state.swaps, pd.DataFrame([rec])], ignore_index=True)
    save_work_data_to_gsheet()
    return True

def apply_cycle_swaps(moves, is_test=False):
    """순환 경로 전체를 정확한 과목과 함께 기록"""
    if not moves:
        return False
    for m in moves:
        a_info = {
            "교사명": m["teacher"],
            "요일": m.get("day_from", WEEKDAY_KR[datetime.strptime(m["from_date"], "%Y-%m-%d").weekday()]),
            "교시": m["from_period"],
            "학급": m["class"],
            "과목": m["subject"]   # 이동하는 교사의 원래 과목
        }
        to_day = WEEKDAY_KR[datetime.strptime(m["to_date"], "%Y-%m-%d").weekday()]
        # 과목B에는 목표 슬롯의 원래 과목을 넣어 명확하게 표시
        do_linked_swap(a_info, m.get("next_teacher", m["teacher"]), 
                       m["from_date"], m["to_date"], to_day, m["to_period"],
                       is_test=is_test, subject_b=m.get("target_subject", m["subject"]))
    return True

# ==========================================================================================
# ★★★ 연계 공강 순환 알고리즘 (속도 최적화 + 과목 정확 기록)
# ==========================================================================================
def find_cycle_linked_swaps(teacher_a, date_a_str, period_a, class_a, subject_a,
                           date_b_str, period_b, max_cycle=4, future_days=7, version=0):
    """학급 시수·담당·과목 완전 보존 2~4인 순환. 속도 최적화됨."""
    original_slot = (normalize_date_str(date_a_str), safe_int(period_a))
    target_slot = (normalize_date_str(date_b_str), safe_int(period_b))

    try:
        da = datetime.strptime(date_a_str, "%Y-%m-%d")
        db = datetime.strptime(date_b_str, "%Y-%m-%d")
        base_min = min(da, db) - timedelta(days=2)
        base_max = max(da, db) + timedelta(days=future_days)
    except Exception:
        return [], "날짜 오류"

    candidates = []
    cur = base_min
    while cur <= base_max:
        if cur.weekday() < 5:
            candidates.append(cur.strftime("%Y-%m-%d"))
        cur += timedelta(days=1)
    candidates = sorted(set(candidates))

    e_cache = {d: get_effective_timetable_for_date(d, version) for d in candidates}

    class_slots = {}
    for d in candidates:
        e = e_cache.get(d)
        if e is None or e.empty:
            continue
        for _, r in e[e["학급"] == class_a].iterrows():
            p = safe_int(r["교시"])
            class_slots[(d, p)] = {
                "teacher": str(r["교사명"]).strip(),
                "subject": str(r["과목"]).strip(),
                "day": str(r["요일"]).strip()
            }

    if original_slot not in class_slots or class_slots[original_slot]["teacher"] != teacher_a:
        return [], "원본 슬롯/교사 불일치"
    if target_slot not in class_slots:
        return [], "목표 슬롯에 학급 수업 없음 (공강 생성 금지)"

    day_b = WEEKDAY_KR[datetime.strptime(date_b_str, "%Y-%m-%d").weekday()]
    if not is_free(teacher_a, day_b, period_b, date_b_str, e_cache.get(date_b_str)):
        return [], "교사A 목표시간 수업 있음"

    slots_list = list(class_slots.keys())
    free_moves = defaultdict(list)
    for from_s, info in class_slots.items():
        t = info["teacher"]
        for to_s in slots_list:
            if to_s == from_s:
                continue
            to_date, to_p = to_s
            to_day = WEEKDAY_KR[datetime.strptime(to_date, "%Y-%m-%d").weekday()]
            if is_free(t, to_day, to_p, to_date, e_cache.get(to_date)):
                free_moves[from_s].append(to_s)

    cycles = []
    def dfs(current, path, visited):
        if len(path) > max_cycle:
            return
        if current == original_slot and len(path) >= 2:
            cycle_slots = [original_slot] + path[:-1]
            moves = []
            n = len(cycle_slots)
            for i in range(n):
                from_s = cycle_slots[i]
                to_s = cycle_slots[(i + 1) % n]
                info = class_slots[from_s]
                tgt_info = class_slots[to_s]
                moves.append({
                    "teacher": info["teacher"],
                    "from_date": from_s[0],
                    "from_period": from_s[1],
                    "to_date": to_s[0],
                    "to_period": to_s[1],
                    "class": class_a,
                    "subject": info["subject"],          # 이동 교사의 원래 과목
                    "target_subject": tgt_info["subject"], # 목표 슬롯의 원래 과목 (표시용)
                    "day_from": info["day"],
                    "next_teacher": tgt_info["teacher"]
                })
            path_parts = []
            for m in moves:
                fd = m["from_date"][5:]
                td = m["to_date"][5:]
                path_parts.append(f"{m['teacher']}({m['class']} {fd} {m['from_period']}→{td} {m['to_period']})")
            cycles.append({
                "length": n,
                "moves": moves,
                "path_desc": " → ".join(path_parts),
                "score": 110 - n * 12
            })
            return
        for nxt in free_moves.get(current, []):
            if nxt not in visited:
                visited.add(nxt)
                path.append(nxt)
                dfs(nxt, path, visited)
                path.pop()
                visited.remove(nxt)

    dfs(target_slot, [target_slot], set([target_slot]))
    cycles.sort(key=lambda x: (x["length"], -x["score"]))
    return cycles[:6], f"{len(cycles)}개 순환 경로 발견" if cycles else "순환 경로 없음"

def get_target_time_recommendations(teacher_a, date_a_str, period_a, class_a, subject_a, date_b_str, period_b, budget_factor=1.0):
    ti = st.session_state.teachers
    norm_a = normalize_date_str(date_a_str)
    norm_b = normalize_date_str(date_b_str)
    ver = st.session_state.get("_data_version", 0)
    e_a = get_effective_timetable_for_date(norm_a, ver)
    e_b = get_effective_timetable_for_date(norm_b, ver)
    if ti.empty:
        return pd.DataFrame(), [], ""
    p_a = safe_int(period_a)
    p_b = safe_int(period_b)
    day_a = WEEKDAY_KR[datetime.strptime(norm_a, "%Y-%m-%d").weekday()]
    day_b = WEEKDAY_KR[datetime.strptime(norm_b, "%Y-%m-%d").weekday()]
    my_class = class_a
    my_grade = grade_of(class_a)
    my_group = subject_group(subject_a)
    cum = cumulative_sub_count(version=ver)
    swap_recs = []

    for t_b in ti["교사명"].tolist():
        if t_b == teacher_a or has_duty(t_b, norm_b):
            continue
        b_lessons = e_b[(e_b["교사명"] == t_b) & (e_b["교시"] == p_b)] if not e_b.empty else pd.DataFrame()
        if not b_lessons.empty:
            for _, b_row in b_lessons.iterrows():
                if is_free(teacher_a, day_b, p_b, norm_b, e_b) and is_free(t_b, day_a, p_a, norm_a, e_a):
                    other_class = b_row["학급"]
                    other_grade = grade_of(other_class)
                    other_group = subject_group(b_row["과목"])
                    score = 0
                    same_class = (other_class == my_class)
                    same_grade = (other_grade == my_grade)
                    if same_class:
                        score += 200
                    elif same_grade:
                        score += 100
                    if other_group == my_group:
                        score += 40
                    if norm_b == norm_a:
                        score += 15
                    score -= cum.get(t_b, 0) * 3
                    score *= budget_factor
                    swap_recs.append({
                        "유형": "1:1", "교사B": t_b,
                        "현재 수업": f"{day_b}{p_b}교시 · {other_class} · {b_row['과목']}",
                        "학급": other_class, "학년": other_grade,
                        "same_class": same_class, "same_grade": same_grade, "점수": score,
                        "b_info": {"교사명": t_b, "일자": norm_b, "요일": day_b, "교시": p_b,
                                   "학급": other_class, "과목": b_row["과목"]}
                    })

    df_swap = (pd.DataFrame(swap_recs)
               .sort_values(["same_class", "same_grade", "점수"], ascending=[False, False, False])
               .reset_index(drop=True) if swap_recs else pd.DataFrame())

    cycles, msg = find_cycle_linked_swaps(
        teacher_a, date_a_str, period_a, class_a, subject_a,
        date_b_str, period_b, max_cycle=4, future_days=7, version=ver
    )
    return df_swap, cycles, msg

# ==========================================================================================
# 뷰 헬퍼
# ==========================================================================================
@st.cache_data(show_spinner=False)
def teacher_matrix(version=0):
    tt = st.session_state.timetable
    if tt.empty:
        return pd.DataFrame()
    teachers = sorted(tt["교사명"].unique())
    rows = []
    for t in teachers:
        row = {"교사명": t}
        for d in DAYS:
            for p in range(1, PERIODS_PER_DAY.get(d, 7) + 1):
                m = tt[(tt["교사명"] == t) & (tt["요일"] == d) & (tt["교시"] == p)]
                row[f"{d}{p}"] = f"{m.iloc[0]['학급']} {m.iloc[0]['과목']}" if not m.empty else ""
        rows.append(row)
    return pd.DataFrame(rows)

@st.cache_data(show_spinner=False)
def class_matrix(version=0):
    tt = st.session_state.timetable
    if tt.empty:
        return pd.DataFrame()
    classes = sorted(tt["학급"].unique())
    rows = []
    for c in classes:
        row = {"학급": c}
        for d in DAYS:
            for p in range(1, PERIODS_PER_DAY.get(d, 7) + 1):
                m = tt[(tt["학급"] == c) & (tt["요일"] == d) & (tt["교시"] == p)]
                row[f"{d}{p}"] = f"{m.iloc[0]['교사명']} {m.iloc[0]['과목']}" if not m.empty else ""
        rows.append(row)
    return pd.DataFrame(rows)

def get_teacher_week_view(teacher: str, ref_date: date, use_test=False):
    weekday = ref_date.weekday()
    monday = ref_date - timedelta(days=weekday)
    week_dates = [monday + timedelta(days=i) for i in range(5)]
    ver = st.session_state.get("_data_version", 0)
    absences = st.session_state.get("absences", pd.DataFrame())
    subs = st.session_state.get("subs", pd.DataFrame())
    pt_df = st.session_state.get("part_time", pd.DataFrame())
    grid = []
    for p in range(1, MAX_PERIOD + 1):
        row = {"교시": p}
        for i, d in enumerate(DAYS):
            on_date = week_dates[i].strftime("%Y-%m-%d")
            e_tt = get_effective_timetable_for_date(on_date, ver, use_test=use_test)
            m = e_tt[(e_tt["교사명"] == teacher) & (e_tt["교시"] == p)] if not e_tt.empty else pd.DataFrame()
            if not m.empty:
                r = m.iloc[0]
                if r.get("원본교사"):
                    cell = f"{r['교사명']}({r['원본교사']}) {r['학급']} {r['과목']}"
                else:
                    cell = f"{r['학급']} {r['과목']}"
                origin = get_swap_origin_info(teacher, on_date, p)
                if origin:
                    cell += f" 🔄 from {origin}"
                if not absences.empty and ((absences["일자"] == on_date) & (absences["교사명"] == teacher) & (absences["교시"] == p)).any():
                    cell = f"[결강] {cell}"
                if not subs.empty and ((subs["일자"] == on_date) & (subs["보강교사"] == teacher) & (subs["교시"] == p)).any():
                    cell = f"[보강] {cell}"
                row[d] = cell
            else:
                cell = ""
                if not pt_df.empty and "시작일" in pt_df.columns:
                    for _, prow in pt_df.iterrows():
                        start = normalize_date_str(prow.get("시작일", ""))
                        end = normalize_date_str(prow.get("종료일", ""))
                        if start and end and start <= on_date <= end and str(prow.get("대체교사", "")).strip() == teacher:
                            pt_name = str(prow.get("시간강사명", "")).strip()
                            if pt_name:
                                m2 = e_tt[(e_tt["교사명"] == pt_name) & (e_tt["교시"] == p)] if not e_tt.empty else pd.DataFrame()
                                if not m2.empty:
                                    r2 = m2.iloc[0]
                                    cell = f"{pt_name}({teacher}) {r2['학급']} {r2['과목']}"
                                    break
                row[d] = cell
        grid.append(row)
    return pd.DataFrame(grid), week_dates

def get_changed_teachers_for_week(ref_date: date):
    weekday = ref_date.weekday()
    monday = ref_date - timedelta(days=weekday)
    week_dates = [(monday + timedelta(days=i)).strftime("%Y-%m-%d") for i in range(5)]
    changed = set()
    for df, cols in [(st.session_state.absences, ["교사명"]), (st.session_state.subs, ["결강교사", "보강교사"])]:
        if not df.empty and "일자" in df.columns:
            mask = df["일자"].isin(week_dates)
            for c in cols:
                if c in df.columns:
                    changed.update(df.loc[mask, c].dropna().tolist())
    if not st.session_state.swaps.empty:
        for col in ["원본일자", "목표일자"]:
            if col in st.session_state.swaps.columns:
                mask = st.session_state.swaps[col].isin(week_dates)
                changed.update(st.session_state.swaps.loc[mask, "교사A"].tolist())
                changed.update(st.session_state.swaps.loc[mask, "교사B"].tolist())
    pt = st.session_state.get("part_time", pd.DataFrame())
    if not pt.empty and "시작일" in pt.columns:
        for _, r in pt.iterrows():
            start = normalize_date_str(r.get("시작일", ""))
            end = normalize_date_str(r.get("종료일", ""))
            if start and end:
                for wd in week_dates:
                    if start <= wd <= end:
                        if r.get("대체교사"):
                            changed.add(str(r["대체교사"]).strip())
                        if r.get("시간강사명"):
                            changed.add(str(r["시간강사명"]).strip())
    return sorted(t for t in changed if t)

def filter_by_owner(df):
    if can_full_data() or df.empty or "입력자" not in df.columns:
        return df
    return df[df["입력자"] == current_user()].copy()

# ==========================================================================================
# ★ 결보강 계획서 (보강수업 칸에 실제 보강교사 표시)
# ==========================================================================================
def build_personal_plan_html(teacher_name: str, on_date: str) -> str:
    """
    공식 양식에 맞게:
    - 결강수업: 결강 정보
    - 보강수업 교사(인): 실제로 보강 배정된 교사
    - 교체수업: 교체 신청/맞교환 정보
    """
    try:
        dt = datetime.strptime(on_date, "%Y-%m-%d")
        day_kr = WEEKDAY_KR[dt.weekday()]
        date_display = f"{dt.year}년 {dt.month}월 {dt.day}일 ({day_kr})"
    except Exception:
        date_display = on_date
        day_kr = ""

    subject_dept = get_teacher_subject(teacher_name)
    dept_line = f"{subject_dept} 과" if subject_dept else "과"

    abs_df = st.session_state.get("absences", pd.DataFrame())
    my_abs = pd.DataFrame()
    if not abs_df.empty:
        my_abs = abs_df[(abs_df["교사명"] == teacher_name) & (abs_df["일자"] == on_date)].copy()

    # 보강 데이터 매칭 (같은 결강ID 또는 일자+교사+교시)
    subs_df = st.session_state.get("subs", pd.DataFrame())
    my_subs = pd.DataFrame()
    if not subs_df.empty:
        my_subs = subs_df[
            (subs_df["결강교사"] == teacher_name) & (subs_df["일자"] == on_date)
        ].copy()

    req_df = load_swap_requests()
    my_reqs = pd.DataFrame()
    if not req_df.empty:
        my_reqs = req_df[
            ((req_df["신청자이름"] == teacher_name) | (req_df["교사A"] == teacher_name)) &
            ((req_df["원본일자"] == on_date) | (req_df["목표일자"] == on_date))
        ].copy()

    swaps = st.session_state.get("swaps", pd.DataFrame())
    my_swaps = pd.DataFrame()
    if not swaps.empty:
        my_swaps = swaps[
            ((swaps["교사A"] == teacher_name) | (swaps["교사B"] == teacher_name)) &
            ((swaps["원본일자"] == on_date) | (swaps["목표일자"] == on_date))
        ].copy()

    reason = str(my_abs.iloc[0].get("사유", "")).strip() if not my_abs.empty else ""

    # 결강 + 보강 매칭
    abs_list = []
    if not my_abs.empty:
        for _, r in my_abs.iterrows():
            p = safe_int(r.get("교시"))
            # 해당 교시에 배정된 보강교사 찾기
            sub_teacher = ""
            if not my_subs.empty:
                match = my_subs[my_subs["교시"] == p]
                if not match.empty:
                    sub_teacher = str(match.iloc[0].get("보강교사", "")).strip()

            abs_list.append({
                "월일": on_date[5:].replace("-", "/"),
                "교시": p,
                "학년반": str(r.get("학급", "")),
                "과목": str(r.get("과목", "")),
                "보강교사": sub_teacher          # ← 여기가 핵심 (신청자가 아닌 보강교사)
            })

    # 교체 목록
    swap_list = []
    for src in [my_reqs, my_swaps]:
        if not src.empty:
            for _, r in src.iterrows():
                target_date = str(r.get("목표일자", r.get("원본일자", "")))
                swap_list.append({
                    "월일": target_date[5:].replace("-", "/") if len(target_date) >= 10 else target_date,
                    "교시": safe_int(r.get("교시B", r.get("교시A", 0))),
                    "과목": str(r.get("과목B", r.get("과목A", ""))),
                    "교사": str(r.get("교사B", r.get("교사A", "")))
                })

    # 9칸 구조로 행 생성
    rows_html = ""
    max_rows = 6
    for i in range(max_rows):
        a = abs_list[i] if i < len(abs_list) else {"월일": "", "교시": "", "학년반": "", "과목": "", "보강교사": ""}
        s = swap_list[i] if i < len(swap_list) else {"월일": "", "교시": "", "과목": "", "교사": ""}
        rows_html += f"""
        <tr>
            <td style="height:29px;">{a['월일']}</td>
            <td>{a['교시']}</td>
            <td>{a['학년반']}</td>
            <td>{a['과목']}</td>
            <td>{a['보강교사']}</td>
            <td>{s['월일']}</td>
            <td>{s['교시']}</td>
            <td>{s['과목']}</td>
            <td>{s['교사']}</td>
        </tr>"""

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>결·보강 계획서 - {teacher_name}</title>
<style>
    @page {{ size: A4; margin: 12mm 14mm; }}
    * {{ box-sizing: border-box; }}
    body {{
        font-family: '맑은 고딕', 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif;
        font-size: 12.5px;
        line-height: 1.25;
        margin: 0;
        padding: 6px 10px;
        color: #000;
        background: #fff;
    }}
    table {{
        border-collapse: collapse;
        width: 100%;
        table-layout: fixed;
    }}
    th, td {{
        border: 1px solid #000;
        padding: 2px 2px;
        text-align: center;
        vertical-align: middle;
        word-break: keep-all;
    }}
    .title {{
        text-align: center;
        font-size: 22px;
        font-weight: bold;
        letter-spacing: 5px;
        margin: 2px 0 8px 0;
        text-decoration: underline;
        text-underline-offset: 3px;
    }}
    .top-right {{
        width: 150px;
        float: right;
        margin-top: -36px;
    }}
    .top-right td {{
        height: 26px;
        font-size: 12px;
        font-weight: bold;
    }}
    .dept-line {{
        font-size: 13.5px;
        margin: 4px 0 6px 2px;
    }}
    .date-box td {{
        height: 32px;
        font-size: 12.5px;
    }}
    .section-header {{
        background-color: #d6e3f0;
        font-weight: bold;
        font-size: 12px;
    }}
    .sub-header {{
        background-color: #eef3f9;
        font-size: 11.5px;
        font-weight: bold;
    }}
    .note-box {{
        border: 1px solid #000;
        min-height: 88px;
        margin-top: 0;
        padding: 6px;
    }}
</style>
</head>
<body>

<div class="title">결 · 보 강  계 획</div>

<table class="top-right">
    <tr>
        <td style="width:50%;">수업계</td>
        <td style="width:50%;">교육과정</td>
    </tr>
    <tr>
        <td style="height:34px;"></td>
        <td></td>
    </tr>
</table>

<div class="dept-line">
    <b>{dept_line}</b> &nbsp;&nbsp; 교 사 : <b>{teacher_name}</b> &nbsp;&nbsp;&nbsp; (인)
</div>

<table class="date-box" style="margin-bottom: 9px;">
    <tr>
        <td style="width: 68px; background:#f0f0f0; font-weight:bold;">결강<br>일자</td>
        <td style="text-align:left; padding-left:10px;">
            {date_display}<br>
            <span style="display:inline-block; margin-top:2px;">사유 : {reason}</span>
        </td>
    </tr>
</table>

<table>
    <thead>
        <tr>
            <th colspan="4" class="section-header">결강수업</th>
            <th class="section-header">보강수업</th>
            <th colspan="4" class="section-header">교체수업</th>
        </tr>
        <tr class="sub-header">
            <th style="width:9.5%;">월일</th>
            <th style="width:7%;">교시</th>
            <th style="width:10%;">학년반</th>
            <th style="width:11%;">과목</th>
            <th style="width:12%;">교사(인)</th>
            <th style="width:9.5%;">월일</th>
            <th style="width:7%;">교시</th>
            <th style="width:11%;">과목</th>
            <th style="width:13%;">교사(인)</th>
        </tr>
    </thead>
    <tbody>
        {rows_html}
    </tbody>
</table>

<div style="margin-top: 13px;">
    <div style="text-align:center; font-weight:bold; border:1px solid #000; border-bottom:none; padding:5px 0;">
        추가 기재 사항
    </div>
    <div class="note-box"></div>
</div>

<div style="margin-top: 10px; font-size: 10.5px; color:#555; text-align:right;">
    {SCHOOL_NAME} · {SCHOOL_YEAR}학년도 &nbsp;|&nbsp; 생성시각 {datetime.now().strftime('%Y-%m-%d %H:%M')}
</div>

</body>
</html>"""
    return html

def build_report_html(norm_date: str) -> str:
    day = WEEKDAY_KR.get(datetime.strptime(norm_date, "%Y-%m-%d").weekday(), "")
    a = st.session_state.absences
    s = st.session_state.subs
    w = st.session_state.swaps
    def rows_abs():
        if a.empty or "일자" not in a.columns:
            return "<tr><td colspan='6'>없음</td></tr>"
        sub = a[a["일자"] == norm_date]
        if sub.empty:
            return "<tr><td colspan='6'>없음</td></tr>"
        return "".join(
            f"<tr><td>{r.get('교사명','')}</td><td>{r.get('사유','')}</td><td>{safe_int(r.get('교시'))}</td>"
            f"<td>{r.get('학급','')}</td><td>{r.get('과목','')}</td><td>{r.get('상세사유','')}</td></tr>"
            for _, r in sub.iterrows()
        )
    def rows_sub():
        if s.empty or "일자" not in s.columns:
            return "<tr><td colspan='7'>없음</td></tr>"
        sub = s[s["일자"] == norm_date]
        if sub.empty:
            return "<tr><td colspan='7'>없음</td></tr>"
        return "".join(
            f"<tr><td>{safe_int(r.get('교시'))}</td><td>{r.get('학급','')}</td><td>{r.get('과목','')}</td>"
            f"<td>{r.get('결강교사','')}</td><td><b>{r.get('보강교사','')}</b></td>"
            f"<td>{r.get('우선순위','')}</td><td>{r.get('비고','')}</td></tr>"
            for _, r in sub.iterrows()
        )
    def rows_swap():
        if w.empty:
            return "<tr><td colspan='4'>없음</td></tr>"
        mask = (w["원본일자"] == norm_date) | (w["목표일자"] == norm_date)
        sub = w[mask]
        if sub.empty:
            return "<tr><td colspan='4'>없음</td></tr>"
        return "".join(
            f"<tr><td>{r.get('교사A','')}</td>"
            f"<td>[{r.get('원본일자')}] {r.get('요일A')} {safe_int(r.get('교시A'))}교시 ↔ "
            f"[{r.get('목표일자')}] {r.get('요일B')} {safe_int(r.get('교시B'))}교시</td>"
            f"<td>{r.get('교사B','')}</td><td>{r.get('유형','')}</td></tr>"
            for _, r in sub.iterrows()
        )
    return f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">
<title>{norm_date} 내역서</title>
<style>
body{{font-family:'맑은 고딕',sans-serif;font-size:12px}}
table{{width:100%;border-collapse:collapse}}
th,td{{border:1px solid #999;padding:5px;text-align:center}}
th{{background:#eef1f6}}
</style></head><body>
<h1 style="text-align:center">결강·보강 변경 내역서</h1>
<div style="text-align:center">{SCHOOL_YEAR} · {SCHOOL_NAME} · {norm_date} ({day})</div>
<h3>1. 결강 현황</h3>
<table><tr><th>교사</th><th>사유</th><th>교시</th><th>학급</th><th>과목</th><th>상세</th></tr>{rows_abs()}</table>
<h3>2. 보강 배정</h3>
<table><tr><th>교시</th><th>학급</th><th>과목</th><th>결강교사</th><th>보강교사</th><th>근거</th><th>비고</th></tr>{rows_sub()}</table>
<h3>3. 맞교환</h3>
<table><tr><th>교사A</th><th>내용</th><th>교사B</th><th>유형</th></tr>{rows_swap()}</table>
</body></html>"""

def to_excel_bytes(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, df in sheets.items():
            (df if isinstance(df, pd.DataFrame) else pd.DataFrame(df)).to_excel(w, sheet_name=name[:31], index=False)
    return buf.getvalue()

# ==========================================================================================
# 로그인 페이지
# ==========================================================================================
def show_login_page():
    if "login_attempts" not in st.session_state:
        st.session_state.login_attempts = 0
    if "login_locked" not in st.session_state:
        st.session_state.login_locked = False

    st.markdown('<div class="login-box">', unsafe_allow_html=True)
    st.title(f"📘 {SCHOOL_NAME}")
    st.subheader("시간표 · 결보강 관리 시스템")
    st.caption(f"{SCHOOL_YEAR}학년도")

    if st.session_state.login_locked:
        st.error(f"🚫 로그인 시도가 {MAX_LOGIN_ATTEMPTS}회를 초과하여 차단되었습니다.")
        st.stop()

    id_input = st.text_input("아이디", placeholder="아이디를 입력하세요", key="login_id")
    col1, col2 = st.columns(2)
    with col1:
        if st.button("로그인", type="primary", use_container_width=True):
            uid = id_input.strip()
            if not uid:
                st.error("아이디를 입력해주세요.")
            else:
                ids = load_id_sheet()
                match = ids[ids["아이디"].astype(str).str.strip() == uid]
                if not match.empty:
                    st.session_state.login_attempts = 0
                    row = match.iloc[0]
                    role = str(row["권한"]).strip() or ROLE_TEACHER
                    name = str(row.get("이름", "")).strip() or uid
                    allowed = str(row.get("허용탭", "")).strip()
                    st.session_state.logged_in = True
                    st.session_state.user_id = uid
                    st.session_state.user_name = name
                    st.session_state.user_role = role
                    st.session_state.user_allowed_tabs = allowed
                    st.rerun()
                else:
                    st.session_state.login_attempts += 1
                    remaining = MAX_LOGIN_ATTEMPTS - st.session_state.login_attempts
                    if st.session_state.login_attempts >= MAX_LOGIN_ATTEMPTS:
                        st.session_state.login_locked = True
                        st.error(f"🚫 로그인 실패 {MAX_LOGIN_ATTEMPTS}회 초과로 차단되었습니다.")
                        st.rerun()
                    else:
                        st.error(f"등록되지 않은 아이디입니다. (남은 시도: {remaining}회)")
    with col2:
        if st.button("게스트로 입장", use_container_width=True):
            st.session_state.logged_in = True
            st.session_state.user_id = "게스트"
            st.session_state.user_name = "게스트"
            st.session_state.user_role = ROLE_GUEST
            st.session_state.user_allowed_tabs = ""
            st.session_state.login_attempts = 0
            st.rerun()

    if st.session_state.login_attempts > 0:
        st.warning(f"현재 로그인 실패 횟수: {st.session_state.login_attempts} / {MAX_LOGIN_ATTEMPTS}")

    st.divider()
    st.markdown("#### 📝 아이디 추가 요청 (게스트용)")
    with st.form("id_request_form"):
        name = st.text_input("이름 *")
        email = st.text_input("이메일 *")
        desired = st.text_input("추가하고 싶은 아이디 *")
        memo = st.text_area("메모 (요청 사유 등)")
        submitted = st.form_submit_button("요청 제출", type="primary")
        if submitted:
            if not (name and email and desired):
                st.error("이름, 이메일, 아이디는 필수입니다.")
            else:
                save_id_request(name, email, desired, memo)
                st.success("요청이 정상적으로 접수되었습니다.")
    st.markdown('</div>', unsafe_allow_html=True)

# ==========================================================================================
# 앱 시작
# ==========================================================================================
if "logged_in" not in st.session_state:
    st.session_state.logged_in = False
    st.session_state.user_id = None
    st.session_state.user_name = ""
    st.session_state.user_role = ROLE_GUEST
    st.session_state.user_allowed_tabs = ""

if not st.session_state.logged_in:
    show_login_page()
    st.stop()

init_state()

# ==========================================================================================
# 사이드바
# ==========================================================================================
with st.sidebar:
    st.markdown(f"**아이디** : `{current_user()}`")
    st.markdown(f"**이름** : `{current_name()}`")
    st.markdown(f"**권한** : `{current_role()}`")
    if st.button("로그아웃", use_container_width=True):
        for k in list(st.session_state.keys()):
            del st.session_state[k]
        st.rerun()

    st.divider()

    if current_role() != ROLE_GUEST:
        st.header("데이터")
        if can_full_data() or is_teacher():
            if st.button("🔄 원본 시간표 다시 불러오기", use_container_width=True):
                load_timetable_from_gsheet.clear()
                ti, tt = load_timetable_from_gsheet()
                st.session_state.teachers = ti
                st.session_state.timetable = tt
                _invalidate_all_caches()
                st.rerun()
            if st.button("🔄 작업 내역 다시 불러오기", use_container_width=True):
                load_work_data_from_gsheet.clear()
                absences, subs, swaps, part_time, cumulative, duties = load_work_data_from_gsheet()
                st.session_state.absences = ensure_input_user(absences)
                st.session_state.subs = ensure_input_user(subs)
                st.session_state.swaps = ensure_input_user(swaps)
                st.session_state.part_time = ensure_part_time_columns(part_time)
                st.session_state.duties = ensure_duty_columns(duties)
                _invalidate_all_caches()
                st.rerun()
            if st.button("💾 현재 작업 저장", use_container_width=True, type="primary"):
                if save_work_data_to_gsheet():
                    st.success("저장 완료")

        st.divider()
        c1, c2 = st.columns(2)
        if c1.button("↩ Undo", use_container_width=True):
            if undo():
                save_work_data_to_gsheet()
                st.rerun()
        if c2.button("↪ Redo", use_container_width=True):
            if redo():
                save_work_data_to_gsheet()
                st.rerun()

        st.divider()
        st.metric("등록 교사", len(st.session_state.teachers))
        st.metric("누적 보강", len(st.session_state.subs))

        if is_edu_or_master():
            try:
                curr_budget = get_current_budget()
                st.metric("보강비 잔액", f"{curr_budget:,.0f}원")
            except Exception:
                pass

        st.divider()
        st.subheader("📄 내역서 / 계획서 출력")

        plan_date = st.date_input("계획서 기준일", value=date.today(), key="plan_date")
        if st.button("📋 결보강 계획서 (본인용)", use_container_width=True, type="primary"):
            html = build_personal_plan_html(current_name(), plan_date.strftime("%Y-%m-%d"))
            st.download_button(
                "HTML 다운로드 (인쇄 → PDF로 저장 추천)",
                html.encode("utf-8"),
                f"결보강계획서_{current_name()}_{plan_date}.html",
                "text/html",
                key="dl_personal"
            )
            st.info("다운로드한 HTML을 브라우저에서 열고 Ctrl+P → PDF로 저장하시면 양식과 거의 동일한 PDF가 생성됩니다.")

        if is_edu_or_master():
            rd = st.date_input("전체 내역서 일자", value=date.today(), key="sidebar_rd")
            if st.button("📊 전체 일일 내역서", use_container_width=True):
                html = build_report_html(rd.strftime("%Y-%m-%d"))
                st.download_button("HTML 다운로드 (전체)", html.encode("utf-8"), f"내역서_{rd}.html", "text/html")
                xls = to_excel_bytes({
                    "결강": st.session_state.absences[st.session_state.absences["일자"] == rd.strftime("%Y-%m-%d")] if not st.session_state.absences.empty else pd.DataFrame(),
                    "보강": st.session_state.subs[st.session_state.subs["일자"] == rd.strftime("%Y-%m-%d")] if not st.session_state.subs.empty else pd.DataFrame(),
                    "맞교환": st.session_state.swaps
                })
                st.download_button("엑셀 다운로드", xls, f"내역서_{rd}.xlsx")

if current_role() == ROLE_GUEST:
    st.title("게스트 모드")
    st.info("현재 게스트로 접속 중입니다. 아이디 추가 요청만 가능합니다.")
    with st.form("guest_request"):
        name = st.text_input("이름 *")
        email = st.text_input("이메일 *")
        desired = st.text_input("추가하고 싶은 아이디 *")
        memo = st.text_area("메모")
        if st.form_submit_button("요청 제출", type="primary"):
            if name and email and desired:
                save_id_request(name, email, desired, memo)
                st.success("요청이 접수되었습니다.")
            else:
                st.error("필수 항목을 입력해주세요.")
    st.stop()

if st.session_state.timetable.empty:
    st.info("시간표를 불러오는 중...")
    st.stop()

# ==========================================================================================
# 메인 화면
# ==========================================================================================
st.title(f"시간표 · 결강/보강 관리  |  {current_name()} ({current_user()}) · {current_role()}")

allowed_tabs = get_user_allowed_tabs()
visible_tabs = [t for t in ALL_TABS if t in allowed_tabs]

if can_manage_ids():
    for t in ["🔑 아이디·권한 관리", "📑 회원별 탭 권한 관리", "🛠️ 다중 출장·전체 조정 추천"]:
        if t not in visible_tabs:
            visible_tabs.append(t)

if not visible_tabs:
    st.warning("접근 가능한 탭이 없습니다.")
    st.stop()

tabs = st.tabs(visible_tabs)
tab_map = {name: tabs[i] for i, name in enumerate(visible_tabs)}

# ------------------------------------------------------------------ 시간표 조회
if "시간표 조회" in tab_map:
    with tab_map["시간표 조회"]:
        view = st.radio("보기 방식", ["교사별 전체 매트릭스", "학급별 매트릭스", "교사 1인 주간표"], horizontal=True)
        ver = st.session_state.get("_data_version", 0)
        if view == "교사별 전체 매트릭스":
            st.caption("※ 원본 시간표 기준. 변경사항은 주간표에서 확인")
            st.dataframe(teacher_matrix(ver), use_container_width=True, height=700, hide_index=True)
        elif view == "학급별 매트릭스":
            st.caption("※ 원본 시간표 기준. 변경사항은 주간표에서 확인")
            st.dataframe(class_matrix(ver), use_container_width=True, height=700, hide_index=True)
        else:
            tlist = get_all_teacher_names()
            t = st.selectbox("교사 선택", tlist, key="view_t")
            ref = st.date_input("기준 날짜", value=date.today(), key="view_ref")
            grid, dates = get_teacher_week_view(t, ref)
            st.caption(f"{dates[0]} ~ {dates[4]}  |  시간강사 대체 시 '시간강사명(원본교사)' 형태")
            st.dataframe(grid, use_container_width=True, hide_index=True)

# ------------------------------------------------------------------ 시간강사 관리
if "시간강사 관리" in tab_map:
    with tab_map["시간강사 관리"]:
        st.subheader("시간강사 등록 및 가능 시간표")
        st.info("시작일·종료일·대체교사를 입력하면 해당 기간 동안 원본 교사 수업이 시간강사로 대체됩니다.")
        if can_full_data() or is_teacher():
            ed = st.data_editor(
                st.session_state.part_time,
                num_rows="dynamic",
                use_container_width=True,
                height=450,
                key="pt_ed",
                hide_index=True,
                column_config={
                    "시작일": st.column_config.TextColumn("시작일 (YYYY-MM-DD)"),
                    "종료일": st.column_config.TextColumn("종료일 (YYYY-MM-DD)"),
                    "대체교사": st.column_config.TextColumn("대체할 정규교사명"),
                }
            )
            if st.button("시간강사 정보 저장", type="primary"):
                st.session_state.part_time = ensure_part_time_columns(ed)
                push_history("시간강사 수정")
                save_work_data_to_gsheet()
                st.success("저장 완료")
                st.rerun()
        else:
            st.dataframe(st.session_state.part_time, use_container_width=True)

# ------------------------------------------------------------------ 결강·보강
if "결강·보강" in tab_map:
    with tab_map["결강·보강"]:
        st.subheader("결강 등록 & 보강 배정")
        left, right = st.columns(2)

        with left:
            st.markdown("### 📌 결강 등록")
            d_sel = st.date_input("결강 일자", value=date.today(), key="abs_d")
            on_date = d_sel.strftime("%Y-%m-%d")
            day = WEEKDAY_KR[d_sel.weekday()]
            who = st.selectbox("결강 교사", st.session_state.teachers["교사명"].tolist(), key="abs_who")
            reason = st.selectbox("사유", ABSENCE_REASONS, key="abs_reason")
            detail = st.text_input("상세 사유", key="abs_detail")

            ver = st.session_state.get("_data_version", 0)
            e_tt = get_effective_timetable_for_date(on_date, ver)
            todays = e_tt[(e_tt["교사명"] == who) & (e_tt["요일"] == day)].sort_values("교시")
            if not todays.empty:
                opts = [f"{safe_int(r.교시)}교시 · {r.학급} · {r.과목}" for r in todays.itertuples()]
                picked = st.multiselect("결강 교시 선택", opts, default=opts, key="abs_pick")
                if st.button("결강 등록", type="primary", key="btn_abs"):
                    push_history(f"결강 ({who})")
                    cid = f"{on_date}-{who}"
                    sel_p = [safe_int(o.split("교시")[0]) for o in picked]
                    rows = todays[todays["교시"].isin(sel_p)]
                    a = st.session_state.absences
                    if not a.empty:
                        a = a[~((a["일자"] == on_date) & (a["교사명"] == who))]
                    new_rows = [{
                        "결강ID": cid, "일자": on_date, "요일": day, "교사명": who,
                        "사유": reason, "상세사유": detail, "교시": safe_int(r.교시),
                        "학급": r.학급, "과목": r.과목, "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M"),
                        "입력자": current_user()
                    } for r in rows.itertuples()]
                    st.session_state.absences = pd.concat([a, pd.DataFrame(new_rows)], ignore_index=True)
                    save_work_data_to_gsheet()
                    st.success("결강 등록 완료")
                    st.rerun()
            show_abs = filter_by_owner(st.session_state.absences)
            st.dataframe(
                show_abs[show_abs["일자"] == on_date] if not show_abs.empty else pd.DataFrame(),
                height=250, hide_index=True
            )

        with right:
            st.markdown("### 📌 보강 배정")
            ab = filter_by_owner(st.session_state.absences)
            if ab.empty:
                st.info("먼저 결강을 등록하세요." if can_full_data() else "본인이 등록한 결강이 없습니다.")
            else:
                cids = sorted(ab["결강ID"].dropna().unique(), reverse=True)
                cid = st.selectbox("결강 건 선택", cids, key="sub_cid")
                rows = ab[ab["결강ID"] == cid].sort_values("교시")
                head = rows.iloc[0]
                st.markdown(f"**{head['일자']} · {head['교사명']} · {len(rows)}시간**")
                include_pt = st.checkbox("시간강사 포함", key="sub_pt")
                show_all = st.checkbox("전체 공강 교사 보기 (최대 20명)", value=True, key="show_all_free")

                if st.button("전 교시 자동 배정", type="primary", key="auto_all"):
                    push_history("자동 보강")
                    for r in rows.itertuples():
                        cand = recommend_substitutes(head["요일"], r.교시, r.과목, r.학급, head["교사명"], head["일자"], top_n=1, include_part_time=include_pt)
                        if not cand.empty:
                            add_substitute(cid, head["일자"], head["요일"], r.교시, r.학급, r.과목, head["교사명"], cand.iloc[0]["보강교사"], "자동", cand.iloc[0]["우선순위"], "")
                    st.rerun()

                for r in rows.itertuples():
                    p = safe_int(r.교시)
                    with st.expander(f"{p}교시 · {r.학급} · {r.과목}", expanded=True):
                        cur = st.session_state.subs
                        assigned = None
                        if not cur.empty:
                            m = cur[(cur["결강ID"] == cid) & (cur["교시"] == p)]
                            if not m.empty:
                                assigned = m.iloc[0]["보강교사"]
                        if assigned:
                            st.success(f"현재 배정: **{assigned}**")
                            if st.button("배정 취소", key=f"cancel_{cid}_{p}"):
                                cancel_substitute(cid, p)
                                st.rerun()
                        else:
                            cand = recommend_substitutes(
                                head["요일"], p, r.과목, r.학급, head["교사명"], head["일자"],
                                top_n=20 if show_all else 8,
                                include_part_time=include_pt
                            )
                            if cand.empty:
                                st.warning("해당 시간대에 비어있는 교사가 없습니다.")
                            else:
                                st.dataframe(cand, hide_index=True, height=220)
                                pick = st.selectbox("보강 교사 선택", cand["보강교사"], key=f"pick_{cid}_{p}")
                                if st.button("배정", key=f"btn_{cid}_{p}"):
                                    pr = cand.loc[cand["보강교사"] == pick, "우선순위"].iloc[0]
                                    add_substitute(cid, head["일자"], head["요일"], p, r.학급, r.과목, head["교사명"], pick, "수동", pr, "")
                                    st.rerun()
# ------------------------------------------------------------------ 시간표 맞교환
if "시간표 맞교환 & 변경 추천" in tab_map:
    with tab_map["시간표 맞교환 & 변경 추천"]:
        st.markdown("### 🔄 스마트 시간표 변경 & 맞교환")
        col_a, col_b = st.columns(2)
        tlist = st.session_state.teachers["교사명"].tolist()

        with col_a:
            st.markdown("#### 1️⃣ 원본 수업")
            date_a = st.date_input("원본 날짜", value=date.today(), key="sw_da")
            date_a_str = date_a.strftime("%Y-%m-%d")
            day_a = WEEKDAY_KR[date_a.weekday()]
            t_a = st.selectbox("교사 A", tlist, key="sw_ta")
            ver = st.session_state.get("_data_version", 0)
            e_a = get_effective_timetable_for_date(date_a_str, ver)
            sub_a = e_a[(e_a["교사명"] == t_a) & (e_a["요일"] == day_a)].sort_values("교시")
            pick_a = None
            if not sub_a.empty:
                opts = [f"{safe_int(r.교시)}교시 · {r.학급} · {r.과목}" for r in sub_a.itertuples()]
                sel = st.selectbox("변경할 수업", opts, key="sw_la")
                row = sub_a.iloc[opts.index(sel)]
                pick_a = {"교사명": t_a, "일자": date_a_str, "요일": day_a, "교시": safe_int(row.교시),
                          "학급": row.학급, "과목": row.과목}

        with col_b:
            st.markdown("#### 2️⃣ 이동 희망 시간")
            date_b = st.date_input("이동 희망 날짜", value=date.today(), key="sw_db")
            date_b_str = date_b.strftime("%Y-%m-%d")
            day_b = WEEKDAY_KR[date_b.weekday()]
            p_b = st.selectbox("희망 교시", list(range(1, PERIODS_PER_DAY.get(day_b, 7)+1)), key="sw_pb")

        is_pt = st.checkbox("시간강사 구인 목적")

        if pick_a:
            st.subheader(f"{pick_a['교시']}교시 → {date_b_str} {p_b}교시")
            df_swap, cycles, cycle_msg = get_target_time_recommendations(
                pick_a["교사명"], pick_a["일자"], pick_a["교시"], pick_a["학급"], pick_a["과목"],
                date_b_str, p_b
            )
            t1, t2 = st.tabs(["1:1 맞교환 (동일 학급 최우선)", "연계 순환 교환 (2~4인)"])
            with t1:
                if df_swap.empty:
                    st.info("추천 없음")
                else:
                    for idx, row in df_swap.iterrows():
                        mark = "🏆 동일학급" if row.get("same_class") else ("⚠ 같은학년" if row.get("same_grade") else "⚠ 다른학년")
                        with st.expander(f"{mark} {row['교사B']} ({row['현재 수업']})", expanded=(idx==0)):
                            if st.button(f"{row['교사B']}와 맞교환", key=f"sw_{idx}"):
                                do_swap(pick_a, row["b_info"], date_a_str, date_b_str, is_pt)
                                st.rerun()
            with t2:
                st.caption(cycle_msg)
                if not cycles:
                    st.info("조건을 만족하는 순환 경로가 없습니다.")
                else:
                    for idx, cyc in enumerate(cycles):
                        with st.expander(f"{'✅' if cyc['length']==2 else '🔗'} {cyc['length']}인 순환 · 점수 {cyc['score']}", expanded=(idx==0)):
                            st.markdown(f"**경로**: `{cyc['path_desc']}`")
                            st.caption("학급의 담당교사·과목·시수가 완전히 보존됩니다.")
                            if st.button(f"이 순환 적용하기", key=f"cyc_{idx}"):
                                apply_cycle_swaps(cyc["moves"], is_test=False)
                                st.success(f"{cyc['length']}인 순환 교환이 등록되었습니다.")
                                st.rerun()

        show_swaps = filter_by_owner(st.session_state.swaps)
        st.dataframe(show_swaps, use_container_width=True, hide_index=True)


# ------------------------------------------------------------------ 통계
if "통계" in tab_map:
    with tab_map["통계"]:
        st.subheader("보강 통계")
        c1, c2, c3 = st.columns(3)
        start = c1.date_input("시작일", value=date(2026, 3, 1), key="st_s")
        end = c2.date_input("종료일", value=date.today(), key="st_e")
        period = c3.selectbox("빠른 선택", ["전체", "1학기", "2학기", "이번 달"], key="st_p")
        if period == "1학기":
            start, end = date(2026, 3, 1), date(2026, 7, 31)
        elif period == "2학기":
            start, end = date(2026, 8, 1), date(2027, 2, 28)
        elif period == "이번 달":
            start = date.today().replace(day=1)
        cum = cumulative_sub_count(start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d"), st.session_state.get("_data_version", 0))
        df = pd.DataFrame({"교사명": list(cum.keys()), "누적보강": list(cum.values())}).sort_values("누적보강", ascending=False)
        st.dataframe(df, use_container_width=True, hide_index=True)
        if not df.empty:
            st.bar_chart(df.set_index("교사명")["누적보강"])

# ------------------------------------------------------------------ 시간표 변경 테스트용
if "시간표 변경 테스트용" in tab_map:
    with tab_map["시간표 변경 테스트용"]:
        st.subheader("🧪 시간표 변경 테스트용 (저장 안 됨 · 샌드박스)")
        st.info("연계공강(순환) 시 **수업계 선생님에게 연락해주세요**")

        if st.button("🔄 테스트 상태 초기화", type="secondary"):
            st.session_state.test_swaps = pd.DataFrame()
            st.success("테스트 상태가 초기화되었습니다.")
            st.rerun()

        col_a, col_b = st.columns(2)
        tlist = st.session_state.teachers["교사명"].tolist()

        with col_a:
            st.markdown("#### 1️⃣ 원본 수업 (테스트)")
            date_a = st.date_input("원본 날짜", value=date.today(), key="test_sw_da")
            date_a_str = date_a.strftime("%Y-%m-%d")
            day_a = WEEKDAY_KR[date_a.weekday()]
            t_a = st.selectbox("교사 A", tlist, key="test_sw_ta")
            ver = st.session_state.get("_data_version", 0)
            e_a = get_effective_timetable_for_date(date_a_str, ver, use_test=True)
            sub_a = e_a[(e_a["교사명"] == t_a) & (e_a["요일"] == day_a)].sort_values("교시")
            pick_a = None
            if not sub_a.empty:
                opts = [f"{safe_int(r.교시)}교시 · {r.학급} · {r.과목}" for r in sub_a.itertuples()]
                sel = st.selectbox("변경할 수업", opts, key="test_sw_la")
                row = sub_a.iloc[opts.index(sel)]
                pick_a = {"교사명": t_a, "일자": date_a_str, "요일": day_a, "교시": safe_int(row.교시),
                          "학급": row.학급, "과목": row.과목}

        with col_b:
            st.markdown("#### 2️⃣ 이동 희망 시간 (테스트)")
            date_b = st.date_input("이동 희망 날짜", value=date.today(), key="test_sw_db")
            date_b_str = date_b.strftime("%Y-%m-%d")
            day_b = WEEKDAY_KR[date_b.weekday()]
            p_b = st.selectbox("희망 교시", list(range(1, PERIODS_PER_DAY.get(day_b, 7)+1)), key="test_sw_pb")

        if pick_a:
            st.subheader(f"[테스트] {pick_a['교시']}교시 → {date_b_str} {p_b}교시")
            df_swap, cycles, cycle_msg = get_target_time_recommendations(
                pick_a["교사명"], pick_a["일자"], pick_a["교시"], pick_a["학급"], pick_a["과목"],
                date_b_str, p_b
            )
            t1, t2 = st.tabs(["1:1 맞교환 테스트", "연계 순환 교환 테스트"])
            with t1:
                if df_swap.empty:
                    st.info("추천 없음")
                else:
                    for idx, row in df_swap.iterrows():
                        mark = "🏆 동일학급" if row.get("same_class") else ("⚠ 같은학년" if row.get("same_grade") else "⚠ 다른학년")
                        with st.expander(f"{mark} {row['교사B']} ({row['현재 수업']})", expanded=(idx==0)):
                            if st.button(f"[테스트] {row['교사B']}와 맞교환", key=f"test_sw_{idx}"):
                                do_swap(pick_a, row["b_info"], date_a_str, date_b_str, is_test=True)
                                st.success("테스트 맞교환이 적용되었습니다. (저장 안 됨)")
                                st.rerun()
            with t2:
                st.caption(cycle_msg)
                if not cycles:
                    st.info("테스트용 순환 경로 없음")
                else:
                    for idx, cyc in enumerate(cycles):
                        with st.expander(f"{'✅' if cyc['length']==2 else '🔗'} {cyc['length']}인 순환", expanded=(idx==0)):
                            st.markdown(f"**경로**: `{cyc['path_desc']}`")
                            if st.button(f"[테스트] 이 순환 적용", key=f"test_cyc_{idx}"):
                                apply_cycle_swaps(cyc["moves"], is_test=True)
                                st.success(f"테스트 {cyc['length']}인 순환이 적용되었습니다. (저장 안 됨)")
                                st.rerun()


        st.markdown("#### 테스트 적용 후 주간표 미리보기")
        t_preview = st.selectbox("미리볼 교사", tlist, key="test_preview_t")
        ref_preview = st.date_input("미리보기 기준일", value=date.today(), key="test_preview_d")
        grid, dates = get_teacher_week_view(t_preview, ref_preview, use_test=True)
        st.caption(f"{dates[0]} ~ {dates[4]}  (테스트 반영됨)")
        st.dataframe(grid, use_container_width=True, height=350, hide_index=True)

        if not st.session_state.get("test_swaps", pd.DataFrame()).empty:
            st.markdown("#### 현재 테스트 중인 맞교환 목록")
            st.dataframe(st.session_state.test_swaps, use_container_width=True, hide_index=True)

            if st.button("✅ 교체 신청하기 (수업교체신청 시트로 저장)", type="primary"):
                for _, row in st.session_state.test_swaps.iterrows():
                    rec = {
                        "신청ID": f"REQ-{datetime.now().strftime('%Y%m%d%H%M%S')}-{current_user()}",
                        "신청자": current_user(),
                        "신청자이름": current_name(),
                        "원본일자": row.get("원본일자", ""),
                        "교사A": row.get("교사A", ""),
                        "요일A": row.get("요일A", ""),
                        "교시A": row.get("교시A", ""),
                        "학급A": row.get("학급A", ""),
                        "과목A": row.get("과목A", ""),
                        "목표일자": row.get("목표일자", ""),
                        "교사B": row.get("교사B", ""),
                        "요일B": row.get("요일B", ""),
                        "교시B": row.get("교시B", ""),
                        "학급B": row.get("학급B", ""),
                        "과목B": row.get("과목B", ""),
                        "유형": row.get("유형", "1:1 맞교환"),
                        "신청시각": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                        "상태": "신청"
                    }
                    save_swap_request(rec)
                st.success("교체 신청이 「수업교체신청」 시트에 저장되었습니다.")
                st.session_state.test_swaps = pd.DataFrame()
                st.rerun()

# ------------------------------------------------------------------ 변경된 교사 주간표
if "변경된 교사 주간표" in tab_map:
    with tab_map["변경된 교사 주간표"]:
        st.subheader("📅 변경된 교사 주간 시간표")
        ref = st.date_input("기준 날짜", value=date.today(), key="chg_ref")
        changed = get_changed_teachers_for_week(ref)
        if not changed:
            st.success("이번 주차 변경 교사 없음")
        else:
            st.info(f"변경 교사 {len(changed)}명: {', '.join(changed)}")
            for t in changed:
                with st.expander(f"👤 {t}", expanded=False):
                    grid, _ = get_teacher_week_view(t, ref)
                    st.dataframe(grid, use_container_width=True, hide_index=True)

# ------------------------------------------------------------------ 복무 관리 & 판단
if "📋 복무 관리 & 판단" in tab_map:
    with tab_map["📋 복무 관리 & 판단"]:
        st.subheader("📋 복무 관리 & 판단")
        left, right = st.columns([1, 1.4])

        with left:
            st.markdown("### 복무 등록")
            if can_full_data():
                teacher_options = st.session_state.teachers["교사명"].tolist()
            else:
                my_name = current_name()
                if my_name and my_name in st.session_state.teachers["교사명"].tolist():
                    teacher_options = [my_name]
                else:
                    teacher_options = [my_name] if my_name else []
                    st.warning("등록된 이름이 교사 목록에 없습니다. 관리자에게 문의하세요.")

            t = st.selectbox("교사 (본인 이름만 선택 가능)", teacher_options, key="duty_t")
            d = st.date_input("일자", value=date.today(), key="duty_d")
            reason = st.selectbox("사유", ABSENCE_REASONS, key="duty_r")
            detail = st.text_input("상세", key="duty_det")
            all_day = st.checkbox("하루 전체", key="duty_all")
            periods = [] if all_day else st.multiselect("교시 선택", list(range(1, 8)), key="duty_ps")

            if st.button("복무 등록", type="primary", key="duty_reg"):
                if all_day or periods:
                    push_history(f"복무 ({t})")
                    duties = ensure_duty_columns(st.session_state.duties)
                    duties = duties[~((duties["교사명"] == t) & (duties["일자"] == d.strftime("%Y-%m-%d")))]
                    news = []
                    if all_day:
                        news.append({"교사명": t, "일자": d.strftime("%Y-%m-%d"), "교시": 0,
                                     "사유": reason, "상세사유": detail,
                                     "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M"),
                                     "입력자": current_user()})
                    else:
                        for p in periods:
                            news.append({"교사명": t, "일자": d.strftime("%Y-%m-%d"), "교시": p,
                                         "사유": reason, "상세사유": detail,
                                         "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M"),
                                         "입력자": current_user()})
                    st.session_state.duties = pd.concat([duties, pd.DataFrame(news)], ignore_index=True)
                    save_work_data_to_gsheet()
                    st.success("등록 완료")
                    st.rerun()

            st.markdown("#### 등록된 복무")
            duties = ensure_duty_columns(st.session_state.duties)
            show_duties = filter_by_owner(duties)
            if not show_duties.empty:
                grouped = show_duties.groupby(["교사명", "일자", "사유"]).agg({"교시": list}).reset_index()
                grouped["교시표시"] = grouped["교시"].apply(format_periods)
                show_df = grouped[["교사명", "일자", "교시표시", "사유"]]
                st.dataframe(show_df, height=220, hide_index=True, use_container_width=True)

                options = [f"{row['교사명']} · {row['일자']} · {row['교시표시']} · {row['사유']}" for _, row in show_df.iterrows()]
                selected_label = st.selectbox("복무 선택 (이름·일자 기준)", options, key="duty_select_name")

                if st.button("이 복무 선택", key="duty_sel"):
                    sel_idx = options.index(selected_label)
                    sel_row = show_df.iloc[sel_idx]
                    st.session_state["_sel_duty"] = {
                        "교사명": sel_row["교사명"], "일자": sel_row["일자"],
                        "교시표시": sel_row["교시표시"], "사유": sel_row["사유"]
                    }
                    st.session_state.pop("_duty_search_cache", None)
                    st.rerun()

                if st.button("선택 행 삭제", key="duty_del"):
                    sel_idx = options.index(selected_label)
                    sel_row = show_df.iloc[sel_idx]
                    st.session_state.duties = duties[~((duties["교사명"] == sel_row["교사명"]) & 
                                                       (duties["일자"] == sel_row["일자"]))].reset_index(drop=True)
                    save_work_data_to_gsheet()
                    st.session_state.pop("_sel_duty", None)
                    st.session_state.pop("_duty_search_cache", None)
                    st.success("삭제 완료")
                    st.rerun()
            else:
                st.info("등록된 복무 없음")

        with right:
            st.markdown("### 선택 교사 처리 + 스마트 교환 검색")
            sel = st.session_state.get("_sel_duty")
            if not sel:
                st.info("왼쪽에서 복무를 선택하세요.")
            else:
                t_name = sel["교사명"]
                d_str = sel["일자"]
                day_kr = WEEKDAY_KR[datetime.strptime(d_str, "%Y-%m-%d").weekday()]
                st.markdown(f"**{t_name}** · {d_str} ({day_kr}) · {sel.get('교시표시', '')}")
                future_days = st.slider("미래 검색 일수", 7, 21, 14, key="future_days")

                if st.button("🔍 동일 학급 최우선 1:1 / 연계 검색 실행", type="primary", key="run_smart_search"):
                    with st.spinner("검색 중..."):
                        base_date = datetime.strptime(d_str, "%Y-%m-%d").date()
                        date_list = [base_date + timedelta(days=i) for i in range(0, future_days+1) if (base_date + timedelta(days=i)).weekday() < 5]
                        ver = st.session_state.get("_data_version", 0)
                        e_cache = {td.strftime("%Y-%m-%d"): get_effective_timetable_for_date(td.strftime("%Y-%m-%d"), ver) for td in date_list}
                        e_today = e_cache.get(d_str, pd.DataFrame())
                        my_lessons = e_today[e_today["교사명"] == t_name] if not e_today.empty else pd.DataFrame()
                        duty_periods = set()
                        if "전체" in sel.get("교시표시", ""):
                            duty_periods = set(range(1, 8))
                        else:
                            raw = st.session_state.duties
                            mask = (raw["교사명"] == t_name) & (raw["일자"] == d_str)
                            duty_periods = set(raw.loc[mask, "교시"].tolist())
                            if 0 in duty_periods:
                                duty_periods = set(range(1, 8))
                        results = {}
                        for _, lesson in my_lessons.iterrows():
                            p = safe_int(lesson["교시"])
                            if p not in duty_periods:
                                continue
                            my_class = lesson["학급"]
                            my_grade = grade_of(my_class)
                            my_group = subject_group(lesson["과목"])
                            my_subject = lesson["과목"]
                            candidates_1to1 = []
                            candidates_linked = []
                            for td in date_list:
                                tds = td.strftime("%Y-%m-%d")
                                tday = WEEKDAY_KR[td.weekday()]
                                e_tt = e_cache[tds]
                                others = e_tt[(e_tt["교시"] == p) & (e_tt["교사명"] != t_name)] if not e_tt.empty else pd.DataFrame()
                                for o in others.itertuples():
                                    if is_free(t_name, tday, p, tds, e_tt) and is_free(o.교사명, day_kr, p, d_str, e_today):
                                        other_class = o.학급
                                        other_grade = grade_of(other_class)
                                        same_class = (other_class == my_class)
                                        same_grade = (other_grade == my_grade)
                                        score = 0
                                        if same_class: score += 200
                                        elif same_grade: score += 100
                                        if subject_group(o.과목) == my_group: score += 40
                                        if tds == d_str: score += 15
                                        candidates_1to1.append({
                                            "type": "1:1", "date": tds, "day": tday, "period": p,
                                            "teacher": o.교사명, "lesson": f"{other_class} {o.과목}",
                                            "same_class": same_class, "same_grade": same_grade, "score": score
                                        })
                                if is_free(t_name, tday, p, tds, e_tt):
                                    for ot in st.session_state.teachers["교사명"].tolist()[:30]:
                                        if ot != t_name and is_free(ot, day_kr, p, d_str, e_today):
                                            candidates_linked.append({
                                                "type": "연계", "date": tds, "day": tday, "period": p,
                                                "teacher": ot, "lesson": "공강", "score": 30
                                            })
                                            break
                            candidates_1to1 = sorted(candidates_1to1, key=lambda x: (-x["same_class"], -x["same_grade"], -x["score"]))[:6]
                            candidates_linked = sorted(candidates_linked, key=lambda x: -x["score"])[:4]
                            results[p] = {
                                "my_class": my_class, "my_subject": my_subject, "my_grade": my_grade,
                                "one_to_one": candidates_1to1, "linked": candidates_linked
                            }
                        st.session_state["_duty_search_cache"] = results
                        st.success("검색 완료!")

                cache = st.session_state.get("_duty_search_cache")
                if cache:
                    for p, data in cache.items():
                        with st.expander(f"{p}교시 · {data['my_class']} {data['my_subject']} (학년 {data['my_grade']})", expanded=True):
                            st.markdown("#### 🏆 1:1 맞교환 (동일 학급 최우선)")
                            if not data["one_to_one"]:
                                st.info("조건에 맞는 1:1 대상 없음")
                            else:
                                for i, c in enumerate(data["one_to_one"]):
                                    mark = "🏆 동일학급" if c["same_class"] else ("⚠ 같은학년" if c["same_grade"] else "⚠ 다른학년")
                                    st.write(f"{mark} | {c['date']} ({c['day']}) {c['period']}교시 - **{c['teacher']}** ({c['lesson']})")
                                    if st.button("이 수업과 1:1 맞교환 실행", key=f"o2o_{p}_{i}"):
                                        a_info = {"교사명": t_name, "일자": d_str, "요일": day_kr, "교시": p,
                                                  "학급": data["my_class"], "과목": data["my_subject"]}
                                        b_info = {"교사명": c["teacher"], "일자": c["date"], "요일": c["day"], "교시": c["period"],
                                                  "학급": "", "과목": ""}
                                        do_swap(a_info, b_info, d_str, c["date"])
                                        st.success("1:1 맞교환 등록 완료!")
                                        st.session_state.pop("_duty_search_cache", None)
                                        st.rerun()
                            st.markdown("#### 🔗 연계 교환")
                            if not data["linked"]:
                                st.info("연계 교환 대상 없음")
                            else:
                                for i, c in enumerate(data["linked"]):
                                    st.write(f"• {c['date']} ({c['day']}) {c['period']}교시 - **{c['teacher']}**")
                                    if st.button("연계 교환 실행", key=f"lnk_{p}_{i}"):
                                        a_info = {"교사명": t_name, "일자": d_str, "요일": day_kr, "교시": p,
                                                  "학급": data["my_class"], "과목": data["my_subject"]}
                                        do_linked_swap(a_info, c["teacher"], d_str, c["date"], c["day"], c["period"])
                                        st.success("연계 교환 등록 완료!")
                                        st.session_state.pop("_duty_search_cache", None)
                                        st.rerun()
                else:
                    st.info("위에서 **검색 실행** 버튼을 눌러주세요.")
                    
# ------------------------------------------------------------------ 다중 출장·전체 조정 추천
if "🛠️ 다중 출장·전체 조정 추천" in tab_map:
    with tab_map["🛠️ 다중 출장·전체 조정 추천"]:
        st.subheader("🛠️ 다중 출장·전체 조정 추천 (마스터/교육과정부 전용)")
        st.info("여러 교사가 동시에 출장·복무일 때 사용합니다. **시작일(교시) ~ 종료일(교시)** 범위를 지정할 수 있습니다.")

        remaining_budget = get_current_budget()
        st.metric("현재 보강비 잔액", f"{remaining_budget:,.0f}원")

        if remaining_budget > 1000000:
            budget_factor = 0.7
            budget_msg = "예산 충분 → 보강 허용 비중 높음"
        elif remaining_budget > 300000:
            budget_factor = 1.0
            budget_msg = "예산 보통 → 균형 추천"
        else:
            budget_factor = 1.8
            budget_msg = "예산 부족 → 맞교환 우선 추천"
        st.caption(f"추천 모드: {budget_msg} (교환 가중치 ×{budget_factor})")

        with st.expander("예산 변경 이력 보기"):
            budget_df = load_budget_df()
            st.dataframe(budget_df.tail(20), use_container_width=True, hide_index=True)

        st.markdown("### 1. 불가능한 교사·기간 선택 (시작일 ~ 종료일)")

        if "multi_absent" not in st.session_state:
            st.session_state.multi_absent = []

        abs_teacher = st.selectbox("교사", st.session_state.teachers["교사명"].tolist(), key="multi_t")

        col_start, col_end = st.columns(2)
        with col_start:
            st.markdown("**시작일**")
            start_date = st.date_input("시작일", value=date.today(), key="multi_start_d")
            start_periods = st.multiselect("시작일 교시", list(range(1, 8)), key="multi_start_p")
            start_all = st.checkbox("시작일 하루 전체", key="multi_start_all")
            if start_all:
                start_periods = [0]

        with col_end:
            st.markdown("**종료일**")
            end_date = st.date_input("종료일", value=date.today(), key="multi_end_d")
            end_periods = st.multiselect("종료일 교시", list(range(1, 8)), key="multi_end_p")
            end_all = st.checkbox("종료일 하루 전체", key="multi_end_all")
            if end_all:
                end_periods = [0]

        if st.button("기간 추가 (시작일~종료일)", type="primary"):
            if start_date > end_date:
                st.error("시작일이 종료일보다 늦을 수 없습니다.")
            else:
                current = start_date
                added_count = 0
                while current <= end_date:
                    if current.weekday() < 5:
                        d_str = current.strftime("%Y-%m-%d")
                        if current == start_date:
                            periods_to_add = start_periods if start_periods else [0]
                        elif current == end_date:
                            periods_to_add = end_periods if end_periods else [0]
                        else:
                            periods_to_add = [0]

                        for p in periods_to_add:
                            item = {"교사": abs_teacher, "일자": d_str, "교시": p}
                            if item not in st.session_state.multi_absent:
                                st.session_state.multi_absent.append(item)
                                added_count += 1
                    current += timedelta(days=1)
                st.success(f"{added_count}건이 추가되었습니다.")
                st.rerun()

        if st.session_state.multi_absent:
            st.write("**현재 선택된 불가능 목록**")
            st.dataframe(pd.DataFrame(st.session_state.multi_absent), use_container_width=True, hide_index=True)
            if st.button("목록 초기화"):
                st.session_state.multi_absent = []
                st.rerun()

            if st.button("🔍 전체 조정 추천 실행", type="primary"):
                with st.spinner("다중 조건으로 추천 검색 중..."):
                    all_recs = []
                    for item in st.session_state.multi_absent:
                        t_name = item["교사"]
                        d_str = item["일자"]
                        p = item["교시"]
                        day_kr = WEEKDAY_KR[datetime.strptime(d_str, "%Y-%m-%d").weekday()]
                        e_tt = get_effective_timetable_for_date(d_str, st.session_state.get("_data_version", 0))
                        if p == 0:
                            lessons = e_tt[e_tt["교사명"] == t_name]
                        else:
                            lessons = e_tt[(e_tt["교사명"] == t_name) & (e_tt["교시"] == p)]
                        if lessons.empty:
                            continue
                        for _, lesson in lessons.iterrows():
                            actual_p = safe_int(lesson["교시"])
                            base_date = datetime.strptime(d_str, "%Y-%m-%d").date()
                            for i in range(0, 15):
                                td = base_date + timedelta(days=i)
                                if td.weekday() >= 5:
                                    continue
                                tds = td.strftime("%Y-%m-%d")
                                tday = WEEKDAY_KR[td.weekday()]
                                df_swap, cycles, _ = get_target_time_recommendations(
                                    t_name, d_str, actual_p, lesson["학급"], lesson["과목"],
                                    tds, actual_p, budget_factor=budget_factor
                                )
                                for _, row in df_swap.head(3).iterrows():
                                    all_recs.append({
                                        "원본교사": t_name, "원본일자": d_str, "원본교시": actual_p,
                                        "유형": "1:1", "추천교사": row["교사B"],
                                        "추천내용": row["현재 수업"], "점수": row["점수"],
                                        "동일학급": row.get("same_class", False)
                                    })
                                for cyc in cycles[:2]:
                                    all_recs.append({
                                        "원본교사": t_name, "원본일자": d_str, "원본교시": actual_p,
                                        "유형": f"{cyc['length']}인순환", "추천교사": cyc["path_desc"][:40]+"...",
                                        "추천내용": cyc["path_desc"], "점수": cyc["score"],
                                        "동일학급": True
                                    })
                    if all_recs:
                        rec_df = pd.DataFrame(all_recs).sort_values(["동일학급", "점수"], ascending=[False, False])
                        st.session_state["_multi_recs"] = rec_df
                        st.success(f"총 {len(rec_df)}건의 추천이 생성되었습니다.")
                    else:
                        st.warning("조건에 맞는 추천이 없습니다.")
                        st.session_state["_multi_recs"] = pd.DataFrame()

            if "_multi_recs" in st.session_state and not st.session_state["_multi_recs"].empty:
                st.markdown("### 추천 결과 (동일학급 우선 + 예산 반영)")
                st.dataframe(st.session_state["_multi_recs"], use_container_width=True, hide_index=True)
                st.caption("실제 적용은 「시간표 맞교환 & 변경 추천」 탭이나 「복무 관리」 탭에서 진행하세요.")

# ------------------------------------------------------------------ 아이디·권한 관리
if "🔑 아이디·권한 관리" in tab_map:
    with tab_map["🔑 아이디·권한 관리"]:
        st.subheader("🔑 아이디 · 권한 관리")
        st.caption("마스터와 교육과정부만 접근 가능합니다.")

        ids_df = load_id_sheet()
        st.markdown("### 현재 등록 아이디")
        edited = st.data_editor(
            ids_df,
            num_rows="dynamic",
            use_container_width=True,
            key="id_editor",
            column_config={
                "아이디": st.column_config.TextColumn("아이디", required=True),
                "이름": st.column_config.TextColumn("등록 이름", required=True),
                "권한": st.column_config.SelectboxColumn(
                    "권한",
                    options=[ROLE_MASTER, ROLE_EDU, ROLE_TEACHER],
                    required=True
                ),
                "허용탭": st.column_config.TextColumn("허용탭 (쉼표로 구분, 비워두면 기본값)")
            }
        )

        col1, col2, col3 = st.columns(3)
        with col1:
            if st.button("아이디 목록 저장", type="primary"):
                masters = edited[edited["권한"] == ROLE_MASTER]
                if len(masters) == 0:
                    st.error("마스터 권한 아이디가 최소 1명은 있어야 합니다.")
                else:
                    save_id_sheet(edited)
                    st.success("저장 완료")
                    st.rerun()
        with col2:
            if is_master():
                st.info("마스터 양도: 원하는 아이디의 권한을 '마스터'로 변경 후 저장")
        with col3:
            if st.button("아이디 추가 요청 목록 보기"):
                req_ws = get_worksheet(WORK_SHEET_ID, "아이디추가요청")
                req_df = df_from_worksheet(req_ws)
                st.dataframe(req_df, use_container_width=True)

        st.divider()
        st.markdown("### 권한 설명")
        st.markdown("""
        | 권한 | 설명 |
        |------|------|
        | **마스터** | 모든 권한 + 아이디 관리/삭제/양도 + 마스터 양도 + 전체 데이터 관리 |
        | **교육과정부** | 아이디 저장/삭제 + 전체 탭/데이터 관리 |
        | **일반교사** | 본인이 입력한 데이터만 조회·저장·삭제 + 본인 이름으로만 복무 등록 |
        | **게스트** | 아이디 추가요청만 가능 |
        """)

# ------------------------------------------------------------------ 회원별 탭 권한 관리
if "📑 회원별 탭 권한 관리" in tab_map:
    with tab_map["📑 회원별 탭 권한 관리"]:
        st.subheader("📑 회원별 탭 권한 관리")
        st.caption("각 회원에게 보여줄 탭을 개별적으로 설정할 수 있습니다.")

        ids_df = load_id_sheet()
        if ids_df.empty:
            st.warning("등록된 아이디가 없습니다.")
        else:
            selected_user = st.selectbox(
                "회원 선택",
                options=ids_df["아이디"].tolist(),
                format_func=lambda x: f"{x} ({ids_df.loc[ids_df['아이디']==x, '이름'].values[0] if len(ids_df.loc[ids_df['아이디']==x])>0 else ''})"
            )

            user_row = ids_df[ids_df["아이디"] == selected_user].iloc[0]
            current_allowed = str(user_row.get("허용탭", "")).strip()
            if current_allowed:
                current_list = [t.strip() for t in current_allowed.split(",") if t.strip()]
            else:
                role = str(user_row.get("권한", ROLE_TEACHER))
                current_list = DEFAULT_TABS.get(role, [])

            st.markdown(f"**현재 선택 회원**: `{selected_user}` / 이름: `{user_row.get('이름','')}` / 권한: `{user_row.get('권한','')}`")

            st.markdown("#### 허용할 탭 선택")
            new_allowed = []
            cols = st.columns(2)
            for idx, tab_name in enumerate(ALL_TABS):
                disabled = False
                if tab_name in ("🔑 아이디·권한 관리", "📑 회원별 탭 권한 관리", "🛠️ 다중 출장·전체 조정 추천") and not can_manage_ids():
                    disabled = True
                with cols[idx % 2]:
                    checked = st.checkbox(
                        tab_name,
                        value=(tab_name in current_list),
                        key=f"tab_check_{selected_user}_{idx}",
                        disabled=disabled
                    )
                    if checked:
                        new_allowed.append(tab_name)

            if st.button("이 회원의 탭 권한 저장", type="primary"):
                ids_df.loc[ids_df["아이디"] == selected_user, "허용탭"] = ",".join(new_allowed)
                save_id_sheet(ids_df)
                st.success(f"{selected_user} 님의 탭 권한이 저장되었습니다.")
                if selected_user == current_user():
                    st.session_state.user_allowed_tabs = ",".join(new_allowed)
                st.rerun()

            st.divider()
            st.markdown("#### 빠른 설정")
            c1, c2, c3 = st.columns(3)
            with c1:
                if st.button("기본 탭으로 초기화"):
                    role = str(user_row.get("권한", ROLE_TEACHER))
                    default = DEFAULT_TABS.get(role, [])
                    ids_df.loc[ids_df["아이디"] == selected_user, "허용탭"] = ",".join(default)
                    save_id_sheet(ids_df)
                    st.success("기본값으로 초기화됨")
                    st.rerun()
            with c2:
                if st.button("모든 탭 허용"):
                    ids_df.loc[ids_df["아이디"] == selected_user, "허용탭"] = ",".join(ALL_TABS)
                    save_id_sheet(ids_df)
                    st.success("모든 탭 허용됨")
                    st.rerun()
            with c3:
                if st.button("모든 탭 차단"):
                    ids_df.loc[ids_df["아이디"] == selected_user, "허용탭"] = ""
                    save_id_sheet(ids_df)
                    st.success("모든 탭 차단됨")
                    st.rerun()

st.caption(f"서라벌여중 시간표 관리 시스템 · {current_name()} ({current_user()}) · {current_role()}")
