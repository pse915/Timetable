# -*- coding: utf-8 -*-
"""
서라벌여중 시간표·결보강 관리 프로그램
2026 최종 완성판 (속도 최적화 + 버그 수정 버전)
- 결보강 계획서: 보강수업 칸에 '보강 배정된 교사'가 나오도록 수정
- 연계 순환 알고리즘 대폭 가속
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

MAX_HISTORY = 5
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
# 히스토리 / 트랜잭션
# ==========================================================================================
def _deepcopy_df(df):
    return df.copy(deep=True) if isinstance(df, pd.DataFrame) else pd.DataFrame()


def _snapshot_state(action_name="작업"):
    budget = pd.DataFrame()
    try:
        budget = _deepcopy_df(load_budget_df())
    except Exception:
        pass
    return {
        "action": action_name,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "absences": _deepcopy_df(st.session_state.get("absences", pd.DataFrame())),
        "subs": _deepcopy_df(st.session_state.get("subs", pd.DataFrame())),
        "swaps": _deepcopy_df(st.session_state.get("swaps", pd.DataFrame())),
        "part_time": _deepcopy_df(st.session_state.get("part_time", pd.DataFrame())),
        "duties": _deepcopy_df(st.session_state.get("duties", pd.DataFrame())),
        "budget": budget,
    }


def record_history(action_name="작업"):
    """작업 완료 후 현재 상태를 저장한다. Undo/Redo가 실제 상태 스냅샷으로 동작한다."""
    if "history" not in st.session_state:
        st.session_state.history=[]
        st.session_state.history_index=-1
    st.session_state.history=st.session_state.history[:st.session_state.history_index+1]
    st.session_state.history.append(_snapshot_state(action_name))
    if len(st.session_state.history)>MAX_HISTORY:
        st.session_state.history.pop(0)
    st.session_state.history_index=len(st.session_state.history)-1


def push_history(action_name="작업"):
    # 하위 호환용. 새 코드에서는 record_history()를 작업 완료 후 호출한다.
    record_history(action_name)


def _restore_snapshot(snap):
    for k in ["absences","subs","swaps","part_time","duties"]:
        st.session_state[k]=_deepcopy_df(snap.get(k,pd.DataFrame()))
    budget=snap.get("budget",pd.DataFrame())
    if isinstance(budget,pd.DataFrame) and not budget.empty:
        try:
            df_to_worksheet(get_worksheet(WORK_SHEET_ID,"예산"),budget)
            load_budget_df.clear()
        except Exception as e:
            st.warning(f"예산 상태 복원 실패: {e}")
    _invalidate_all_caches()


def undo():
    idx=st.session_state.get("history_index",-1)
    if idx<=0:
        return False
    st.session_state.history_index=idx-1
    _restore_snapshot(st.session_state.history[st.session_state.history_index])
    return True


def redo():
    idx=st.session_state.get("history_index",-1)
    hist=st.session_state.get("history",[])
    if idx>=len(hist)-1:
        return False
    st.session_state.history_index=idx+1
    _restore_snapshot(hist[st.session_state.history_index])
    return True


def _invalidate_all_caches():
    st.session_state._data_version=st.session_state.get("_data_version",0)+1
    for name in [
        "get_effective_timetable_for_date","get_single_lesson_1to1_candidates",
        "get_single_lesson_linked_cycles","effective_teacher_matrix","teacher_matrix",
        "class_matrix","cumulative_sub_count","weekly_load","load_budget_df",
        "load_swap_requests"
    ]:
        fn=globals().get(name)
        if fn is not None and hasattr(fn,"clear"):
            try: fn.clear()
            except Exception: pass

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

def save_work_data_to_gsheet(changed_sheets=None):
    """변경된 시트만 저장. None이면 전체 시트를 저장한다."""
    if not can_full_data() and not is_teacher():
        st.warning("저장 권한이 없습니다.")
        return False
    try:
        sheet_map={"absences":"결강","subs":"보강","swaps":"맞교환","part_time":"시간강사","duties":"복무"}
        names=list(sheet_map) if changed_sheets is None else [n for n in changed_sheets if n in sheet_map]
        for name in names:
            df=st.session_state.get(name,pd.DataFrame())
            if name=="part_time":
                df=ensure_part_time_columns(df); st.session_state.part_time=df
            elif name=="duties":
                df=ensure_duty_columns(df); st.session_state.duties=df
            df_to_worksheet(get_worksheet(WORK_SHEET_ID,sheet_map[name]),df)
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
    record_history("초기 상태")

# ==========================================================================================
# 핵심 로직
# ==========================================================================================
# ==========================================================================================
# 핵심 시간표 엔진 / 교환 검증 / 고속 후보 검색
# ==========================================================================================
SCHEDULE_COLUMNS=[
    "교사명","요일","교시","과목","학급","과목군","원본교사",
    "원본일자","원본교시","변경유형","변경이력","변경ID"
]


def _blank_schedule_df():
    return pd.DataFrame(columns=SCHEDULE_COLUMNS)


def _clean(v):
    if v is None: return ""
    try:
        if pd.isna(v): return ""
    except Exception: pass
    return str(v).strip()


def _get(rec,key,default=""):
    return rec.get(key,default) if isinstance(rec,dict) else getattr(rec,key,default)


def _swap_type(rec):
    typ=_clean(_get(rec,"유형",""))
    if typ in {"1:1","1:1 맞교환","1:1맞교환","직접1:1"}: return "1:1"
    if "연계" in typ: return "연계"
    return typ


def _make_lesson(teacher,day,period,subject,class_name,original_teacher=None,original_date="",original_period=0,change_type="",change_history="",change_id=""):
    teacher=_clean(teacher); subject=_clean(subject); class_name=_clean(class_name)
    ot=_clean(original_teacher) or teacher
    return {
        "교사명":teacher,"요일":day,"교시":safe_int(period),"과목":subject,"학급":class_name,
        "과목군":subject_group(subject),"원본교사":ot if ot!=teacher else "",
        "원본일자":_clean(original_date),"원본교시":safe_int(original_period),
        "변경유형":_clean(change_type),"변경이력":_clean(change_history),"변경ID":_clean(change_id)
    }


def _part_time_slot_available(prow,day,period):
    key=f"{day}{safe_int(period)}"
    if key not in prow.index: return True
    cols=[f"{d}{p}" for d in DAYS for p in range(1,8)]
    vals=[_clean(prow.get(c,"")) for c in cols]
    # 구형 데이터가 가능시간을 하나도 입력하지 않았다면 기존 동작을 유지한다.
    if not any(vals): return True
    v=_clean(prow.get(key,"")).lower()
    return v in {"y","yes","1","가능","가능함","가능o","o","○","ㅇ","true","예"}



def _next_change_id(prefix="CHG"):
    """충돌 없는 변경 ID 생성. Streamlit 세션 안에서는 증가 카운터를 유지한다."""
    n=safe_int(st.session_state.get("_change_seq",0))+1
    st.session_state["_change_seq"]=n
    return f"{prefix}-{datetime.now().strftime('%Y%m%d%H%M%S%f')}-{n}"


def _apply_swap_record(current,sw,norm,day,source_label="실제"):
    typ=_swap_type(sw)
    ta,tb=_clean(_get(sw,"교사A")),_clean(_get(sw,"교사B"))
    pa,pb=safe_int(_get(sw,"교시A")),safe_int(_get(sw,"교시B"))
    da,db=normalize_date_str(_get(sw,"원본일자")),normalize_date_str(_get(sw,"목표일자"))
    rid=_clean(_get(sw,"변경ID", "")) or _clean(_get(sw,"등록시각", ""))
    if typ=="1:1":
        if norm==da:
            current.pop((ta,pa),None)
            if tb and pa>0:
                sb=_clean(_get(sw,"과목B")); cb=_clean(_get(sw,"학급B"))
                otb=_clean(_get(sw,"원본교사B")) or tb
                odb=_clean(_get(sw,"원본일자B")) or db
                opb=safe_int(_get(sw,"원본교시B",pb)) or pb
                prev_hist=_clean(_get(sw,"변경이력B"))
                history=f"{tb}의 {db} {pb}교시 수업 → {ta}의 {da} {pa}교시"
                if prev_hist: history=f"{prev_hist} | {history}"
                current[(tb,pa)]=_make_lesson(tb,day,pa,sb,cb,otb,odb,opb,"1:1 맞교환",history,rid)
        if norm==db:
            current.pop((tb,pb),None)
            if ta and pb>0:
                sa=_clean(_get(sw,"과목A")); ca=_clean(_get(sw,"학급A"))
                ota=_clean(_get(sw,"원본교사A")) or ta
                oda=_clean(_get(sw,"원본일자A")) or da
                opa=safe_int(_get(sw,"원본교시A",pa)) or pa
                prev_hist=_clean(_get(sw,"변경이력A"))
                history=f"{ta}의 {da} {pa}교시 수업 → {tb}의 {db} {pb}교시"
                if prev_hist: history=f"{prev_hist} | {history}"
                current[(ta,pb)]=_make_lesson(ta,day,pb,sa,ca,ota,oda,opa,"1:1 맞교환",history,rid)
    elif typ=="연계":
        if norm==da:
            # A의 원래 슬롯은 비운다. 목표 슬롯의 기존 담당교사는
            # 순환의 다음 move에서 자신의 source 슬롯을 비우므로 여기서는 건드리지 않는다.
            current.pop((ta,pa),None)
        if norm==db:
            if ta and pb>0:
                sa=_clean(_get(sw,"과목A")); ca=_clean(_get(sw,"학급A"))
                history=f"{ta}의 {da} {pa}교시 수업 → {tb}의 {db} {pb}교시"
                current[(ta,pb)]=_make_lesson(ta,day,pb,sa,ca,ta,da,pa,"연계 공강 교환",history,rid)


@st.cache_data(show_spinner=False,ttl=180)
def get_effective_timetable_for_date(on_date:str,version:int=0,use_test:bool=False)->pd.DataFrame:
    norm=normalize_date_str(on_date)
    if not norm: return st.session_state.timetable.copy()
    try: day=WEEKDAY_KR[datetime.strptime(norm,"%Y-%m-%d").weekday()]
    except Exception: return st.session_state.timetable.copy()
    tt=st.session_state.get("timetable",pd.DataFrame())
    if tt.empty: return _blank_schedule_df()
    current={}
    for r in tt.loc[tt["요일"].eq(day)].itertuples(index=False):
        p=safe_int(r.교시); t=_clean(r.교사명); s=_clean(r.과목); c=_clean(r.학급)
        current[(t,p)]=_make_lesson(t,day,p,s,c,t)
    for _,sw,source in _get_swap_records_for_date(norm,use_test):
        _apply_swap_record(current,sw,norm,day,source)
    # 보강은 교환 결과 위에 적용한다.
    subs=st.session_state.get("subs",pd.DataFrame())
    if isinstance(subs,pd.DataFrame) and not subs.empty and "일자" in subs.columns:
        for r in subs.loc[subs["일자"].eq(norm)].itertuples(index=False):
            p=safe_int(getattr(r,"교시",0)); abs_t=_clean(getattr(r,"결강교사","")); sub_t=_clean(getattr(r,"보강교사",""))
            if not sub_t or p<=0: continue
            old_abs=current.get((abs_t,p))
            if old_abs is None:
                continue
            occupied_sub=current.get((sub_t,p))
            if occupied_sub is not None and sub_t!=abs_t:
                conflict=dict(old_abs)
                conflict["변경유형"]="보강 충돌"
                conflict["변경이력"]=f"보강 충돌: {abs_t} 결강 → {sub_t} 보강 예정 ({norm} {p}교시), 그러나 {sub_t}에게 기존 수업이 있음"
                conflict["변경ID"]=_clean(getattr(r,"결강ID",""))
                current[(abs_t,p)]=conflict
                continue
            current.pop((abs_t,p),None)
            subject=_clean(getattr(r,"과목","")) or _clean(old_abs.get("과목",""))
            cls=_clean(getattr(r,"학급","")) or _clean(old_abs.get("학급",""))
            original_teacher=_clean(old_abs.get("원래교사")) or _clean(old_abs.get("원본교사")) or abs_t
            original_date=old_abs.get("원본일자",norm) or norm
            original_period=safe_int(old_abs.get("원본교시",p)) or p
            rid=_clean(getattr(r,"결강ID",""))
            hist=f"{abs_t} 결강 → {sub_t} 보강 ({norm} {p}교시)"
            current[(sub_t,p)]=_make_lesson(sub_t,day,p,subject,cls,original_teacher,original_date,original_period,"보강",hist,rid)
    # 시간강사: 기간과 가능시간표를 모두 반영한다.
    pt=st.session_state.get("part_time",pd.DataFrame())
    if isinstance(pt,pd.DataFrame) and not pt.empty:
        for _,prow in pt.iterrows():
            start=normalize_date_str(prow.get("시작일","")); end=normalize_date_str(prow.get("종료일",""))
            if not(start and end and start<=norm<=end): continue
            orig=_clean(prow.get("대체교사","")); pt_name=_clean(prow.get("시간강사명",""))
            if not orig or not pt_name or orig==pt_name: continue
            for key in list(current.keys()):
                if key[0]!=orig or not _part_time_slot_available(prow,day,key[1]): continue
                if (pt_name,key[1]) in current and (pt_name,key[1]) != key:
                    conflict=dict(current[key]); conflict["변경유형"]="시간강사 충돌"
                    conflict["변경이력"]=f"시간강사 충돌: {orig} → {pt_name} 대체 예정 ({norm} {day}{key[1]}교시), 그러나 시간강사에게 기존 수업이 있음"
                    current[key]=conflict
                    continue
                lesson=dict(current.pop(key)); lesson["교사명"]=pt_name; lesson["원본교사"]=orig; lesson["변경유형"]="시간강사"
                lesson["변경이력"]=f"{orig} → {pt_name} 시간강사 대체 ({norm} {day}{key[1]}교시)"
                lesson["변경ID"]=f"PT:{pt_name}:{orig}:{start}:{end}"
                current[(pt_name,key[1])]=lesson
    return pd.DataFrame(list(current.values()),columns=SCHEDULE_COLUMNS) if current else _blank_schedule_df()


def _get_swap_records_for_date(norm,use_test=False):
    result=[]
    datasets=[("실제",st.session_state.get("swaps",pd.DataFrame()))]
    if use_test: datasets.append(("테스트",st.session_state.get("test_swaps",pd.DataFrame())))
    for label,df in datasets:
        if not isinstance(df,pd.DataFrame) or df.empty or "원본일자" not in df.columns: continue
        mask=df["원본일자"].apply(normalize_date_str).eq(norm) | df["목표일자"].apply(normalize_date_str).eq(norm)
        for idx,r in df.loc[mask].iterrows(): result.append((idx,r,label))
    result.sort(key=lambda x:str(_get(x[1],"등록시각","")))
    return result


def _duty_period_map():
    out=defaultdict(set); df=st.session_state.get("duties",pd.DataFrame())
    if isinstance(df,pd.DataFrame) and not df.empty:
        for r in df.itertuples(index=False):
            t=_clean(getattr(r,"교사명","")); d=normalize_date_str(getattr(r,"일자","")); p=safe_int(getattr(r,"교시",0))
            if t and d: out[(d,t)].add(p)
    return out


def has_duty(teacher,on_date,period=None):
    t=_clean(teacher); d=normalize_date_str(on_date); df=st.session_state.get("duties",pd.DataFrame())
    if not isinstance(df,pd.DataFrame) or df.empty or "일자" not in df.columns:return False
    m=(df["교사명"].astype(str).str.strip()==t)&(df["일자"].astype(str).str.strip()==d)
    if not m.any(): return False
    if period is None: return True
    ps={safe_int(x) for x in df.loc[m,"교시"].tolist()}
    return 0 in ps or safe_int(period) in ps


def _blocked_period_map():
    """교무상 사용 불가 시간(복무 + 결강)을 후보 검색용으로 빠르게 제공한다."""
    out=_duty_period_map()
    df=st.session_state.get("absences",pd.DataFrame())
    if isinstance(df,pd.DataFrame) and not df.empty and "일자" in df.columns:
        for r in df.itertuples(index=False):
            t=_clean(getattr(r,"교사명","")); d=normalize_date_str(getattr(r,"일자","")); p=safe_int(getattr(r,"교시",0))
            if t and d: out[(d,t)].add(p)
    return out


def _schedule_lookup(e_tt):
    lookup={}; busy=defaultdict(set)
    if e_tt is None or e_tt.empty: return lookup,busy
    for r in e_tt.to_dict("records"):
        t=_clean(r.get("교사명")); p=safe_int(r.get("교시"))
        if t and p: lookup[(t,p)]=r; busy[t].add(p)
    return lookup,busy


def has_absence(teacher,on_date,period=None):
    """결강 등록 상태 확인. 교시 0은 하루 전체 결강이다."""
    t=_clean(teacher); d=normalize_date_str(on_date); df=st.session_state.get("absences",pd.DataFrame())
    if not isinstance(df,pd.DataFrame) or df.empty or "일자" not in df.columns:return False
    mask=(df["교사명"].astype(str).str.strip()==t)&(df["일자"].astype(str).str.strip()==d)
    if not mask.any():return False
    if period is None:return True
    ps={safe_int(x) for x in df.loc[mask,"교시"].tolist()}
    return 0 in ps or safe_int(period) in ps


def is_free(teacher,day,period,on_date=None,e_tt=None):
    p=safe_int(period); d=normalize_date_str(on_date)
    if not d or p<=0:return False
    if has_duty(teacher,d,p) or has_absence(teacher,d,p):return False
    if e_tt is None:e_tt=get_effective_timetable_for_date(d,st.session_state.get("_data_version",0))
    if e_tt is None or e_tt.empty:return True
    lookup,_=_schedule_lookup(e_tt)
    return (_clean(teacher),p) not in lookup


def get_change_info(teacher,on_date,period,use_test=False):
    t=_clean(teacher); d=normalize_date_str(on_date); p=safe_int(period)
    subs=st.session_state.get("subs",pd.DataFrame())
    if isinstance(subs,pd.DataFrame) and not subs.empty:
        m=subs[(subs["일자"].apply(normalize_date_str)==d)&(subs["교시"].apply(safe_int)==p)]
        if not m.empty:
            r=m.iloc[-1]; a=_clean(r.get("결강교사")); b=_clean(r.get("보강교사"))
            if t==b: return f"🟢 보강: {a} 결강 → {b} 보강"
            if t==a: return f"🔴 결강: {a} → {b} 보강"
    for _,sw,_ in reversed(_get_swap_records_for_date(d,use_test)):
        typ=_swap_type(sw); a=_clean(_get(sw,"교사A")); b=_clean(_get(sw,"교사B")); da=normalize_date_str(_get(sw,"원본일자")); db=normalize_date_str(_get(sw,"목표일자")); pa=safe_int(_get(sw,"교시A")); pb=safe_int(_get(sw,"교시B"))
        if typ=="1:1":
            if d==da and t==b and p==pa: return f"🔄 {b} ← {a}의 {da} {pa}교시"
            if d==db and t==a and p==pb: return f"🔄 {a} ← {b}의 {db} {pb}교시"
            if d==da and t==a and p==pa: return f"🔄 {a} ↔ {b} 맞교환"
            if d==db and t==b and p==pb: return f"🔄 {b} ↔ {a} 맞교환"
        else:
            if d==db and t==a and p==pb: return f"🔗 {a} → {b} 연계교환"
    return ""


def get_swap_origin_info(teacher,on_date,period):
    return get_change_info(teacher,on_date,period,use_test=False)


@st.cache_data(show_spinner=False)
def cumulative_sub_count(start_date=None,end_date=None,version=0):
    s=st.session_state.get("subs",pd.DataFrame()); base={t:0 for t in st.session_state.teachers["교사명"].tolist()} if not st.session_state.teachers.empty else {}
    if s.empty or "보강교사" not in s.columns: return base
    if start_date and end_date: s=s[(s["일자"]>=normalize_date_str(start_date))&(s["일자"]<=normalize_date_str(end_date))]
    for k,v in s["보강교사"].value_counts().items():
        if k in base: base[k]=int(v)
    return base


@st.cache_data(show_spinner=False)
def weekly_load(version=0):
    tt=st.session_state.timetable
    return tt["교사명"].value_counts().to_dict() if not tt.empty else {}



def recommend_substitutes(day,period,subject,class_name,absent_teacher,on_date,top_n=20,include_part_time=False,e_tt=None):
    teachers=st.session_state.get("teachers",pd.DataFrame())
    if teachers.empty:return pd.DataFrame()
    norm=normalize_date_str(on_date); e_tt=e_tt if e_tt is not None else get_effective_timetable_for_date(norm,st.session_state.get("_data_version",0))
    lookup,busy=_schedule_lookup(e_tt); duty=_blocked_period_map(); groups_by_teacher=defaultdict(set); grades_by_teacher=defaultdict(set)
    for (t,_),row in lookup.items():
        groups_by_teacher[t].add(subject_group(_clean(row.get("과목")))); grades_by_teacher[t].add(grade_of(_clean(row.get("학급"))))
    cum=cumulative_sub_count(version=st.session_state.get("_data_version",0)); load=weekly_load(version=st.session_state.get("_data_version",0)); max_cum=max(cum.values()) if cum else 0
    grp=subject_group(subject); grade=grade_of(class_name); rows=[]
    for t in teachers["교사명"].astype(str).str.strip().tolist():
        if not t or t==absent_teacher:continue
        blocked=duty.get((norm,t),set())
        if safe_int(period) in blocked or 0 in blocked or safe_int(period) in busy.get(t,set()):continue
        if grp in groups_by_teacher[t] and grade in grades_by_teacher[t]:prio,label,score=1,"1순위 · 동일 과목 & 동일 학년",100
        elif grp in groups_by_teacher[t]:prio,label,score=2,"2순위 · 동일 과목",70
        elif grade in grades_by_teacher[t]:prio,label,score=3,"3순위 · 동일 학년",45
        else:prio,label,score=4,"4순위 · 전체 공강",20
        score += (max_cum-cum.get(t,0))*2 + max(0,22-load.get(t,0))*0.3
        tr=teachers[teachers["교사명"]==t]
        rows.append({"보강교사":t,"유형":"정규교사","우선순위":label,"_prio":prio,"담당과목":tr["담당과목"].iloc[0] if not tr.empty and "담당과목" in tr.columns else "","주당시수":load.get(t,0),"누적보강":cum.get(t,0),"추천점수":round(score,1)})
    if include_part_time:
        pt=st.session_state.get("part_time",pd.DataFrame())
        if isinstance(pt,pd.DataFrame) and not pt.empty:
            for _,r in pt.iterrows():
                name=_clean(r.get("시간강사명"))
                if name and _part_time_slot_available(r,day,period): rows.append({"보강교사":name,"유형":"시간강사","우선순위":"시간강사","_prio":5,"담당과목":_clean(r.get("담당과목")),"주당시수":0,"누적보강":0,"추천점수":5.0})
    return pd.DataFrame(rows).sort_values(["_prio","추천점수"],ascending=[True,False]).drop(columns=["_prio"]).head(top_n).reset_index(drop=True) if rows else pd.DataFrame()


def add_substitute(cid,on_date,day,period,class_name,subject,absent_teacher,sub_teacher,method,priority,memo,record=True,save=True,charge=True):
    p=safe_int(period); norm=normalize_date_str(on_date)
    if not sub_teacher or not norm or p<=0:return False
    e_tt=get_effective_timetable_for_date(norm,st.session_state.get("_data_version",0))
    # 동일 보강 건 수정 시 기존 배정을 잠시 제외해야 한다.
    existing=st.session_state.get("subs",pd.DataFrame()); e_check=e_tt
    if not is_free(sub_teacher,day,p,norm,e_check):
        st.warning(f"{sub_teacher} 선생님은 {p}교시에 이미 수업/보강이 있습니다."); return False
    old=existing.copy(deep=True)
    if not existing.empty:
        conflict=existing[(existing["일자"].apply(normalize_date_str)==norm)&(existing["교시"].apply(safe_int)==p)&(existing["보강교사"].astype(str).str.strip()==str(sub_teacher).strip())&~((existing["결강ID"]==cid)&(existing["교시"].apply(safe_int)==p))]
        if not conflict.empty:
            st.warning(f"{sub_teacher} 선생님은 {norm} {p}교시에 이미 다른 보강이 배정되어 있습니다."); return False
        existing=existing[~((existing["결강ID"]==cid)&(existing["교시"].apply(safe_int)==p))].copy()
    new=pd.DataFrame([{"결강ID":cid,"일자":norm,"요일":day,"교시":p,"학급":class_name,"과목":subject,"결강교사":absent_teacher,"보강교사":sub_teacher,"배정방식":method,"우선순위":priority,"비고":memo,"등록시각":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"입력자":current_user()}])
    st.session_state.subs=pd.concat([existing,new],ignore_index=True)
    if save and not save_work_data_to_gsheet(["subs"]):st.session_state.subs=old;_invalidate_all_caches();return False
    if charge:update_budget(-SUB_COST,f"보강 1건 ({sub_teacher} ← {absent_teacher})")
    if record:record_history(f"보강 배정 ({sub_teacher})")
    return True


def add_substitutes_batch(assignments,action_name="자동 보강"):
    if not assignments:return 0
    old=st.session_state.get("subs",pd.DataFrame()).copy(deep=True); count=0
    for args in assignments:
        if add_substitute(*args,record=False,save=False,charge=False):count+=1
    if count:
        if not save_work_data_to_gsheet(["subs"]):st.session_state.subs=old;_invalidate_all_caches();return 0
        update_budget(-SUB_COST*count,f"{action_name} {count}건");record_history(action_name)
    return count


def cancel_substitute(cid,period):
    if not can_full_data():
        s=st.session_state.subs;p=safe_int(period);m=s[(s["결강ID"]==cid)&(s["교시"].apply(safe_int)==p)]
        if not m.empty and m.iloc[0].get("입력자")!=current_user():st.warning("본인이 입력한 데이터만 취소할 수 있습니다.");return False
    p=safe_int(period);old=st.session_state.subs.copy(deep=True);st.session_state.subs=old[~((old["결강ID"]==cid)&(old["교시"].apply(safe_int)==p))].reset_index(drop=True)
    if not save_work_data_to_gsheet(["subs"]):st.session_state.subs=old;_invalidate_all_caches();return False
    update_budget(+SUB_COST,f"보강 취소 복구 ({p}교시)");record_history(f"보강 취소 ({p}교시)");return True

def _current_lesson_info(teacher,on_date,period,use_test=False):
    d=normalize_date_str(on_date); t=_clean(teacher); p=safe_int(period)
    if not d or not t or p<=0:return None
    e=get_effective_timetable_for_date(d,st.session_state.get("_data_version",0),use_test=use_test)
    lookup,_=_schedule_lookup(e); row=lookup.get((t,p))
    if row is None:return None
    return {"교사명":t,"일자":d,"요일":_clean(row.get("요일")) or WEEKDAY_KR[datetime.strptime(d,"%Y-%m-%d").weekday()],"교시":p,
            "학급":_clean(row.get("학급")),"과목":_clean(row.get("과목")),"원본교사":_clean(row.get("원본교사")),
            "원본일자":_clean(row.get("원본일자")),"원본교시":safe_int(row.get("원본교시",0)),
            "변경유형":_clean(row.get("변경유형")),"변경이력":_clean(row.get("변경이력")),"변경ID":_clean(row.get("변경ID"))}


def validate_swap(a,b,date_a,date_b,*,use_test=False,allow_linked=False):
    da,db=normalize_date_str(date_a),normalize_date_str(date_b)
    ta,tb=_clean(a.get("교사명")),_clean(b.get("교사명")); pa,pb=safe_int(a.get("교시")),safe_int(b.get("교시"))
    if not all([da,db,ta,tb,pa,pb]):return False,"교환 날짜·교사·교시 정보가 올바르지 않습니다."
    if ta==tb:return False,"같은 교사끼리는 교환할 수 없습니다."
    if (da,ta,pa)==(db,tb,pb):return False,"동일 슬롯을 자기 자신과 교환할 수 없습니다."
    ver=st.session_state.get("_data_version",0)
    ea=get_effective_timetable_for_date(da,ver,use_test=use_test); eb=get_effective_timetable_for_date(db,ver,use_test=use_test)
    la,_=_schedule_lookup(ea); lb,_=_schedule_lookup(eb)
    if (ta,pa) not in la:return False,f"현재 변경된 시간표에서 {ta}의 {da} {pa}교시 수업을 찾을 수 없습니다."
    if (tb,pb) not in lb:
        if allow_linked:return False,f"현재 변경된 시간표에서 목표 교사 {tb}의 {db} {pb}교시 상태를 확인할 수 없습니다."
        return False,f"현재 변경된 시간표에서 {tb}의 {db} {pb}교시 수업을 찾을 수 없습니다."
    if has_duty(ta,da,pa) or has_duty(tb,db,pb):return False,"복무 등록된 시간에는 교환할 수 없습니다."
    if has_absence(ta,da,pa) or has_absence(tb,db,pb):return False,"결강 등록된 시간에는 교환할 수 없습니다."
    if not allow_linked:
        if da==db and pa==pb:return False,"같은 날짜·같은 교시의 맞교환은 지원하지 않습니다. 다른 교시 또는 날짜를 선택하세요."
        if (tb,pa) in la:return False,f"{tb} 선생님이 {da} {pa}교시에 이미 수업이 있어 교환할 수 없습니다."
        if (ta,pb) in lb:return False,f"{ta} 선생님이 {db} {pb}교시에 이미 수업이 있어 교환할 수 없습니다."
    return True,""


def do_swap(a,b,date_a,date_b,is_part_time_purpose=False,is_test=False):
    cur_a=_current_lesson_info(a.get("교사명"),date_a,a.get("교시"),use_test=is_test)
    cur_b=_current_lesson_info(b.get("교사명"),date_b,b.get("교시"),use_test=is_test)
    if cur_a is None or cur_b is None:
        st.warning("현재 변경된 시간표에서 교환할 두 수업을 확인할 수 없습니다. 화면을 새로 고친 뒤 다시 선택하세요.")
        return False
    ok,msg=validate_swap(cur_a,cur_b,date_a,date_b,use_test=is_test)
    if not ok: st.warning(msg); return False
    rec={"변경ID":_next_change_id("TEST" if is_test else "SWAP"),
         "원본일자":cur_a["일자"],"교사A":cur_a["교사명"],"요일A":cur_a["요일"],"교시A":cur_a["교시"],"학급A":cur_a.get("학급",""),"과목A":cur_a.get("과목",""),
         "원본교사A":cur_a.get("원본교사",""),"원본일자A":cur_a.get("원본일자",""),"원본교시A":safe_int(cur_a.get("원본교시",0)),"변경이력A":cur_a.get("변경이력",""),
         "목표일자":cur_b["일자"],"교사B":cur_b["교사명"],"요일B":cur_b["요일"],"교시B":cur_b["교시"],"학급B":cur_b.get("학급",""),"과목B":cur_b.get("과목",""),
         "원본교사B":cur_b.get("원본교사",""),"원본일자B":cur_b.get("원본일자",""),"원본교시B":safe_int(cur_b.get("원본교시",0)),"변경이력B":cur_b.get("변경이력",""),
         "유형":"1:1 맞교환","시간강사구인":"Y" if is_part_time_purpose else "N","등록시각":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"입력자":current_user()}
    if is_test:
        st.session_state.test_swaps=pd.concat([st.session_state.get("test_swaps",pd.DataFrame()),pd.DataFrame([rec])],ignore_index=True); _invalidate_all_caches(); return True
    old=st.session_state.swaps.copy(deep=True); st.session_state.swaps=pd.concat([old,pd.DataFrame([rec])],ignore_index=True)
    if not save_work_data_to_gsheet(["swaps"]): st.session_state.swaps=old; _invalidate_all_caches(); return False
    record_history(f"맞교환 ({cur_a['교사명']} ↔ {cur_b['교사명']})"); return True

def validate_cycle_moves(moves,use_test=False):
    if not moves:return False,"순환 교환 항목이 없습니다."
    src=[(normalize_date_str(m["from_date"]),_clean(m["teacher"]),safe_int(m["from_period"])) for m in moves]
    tgt=[(normalize_date_str(m["to_date"]),_clean(m.get("next_teacher",m["teacher"])),safe_int(m["to_period"])) for m in moves]
    if len(set(src))!=len(src) or len(set(tgt))!=len(tgt):return False,"순환 교환 안에 중복 슬롯이 있습니다."
    ver=st.session_state.get("_data_version",0)
    src_set=set(src)
    for m in moves:
        fd=normalize_date_str(m["from_date"]); td=normalize_date_str(m["to_date"]); ft=_clean(m["teacher"]); nt=_clean(m.get("next_teacher",m["teacher"])); fp=safe_int(m["from_period"]); tp=safe_int(m["to_period"])
        lf,_=_schedule_lookup(get_effective_timetable_for_date(fd,ver,use_test=use_test)); lt,_=_schedule_lookup(get_effective_timetable_for_date(td,ver,use_test=use_test))
        if (ft,fp) not in lf:return False,f"{ft}의 {fd} {fp}교시 현재 수업이 없습니다."
        if has_duty(ft,fd,fp) or has_absence(ft,fd,fp):return False,"복무/결강 등록된 시간에는 순환 교환할 수 없습니다."
        if (nt,tp) not in lt and (td,nt,tp) not in src_set:return False,f"순환 목표 슬롯 {nt} {td} {tp}교시를 확인할 수 없습니다."
    return True,""

def do_linked_swap(a,teacher_b,date_a,date_b,day_b,period_b,is_part_time_purpose=False,is_test=False,subject_b=None,_skip_validation=False):
    b={"교사명":teacher_b,"일자":date_b,"요일":day_b,"교시":period_b,"학급":a.get("학급",""),"과목":subject_b if subject_b is not None else a.get("과목","")}
    if not _skip_validation:
        ok,msg=validate_swap(a,b,date_a,date_b,use_test=is_test,allow_linked=True)
        if not ok: st.warning(msg); return False
    rec={"변경ID":_next_change_id("TEST-LINK" if is_test else "LINK"),"원본일자":normalize_date_str(date_a),"교사A":a["교사명"],"요일A":a["요일"],"교시A":safe_int(a["교시"]),"학급A":a.get("학급",""),"과목A":a.get("과목",""),"목표일자":normalize_date_str(date_b),"교사B":teacher_b,"요일B":day_b,"교시B":safe_int(period_b),"학급B":a.get("학급",""),"과목B":subject_b if subject_b is not None else a.get("과목",""),"유형":"연계 공강 교환","시간강사구인":"Y" if is_part_time_purpose else "N","등록시각":datetime.now().strftime("%Y-%m-%d %H:%M:%S"),"입력자":current_user()}
    if is_test:
        st.session_state.test_swaps=pd.concat([st.session_state.get("test_swaps",pd.DataFrame()),pd.DataFrame([rec])],ignore_index=True); _invalidate_all_caches(); return True
    old=st.session_state.swaps.copy(deep=True); st.session_state.swaps=pd.concat([old,pd.DataFrame([rec])],ignore_index=True)
    if not save_work_data_to_gsheet(["swaps"]): st.session_state.swaps=old; _invalidate_all_caches(); return False
    record_history(f"연계교환 ({a['교사명']} → {teacher_b})"); return True


def apply_cycle_swaps(moves,is_test=False):
    ok,msg=validate_cycle_moves(moves,use_test=is_test)
    if not ok: st.warning(msg); return False
    rows=[]
    now=datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    for m in moves:
        rows.append({"변경ID":_next_change_id("TEST-CYCLE" if is_test else "CYCLE"),"원본일자":normalize_date_str(m["from_date"]),"교사A":m["teacher"],"요일A":m.get("day_from",WEEKDAY_KR[datetime.strptime(m["from_date"],"%Y-%m-%d").weekday()]),"교시A":safe_int(m["from_period"]),"학급A":m["class"],"과목A":m["subject"],"목표일자":normalize_date_str(m["to_date"]),"교사B":m.get("next_teacher",m["teacher"]),"요일B":WEEKDAY_KR[datetime.strptime(m["to_date"],"%Y-%m-%d").weekday()],"교시B":safe_int(m["to_period"]),"학급B":m["class"],"과목B":m.get("target_subject",m["subject"]),"유형":"연계 공강 교환","시간강사구인":"N","등록시각":now,"입력자":current_user()})
    add=pd.DataFrame(rows)
    if is_test: st.session_state.test_swaps=pd.concat([st.session_state.get("test_swaps",pd.DataFrame()),add],ignore_index=True); st.session_state["test_has_cycle"]=True; _invalidate_all_caches(); return True
    old=st.session_state.swaps.copy(deep=True); st.session_state.swaps=pd.concat([old,add],ignore_index=True)
    if not save_work_data_to_gsheet(["swaps"]): st.session_state.swaps=old; _invalidate_all_caches(); return False
    record_history(f"{len(moves)}인 연계 순환 교환"); return True


def _build_schedule_context(dates,version=0,use_test=False):
    tables={}; lookups={}; busy={}; class_period=defaultdict(list); duty=_blocked_period_map()
    for d in sorted(set(normalize_date_str(x) for x in dates if normalize_date_str(x))):
        e=get_effective_timetable_for_date(d,version,use_test=use_test); tables[d]=e
        lk,bs=_schedule_lookup(e); lookups[d]=lk; busy[d]=bs
        for (t,p),row in lk.items(): class_period[d,p,_clean(row.get("학급"))].append((t,row))
    return tables,lookups,busy,class_period,duty


def find_cycle_linked_swaps(teacher_a,date_a_str,period_a,class_a,subject_a,date_b_str,period_b,min_cycle=2,max_cycle=3,future_days=7,version=0,use_test=False):
    da=datetime.strptime(normalize_date_str(date_a_str),"%Y-%m-%d").date(); db=datetime.strptime(normalize_date_str(date_b_str),"%Y-%m-%d").date()
    origin=(da.strftime("%Y-%m-%d"),safe_int(period_a)); target=(db.strftime("%Y-%m-%d"),safe_int(period_b))
    dates=[]; cur=min(da,db)-timedelta(days=2); last=max(da,db)+timedelta(days=future_days)
    while cur<=last:
        if cur.weekday()<5: dates.append(cur.strftime("%Y-%m-%d"))
        cur+=timedelta(days=1)
    tables,lookups,busy,class_period,duty=_build_schedule_context(dates,version,use_test)
    class_slots={}
    for d in dates:
        for p in range(1,PERIODS_PER_DAY.get(WEEKDAY_KR[datetime.strptime(d,"%Y-%m-%d").weekday()],7)+1):
            for t,row in class_period.get((d,p,_clean(class_a)),[]):
                class_slots[(d,p)]={"teacher":t,"subject":_clean(row.get("과목")),"day":WEEKDAY_KR[datetime.strptime(d,"%Y-%m-%d").weekday()]}
    if class_slots.get(origin,{}).get("teacher")!=_clean(teacher_a):return [],"원본 슬롯/교사 불일치"
    if target not in class_slots:return [],"목표 슬롯에 학급 수업 없음 (공강 생성 금지)"
    if target==origin:return [],"원본과 목표 슬롯이 같습니다."
    if safe_int(period_b) in busy.get(target[0],{}).get(_clean(teacher_a),set()):return [],"교사A 목표시간 수업 있음"
    if origin not in class_slots or target not in class_slots:return [],"원본 또는 목표 학급 슬롯을 확인할 수 없습니다."
    all_teachers={v["teacher"] for v in class_slots.values()}; free={}
    for t in all_teachers:
        fs=set()
        for d in dates:
            day=WEEKDAY_KR[datetime.strptime(d,"%Y-%m-%d").weekday()]
            blocked=busy.get(d,{}).get(t,set())|duty.get((d,t),set())
            if 0 in duty.get((d,t),set()): blocked=set(range(1,PERIODS_PER_DAY.get(day,7)+1))
            for p in range(1,PERIODS_PER_DAY.get(day,7)+1):
                if p not in blocked: fs.add((d,p))
        free[t]=fs
    graph=defaultdict(list)
    class_keys=set(class_slots)
    for s,info in class_slots.items(): graph[s]=[x for x in free.get(info["teacher"],set()).intersection(class_keys) if x!=s]
    cycles=[]; seen=set()
    def dfs(cur,path,visited):
        if len(cycles)>=6 or len(path)>max_cycle:return
        if cur==origin:
            if len(path)<min_cycle:return
            slots=[origin]+path[:-1]; moves=[]
            for i,fr in enumerate(slots):
                to=slots[(i+1)%len(slots)]; inf=class_slots[fr]; tgt=class_slots[to]
                moves.append({"teacher":inf["teacher"],"from_date":fr[0],"from_period":fr[1],"to_date":to[0],"to_period":to[1],"class":class_a,"subject":inf["subject"],"target_subject":tgt["subject"],"day_from":inf["day"],"next_teacher":tgt["teacher"]})
            key=tuple((m["teacher"],m["from_date"],m["from_period"],m["to_date"],m["to_period"]) for m in moves)
            if key not in seen:
                seen.add(key); parts=[f"{m['teacher']}({m['class']} {m['from_date'][5:]} {m['from_period']}→{m['to_date'][5:]} {m['to_period']})" for m in moves]; cycles.append({"length":len(moves),"moves":moves,"path_desc":" → ".join(parts),"score":110-len(moves)*12})
            return
        for nxt in graph.get(cur,[]):
            if nxt not in visited:
                visited.add(nxt); path.append(nxt); dfs(nxt,path,visited); path.pop(); visited.remove(nxt)
    dfs(target,[target],{target}); cycles.sort(key=lambda c:(c["length"],-c["score"]))
    return cycles[:6],(f"{len(cycles)}개 순환 경로 발견" if cycles else "순환 경로 없음")


def get_target_time_recommendations(teacher_a,date_a_str,period_a,class_a,subject_a,date_b_str,period_b,budget_factor=1.0):
    da,db=normalize_date_str(date_a_str),normalize_date_str(date_b_str); ver=st.session_state.get("_data_version",0); ti=st.session_state.get("teachers",pd.DataFrame())
    if ti.empty:return pd.DataFrame(),[],""
    dates=[da,db]; _,lookups,busy,_,duty=_build_schedule_context(dates,ver,False); pa,pb=safe_int(period_a),safe_int(period_b); day_a=WEEKDAY_KR[datetime.strptime(da,"%Y-%m-%d").weekday()]; day_b=WEEKDAY_KR[datetime.strptime(db,"%Y-%m-%d").weekday()]
    source_row=lookups.get(da,{}).get((teacher_a,pa))
    if source_row is None:return pd.DataFrame(),[],"원본 수업을 찾을 수 없습니다."
    class_a=_clean(source_row.get("학급")) or _clean(class_a)
    subject_a=_clean(source_row.get("과목")) or _clean(subject_a)
    if pb in busy.get(db,{}).get(teacher_a,set()) or pb in duty.get((db,teacher_a),set()) or 0 in duty.get((db,teacher_a),set()):return pd.DataFrame(),[],"교사A 목표시간이 공강이 아닙니다."
    cum=cumulative_sub_count(version=ver); recs=[]
    for tb in ti["교사명"].astype(str).str.strip():
        if not tb or tb==teacher_a:continue
        row=lookups.get(db,{}).get((tb,pb))
        if row is None:continue
        if pa in busy.get(da,{}).get(tb,set()) or pa in duty.get((da,tb),set()) or 0 in duty.get((da,tb),set()):continue
        oc=_clean(row.get("학급")); os=_clean(row.get("과목")); same=oc==class_a; sg=grade_of(oc)==grade_of(class_a); score=(200 if same else 100 if sg else 0)+(40 if subject_group(os)==subject_group(subject_a) else 0)+(15 if da==db else 0)-cum.get(tb,0)*3; score*=budget_factor
        recs.append({"유형":"1:1","교사B":tb,"현재 수업":f"{day_b}{pb}교시 · {oc} · {os}","학급":oc,"학년":grade_of(oc),"same_class":same,"same_grade":sg,"점수":score,"b_info":{"교사명":tb,"일자":db,"요일":day_b,"교시":pb,"학급":oc,"과목":os}})
    df=pd.DataFrame(recs).sort_values(["same_class","same_grade","점수"],ascending=[False,False,False]).reset_index(drop=True) if recs else pd.DataFrame()
    cycles,msg=find_cycle_linked_swaps(teacher_a,da,pa,class_a,subject_a,db,pb,max_cycle=3,future_days=7,version=ver)
    return df,cycles,msg


def get_weekly_1to1_swap_table(teacher,ref_date,future_days=0,version=0):
    monday=ref_date-timedelta(days=ref_date.weekday()); days=[(monday+timedelta(days=i)).strftime("%Y-%m-%d") for i in range(5)]; search=list(days)
    if future_days>0:
        for i in range(1,future_days+1):
            d=monday+timedelta(days=4+i)
            if d.weekday()<5:search.append(d.strftime("%Y-%m-%d"))
    _,lookups,busy,_,duty=_build_schedule_context(search,version,False); out=[]; seen=set()
    for d in days:
        day=WEEKDAY_KR[datetime.strptime(d,"%Y-%m-%d").weekday()]
        for (t,p),lesson in lookups.get(d,{}).items():
            if t!=teacher:continue
            my_class=_clean(lesson.get("학급")); my_subj=_clean(lesson.get("과목"))
            for td in search:
                if td==d:continue
                tday=WEEKDAY_KR[datetime.strptime(td,"%Y-%m-%d").weekday()]
                if p in busy.get(td,{}).get(teacher,set()) or p in duty.get((td,teacher),set()) or 0 in duty.get((td,teacher),set()):continue
                for (ot,op),other in lookups.get(td,{}).items():
                    if op!=p or ot==teacher:continue
                    if p in busy.get(d,{}).get(ot,set()) or p in duty.get((d,ot),set()) or 0 in duty.get((d,ot),set()):continue
                    key=(d,p,td,ot)
                    if key in seen:continue
                    seen.add(key); oc=_clean(other.get("학급")); os=_clean(other.get("과목")); same=oc==my_class; sg=grade_of(oc)==grade_of(my_class); score=(200 if same else 100 if sg else 0)+(40 if subject_group(os)==subject_group(my_subj) else 0)+(10 if td[:7]==d[:7] else 0)
                    out.append({"원본일자":d,"원본요일":day,"원본교시":p,"원본학급":my_class,"원본과목":my_subj,"이동희망일":td,"이동요일":tday,"이동교시":p,"상대교사":ot,"상대수업":f"{oc} {os}","동일학급":"🏆" if same else "","동학년":"⚠" if sg and not same else "","점수":score,"_sort":(0 if same else 1,0 if sg else 1,-score)})
    return pd.DataFrame(out).sort_values("_sort").drop(columns=["_sort"]).reset_index(drop=True) if out else pd.DataFrame()


@st.cache_data(show_spinner=False,ttl=180)
def get_single_lesson_1to1_candidates(teacher,orig_date_str,orig_period,orig_class,orig_subject,future_days=0,version=0,use_test=False):
    src=datetime.strptime(normalize_date_str(orig_date_str),"%Y-%m-%d").date(); src_s=src.strftime("%Y-%m-%d"); monday=src-timedelta(days=src.weekday()); dates=[monday+timedelta(days=i) for i in range(5)]
    if future_days>0:
        friday=dates[-1]; dates.extend(friday+timedelta(days=i) for i in range(1,future_days+1) if (friday+timedelta(days=i)).weekday()<5)
    ds=[d.strftime("%Y-%m-%d") for d in dates]; ver=version or st.session_state.get("_data_version",0); _,lookups,busy,class_period,duty=_build_schedule_context(ds,ver,use_test)
    p=safe_int(orig_period); src_row=lookups.get(src_s,{}).get((_clean(teacher),p))
    if src_row is None:return pd.DataFrame()
    orig_class=_clean(src_row.get("학급")) or _clean(orig_class)
    orig_subject=_clean(src_row.get("과목")) or _clean(orig_subject)
    results=[]; seen=set(); sg=subject_group(orig_subject); source_day=WEEKDAY_KR[src.weekday()]
    for td in dates:
        td_s=td.strftime("%Y-%m-%d")
        if td_s==src_s:continue
        tday=WEEKDAY_KR[td.weekday()]
        if p in busy.get(td_s,{}).get(_clean(teacher),set()) or p in duty.get((td_s,_clean(teacher)),set()) or 0 in duty.get((td_s,_clean(teacher)),set()):continue
        for ot,row in class_period.get((td_s,p,_clean(orig_class)),[]):
            if ot==_clean(teacher):continue
            if p in busy.get(src_s,{}).get(ot,set()) or p in duty.get((src_s,ot),set()) or 0 in duty.get((src_s,ot),set()):continue
            key=(td_s,ot)
            if key in seen:continue
            seen.add(key); os=_clean(row.get("과목")); score=200+(40 if subject_group(os)==sg else 0)+(10 if td_s[:7]==src_s[:7] else 0)
            results.append({"원본일자":src_s,"원본요일":source_day,"원본교시":p,"원본학급":orig_class,"원본과목":orig_subject,"이동희망일":td_s,"이동요일":tday,"이동교시":p,"상대교사":ot,"상대수업":f"{orig_class} {os}","동일학급":"🏆","동학년":"","점수":score,"상대학급":orig_class,"상대과목":os})
    return pd.DataFrame(results).sort_values(["점수","이동희망일","상대교사"],ascending=[False,True,True]).reset_index(drop=True) if results else pd.DataFrame()


@st.cache_data(show_spinner=False,ttl=180)
def get_single_lesson_linked_cycles(teacher,orig_date_str,orig_period,orig_class,orig_subject,future_days=0,version=0,min_cycle=2,max_cycle=3,use_test=False):
    src=datetime.strptime(normalize_date_str(orig_date_str),"%Y-%m-%d").date(); src_s=src.strftime("%Y-%m-%d"); monday=src-timedelta(days=src.weekday()); dates=[monday+timedelta(days=i) for i in range(5)]
    if future_days>0:
        friday=dates[-1]; dates.extend(friday+timedelta(days=i) for i in range(1,future_days+1) if (friday+timedelta(days=i)).weekday()<5)
    ver=version or st.session_state.get("_data_version",0); _,lookups,busy,class_period,duty=_build_schedule_context([d.strftime("%Y-%m-%d") for d in dates],ver,use_test); candidates=[]
    for d in dates:
        ds=d.strftime("%Y-%m-%d")
        if ds==src_s:continue
        day=WEEKDAY_KR[d.weekday()]
        for t,p in [(t,p) for (t,p),row in lookups.get(ds,{}).items() if _clean(row.get("학급"))==_clean(orig_class)]:
            if p in busy.get(ds,{}).get(_clean(teacher),set()) or p in duty.get((ds,_clean(teacher)),set()) or 0 in duty.get((ds,_clean(teacher)),set()):continue
            row=lookups[ds][(t,p)]; pri=(0 if p==safe_int(orig_period) else 1,abs((d-src).days),0 if subject_group(_clean(row.get("과목")))==subject_group(orig_subject) else 1); candidates.append((pri,ds,p))
    if not candidates:return [],"연계 순환을 시작할 수 있는 빈 시간대가 없습니다."
    candidates.sort(); all_cycles=[]; seen=set()
    for _,ds,p in candidates[:10]:
        cyc,_=find_cycle_linked_swaps(teacher,src_s,orig_period,orig_class,orig_subject,ds,p,min_cycle=min_cycle,max_cycle=max_cycle,future_days=future_days,version=ver,use_test=use_test)
        for c in cyc:
            key=tuple((m["teacher"],m["from_date"],m["from_period"],m["to_date"],m["to_period"]) for m in c["moves"])
            if key not in seen:seen.add(key);all_cycles.append(c)
        if len(all_cycles)>=6:break
    all_cycles.sort(key=lambda c:(c["length"],-c["score"]))
    return (all_cycles[:6],f"{len(all_cycles)}개 연계 순환 경로 발견") if all_cycles else ([],"조건을 만족하는 연계 순환 경로가 없습니다.")


# ==========================================================================================
# 뷰 헬퍼 / 고속 인덱스
# ==========================================================================================
def _format_effective_cell(row,teacher,on_date,period,use_test=False):
    if row is None:return ""
    cell=f"{_clean(row.get('학급',''))} {_clean(row.get('과목',''))}".strip()
    typ=_clean(row.get("변경유형",""))
    if typ=="보강": return f"🟢 {cell}"
    if typ=="시간강사": return f"🟡 {cell}"
    if typ in ("1:1 맞교환","연계 공강 교환"): return f"🔄 {cell}"
    return cell


@st.cache_data(show_spinner=False)
def teacher_matrix(version=0):
    tt=st.session_state.timetable
    if tt.empty:return pd.DataFrame()
    idx={( _clean(r.교사명),r.요일,safe_int(r.교시)):f"{r.학급} {r.과목}" for r in tt.itertuples(index=False)}
    rows=[]
    for t in sorted(tt["교사명"].astype(str).str.strip().unique()):
        row={"교사명":t}
        for d in DAYS:
            for p in range(1,PERIODS_PER_DAY.get(d,7)+1):row[f"{d}{p}"]=idx.get((t,d,p),"")
        rows.append(row)
    return pd.DataFrame(rows)


@st.cache_data(show_spinner=False,ttl=180)
def effective_teacher_matrix(ref_date,version=0,use_test=False):
    monday=ref_date-timedelta(days=ref_date.weekday()); daily={}; names=set()
    base=st.session_state.get("timetable",pd.DataFrame())
    if not base.empty:names.update(base["교사명"].dropna().astype(str).str.strip())
    for i,d in enumerate(DAYS):
        ds=(monday+timedelta(days=i)).strftime("%Y-%m-%d"); e=get_effective_timetable_for_date(ds,version,use_test=use_test); daily[d]=e
        if not e.empty:names.update(e["교사명"].dropna().astype(str).str.strip())
    rows=[]
    for t in sorted(x for x in names if x):
        row={"교사명":t}
        for d in DAYS:
            lk,_=_schedule_lookup(daily[d])
            for p in range(1,PERIODS_PER_DAY.get(d,7)+1):row[f"{d}{p}"]=_format_effective_cell(lk.get((t,p)),t,"",p,use_test)
        rows.append(row)
    return pd.DataFrame(rows)


@st.cache_data(show_spinner=False)
def class_matrix(version=0):
    tt=st.session_state.timetable
    if tt.empty:return pd.DataFrame()
    idx={( _clean(r.학급),r.요일,safe_int(r.교시)):f"{r.교사명} {r.과목}" for r in tt.itertuples(index=False)}
    rows=[]
    for c in sorted(tt["학급"].astype(str).str.strip().unique()):
        row={"학급":c}
        for d in DAYS:
            for p in range(1,PERIODS_PER_DAY.get(d,7)+1):row[f"{d}{p}"]=idx.get((c,d,p),"")
        rows.append(row)
    return pd.DataFrame(rows)


def get_teacher_week_view(teacher,ref_date,use_test=False):
    monday=ref_date-timedelta(days=ref_date.weekday()); dates=[monday+timedelta(days=i) for i in range(5)]; ver=st.session_state.get("_data_version",0); lookups={}
    for d in dates:
        ds=d.strftime("%Y-%m-%d"); lookups[ds],_=_schedule_lookup(get_effective_timetable_for_date(ds,ver,use_test=use_test))
    abs_df=st.session_state.get("absences",pd.DataFrame()); subs=st.session_state.get("subs",pd.DataFrame()); pt=st.session_state.get("part_time",pd.DataFrame())
    abs_keys=set(); sub_keys=set()
    if isinstance(abs_df,pd.DataFrame) and not abs_df.empty:
        for rr in abs_df.itertuples(index=False): abs_keys.add((normalize_date_str(getattr(rr,"일자","")),_clean(getattr(rr,"교사명","")),safe_int(getattr(rr,"교시",0))))
    if isinstance(subs,pd.DataFrame) and not subs.empty:
        for rr in subs.itertuples(index=False): sub_keys.add((normalize_date_str(getattr(rr,"일자","")),_clean(getattr(rr,"보강교사","")),safe_int(getattr(rr,"교시",0))))
    grid=[]
    for p in range(1,MAX_PERIOD+1):
        row={"교시":p}
        for i,d in enumerate(DAYS):
            ds=dates[i].strftime("%Y-%m-%d"); r=lookups[ds].get((_clean(teacher),p)); cell=_format_effective_cell(r,teacher,ds,p,use_test) if r else ""
            if r:
                if _clean(r.get("변경이력")):cell+=f" · {r['변경이력']}"
                if (ds,_clean(teacher),p) in abs_keys:cell=f"[결강] {cell}"
                if (ds,_clean(teacher),p) in sub_keys:cell=f"[보강] {cell}"
            else:
                # 시간강사로 인해 원 교사 자리에 공백이 된 경우
                if not pt.empty and "시작일" in pt.columns:
                    for _,prow in pt.iterrows():
                        stt=normalize_date_str(prow.get("시작일","")); en=normalize_date_str(prow.get("종료일",""))
                        if stt and en and stt<=ds<=en and _clean(prow.get("대체교사"))==teacher:
                            pt_name=_clean(prow.get("시간강사명")); rr=lookups[ds].get((pt_name,p)) if pt_name else None
                            if rr:cell=f"🟡 {pt_name}({teacher}) {rr.get('학급','')} {rr.get('과목','')}"
                            break
            row[d]=cell
        grid.append(row)
    return pd.DataFrame(grid),dates


def get_changed_teachers_for_week(ref_date):
    monday=ref_date-timedelta(days=ref_date.weekday()); week=[(monday+timedelta(days=i)).strftime("%Y-%m-%d") for i in range(5)]; changed=set()
    datasets=[(st.session_state.absences,["교사명"],"일자"),(st.session_state.subs,["결강교사","보강교사"],"일자"),(st.session_state.swaps,["교사A","교사B"],None)]
    for df,cols,dcol in datasets:
        if not isinstance(df,pd.DataFrame) or df.empty:continue
        if dcol:
            mask=df[dcol].apply(normalize_date_str).isin(week)
            for c in cols:
                if c in df.columns:changed.update(_clean(x) for x in df.loc[mask,c])
        else:
            mask=df["원본일자"].apply(normalize_date_str).isin(week)|df["목표일자"].apply(normalize_date_str).isin(week)
            for c in cols:
                if c in df.columns:changed.update(_clean(x) for x in df.loc[mask,c])
    pt=st.session_state.get("part_time",pd.DataFrame())
    if isinstance(pt,pd.DataFrame) and not pt.empty:
        for _,r in pt.iterrows():
            stt=normalize_date_str(r.get("시작일","")); en=normalize_date_str(r.get("종료일",""))
            if stt and en and any(stt<=d<=en for d in week):
                changed.add(_clean(r.get("대체교사","")));changed.add(_clean(r.get("시간강사명","")))
    return sorted(x for x in changed if x)


def filter_by_owner(df):
    if can_full_data() or df.empty or "입력자" not in df.columns:return df
    return df[df["입력자"]==current_user()].copy()


# ==========================================================================================
# ★ 결보강 계획서 (보강수업 칸에 실제 보강교사 표시)
# ==========================================================================================
def build_personal_plan_html(teacher_name: str, on_date: str) -> str:
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

    abs_list = []
    if not my_abs.empty:
        for _, r in my_abs.iterrows():
            p = safe_int(r.get("교시"))
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
                "보강교사": sub_teacher
            })

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

def build_test_swaps_report_html(test_swaps: pd.DataFrame) -> str:
    if test_swaps.empty:
        return "<html><body><p>테스트 중인 맞교환 내역이 없습니다.</p></body></html>"

    my_name = current_name().strip()
    if not my_name:
        my_name = current_user()

    my_swaps = test_swaps[
        (test_swaps["교사A"] == my_name) | (test_swaps["교사B"] == my_name)
    ].copy()

    if my_swaps.empty:
        return f"<html><body><p>{my_name} 선생님의 테스트 맞교환 내역이 없습니다.</p></body></html>"

    origin_dates = []
    for _, r in my_swaps.iterrows():
        da = str(r.get("원본일자", "")).strip()
        if da and str(r.get("교사A", "")).strip() == my_name:
            origin_dates.append(da)

    if not origin_dates:
        for _, r in my_swaps.iterrows():
            da = str(r.get("원본일자", "")).strip()
            if da:
                origin_dates.append(da)

    if not origin_dates:
        return f"<html><body><p>{my_name} 선생님의 결강 대상 일자가 없습니다.</p></body></html>"

    on_date = sorted(set(origin_dates))[0]
    teacher_name = my_name

    try:
        dt = datetime.strptime(on_date, "%Y-%m-%d")
        day_kr = WEEKDAY_KR[dt.weekday()]
        date_display = f"{dt.year}년 {dt.month}월 {dt.day}일 ({day_kr})"
    except Exception:
        date_display = on_date
        day_kr = ""

    subject_dept = get_teacher_subject(teacher_name)
    dept_line = f"{subject_dept} 과" if subject_dept else "과"

    day_swaps = my_swaps[my_swaps["원본일자"] == on_date].copy()

    absence_list = []
    swap_list = []

    for _, r in day_swaps.iterrows():
        if str(r.get("교사A", "")) == teacher_name:
            absence_list.append({
                "월일": on_date[5:].replace("-", "/") if len(on_date) >= 10 else on_date,
                "교시": safe_int(r.get("교시A", 0)),
                "학년반": str(r.get("학급A", "")),
                "과목": str(r.get("과목A", "")),
                "보강교사": ""
            })
            target_date = str(r.get("목표일자", ""))
            swap_list.append({
                "월일": target_date[5:].replace("-", "/") if len(target_date) >= 10 else target_date,
                "교시": safe_int(r.get("교시B", 0)),
                "과목": str(r.get("과목B", "")),
                "교사": str(r.get("교사B", ""))
            })

    rows_html = ""
    max_rows = 6
    for i in range(max_rows):
        a = absence_list[i] if i < len(absence_list) else {
            "월일": "", "교시": "", "학년반": "", "과목": "", "보강교사": ""
        }
        s = swap_list[i] if i < len(swap_list) else {
            "월일": "", "교시": "", "과목": "", "교사": ""
        }
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

    page_html = f"""
<div class="page">
    <div class="title">결 · 보 강  계 획 </div>

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
            <td style="width: 68px; background:#f0f0f0; font-weight:bold;">해당<br>일자</td>
            <td style="text-align:left; padding-left:10px;">
                {date_display}<br>
                <span style="display:inline-block; margin-top:2px;">사유 : </span>
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
</div>"""

    html = f"""<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8">
<title>결보강 계획서</title>
<style>
    @page {{ size: A4; margin: 12mm 14mm; }}
    * {{ box-sizing: border-box; }}
    body {{
        font-family: '맑은 고딕', 'Malgun Gothic', 'Apple SD Gothic Neo', sans-serif;
        font-size: 12.5px;
        line-height: 1.25;
        margin: 0;
        padding: 0;
        color: #000;
        background: #fff;
    }}
    .page {{
        padding: 6px 10px;
        page-break-after: always;
    }}
    .page:last-child {{
        page-break-after: auto;
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
{page_html}
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

    # ===== 구글 드라이브 이미지 =====
    IMAGE_URL = "https://i.imgur.com/Gl0YDO3.jpeg"
    st.markdown(
        f"""
        <div style="background-color: #f0f2f6; padding: 20px; border-radius: 10px; text-align: center; margin-bottom: 20px;">
            <img src="{IMAGE_URL}" width="150" style="object-fit: contain;">
        </div>
        """, 
        unsafe_allow_html=True
    )
    # ==============================

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
        st.info("시작일·종료일·대체교사 + 가능시간표를 입력하면 해당 기간의 해당 교시만 시간강사로 대체됩니다. 가능시간표를 비워두면 기존 방식처럼 전체 교시가 대상입니다.")
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
                if save_work_data_to_gsheet(["part_time"]):
                    record_history("시간강사 수정")
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
                    if save_work_data_to_gsheet(["absences"]):
                        record_history(f"결강 ({who})")
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
                    assignments=[]
                    e_auto=get_effective_timetable_for_date(head["일자"],st.session_state.get("_data_version",0))
                    for r in rows.itertuples():
                        cand=recommend_substitutes(head["요일"],r.교시,r.과목,r.학급,head["교사명"],head["일자"],top_n=1,include_part_time=include_pt,e_tt=e_auto)
                        if not cand.empty:
                            assignments.append((cid,head["일자"],head["요일"],r.교시,r.학급,r.과목,head["교사명"],cand.iloc[0]["보강교사"],"자동",cand.iloc[0]["우선순위"],""))
                    count=add_substitutes_batch(assignments,"자동 보강")
                    st.success(f"{count}건 자동 배정 완료")
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
        # ========== 단일 수업 → 주간 1:1 (클릭 가능한 버튼 버전) ==========
        st.markdown("### 📅 단일 수업 → 주간 1:1 가능 위치")

        st.markdown("#### 현재 수업 선택 — 변경 반영 매트릭스에서 수업 셀 하나를 클릭")
        st.caption("수업이 적힌 셀을 클릭하면 교사·요일·교시를 자동으로 읽어, 그 날짜의 동일 학급 1:1 후보를 바로 표시합니다.")
        week_anchor = st.date_input("검색 기준 주", value=date.today(), key="week_1to1_week_anchor")
        ver = st.session_state.get("_data_version", 0)
        lesson_matrix = effective_teacher_matrix(week_anchor, ver).copy()

        if lesson_matrix.empty:
            st.warning("교사 시간표 데이터가 없습니다.")
        else:
            matrix_event = st.dataframe(
                lesson_matrix,
                hide_index=True,
                use_container_width=True,
                height=520,
                key="week_1to1_lesson_matrix",
                on_select="rerun",
                selection_mode="single-cell"
            )
            selected_cells = matrix_event.selection.cells
            extra_days = st.slider("미래 추가 검색 일수", 0, 14, 7, key="week_extra_days")

            if selected_cells:
                row_idx, column_name = selected_cells[0]
                day_kr = str(column_name)[:1]
                orig_period = safe_int(str(column_name)[1:])
                cell_value = str(lesson_matrix.iloc[row_idx][column_name]).strip()

                if day_kr in DAYS and orig_period >= 1 and cell_value:
                    week_teacher = str(lesson_matrix.iloc[row_idx]["교사명"]).strip()
                    monday = week_anchor - timedelta(days=week_anchor.weekday())
                    orig_date = monday + timedelta(days=DAYS.index(day_kr))
                    e_orig = get_effective_timetable_for_date(orig_date.strftime("%Y-%m-%d"), ver)
                    selected_lesson = e_orig[
                        (e_orig["교사명"] == week_teacher)
                        & (e_orig["요일"] == day_kr)
                        & (e_orig["교시"] == orig_period)
                    ]

                    if selected_lesson.empty:
                        st.warning("선택한 셀의 수업을 해당 날짜 시간표에서 찾을 수 없습니다.")
                        st.session_state.pop("_single_week_df", None)
                        st.session_state.pop("_single_orig", None)
                        st.session_state.pop("_single_linked_cycles", None)
                        st.session_state.pop("_single_linked_cycle_msg", None)
                    else:
                        lesson = selected_lesson.iloc[0]
                        cycle_context = (
                            week_teacher, orig_date.strftime("%Y-%m-%d"), orig_period,
                            str(lesson["학급"]), str(lesson["과목"]), extra_days, ver
                        )
                        if st.session_state.get("_single_cycle_context") != cycle_context:
                            st.session_state["_single_cycle_context"] = cycle_context
                            st.session_state["_single_show_extended_cycles"] = False
                        st.info(f"**현재 수업**: {orig_date.strftime('%Y-%m-%d')} ({day_kr}) {orig_period}교시 · {lesson['학급']} · {lesson['과목']}")
                        if _clean(lesson.get("변경이력", "")):
                            st.caption(f"변경 이력: {lesson['변경이력']}")
                        st.caption("🏆 동일 학급 후보만 자동 검색합니다. 현재 변경 결과를 기준으로 계산합니다.")
                        df_week = get_single_lesson_1to1_candidates(
                            week_teacher, orig_date.strftime("%Y-%m-%d"), orig_period,
                            str(lesson["학급"]), str(lesson["과목"]),
                            future_days=extra_days, version=ver
                        )
                        if df_week.empty:
                            linked_cycles, linked_cycle_msg = get_single_lesson_linked_cycles(
                                week_teacher, orig_date.strftime("%Y-%m-%d"), orig_period,
                                str(lesson["학급"]), str(lesson["과목"]),
                                future_days=extra_days, version=ver, min_cycle=2, max_cycle=3
                            )
                            if st.session_state.get("_single_show_extended_cycles", False):
                                extended_cycles, extended_msg = get_single_lesson_linked_cycles(
                                    week_teacher, orig_date.strftime("%Y-%m-%d"), orig_period,
                                    str(lesson["학급"]), str(lesson["과목"]),
                                    future_days=extra_days, version=ver, min_cycle=4, max_cycle=6
                                )
                                linked_cycles.extend(extended_cycles)
                                linked_cycle_msg = f"기본: {linked_cycle_msg} / 확장: {extended_msg}"
                        else:
                            linked_cycles, linked_cycle_msg = [], ""

                        st.session_state["_single_week_df"] = df_week
                        st.session_state["_single_linked_cycles"] = linked_cycles
                        st.session_state["_single_linked_cycle_msg"] = linked_cycle_msg
                        st.session_state["_single_orig"] = {
                            "teacher": week_teacher,
                            "date": orig_date.strftime("%Y-%m-%d"),
                            "day": day_kr,
                            "period": orig_period,
                            "class": str(lesson["학급"]),
                            "subject": str(lesson["과목"])
                        }
                else:
                    st.info("교사명이나 빈칸이 아닌, 수업 내용이 적힌 셀을 클릭하세요.")
                    st.session_state.pop("_single_week_df", None)
                    st.session_state.pop("_single_orig", None)
                    st.session_state.pop("_single_linked_cycles", None)
                    st.session_state.pop("_single_linked_cycle_msg", None)
            else:
                st.info("시간표 매트릭스에서 원본 수업 셀 하나를 클릭하세요.")

            if "_single_week_df" in st.session_state:
                dfw = st.session_state["_single_week_df"]
                orig = st.session_state["_single_orig"]

                if dfw.empty:
                    st.warning("동일 학급 1:1 교환 후보가 없습니다.")
                    linked_cycles = st.session_state.get("_single_linked_cycles", [])
                    linked_cycle_msg = st.session_state.get("_single_linked_cycle_msg", "")
                    st.markdown("#### 🔗 연계 순환 교환 (2·3인 우선)")
                    st.caption(linked_cycle_msg)
                    if not st.session_state.get("_single_show_extended_cycles", False):
                        if st.button("4~6인 순환도 추가 검색", key="matrix_show_extended_cycles"):
                            st.session_state["_single_show_extended_cycles"] = True
                            st.rerun()
                        st.caption("기본 목록에는 2·3인 순환만 표시됩니다.")
                    else:
                        st.info("확장 검색 결과가 포함되어 있습니다: 4~6인 순환")
                    if not linked_cycles:
                        st.info("조건을 만족하는 연계 순환 경로가 없습니다.")
                    else:
                        for idx, cyc in enumerate(linked_cycles):
                            with st.expander(
                                f"{'✅' if cyc['length'] == 2 else '🔗'} {cyc['length']}인 순환 · 점수 {cyc['score']}",
                                expanded=(idx == 0)
                            ):
                                st.markdown(f"**경로**: `{cyc['path_desc']}`")
                                st.caption("학급의 담당교사·과목·시수가 보존되는 순환 교환입니다.")
                                if st.button("이 순환 적용하기", key=f"matrix_cyc_{idx}"):
                                    apply_cycle_swaps(cyc["moves"], is_test=False)
                                    st.session_state.pop("_single_week_df", None)
                                    st.session_state.pop("_single_linked_cycles", None)
                                    st.success(f"{cyc['length']}인 순환 교환이 등록되었습니다.")
                                    st.rerun()
                else:
                    st.success(f"총 {len(dfw)}건의 가능한 위치")

                    # ---------- 클릭 가능한 버튼 그리드 ----------
                    st.markdown("#### 가능한 이동 위치 (버튼을 클릭하면 바로 적용)")
                    st.caption("원하는 요일·교시 버튼을 누르면 즉시 맞교환이 등록됩니다.")

                    # 요일별로 그룹화
                    day_order = ["월", "화", "수", "목", "금"]
                    for day in day_order:
                        day_df = dfw[dfw["이동요일"] == day]
                        if day_df.empty:
                            continue

                        st.markdown(f"**{day}요일**")
                        cols = st.columns(4)  # 한 줄에 4개씩

                        for idx, (_, row) in enumerate(day_df.iterrows()):
                            with cols[idx % 4]:
                                target_period = row["이동교시"] if "이동교시" in row.index else row["원본교시"]
                                btn_label = (
                                    f"{row['이동희망일']} ({row['이동요일']}) {target_period}교시\n"
                                    f"{row['상대교사']}\n{row['상대수업']}"
                                )
                                if st.button(btn_label, key=f"apply_{day}_{idx}_{row['상대교사']}", use_container_width=True):
                                    a_info = {
                                        "교사명": orig["teacher"],
                                        "일자": orig["date"],
                                        "요일": orig["day"],
                                        "교시": orig["period"],
                                        "학급": orig["class"],
                                        "과목": orig["subject"]
                                    }
                                    b_info = {
                                        "교사명": row["상대교사"],
                                        "일자": row["이동희망일"],
                                        "요일": row["이동요일"],
                                        "교시": safe_int(row["이동교시"] if "이동교시" in row.index else row["원본교시"]),
                                        "학급": row.get("상대학급", "") or (str(row["상대수업"]).split()[0] if " " in str(row["상대수업"]) else ""),
                                        "과목": row.get("상대과목", "") or (" ".join(str(row["상대수업"]).split()[1:]) if " " in str(row["상대수업"]) else str(row["상대수업"]))
                                    }
                                    if do_swap(a_info, b_info, orig["date"], row["이동희망일"]):
                                        st.success(f"✅ {orig['teacher']} ↔ {row['상대교사']} 맞교환 완료!")
                                        st.session_state.pop("_single_week_df", None)
                                        st.rerun()
                                    else:
                                        st.error("등록 실패")

                    # 엑셀 다운로드
                    st.divider()
                    try:
                        xls = to_excel_bytes({"후보": dfw})
                        st.download_button("📥 엑셀 다운로드", data=xls,
                                           file_name=f"1대1_{week_teacher}_{orig_date}.xlsx",
                                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    except:
                        pass

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
            st.session_state["test_has_cycle"] = False
            get_effective_timetable_for_date.clear()
            effective_teacher_matrix.clear()
            get_single_lesson_1to1_candidates.clear()
            get_single_lesson_linked_cycles.clear()
            st.success("테스트 상태가 초기화되었습니다.")
            st.rerun()

        tlist = st.session_state.teachers["교사명"].tolist()
        st.markdown("#### 원본 수업 선택 — 변경 반영 매트릭스에서 수업 셀 하나를 클릭")
        test_week_anchor = st.date_input("테스트 검색 기준 주", value=date.today(), key="test_week_anchor")
        ver = st.session_state.get("_data_version", 0)
        test_matrix = effective_teacher_matrix(test_week_anchor, ver, use_test=True)
        test_pick_a = None
        if test_matrix.empty:
            st.warning("테스트용 시간표 데이터가 없습니다.")
        else:
            test_event = st.dataframe(
                test_matrix, hide_index=True, use_container_width=True, height=520,
                key="test_lesson_matrix", on_select="rerun", selection_mode="single-cell"
            )
            test_cells = test_event.selection.cells
            test_extra_days = st.slider("테스트 미래 추가 검색 일수", 0, 14, 7, key="test_extra_days")
            if test_cells:
                test_row_idx, test_column = test_cells[0]
                test_day = str(test_column)[:1]
                test_period = safe_int(str(test_column)[1:])
                test_value = str(test_matrix.iloc[test_row_idx][test_column]).strip()
                if test_day in DAYS and test_period >= 1 and test_value:
                    test_teacher = str(test_matrix.iloc[test_row_idx]["교사명"]).strip()
                    test_monday = test_week_anchor - timedelta(days=test_week_anchor.weekday())
                    test_date = test_monday + timedelta(days=DAYS.index(test_day))
                    test_date_str = test_date.strftime("%Y-%m-%d")
                    test_tt = get_effective_timetable_for_date(test_date_str, ver, use_test=True)
                    test_lesson = test_tt[
                        (test_tt["교사명"] == test_teacher)
                        & (test_tt["요일"] == test_day)
                        & (test_tt["교시"] == test_period)
                    ]
                    if not test_lesson.empty:
                        test_lesson = test_lesson.iloc[0]
                        test_pick_a = {
                            "교사명": test_teacher, "일자": test_date_str, "요일": test_day,
                            "교시": test_period, "학급": test_lesson["학급"], "과목": test_lesson["과목"]
                        }
                        st.info(f"**현재 테스트 수업**: {test_date_str} ({test_day}) {test_period}교시 · {test_lesson['학급']} · {test_lesson['과목']}")
                        if _clean(test_lesson.get("변경이력", "")):
                            st.caption(f"변경 이력: {test_lesson['변경이력']}")
                else:
                    st.info("교사명이나 빈칸이 아닌 수업 셀을 클릭하세요.")
            else:
                st.info("매트릭스에서 테스트할 원본 수업 셀 하나를 클릭하세요.")

        if test_pick_a:
            df_swap = get_single_lesson_1to1_candidates(
                test_pick_a["교사명"], test_pick_a["일자"], test_pick_a["교시"],
                str(test_pick_a["학급"]), str(test_pick_a["과목"]),
                future_days=test_extra_days, version=ver, use_test=True
            )
            cycles, cycle_msg = ([], "")
            if df_swap.empty:
                cycles, cycle_msg = get_single_lesson_linked_cycles(
                    test_pick_a["교사명"], test_pick_a["일자"], test_pick_a["교시"],
                    str(test_pick_a["학급"]), str(test_pick_a["과목"]),
                    future_days=test_extra_days, version=ver, min_cycle=2, max_cycle=3, use_test=True
                )

            t1, t2 = st.tabs(["1:1 맞교환 테스트", "연계 순환 교환 테스트"])
            with t1:
                if df_swap.empty:
                    st.info("동일 학급 1:1 후보 없음")
                else:
                    for idx, row in df_swap.iterrows():
                        label = f"{row['이동희망일']} ({row['이동요일']}) {row.get('이동교시', row['원본교시'])}교시 · {row['상대교사']} · {row['상대수업']}"
                        if st.button(f"[테스트] {label}", key=f"test_matrix_swap_{idx}"):
                            b_info = {
                                "교사명": row["상대교사"], "일자": row["이동희망일"], "요일": row["이동요일"],
                                "교시": safe_int(row.get("이동교시", row["원본교시"])),
                                "학급": row.get("상대학급", "") or (str(row.get("상대수업", "")).split()[0] if str(row.get("상대수업", "")).strip() else ""),
                                "과목": row.get("상대과목", "") or " ".join(str(row.get("상대수업", "")).split()[1:])
                            }
                            do_swap(test_pick_a, b_info, test_pick_a["일자"], row["이동희망일"], is_test=True)
                            st.success("테스트 맞교환이 적용되었습니다. 저장되지 않습니다.")
                            st.rerun()
            with t2:
                st.caption(cycle_msg or "1:1 후보가 있을 때는 연계 순환을 표시하지 않습니다.")
                if not cycles:
                    st.info("테스트용 2·3인 순환 경로 없음")
                else:
                    for idx, cyc in enumerate(cycles):
                        with st.expander(f"{'✅' if cyc['length']==2 else '🔗'} {cyc['length']}인 순환", expanded=(idx==0)):
                            st.markdown(f"**경로**: `{cyc['path_desc']}`")
                            if st.button("[테스트] 이 순환 적용", key=f"test_matrix_cyc_{idx}"):
                                apply_cycle_swaps(cyc["moves"], is_test=True)
                                st.session_state["test_has_cycle"] = True
                                st.success(f"테스트 {cyc['length']}인 순환이 적용되었습니다. 저장되지 않습니다.")
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

            col_btn1, col_btn2 = st.columns(2)
            with col_btn1:
                if st.session_state.get("test_has_cycle", False):
                    st.warning("연계 순환 교환은 교육과정부로 문의 바랍니다")
                else:
                    if st.button("현재 테스트 중인 맞교환 목록 결보강 계획서 출력하기", type="primary", key="test_report_btn"):
                        html = build_test_swaps_report_html(st.session_state.test_swaps)
                        st.download_button(
                            "HTML 다운로드 (테스트 결보강 계획서)",
                            html.encode("utf-8"),
                            f"테스트_결보강계획서_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html",
                            "text/html",
                            key="test_html_dl"
                        )
                        st.info("이 HTML을 브라우저에서 열고 Ctrl+P → PDF로 저장하시면 양식과 거의 동일한 PDF가 생성됩니다.")
            with col_btn2:
                if st.button("현재 테스트 중인 맞교환 목록 전체 삭제하기", key="test_clear_btn"):
                    st.session_state.test_swaps = pd.DataFrame()
                    st.session_state["test_has_cycle"] = False
                    get_effective_timetable_for_date.clear()
                    effective_teacher_matrix.clear()
                    get_single_lesson_1to1_candidates.clear()
                    get_single_lesson_linked_cycles.clear()
                    st.success("테스트 중인 맞교환 목록이 전체 삭제되었습니다.")
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
                    if save_work_data_to_gsheet(["duties"]):
                        record_history(f"복무 ({t})")
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
                                            "teacher": o.교사명, "class": other_class, "subject": o.과목, "lesson": f"{other_class} {o.과목}",
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
                                                  "학급": c.get("class", ""), "과목": c.get("subject", "")}
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

st.caption(f"서라벌여중 시간표 관리 시스템20260915v2.1.0.0 · {current_name()} ({current_user()}) · {current_role()}")
