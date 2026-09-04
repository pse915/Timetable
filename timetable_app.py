# -*- coding: utf-8 -*-
"""
서라벌여중 시간표·결보강 관리 프로그램
2026 최종 완전판 (초고속 + 결강/보강/이동 원본 표기 완료)
"""

import io
from datetime import date, datetime, timedelta
import numpy as np
import pandas as pd
import streamlit as st
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

SUBJECT_GROUP = {
    "국어1": "국어", "국어2": "국어", "사회": "사회", "사회1": "사회", "사회2": "사회", "사회3": "사회",
    "역사": "역사", "도덕1": "도덕", "도덕2": "도덕", "수학": "수학", "수학1": "수학", "수학2": "수학",
    "과학": "과학", "과학1": "과학", "과학2": "과학", "기가": "기술가정",
    "체육1": "체육", "체육2": "체육", "체육3": "체육", "스포": "스포츠",
    "음악": "음악", "음악1": "음악", "음악2": "음악", "미술": "미술",
    "영어": "영어", "영어1": "영어", "영어2": "영어", "영회": "영어",
    "한문": "한문", "일본어": "일본어", "정보": "정보", "진동": "진로활동",
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

# ==========================================================================================
# 데이터 로드 / 초기화
# ==========================================================================================
DUTY_COLS = ["교사명", "일자", "교시", "사유", "상세사유", "등록시각"]

def ensure_duty_columns(df):
    if df is None or not isinstance(df, pd.DataFrame):
        return pd.DataFrame(columns=DUTY_COLS)
    df = df.copy()
    for c in DUTY_COLS:
        if c not in df.columns:
            df[c] = 0 if c == "교시" else ""
    df["교시"] = df["교시"].apply(safe_int)
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
        absences = df_from_worksheet(get_worksheet(WORK_SHEET_ID, "결강"))
        subs = df_from_worksheet(get_worksheet(WORK_SHEET_ID, "보강"))
        swaps = df_from_worksheet(get_worksheet(WORK_SHEET_ID, "맞교환"))
        part_time = df_from_worksheet(get_worksheet(WORK_SHEET_ID, "시간강사"))
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
    try:
        df_to_worksheet(get_worksheet(WORK_SHEET_ID, "결강"), st.session_state.absences)
        df_to_worksheet(get_worksheet(WORK_SHEET_ID, "보강"), st.session_state.subs)
        df_to_worksheet(get_worksheet(WORK_SHEET_ID, "맞교환"), st.session_state.swaps)
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
        absences = pd.DataFrame(columns=["결강ID", "일자", "요일", "교사명", "사유", "상세사유", "교시", "학급", "과목", "등록시각"])
    if subs.empty:
        subs = pd.DataFrame(columns=["결강ID", "일자", "요일", "교시", "학급", "과목", "결강교사", "보강교사", "배정방식", "우선순위", "비고", "등록시각"])
    if swaps.empty:
        swaps = pd.DataFrame(columns=["원본일자", "교사A", "요일A", "교시A", "학급A", "과목A",
                                      "목표일자", "교사B", "요일B", "교시B", "학급B", "과목B", "유형", "시간강사구인", "등록시각"])
    if part_time.empty:
        part_time = pd.DataFrame(columns=["번호", "시간강사명", "담당과목", "과목군", "비고"] + [f"{d}{p}" for d in DAYS for p in range(1, 8)])

    st.session_state.absences = absences
    st.session_state.subs = subs
    st.session_state.swaps = swaps
    st.session_state.part_time = part_time
    st.session_state.cumulative = cumulative
    st.session_state.duties = ensure_duty_columns(duties)
    st.session_state._data_version = 0
    st.session_state.history = []
    st.session_state.history_index = -1
    push_history("초기 상태")

init_state()

# ==========================================================================================
# 핵심 로직
# ==========================================================================================
@st.cache_data(show_spinner=False, ttl=120)
def get_effective_timetable_for_date(on_date: str, version: int = 0) -> pd.DataFrame:
    norm = normalize_date_str(on_date)
    if not norm:
        return st.session_state.timetable.copy()

    try:
        day = WEEKDAY_KR[datetime.strptime(norm, "%Y-%m-%d").weekday()]
    except Exception:
        return st.session_state.timetable.copy()

    tt = st.session_state.timetable
    if tt.empty:
        return pd.DataFrame(columns=["교사명", "요일", "교시", "과목", "학급", "과목군"])

    base = tt[tt["요일"] == day]
    current = {}
    for r in base.itertuples(index=False):
        p = safe_int(r.교시)
        t = str(r.교사명).strip()
        current[(t, p)] = {
            "교사명": t, "요일": day, "교시": p,
            "과목": str(r.과목).strip(), "학급": str(r.학급).strip(),
            "과목군": str(getattr(r, "과목군", subject_group(r.과목))).strip()
        }

    # 맞교환 반영
    swaps = st.session_state.swaps
    if not swaps.empty:
        mask = (swaps["원본일자"] == norm) | (swaps["목표일자"] == norm)
        for sw in swaps[mask].itertuples(index=False):
            t_a, p_a = str(sw.교사A).strip(), safe_int(sw.교시A)
            t_b, p_b = str(sw.교사B).strip(), safe_int(sw.교시B)
            typ = str(getattr(sw, "유형", "")).strip()
            s_a, c_a = str(getattr(sw, "과목A", "")).strip(), str(getattr(sw, "학급A", "")).strip()
            s_b, c_b = str(getattr(sw, "과목B", "")).strip(), str(getattr(sw, "학급B", "")).strip()

            if sw.원본일자 == norm:
                current.pop((t_a, p_a), None)
                if typ in ["1:1 맞교환", "1:1맞교환", "직접1:1"] and t_b:
                    current[(t_b, p_a)] = {"교사명": t_b, "요일": day, "교시": p_a,
                                           "과목": s_b or s_a, "학급": c_b or c_a, "과목군": subject_group(s_b or s_a)}
            if sw.목표일자 == norm:
                if typ in ["1:1 맞교환", "1:1맞교환", "직접1:1"]:
                    current.pop((t_b, p_b), None)
                    if t_a and p_b:
                        current[(t_a, p_b)] = {"교사명": t_a, "요일": day, "교시": p_b,
                                               "과목": s_a, "학급": c_a, "과목군": subject_group(s_a)}
                elif "연계" in typ and t_a and p_b:
                    current[(t_a, p_b)] = {"교사명": t_a, "요일": day, "교시": p_b,
                                           "과목": s_a, "학급": c_a, "과목군": subject_group(s_a)}

    # 보강 반영
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
                    "과목군": subject_group(str(r.과목))
                }

    df = pd.DataFrame(list(current.values()))
    if df.empty:
        df = pd.DataFrame(columns=["교사명", "요일", "교시", "과목", "학급", "과목군"])
    return df

def get_swap_origin_info(teacher: str, on_date: str, period: int) -> str:
    """이 수업이 어디서 이동해 왔는지 원본 정보 반환"""
    norm_date = normalize_date_str(on_date)
    if not norm_date:
        return ""
    swaps = st.session_state.get("swaps", pd.DataFrame())
    if swaps.empty:
        return ""

    p = safe_int(period)

    # 교사A가 목표 날짜로 이동해 온 경우
    mask1 = (swaps["목표일자"] == norm_date) & (swaps["교사A"] == teacher) & (swaps["교시B"] == p)
    if mask1.any():
        row = swaps[mask1].iloc[0]
        orig_day = str(row.get("요일A", "")).strip()
        orig_p = safe_int(row.get("교시A", 0))
        partner = str(row.get("교사B", "")).strip()
        return f"{orig_day}{orig_p}({partner})" if partner else f"{orig_day}{orig_p}"

    # 상대방이 원본에서 이 자리로 온 경우
    mask2 = (swaps["원본일자"] == norm_date) & (swaps["교사B"] == teacher) & (swaps["교시A"] == p)
    if mask2.any():
        row = swaps[mask2].iloc[0]
        orig_day = str(row.get("요일B", "")).strip()
        orig_p = safe_int(row.get("교시B", 0))
        partner = str(row.get("교사A", "")).strip()
        return f"{orig_day}{orig_p}({partner})" if partner else f"{orig_day}{orig_p}"

    # 연계 교환
    mask3 = (swaps["목표일자"] == norm_date) & (swaps["교사A"] == teacher) & (swaps["교시B"] == p)
    if mask3.any():
        row = swaps[mask3].iloc[0]
        if "연계" in str(row.get("유형", "")):
            orig_day = str(row.get("요일A", "")).strip()
            orig_p = safe_int(row.get("교시A", 0))
            return f"{orig_day}{orig_p}"
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

def is_free(teacher: str, day: str, period: int, on_date: str = None, eff_tt=None) -> bool:
    p = safe_int(period)
    norm = normalize_date_str(on_date)
    if has_duty(teacher, norm, p):
        return False
    if eff_tt is None:
        eff_tt = get_effective_timetable_for_date(norm, st.session_state.get("_data_version", 0))
    if not eff_tt.empty and not eff_tt[(eff_tt["교사명"] == teacher) & (eff_tt["교시"] == p)].empty:
        return False
    subs = st.session_state.subs
    if not subs.empty and norm:
        if ((subs["일자"] == norm) & (subs["교시"] == p) & (subs["보강교사"] == teacher)).any():
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

def recommend_substitutes(day, period, subject, class_name, absent_teacher, on_date, top_n=10, include_part_time=False, eff_tt=None):
    teachers = st.session_state.teachers
    if teachers.empty:
        return pd.DataFrame()

    norm = normalize_date_str(on_date)
    if eff_tt is None:
        eff_tt = get_effective_timetable_for_date(norm, st.session_state.get("_data_version", 0))

    grp = subject_group(subject)
    grade = grade_of(class_name)
    cum = cumulative_sub_count(version=st.session_state.get("_data_version", 0))
    load = weekly_load(version=st.session_state.get("_data_version", 0))
    max_cum = max(cum.values()) if cum else 0

    free = []
    for t in teachers["교사명"].tolist():
        if t == absent_teacher or has_duty(t, norm):
            continue
        if is_free(t, day, period, norm, eff_tt=eff_tt):
            free.append(t)

    rows = []
    for t in free:
        my = eff_tt[eff_tt["교사명"] == t] if not eff_tt.empty else pd.DataFrame()
        my_groups = set(my["과목군"]) if not my.empty else set()
        my_grades = {grade_of(c) for c in my["학급"]} if not my.empty else set()

        if grp in my_groups and grade in my_grades:
            prio, label, score = 1, "1순위 · 동일 과목 & 동일 학년", 120
        elif grp in my_groups:
            prio, label, score = 2, "2순위 · 동일 과목", 90
        elif grade in my_grades:
            prio, label, score = 3, "3순위 · 동일 학년", 60
        else:
            prio, label, score = 4, "4순위 · 전체 공강", 20

        score += (max_cum - cum.get(t, 0)) * 6 + max(0, 22 - load.get(t, 0)) * 0.8
        t_row = teachers[teachers["교사명"] == t]
        rows.append({
            "보강교사": t, "유형": "정규교사", "우선순위": label, "_prio": prio,
            "담당과목": t_row["담당과목"].iloc[0] if not t_row.empty else "",
            "주당시수": load.get(t, 0), "누적보강": cum.get(t, 0), "추천점수": round(score, 1)
        })

    if include_part_time and not st.session_state.part_time.empty:
        for t in st.session_state.part_time["시간강사명"].dropna().unique():
            if is_free(t, day, period, norm):
                rows.append({"보강교사": t, "유형": "시간강사", "우선순위": "시간강사", "_prio": 5,
                             "담당과목": "", "주당시수": 0, "누적보강": 0, "추천점수": 40})

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
        "비고": memo, "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M")
    }])
    st.session_state.subs = pd.concat([s, new], ignore_index=True)
    save_work_data_to_gsheet()

def cancel_substitute(cid, period):
    push_history(f"보강 취소 ({period}교시)")
    s = st.session_state.subs
    p = safe_int(period)
    st.session_state.subs = s[~((s["결강ID"] == cid) & (s["교시"] == p))].reset_index(drop=True)
    save_work_data_to_gsheet()

def do_swap(a, b, date_a, date_b, is_part_time_purpose=False):
    push_history(f"맞교환 ({a['교사명']} ↔ {b['교사명']})")
    rec = {
        "원본일자": normalize_date_str(date_a), "교사A": a["교사명"], "요일A": a["요일"], "교시A": safe_int(a["교시"]),
        "학급A": a["학급"], "과목A": a["과목"],
        "목표일자": normalize_date_str(date_b), "교사B": b["교사명"], "요일B": b["요일"], "교시B": safe_int(b["교시"]),
        "학급B": b["학급"], "과목B": b["과목"],
        "유형": "1:1 맞교환", "시간강사구인": "Y" if is_part_time_purpose else "N",
        "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    st.session_state.swaps = pd.concat([st.session_state.swaps, pd.DataFrame([rec])], ignore_index=True)
    save_work_data_to_gsheet()
    return True

def do_linked_swap(a, teacher_b, date_a, date_b, day_b, period_b, is_part_time_purpose=False):
    push_history(f"연계교환 ({a['교사명']} → {teacher_b})")
    rec = {
        "원본일자": normalize_date_str(date_a), "교사A": a["교사명"], "요일A": a["요일"], "교시A": safe_int(a["교시"]),
        "학급A": a["학급"], "과목A": a["과목"],
        "목표일자": normalize_date_str(date_b), "교사B": teacher_b, "요일B": day_b, "교시B": safe_int(period_b),
        "학급B": a["학급"], "과목B": a["과목"],
        "유형": "연계 공강 교환", "시간강사구인": "Y" if is_part_time_purpose else "N",
        "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    st.session_state.swaps = pd.concat([st.session_state.swaps, pd.DataFrame([rec])], ignore_index=True)
    save_work_data_to_gsheet()
    return True

def get_target_time_recommendations(teacher_a, date_a_str, period_a, class_a, subject_a, date_b_str, period_b, eff_tt_a=None, eff_tt_b=None):
    ti = st.session_state.teachers
    norm_a = normalize_date_str(date_a_str)
    norm_b = normalize_date_str(date_b_str)
    if eff_tt_a is None:
        eff_tt_a = get_effective_timetable_for_date(norm_a, st.session_state.get("_data_version", 0))
    if eff_tt_b is None:
        eff_tt_b = get_effective_timetable_for_date(norm_b, st.session_state.get("_data_version", 0))

    if ti.empty:
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

    p_a = safe_int(period_a)
    p_b = safe_int(period_b)
    day_a = WEEKDAY_KR[datetime.strptime(norm_a, "%Y-%m-%d").weekday()]
    day_b = WEEKDAY_KR[datetime.strptime(norm_b, "%Y-%m-%d").weekday()]

    grp_a = subject_group(subject_a)
    grade_a = grade_of(class_a)
    cum = cumulative_sub_count(version=st.session_state.get("_data_version", 0))
    load = weekly_load(version=st.session_state.get("_data_version", 0))

    swap_recs, linked_recs = [], []

    for t_b in ti["교사명"].tolist():
        if t_b == teacher_a or has_duty(t_b, norm_b):
            continue
        b_lessons = eff_tt_b[(eff_tt_b["교사명"] == t_b) & (eff_tt_b["교시"] == p_b)] if not eff_tt_b.empty else pd.DataFrame()
        if not b_lessons.empty:
            for _, b_row in b_lessons.iterrows():
                b_item = {"교사명": t_b, "일자": norm_b, "요일": day_b, "교시": p_b, "학급": b_row["학급"], "과목": b_row["과목"]}
                is_valid = is_free(teacher_a, day_b, p_b, norm_b, eff_tt_b) and is_free(t_b, day_a, p_a, norm_a, eff_tt_a)
                score = 2000 if is_valid else -1000
                if is_valid:
                    if subject_group(b_row["과목"]) == grp_a: score += 300
                    if grade_of(b_row["학급"]) == grade_a: score += 200
                    score -= cum.get(t_b, 0) * 15
                swap_recs.append({
                    "유형": "직접1:1", "교사B": t_b, "현재 수업": f"{day_b}{p_b}교시 · {b_row['학급']} · {b_row['과목']}",
                    "상태": "✅ 가능" if is_valid else "❌ 충돌", "is_valid": is_valid, "_score": score,
                    "b_info": b_item
                })
        if is_free(t_b, day_b, p_b, norm_b, eff_tt_b) and is_free(teacher_a, day_b, p_b, norm_b, eff_tt_b):
            score = 1500 - cum.get(t_b, 0) * 12
            linked_recs.append({
                "유형": "연계공강교환", "교사B": t_b, "현재 수업": f"{day_b}{p_b}교시 공강",
                "상태": "✅ 가능", "is_valid": True, "_score": score,
                "b_info": {"교사명": t_b, "일자": norm_b, "요일": day_b, "교시": p_b, "학급": class_a, "과목": subject_a}
            })

    df_swap = pd.DataFrame(swap_recs).sort_values("_score", ascending=False).reset_index(drop=True) if swap_recs else pd.DataFrame()
    df_linked = pd.DataFrame(linked_recs).sort_values("_score", ascending=False).reset_index(drop=True) if linked_recs else pd.DataFrame()
    return df_swap, df_linked, pd.DataFrame()

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

def get_teacher_week_view(teacher: str, ref_date: date):
    """결강 / 보강 / 이동 원본이 모두 표시되는 주간표"""
    weekday = ref_date.weekday()
    monday = ref_date - timedelta(days=weekday)
    week_dates = [monday + timedelta(days=i) for i in range(5)]
    ver = st.session_state.get("_data_version", 0)

    absences = st.session_state.get("absences", pd.DataFrame())
    subs = st.session_state.get("subs", pd.DataFrame())

    grid = []
    for p in range(1, MAX_PERIOD + 1):
        row = {"교시": p}
        for i, d in enumerate(DAYS):
            on_date = week_dates[i].strftime("%Y-%m-%d")
            eff = get_effective_timetable_for_date(on_date, ver)

            m = eff[(eff["교사명"] == teacher) & (eff["교시"] == p)] if not eff.empty else pd.DataFrame()

            if not m.empty:
                r = m.iloc[0]
                cell = f"{r['학급']} {r['과목']}"

                # 1. 맞교환/이동 원본 표시
                origin = get_swap_origin_info(teacher, on_date, p)
                if origin:
                    cell += f" 🔄 from {origin}"

                # 2. 결강 표시
                is_absent = False
                if not absences.empty:
                    is_absent = ((absences["일자"] == on_date) &
                                 (absences["교사명"] == teacher) &
                                 (absences["교시"] == p)).any()
                if is_absent:
                    cell = f"[결강] {cell}"

                # 3. 보강 표시
                is_sub = False
                if not subs.empty:
                    is_sub = ((subs["일자"] == on_date) &
                              (subs["보강교사"] == teacher) &
                              (subs["교시"] == p)).any()
                if is_sub:
                    cell = f"[보강] {cell}"

                row[d] = cell
            else:
                row[d] = ""
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
    return sorted(t for t in changed if t)

def build_report_html(norm_date: str) -> str:
    day = WEEKDAY_KR.get(datetime.strptime(norm_date, "%Y-%m-%d").weekday(), "")
    a, s, w = st.session_state.absences, st.session_state.subs, st.session_state.swaps

    def rows_abs():
        if a.empty or "일자" not in a.columns:
            return "<tr><td colspan='6'>없음</td></tr>"
        sub = a[a["일자"] == norm_date]
        if sub.empty:
            return "<tr><td colspan='6'>없음</td></tr>"
        return "".join(f"<tr><td>{r.get('교사명','')}</td><td>{r.get('사유','')}</td><td>{safe_int(r.get('교시'))}</td>"
                       f"<td>{r.get('학급','')}</td><td>{r.get('과목','')}</td><td>{r.get('상세사유','')}</td></tr>" for _, r in sub.iterrows())

    def rows_sub():
        if s.empty or "일자" not in s.columns:
            return "<tr><td colspan='7'>없음</td></tr>"
        sub = s[s["일자"] == norm_date]
        if sub.empty:
            return "<tr><td colspan='7'>없음</td></tr>"
        return "".join(f"<tr><td>{safe_int(r.get('교시'))}</td><td>{r.get('학급','')}</td><td>{r.get('과목','')}</td>"
                       f"<td>{r.get('결강교사','')}</td><td><b>{r.get('보강교사','')}</b></td><td>{r.get('우선순위','')}</td>"
                       f"<td>{r.get('비고','')}</td></tr>" for _, r in sub.iterrows())

    def rows_swap():
        if w.empty:
            return "<tr><td colspan='4'>없음</td></tr>"
        mask = (w["원본일자"] == norm_date) | (w["목표일자"] == norm_date)
        sub = w[mask]
        if sub.empty:
            return "<tr><td colspan='4'>없음</td></tr>"
        return "".join(f"<tr><td>{r.get('교사A','')}</td><td>[{r.get('원본일자')}] {r.get('요일A')} {safe_int(r.get('교시A'))}교시 ↔ "
                       f"[{r.get('목표일자')}] {r.get('요일B')} {safe_int(r.get('교시B'))}교시</td>"
                       f"<td>{r.get('교사B','')}</td><td>{r.get('유형','')}</td></tr>" for _, r in sub.iterrows())

    return f"""<!doctype html><html lang="ko"><head><meta charset="utf-8"><title>{norm_date} 내역서</title>
<style>body{{font-family:'맑은 고딕',sans-serif;font-size:12px}} table{{width:100%;border-collapse:collapse}} th,td{{border:1px solid #999;padding:5px;text-align:center}} th{{background:#eef1f6}}</style></head><body>
<h1 style="text-align:center">결강·보강 변경 내역서</h1>
<div style="text-align:center">{SCHOOL_YEAR} · {SCHOOL_NAME} · {norm_date} ({day})</div>
<h3>1. 결강 현황</h3><table><tr><th>교사</th><th>사유</th><th>교시</th><th>학급</th><th>과목</th><th>상세</th></tr>{rows_abs()}</table>
<h3>2. 보강 배정</h3><table><tr><th>교시</th><th>학급</th><th>과목</th><th>결강교사</th><th>보강교사</th><th>근거</th><th>비고</th></tr>{rows_sub()}</table>
<h3>3. 맞교환</h3><table><tr><th>교사A</th><th>내용</th><th>교사B</th><th>유형</th></tr>{rows_swap()}</table>
</body></html>"""

def to_excel_bytes(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, df in sheets.items():
            (df if isinstance(df, pd.DataFrame) else pd.DataFrame(df)).to_excel(w, sheet_name=name[:31], index=False)
    return buf.getvalue()

# ==========================================================================================
# 사이드바
# ==========================================================================================
with st.sidebar:
    st.header("데이터")
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
        st.session_state.absences = absences
        st.session_state.subs = subs
        st.session_state.swaps = swaps
        st.session_state.part_time = part_time
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

    st.divider()
    st.subheader("📄 일일 내역서")
    rd = st.date_input("출력 일자", value=date.today(), key="sidebar_rd").strftime("%Y-%m-%d")
    if st.button("다운로드", use_container_width=True, type="primary"):
        html = build_report_html(rd)
        st.download_button("HTML", html.encode("utf-8"), f"내역서_{rd}.html", "text/html")
        xls = to_excel_bytes({
            "결강": st.session_state.absences[st.session_state.absences["일자"] == rd] if not st.session_state.absences.empty else pd.DataFrame(),
            "보강": st.session_state.subs[st.session_state.subs["일자"] == rd] if not st.session_state.subs.empty else pd.DataFrame(),
            "맞교환": st.session_state.swaps
        })
        st.download_button("엑셀", xls, f"내역서_{rd}.xlsx")

if st.session_state.timetable.empty:
    st.info("시간표를 불러오는 중...")
    st.stop()

# ==========================================================================================
# 메인 화면
# ==========================================================================================
st.title("시간표 · 결강/보강 관리 (완전판)")

tabs = st.tabs([
    "시간표 조회", "시간강사 관리", "결강·보강", "시간표 맞교환 & 변경 추천",
    "통계", "시간표 변경 테스트용", "변경된 교사 주간표", "📋 복무 관리 & 판단"
])

# ------------------------------------------------------------------ 0. 시간표 조회
with tabs[0]:
    view = st.radio("보기 방식", ["교사별 전체 매트릭스", "학급별 매트릭스", "교사 1인 주간표"], horizontal=True)
    ver = st.session_state.get("_data_version", 0)
    if view == "교사별 전체 매트릭스":
        st.dataframe(teacher_matrix(ver), use_container_width=True, height=700)
    elif view == "학급별 매트릭스":
        st.dataframe(class_matrix(ver), use_container_width=True, height=700)
    else:
        tlist = st.session_state.teachers["교사명"].tolist()
        t = st.selectbox("교사 선택", tlist, key="view_t")
        ref = st.date_input("기준 날짜", value=date.today(), key="view_ref")
        grid, dates = get_teacher_week_view(t, ref)
        st.caption(f"{dates[0]} ~ {dates[4]}  |  [결강] [보강] 🔄 from 원본 표시")
        st.dataframe(grid, use_container_width=True)

# ------------------------------------------------------------------ 1. 시간강사 관리
with tabs[1]:
    st.subheader("시간강사 등록 및 가능 시간표")
    ed = st.data_editor(st.session_state.part_time, num_rows="dynamic", use_container_width=True, height=400, key="pt_ed")
    if st.button("시간강사 정보 저장", type="primary"):
        st.session_state.part_time = ed
        push_history("시간강사 수정")
        save_work_data_to_gsheet()
        st.success("저장 완료")
        st.rerun()

# ------------------------------------------------------------------ 2. 결강·보강
with tabs[2]:
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
        eff = get_effective_timetable_for_date(on_date, ver)
        todays = eff[(eff["교사명"] == who) & (eff["요일"] == day)].sort_values("교시")
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
                    "학급": r.학급, "과목": r.과목, "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M")
                } for r in rows.itertuples()]
                st.session_state.absences = pd.concat([a, pd.DataFrame(new_rows)], ignore_index=True)
                save_work_data_to_gsheet()
                st.success("결강 등록 완료")
                st.rerun()
        st.dataframe(st.session_state.absences[st.session_state.absences["일자"] == on_date] if not st.session_state.absences.empty else pd.DataFrame(), height=250)

    with right:
        st.markdown("### 📌 보강 배정")
        ab = st.session_state.absences
        if ab.empty:
            st.info("먼저 결강을 등록하세요.")
        else:
            cids = sorted(ab["결강ID"].dropna().unique(), reverse=True)
            cid = st.selectbox("결강 건 선택", cids, key="sub_cid")
            rows = ab[ab["결강ID"] == cid].sort_values("교시")
            head = rows.iloc[0]
            st.markdown(f"**{head['일자']} · {head['교사명']} · {len(rows)}시간**")
            include_pt = st.checkbox("시간강사 포함", key="sub_pt")
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
                        cand = recommend_substitutes(head["요일"], p, r.과목, r.학급, head["교사명"], head["일자"], include_part_time=include_pt)
                        if cand.empty:
                            st.warning("추천 가능한 교사가 없습니다.")
                        else:
                            st.dataframe(cand, hide_index=True, height=160)
                            pick = st.selectbox("보강 교사", cand["보강교사"], key=f"pick_{cid}_{p}")
                            if st.button("배정", key=f"btn_{cid}_{p}"):
                                pr = cand.loc[cand["보강교사"] == pick, "우선순위"].iloc[0]
                                add_substitute(cid, head["일자"], head["요일"], p, r.학급, r.과목, head["교사명"], pick, "수동", pr, "")
                                st.rerun()

# ------------------------------------------------------------------ 3. 시간표 맞교환
with tabs[3]:
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
        eff_a = get_effective_timetable_for_date(date_a_str, ver)
        sub_a = eff_a[(eff_a["교사명"] == t_a) & (eff_a["요일"] == day_a)].sort_values("교시")
        pick_a = None
        if not sub_a.empty:
            opts = [f"{safe_int(r.교시)}교시 · {r.학급} · {r.과목}" for r in sub_a.itertuples()]
            sel = st.selectbox("변경할 수업", opts, key="sw_la")
            row = sub_a.iloc[opts.index(sel)]
            pick_a = {"교사명": t_a, "일자": date_a_str, "요일": day_a, "교시": safe_int(row.교시), "학급": row.학급, "과목": row.과목}

    with col_b:
        st.markdown("#### 2️⃣ 이동 희망 시간")
        date_b = st.date_input("이동 희망 날짜", value=date.today(), key="sw_db")
        date_b_str = date_b.strftime("%Y-%m-%d")
        day_b = WEEKDAY_KR[date_b.weekday()]
        p_b = st.selectbox("희망 교시", list(range(1, PERIODS_PER_DAY.get(day_b, 7)+1)), key="sw_pb")

    is_pt = st.checkbox("시간강사 구인 목적")

    if pick_a:
        st.subheader(f"{pick_a['교시']}교시 → {date_b_str} {p_b}교시")
        df_swap, df_linked, _ = get_target_time_recommendations(
            pick_a["교사명"], pick_a["일자"], pick_a["교시"], pick_a["학급"], pick_a["과목"],
            date_b_str, p_b
        )
        t1, t2 = st.tabs(["1:1 맞교환", "연계 교환"])
        with t1:
            if df_swap.empty:
                st.info("추천 없음")
            else:
                for idx, row in df_swap.iterrows():
                    with st.expander(f"{row['상태']} {row['교사B']} ({row['현재 수업']})", expanded=(idx==0)):
                        if row["is_valid"] and st.button(f"{row['교사B']}와 맞교환", key=f"sw_{idx}"):
                            do_swap(pick_a, row["b_info"], date_a_str, date_b_str, is_pt)
                            st.rerun()
        with t2:
            if df_linked.empty:
                st.info("추천 없음")
            else:
                for idx, row in df_linked.iterrows():
                    with st.expander(f"{row['상태']} {row['교사B']}", expanded=(idx==0)):
                        if st.button(f"{row['교사B']}와 연계 교환", key=f"lk_{idx}"):
                            do_linked_swap(pick_a, row["교사B"], date_a_str, date_b_str, day_b, p_b, is_pt)
                            st.rerun()

    st.dataframe(st.session_state.swaps, use_container_width=True)

# ------------------------------------------------------------------ 4. 통계
with tabs[4]:
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
    st.dataframe(df, use_container_width=True)
    if not df.empty:
        st.bar_chart(df.set_index("교사명")["누적보강"])

# ------------------------------------------------------------------ 5. 테스트용
with tabs[5]:
    st.subheader("🧪 시간표 변경 테스트용 (저장 안 됨)")
    tlist = st.session_state.teachers["교사명"].tolist()
    t = st.selectbox("테스트 교사", tlist, key="test_t")
    ref = st.date_input("기준일", value=date.today(), key="test_d")
    grid, dates = get_teacher_week_view(t, ref)
    st.caption(f"{dates[0]} ~ {dates[4]}")
    st.dataframe(grid, use_container_width=True, height=400)

# ------------------------------------------------------------------ 6. 변경된 교사 주간표
with tabs[6]:
    st.subheader("📅 변경된 교사 주간 시간표")
    ref = st.date_input("기준 날짜", value=date.today(), key="chg_ref")
    changed = get_changed_teachers_for_week(ref)
    if not changed:
        st.success("이번 주차 변경 교사 없음")
    else:
        st.info(f"변경 교사 {len(changed)}명: {', '.join(changed)}")
        for t in changed:
            with st.expander(f"👤 {t}", expanded=True):
                grid, _ = get_teacher_week_view(t, ref)
                st.dataframe(grid, use_container_width=True)

# ------------------------------------------------------------------ 7. 복무 관리 & 판단
with tabs[7]:
    st.subheader("📋 복무 관리 & 판단")
    left, right = st.columns([1, 1.4])

    with left:
        st.markdown("### 복무 등록")
        t = st.selectbox("교사", st.session_state.teachers["교사명"].tolist(), key="duty_t")
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
                    news.append({"교사명": t, "일자": d.strftime("%Y-%m-%d"), "교시": 0, "사유": reason, "상세사유": detail, "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M")})
                else:
                    for p in periods:
                        news.append({"교사명": t, "일자": d.strftime("%Y-%m-%d"), "교시": p, "사유": reason, "상세사유": detail, "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M")})
                st.session_state.duties = pd.concat([duties, pd.DataFrame(news)], ignore_index=True)
                save_work_data_to_gsheet()
                st.success("등록 완료")
                st.rerun()

        st.markdown("#### 등록된 복무")
        duties = ensure_duty_columns(st.session_state.duties)
        if not duties.empty:
            show = duties.copy()
            show["교시표시"] = show["교시"].apply(lambda x: "전체" if safe_int(x) == 0 else f"{safe_int(x)}교시")
            st.dataframe(show[["교사명", "일자", "교시표시", "사유"]], height=220)
            idx = st.number_input("선택 행 번호", 0, max(0, len(show)-1), 0, key="duty_idx")
            if st.button("이 복무 선택", key="duty_sel"):
                st.session_state["_sel_duty"] = show.iloc[int(idx)].to_dict()
                st.rerun()
            if st.button("선택 행 삭제", key="duty_del"):
                st.session_state.duties = show.drop(show.index[int(idx)]).reset_index(drop=True)
                save_work_data_to_gsheet()
                st.rerun()
        else:
            st.info("등록된 복무 없음")

    with right:
        st.markdown("### 선택 교사 당일 처리")
        sel = st.session_state.get("_sel_duty")
        if not sel:
            st.info("왼쪽에서 복무를 선택하세요.")
        else:
            t_name = sel["교사명"]
            d_str = sel["일자"]
            duty_p = safe_int(sel.get("교시", 0))
            day_kr = WEEKDAY_KR[datetime.strptime(d_str, "%Y-%m-%d").weekday()]
            st.markdown(f"**{t_name}** · {d_str} ({day_kr}) · {'전체' if duty_p == 0 else f'{duty_p}교시'}")

            ver = st.session_state.get("_data_version", 0)
            eff = get_effective_timetable_for_date(d_str, ver)
            lessons = eff[eff["교사명"] == t_name].sort_values("교시")

            for r in lessons.itertuples():
                p = safe_int(r.교시)
                is_target = (duty_p == 0) or (duty_p == p)
                with st.expander(f"{p}교시 · {r.학급} · {r.과목}", expanded=is_target):
                    if not is_target:
                        st.caption("복무 대상 아님")
                        continue

                    # 맞교환 한눈에
                    possible = []
                    for op in range(1, PERIODS_PER_DAY.get(day_kr, 7) + 1):
                        if op != p and is_free(t_name, day_kr, op, d_str, eff):
                            others = eff[(eff["교시"] == op) & (eff["교사명"] != t_name)]
                            for o in others.itertuples():
                                possible.append({"교사": o.교사명, "수업": f"{o.학급} {o.과목}", "교시": op})
                    if possible:
                        st.markdown("**🔄 맞교환 가능 목록**")
                        for i, sw in enumerate(possible[:4]):
                            st.write(f"• {sw['교사']} ({sw['교시']}교시 {sw['수업']})")
                    else:
                        st.info("같은 날 맞교환 가능 수업 없음")

                    if st.button("결강으로 등록", key=f"dabs_{p}"):
                        push_history(f"복무→결강 ({t_name} {p})")
                        cid = f"{d_str}-{t_name}"
                        a = st.session_state.absences
                        if not a.empty:
                            a = a[~((a["일자"] == d_str) & (a["교사명"] == t_name) & (a["교시"] == p))]
                        new = {
                            "결강ID": cid, "일자": d_str, "요일": day_kr, "교사명": t_name,
                            "사유": sel.get("사유", ""), "상세사유": sel.get("상세사유", ""),
                            "교시": p, "학급": r.학급, "과목": r.과목,
                            "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M")
                        }
                        st.session_state.absences = pd.concat([a, pd.DataFrame([new])], ignore_index=True)
                        save_work_data_to_gsheet()
                        st.success("결강 등록 완료")
                        st.rerun()

st.caption("완전판 · 결강/보강/이동 원본 표기 완료 · 초고속 최적화")
