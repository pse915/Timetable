# -*- coding: utf-8 -*-
"""
==========================================================================================
 중학교 전체 시간표 관리 및 결강·보강 자동 처리 프로그램 (Streamlit Web App)
 2026 서라벌여자중학교용 – 최종 성능·가시성 강화 버전
 (Undo/Redo + 시간강사 + 누적보강 + 작업내역 + 배정취소 + 복무기반 자동 판단 탭)
 + 결강·보강 통합 탭 / 기본정보 탭 제거 / 복무 시트 연동 / 내역서 사이드바 이동
==========================================================================================
"""

import io
from datetime import date, datetime, timedelta
from copy import deepcopy

import numpy as np
import pandas as pd
import streamlit as st
import gspread
from gspread.exceptions import APIError, WorksheetNotFound
from google.oauth2.service_account import Credentials

# ==========================================================================================
# 0. 기본 설정 / 상수 / 유틸리티 & Custom CSS
# ==========================================================================================
st.set_page_config(page_title="시간표·결보강 관리", page_icon="📘", layout="wide")

st.markdown("""
<style>
    [data-testid="stDataFrame"] {
        border: 1px solid #94a3b8 !important;
        border-radius: 6px;
    }
    [data-testid="stDataFrame"] [role="gridcell"], 
    [data-testid="stDataFrame"] [role="columnheader"] {
        border-right: 1px solid #cbd5e1 !important;
        border-bottom: 1px solid #cbd5e1 !important;
    }
    div[role="grid"] div[role="row"] > div {
        border-right: 1px solid #cbd5e1 !important;
        border-bottom: 1px solid #cbd5e1 !important;
    }
    .stTable table, div[data-testid="stTable"] table {
        border-collapse: collapse !important;
        width: 100% !important;
        border: 1px solid #64748b !important;
    }
    .stTable th, div[data-testid="stTable"] th {
        background-color: #f1f5f9 !important;
        color: #0f172a !important;
        font-weight: 700 !important;
        border: 1px solid #64748b !important;
        text-align: center !important;
    }
    .stTable td, div[data-testid="stTable"] td {
        border: 1px solid #cbd5e1 !important;
        padding: 6px 10px !important;
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

MAX_HISTORY = 8
ABSENCE_REASONS = ["병가", "연가", "출장", "공가", "조퇴", "외출", "연수", "특별휴가", "기타"]

SUBJECT_GROUP = {
    "국어1": "국어", "국어2": "국어",
    "사회": "사회", "사회1": "사회", "사회2": "사회", "사회3": "사회",
    "역사": "역사",
    "도덕1": "도덕", "도덕2": "도덕",
    "수학": "수학", "수학1": "수학", "수학2": "수학",
    "과학": "과학", "과학1": "과학", "과학2": "과학",
    "기가": "기술가정",
    "체육1": "체육", "체육2": "체육", "체육3": "체육", "스포": "스포츠",
    "음악": "음악", "음악1": "음악", "음악2": "음악",
    "미술": "미술",
    "영어": "영어", "영어1": "영어", "영어2": "영어", "영회": "영어",
    "한문": "한문", "일본어": "일본어", "정보": "정보", "진동": "진로활동",
}

def safe_int(val, default=0) -> int:
    if pd.isna(val) or val is None:
        return default
    try:
        return int(float(str(val).strip()))
    except (ValueError, TypeError):
        return default

def normalize_date_str(d_str) -> str:
    if pd.isna(d_str) or not d_str or str(d_str).strip() in ("", "nan", "None"):
        return ""
    try:
        dt = pd.to_datetime(str(d_str).strip(), errors="coerce")
        if pd.isna(dt):
            return str(d_str).strip()
        return dt.strftime("%Y-%m-%d")
    except Exception:
        return str(d_str).strip()

def subject_group(subject: str) -> str:
    if not isinstance(subject, str) or not subject.strip():
        return ""
    s = str(subject).strip()
    return SUBJECT_GROUP.get(s, s.rstrip("0123456789"))

def grade_of(class_name: str) -> str:
    if isinstance(class_name, str) and "-" in class_name:
        return class_name.split("-")[0]
    return ""

# ==========================================================================================
# Google Sheets 연결
# ==========================================================================================
@st.cache_resource(show_spinner=False)
def get_gspread_client():
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive"
    ]
    if "gcp_service_account" not in st.secrets:
        st.error("GCP Secrets 인증 오류: `.streamlit/secrets.toml` 파일에 'gcp_service_account'가 설정되지 않았습니다.")
        st.stop()
    try:
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], scopes=scopes
        )
        return gspread.authorize(creds)
    except Exception as e:
        st.error(f"Google 계정 인증 실패: {e}")
        st.stop()

@st.cache_resource(show_spinner=False)
def get_spreadsheet(spreadsheet_id: str):
    client = get_gspread_client()
    return client.open_by_key(spreadsheet_id)

def get_worksheet(spreadsheet_id: str, sheet_name: str):
    try:
        sh = get_spreadsheet(spreadsheet_id)
        return sh.worksheet(sheet_name)
    except WorksheetNotFound:
        try:
            sh = get_spreadsheet(spreadsheet_id)
            return sh.add_worksheet(title=sheet_name, rows=2000, cols=40)
        except Exception as e:
            st.error(f"워크시트 생성 실패 [{sheet_name}]: {e}")
            return None
    except APIError as e:
        st.error(f"Google API 호출 할당량 초과 또는 네트워크 오류 ({sheet_name}): {e}")
        return None
    except Exception as e:
        st.error(f"시트 연동 에러 [{sheet_name}]: {e}")
        return None

def df_from_worksheet(ws) -> pd.DataFrame:
    if ws is None:
        return pd.DataFrame()
    try:
        data = ws.get_all_values()
        if not data or len(data) < 2:
            return pd.DataFrame()
        headers = [str(h).strip() for h in data[0]]
        cleaned_rows = []
        for row in data[1:]:
            row = list(row) + [""] * max(0, len(headers) - len(row))
            cleaned_rows.append(["" if c is None else str(c).strip() for c in row[:len(headers)]])
        df = pd.DataFrame(cleaned_rows, columns=headers)
        return df.replace({"nan": "", "None": "", "NaN": "", "<NA>": ""})
    except Exception as e:
        st.warning(f"데이터 변환 처리 중 오류 발생: {e}")
        return pd.DataFrame()

def df_to_worksheet(ws, df: pd.DataFrame):
    if ws is None:
        return
    try:
        ws.clear()
        if df is None or df.empty:
            return
        df_clean = df.fillna("").astype(str)
        values = [df_clean.columns.tolist()] + df_clean.values.tolist()
        ws.update("A1", values, value_input_option="USER_ENTERED")
    except APIError as e:
        st.error(f"Google Sheets API 할당량 초과 또는 쓰기 에러: {e}")
    except Exception as e:
        st.error(f"시트 저장 처리 실패: {e}")

# ==========================================================================================
# 1. 히스토리 (경량화)
# ==========================================================================================
def _snapshot_df(df):
    if df is None or not isinstance(df, pd.DataFrame) or df.empty:
        return pd.DataFrame()
    return df.copy(deep=False)

def push_history(action_name: str = "작업"):
    if "history" not in st.session_state:
        st.session_state.history = []
        st.session_state.history_index = -1

    st.session_state.history = st.session_state.history[:st.session_state.history_index + 1]

    snapshot = {
        "action": action_name,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "absences": _snapshot_df(st.session_state.get("absences")),
        "subs": _snapshot_df(st.session_state.get("subs")),
        "swaps": _snapshot_df(st.session_state.get("swaps")),
        "part_time": _snapshot_df(st.session_state.get("part_time")),
        "duties": _snapshot_df(st.session_state.get("duties")),
    }
    st.session_state.history.append(snapshot)

    if len(st.session_state.history) > MAX_HISTORY:
        st.session_state.history.pop(0)
    else:
        st.session_state.history_index += 1

def undo():
    if "history" not in st.session_state or st.session_state.history_index <= 0:
        return False
    st.session_state.history_index -= 1
    snap = st.session_state.history[st.session_state.history_index]
    st.session_state.absences = snap["absences"].copy(deep=False)
    st.session_state.subs = snap["subs"].copy(deep=False)
    st.session_state.swaps = snap["swaps"].copy(deep=False)
    st.session_state.part_time = snap.get("part_time", st.session_state.part_time).copy(deep=False)
    st.session_state.duties = snap.get("duties", st.session_state.get("duties", pd.DataFrame())).copy(deep=False)
    _invalidate_all_caches()
    return True

def redo():
    if "history" not in st.session_state:
        return False
    if st.session_state.history_index >= len(st.session_state.history) - 1:
        return False
    st.session_state.history_index += 1
    snap = st.session_state.history[st.session_state.history_index]
    st.session_state.absences = snap["absences"].copy(deep=False)
    st.session_state.subs = snap["subs"].copy(deep=False)
    st.session_state.swaps = snap["swaps"].copy(deep=False)
    st.session_state.part_time = snap.get("part_time", st.session_state.part_time).copy(deep=False)
    st.session_state.duties = snap.get("duties", st.session_state.get("duties", pd.DataFrame())).copy(deep=False)
    _invalidate_all_caches()
    return True

def init_history_if_needed():
    if "history" not in st.session_state:
        st.session_state.history = []
        st.session_state.history_index = -1
        push_history("초기 상태")

def save_history_log_to_gsheet():
    try:
        if "history" not in st.session_state or not st.session_state.history:
            return
        logs = [{
            "시각": h.get("time", ""),
            "작업내용": h.get("action", ""),
            "현재인덱스": st.session_state.history_index
        } for h in st.session_state.history[-MAX_HISTORY:]]
        df_log = pd.DataFrame(logs)
        ws = get_worksheet(WORK_SHEET_ID, "작업내역")
        if ws:
            df_to_worksheet(ws, df_log)
    except Exception as e:
        st.warning(f"작업내역 로그 저장 실패: {e}")

def _invalidate_all_caches():
    """데이터 변경 시 모든 캐시 초기화"""
    st.session_state._eff_tt_cache = {}
    st.session_state._cum_cache = None
    st.session_state._load_cache = None
    st.session_state._duty_cache = None
    st.session_state._data_version = st.session_state.get("_data_version", 0) + 1

# ==========================================================================================
# 2. 데이터 로드 / 저장
# ==========================================================================================
REQUIRED_TT_COLS = ["교사명", "요일", "교시", "과목", "학급"]

def make_empty_frames():
    ti = pd.DataFrame(columns=["번호", "교사명", "담당과목", "과목군", "담당학년", "주당시수", "담임학급", "비고"])
    tt = pd.DataFrame(columns=["교사명", "요일", "교시", "과목", "학급", "과목군"])
    pt_cols = ["번호", "시간강사명", "담당과목", "과목군", "비고"] + [f"{d}{p}" for d in DAYS for p in range(1, 8)]
    pt = pd.DataFrame(columns=pt_cols)
    return ti, tt, pt

def normalize_frames(ti: pd.DataFrame, tt: pd.DataFrame):
    if tt.empty:
        return ti, pd.DataFrame(columns=REQUIRED_TT_COLS + ["과목군"])

    tt = tt.copy()
    for c in REQUIRED_TT_COLS:
        if c not in tt.columns:
            tt[c] = ""

    tt["교시"] = tt["교시"].apply(safe_int)
    for c in ["교사명", "요일", "과목", "학급"]:
        tt[c] = tt[c].astype(str).str.strip()

    if "과목군" not in tt.columns or tt["과목군"].isna().all() or (tt["과목군"] == "").all():
        tt["과목군"] = tt["과목"].map(subject_group)
    tt["과목군"] = tt["과목군"].astype(str).str.strip()

    tt = tt[tt["요일"].isin(DAYS)]
    tt = tt[(tt["교시"] >= 1) & (tt["교시"] <= MAX_PERIOD)]
    tt = tt.drop_duplicates(subset=["교사명", "요일", "교시"]).reset_index(drop=True)

    ti = ti.copy() if not ti.empty else pd.DataFrame()
    if "교사명" not in ti.columns:
        ti = pd.DataFrame({"교사명": sorted(tt["교사명"].unique())})

    ti["교사명"] = ti["교사명"].astype(str).str.strip()
    for c in ["번호", "담당과목", "과목군", "담당학년", "주당시수", "담임학급", "비고"]:
        if c not in ti.columns:
            ti[c] = ""
        else:
            ti[c] = ti[c].astype(str).str.strip()

    missing = sorted(set(tt["교사명"]) - set(ti["교사명"]))
    if missing:
        add = pd.DataFrame({"교사명": missing})
        for c in ti.columns:
            if c not in add.columns:
                add[c] = ""
        ti = pd.concat([ti, add[ti.columns]], ignore_index=True)

    for i, row in ti.iterrows():
        sub = tt.loc[tt["교사명"] == row["교사명"]]
        if str(row.get("담당과목", "")).strip() in ("", "nan"):
            ti.at[i, "담당과목"] = "/".join(sub["과목"].value_counts().index.tolist())
        if str(row.get("과목군", "")).strip() in ("", "nan"):
            ti.at[i, "과목군"] = "/".join(sorted(set(sub["과목군"])))
        if str(row.get("담당학년", "")).strip() in ("", "nan"):
            ti.at[i, "담당학년"] = "/".join(sorted({grade_of(c) for c in sub["학급"] if grade_of(c)}))
        if str(row.get("주당시수", "")).strip() in ("", "nan", "0"):
            ti.at[i, "주당시수"] = str(len(sub))
    return ti.reset_index(drop=True), tt

@st.cache_data(ttl=300, show_spinner="시간표 불러오는 중...")
def load_timetable_from_gsheet():
    try:
        ws_ti = get_worksheet(TIMETABLE_SHEET_ID, "교사정보")
        ws_tt = get_worksheet(TIMETABLE_SHEET_ID, "시간표")
        ti = df_from_worksheet(ws_ti)
        tt = df_from_worksheet(ws_tt)
        return normalize_frames(ti, tt)
    except Exception as e:
        st.error(f"원본 시간표 로드 실패: {e}")
        return make_empty_frames()[:2]

@st.cache_data(ttl=60, show_spinner="작업 내역 불러오는 중...")
def load_work_data_from_gsheet():
    try:
        absences = df_from_worksheet(get_worksheet(WORK_SHEET_ID, "결강"))
        subs = df_from_worksheet(get_worksheet(WORK_SHEET_ID, "보강"))
        swaps = df_from_worksheet(get_worksheet(WORK_SHEET_ID, "맞교환"))
        part_time = df_from_worksheet(get_worksheet(WORK_SHEET_ID, "시간강사"))
        cumulative = df_from_worksheet(get_worksheet(WORK_SHEET_ID, "누적보강"))
        duties = df_from_worksheet(get_worksheet(WORK_SHEET_ID, "복무"))

        if not absences.empty:
            if "교시" in absences.columns:
                absences["교시"] = absences["교시"].apply(safe_int)
            if "일자" in absences.columns:
                absences["일자"] = absences["일자"].apply(normalize_date_str)
        if not subs.empty:
            if "교시" in subs.columns:
                subs["교시"] = subs["교시"].apply(safe_int)
            if "일자" in subs.columns:
                subs["일자"] = subs["일자"].apply(normalize_date_str)
        if not swaps.empty:
            for col in ["원본일자", "목표일자"]:
                if col in swaps.columns:
                    swaps[col] = swaps[col].apply(normalize_date_str)
            for col in ["교시A", "교시B"]:
                if col in swaps.columns:
                    swaps[col] = swaps[col].apply(safe_int)
        if not duties.empty and "일자" in duties.columns:
            duties["일자"] = duties["일자"].apply(normalize_date_str)

        return absences, subs, swaps, part_time, cumulative, duties
    except Exception as e:
        st.error(f"작업 내역 로드 실패: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

def save_work_data_to_gsheet():
    try:
        df_to_worksheet(get_worksheet(WORK_SHEET_ID, "결강"), st.session_state.absences)
        df_to_worksheet(get_worksheet(WORK_SHEET_ID, "보강"), st.session_state.subs)
        df_to_worksheet(get_worksheet(WORK_SHEET_ID, "맞교환"), st.session_state.swaps)
        df_to_worksheet(get_worksheet(WORK_SHEET_ID, "시간강사"), st.session_state.part_time)
        df_to_worksheet(get_worksheet(WORK_SHEET_ID, "복무"), st.session_state.duties)
        save_history_log_to_gsheet()
        _invalidate_all_caches()
        return True
    except Exception as e:
        st.error(f"저장 실패: {e}")
        return False

def save_cumulative_to_gsheet(df: pd.DataFrame):
    try:
        df_to_worksheet(get_worksheet(WORK_SHEET_ID, "누적보강"), df)
        return True
    except Exception as e:
        st.error(f"누적보강 저장 실패: {e}")
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
                                      "목표일자", "교사B", "요일B", "교시B", "학급B", "과목B",
                                      "유형", "시간강사구인", "등록시각"])
    if part_time.empty:
        pt_cols = ["번호", "시간강사명", "담당과목", "과목군", "비고"] + [f"{d}{p}" for d in DAYS for p in range(1, 8)]
        part_time = pd.DataFrame(columns=pt_cols)
    if duties.empty:
        duties = pd.DataFrame(columns=["교사명", "일자", "사유", "상세사유", "등록시각"])

    st.session_state.absences = absences
    st.session_state.subs = subs
    st.session_state.swaps = swaps
    st.session_state.part_time = part_time
    st.session_state.cumulative = cumulative
    st.session_state.duties = duties
    st.session_state._eff_tt_cache = {}
    st.session_state._cum_cache = None
    st.session_state._load_cache = None
    st.session_state._duty_cache = None
    st.session_state._data_version = 0

    init_history_if_needed()

init_state()

# ==========================================================================================
# 3. 핵심 로직 (이동 원본 표시 + 속도 강화 + 복무 연동)
# ==========================================================================================
def get_effective_timetable_for_date(on_date: str) -> pd.DataFrame:
    """맞교환 반영 최종 시간표 (캐시 + 버전 관리)"""
    norm_on_date = normalize_date_str(on_date)
    if not norm_on_date:
        return st.session_state.timetable.copy()

    cache = st.session_state.setdefault("_eff_tt_cache", {})
    ver = st.session_state.get("_data_version", 0)
    cache_key = f"{norm_on_date}_{ver}"
    if cache_key in cache:
        return cache[cache_key].copy()

    try:
        dt = datetime.strptime(norm_on_date, "%Y-%m-%d")
        day = WEEKDAY_KR[dt.weekday()]
    except Exception:
        return st.session_state.timetable.copy()

    tt = st.session_state.timetable
    if tt.empty:
        empty = pd.DataFrame(columns=REQUIRED_TT_COLS + ["과목군"])
        cache[cache_key] = empty
        return empty.copy()

    base = tt[tt["요일"] == day]
    current_map = {}
    for r in base.itertuples(index=False):
        p = safe_int(r.교시)
        t = str(r.교사명).strip()
        current_map[(t, p)] = {
            "교사명": t, "요일": day, "교시": p,
            "과목": str(r.과목).strip(),
            "학급": str(r.학급).strip(),
            "과목군": str(getattr(r, "과목군", subject_group(r.과목))).strip()
        }

    swaps = st.session_state.get("swaps", pd.DataFrame())
    if not swaps.empty and "원본일자" in swaps.columns:
        relevant = swaps[(swaps["원본일자"] == norm_on_date) | (swaps["목표일자"] == norm_on_date)]
        for sw in relevant.itertuples(index=False):
            orig_date = getattr(sw, "원본일자", "")
            target_date = getattr(sw, "목표일자", "")
            swap_type = str(getattr(sw, "유형", "")).strip()

            t_a = str(getattr(sw, "교사A", "")).strip()
            p_a = safe_int(getattr(sw, "교시A", 0))
            c_a = str(getattr(sw, "학급A", "")).strip()
            s_a = str(getattr(sw, "과목A", "")).strip()

            t_b = str(getattr(sw, "교사B", "")).strip()
            p_b = safe_int(getattr(sw, "교시B", 0))
            c_b = str(getattr(sw, "학급B", "")).strip()
            s_b = str(getattr(sw, "과목B", "")).strip()

            if orig_date == norm_on_date:
                current_map.pop((t_a, p_a), None)
                if swap_type in ["1:1 맞교환", "1:1맞교환", "직접1:1"] and t_b:
                    current_map[(t_b, p_a)] = {
                        "교사명": t_b, "요일": day, "교시": p_a,
                        "과목": s_b if s_b else s_a,
                        "학급": c_b if c_b else c_a,
                        "과목군": subject_group(s_b if s_b else s_a)
                    }

            if target_date == norm_on_date:
                if swap_type in ["1:1 맞교환", "1:1맞교환", "직접1:1"]:
                    if t_b and p_b > 0:
                        current_map.pop((t_b, p_b), None)
                    if t_a and p_b > 0:
                        current_map[(t_a, p_b)] = {
                            "교사명": t_a, "요일": day, "교시": p_b,
                            "과목": s_a, "학급": c_a, "과목군": subject_group(s_a)
                        }
                elif swap_type in ["연계 공강 교환", "연계공강교환"]:
                    if t_a and p_b > 0:
                        current_map[(t_a, p_b)] = {
                            "교사명": t_a, "요일": day, "교시": p_b,
                            "과목": s_a, "학급": c_a, "과목군": subject_group(s_a)
                        }

    df_eff = pd.DataFrame(list(current_map.values()))
    if df_eff.empty:
        df_eff = pd.DataFrame(columns=REQUIRED_TT_COLS + ["과목군"])
    cache[cache_key] = df_eff
    return df_eff.copy()

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
        orig_day = str(row.get("요일A", "")).strip()
        orig_p = safe_int(row.get("교시A", 0))
        partner = str(row.get("교사B", "")).strip()
        return f"{orig_day}{orig_p}({partner})" if partner else f"{orig_day}{orig_p}"

    mask2 = (swaps["원본일자"] == norm_date) & (swaps["교사B"] == teacher) & (swaps["교시A"] == p)
    if mask2.any():
        row = swaps[mask2].iloc[0]
        orig_day = str(row.get("요일B", "")).strip()
        orig_p = safe_int(row.get("교시B", 0))
        partner = str(row.get("교사A", "")).strip()
        return f"{orig_day}{orig_p}({partner})" if partner else f"{orig_day}{orig_p}"

    mask3 = (swaps["목표일자"] == norm_date) & (swaps["교사A"] == teacher) & (swaps["교시B"] == p) & \
            (swaps["유형"].str.contains("연계", na=False))
    if mask3.any():
        row = swaps[mask3].iloc[0]
        orig_day = str(row.get("요일A", "")).strip()
        orig_p = safe_int(row.get("교시A", 0))
        return f"{orig_day}{orig_p}"

    return ""

def is_teacher_available(teacher: str, day: str, period: int, is_part_time: bool = False) -> bool:
    source = st.session_state.part_time if is_part_time else st.session_state.teachers
    name_col = "시간강사명" if is_part_time else "교사명"
    if source.empty or name_col not in source.columns:
        return True
    row = source[source[name_col] == teacher]
    if row.empty:
        return True
    col_candidates = [f"{day}{period}", f"{day}요일{period}", f"{day}{period}교시", f"{day}요일{period}교시"]
    for col in col_candidates:
        if col in source.columns:
            val = row[col].values[0]
            if pd.isna(val):
                continue
            str_val = str(val).strip().lower()
            if str_val in {"0", "0.0", "false", "불가", "불가능", "n", "x", "없음"}:
                return False
            try:
                num_val = pd.to_numeric(val, errors="coerce")
                if pd.notna(num_val) and num_val <= 0:
                    return False
            except Exception:
                pass
    return True

def lesson_of(teacher: str, day: str, period: int):
    tt = st.session_state.timetable
    if tt.empty:
        return None
    m = tt[(tt["교사명"] == teacher) & (tt["요일"] == day) & (tt["교시"] == safe_int(period))]
    return None if m.empty else m.iloc[0].to_dict()

def has_duty_on_date(teacher: str, on_date: str) -> bool:
    """복무 시트에 해당 교사가 해당 일자에 등록되어 있으면 True (수업 불가능)"""
    duties = st.session_state.get("duties", pd.DataFrame())
    if duties.empty or "교사명" not in duties.columns or "일자" not in duties.columns:
        return False
    norm_date = normalize_date_str(on_date)
    return ((duties["교사명"] == teacher) & (duties["일자"] == norm_date)).any()

def is_free(teacher: str, day: str, period: int, on_date: str = None,
            is_part_time: bool = False, eff_tt: pd.DataFrame = None) -> bool:
    p_int = safe_int(period)
    norm_date = normalize_date_str(on_date)
    if not is_teacher_available(teacher, day, p_int, is_part_time):
        return False
    # 복무 등록된 날은 무조건 불가능
    if norm_date and has_duty_on_date(teacher, norm_date):
        return False
    if not is_part_time:
        if norm_date:
            eff = eff_tt if eff_tt is not None else get_effective_timetable_for_date(norm_date)
            if not eff.empty and not eff[(eff["교사명"] == teacher) & (eff["교시"] == p_int)].empty:
                return False
        else:
            if lesson_of(teacher, day, p_int) is not None:
                return False
    if norm_date and not is_part_time:
        s = st.session_state.subs
        if not s.empty and "일자" in s.columns and "교시" in s.columns:
            if ((s["일자"] == norm_date) & (s["교시"] == p_int) & (s["보강교사"] == teacher)).any():
                return False
        a = st.session_state.absences
        if not a.empty and "일자" in a.columns and "교시" in a.columns:
            if ((a["일자"] == norm_date) & (a["교사명"] == teacher) & (a["교시"] == p_int)).any():
                return False
    return True

def is_class_free(class_name: str, day: str, period: int) -> bool:
    tt = st.session_state.timetable
    if tt.empty:
        return True
    return tt[(tt["학급"] == class_name) & (tt["요일"] == day) & (tt["교시"] == safe_int(period))].empty

def can_teacher_take_slot(teacher: str, day: str, period: int, on_date: str = None,
                          is_part_time: bool = False, eff_tt: pd.DataFrame = None) -> bool:
    return is_teacher_available(teacher, day, period, is_part_time) and \
           is_free(teacher, day, period, on_date, is_part_time, eff_tt=eff_tt)

def absent_all_day(teacher: str, on_date: str) -> bool:
    a = st.session_state.absences
    if a.empty or "일자" not in a.columns:
        # 복무도 하루 종일 불가능으로 취급
        return has_duty_on_date(teacher, on_date)
    norm_date = normalize_date_str(on_date)
    if ((a["일자"] == norm_date) & (a["교사명"] == teacher)).any():
        return True
    return has_duty_on_date(teacher, norm_date)

def cumulative_sub_count(start_date: str = None, end_date: str = None) -> dict:
    if start_date is None and end_date is None:
        if st.session_state.get("_cum_cache") is not None:
            return st.session_state._cum_cache

    s = st.session_state.subs
    base = {t: 0 for t in st.session_state.teachers["교사명"].tolist()} if not st.session_state.teachers.empty else {}
    if not s.empty and "보강교사" in s.columns:
        if start_date and end_date and "일자" in s.columns:
            s = s[(s["일자"] >= normalize_date_str(start_date)) & (s["일자"] <= normalize_date_str(end_date))]
        counts = s["보강교사"].value_counts()
        for k, v in counts.items():
            if k in base:
                base[k] = safe_int(v)

    if start_date is None and end_date is None:
        st.session_state._cum_cache = base
    return base

def weekly_load() -> dict:
    if st.session_state.get("_load_cache") is not None:
        return st.session_state._load_cache
    tt = st.session_state.timetable
    result = tt["교사명"].value_counts().to_dict() if not tt.empty else {}
    st.session_state._load_cache = result
    return result

def recommend_substitutes(day: str, period: int, subject: str, class_name: str,
                          absent_teacher: str, on_date: str, top_n: int = 12,
                          include_part_time: bool = False, eff_tt: pd.DataFrame = None) -> pd.DataFrame:
    teachers = st.session_state.teachers
    part = st.session_state.part_time
    norm_date = normalize_date_str(on_date)
    if eff_tt is None:
        eff_tt = get_effective_timetable_for_date(norm_date) if norm_date else st.session_state.timetable

    if teachers.empty or "교사명" not in teachers.columns:
        return pd.DataFrame()

    grp = subject_group(subject)
    grade = grade_of(class_name)
    cum = cumulative_sub_count()
    load = weekly_load()
    max_cum = max(cum.values()) if cum else 0
    rows = []

    free_teachers = []
    for t in teachers["교사명"].tolist():
        if t == absent_teacher or absent_all_day(t, norm_date):
            continue
        if is_free(t, day, period, norm_date, eff_tt=eff_tt):
            free_teachers.append(t)

    for t in free_teachers:
        sub_tt = eff_tt[eff_tt["교사명"] == t] if not eff_tt.empty else pd.DataFrame()
        my_groups = set(sub_tt["과목군"]) if not sub_tt.empty and "과목군" in sub_tt.columns else set()
        my_grades = {grade_of(c) for c in sub_tt["학급"]} if not sub_tt.empty and "학급" in sub_tt.columns else set()
        my_classes = set(sub_tt["학급"]) if not sub_tt.empty and "학급" in sub_tt.columns else set()

        is_same_subject = grp in my_groups
        is_same_grade = grade in my_grades

        if is_same_subject and is_same_grade:
            prio, prio_label, score = 1, "1순위 · 동일 과목 & 동일 학년", 120
        elif is_same_subject:
            prio, prio_label, score = 2, "2순위 · 동일 과목", 90
        elif is_same_grade:
            prio, prio_label, score = 3, "3순위 · 동일 학년", 60
        else:
            prio, prio_label, score = 4, "4순위 · 전체 공강", 20

        c = cum.get(t, 0)
        score += (max_cum - c) * 6 + max(0, (22 - safe_int(load.get(t, 0)))) * 0.8
        if class_name in my_classes:
            score += 5

        t_row = teachers[teachers["교사명"] == t]
        rows.append({
            "보강교사": t, "유형": "정규교사", "우선순위": prio_label, "_prio": prio,
            "담당과목": t_row["담당과목"].iloc[0] if not t_row.empty and "담당과목" in t_row.columns else "",
            "주당시수": safe_int(load.get(t, 0)), "누적보강": c, "추천점수": round(score, 1)
        })

    if include_part_time and not part.empty and "시간강사명" in part.columns:
        for t in part["시간강사명"].dropna().tolist():
            if not is_free(t, day, period, norm_date, is_part_time=True):
                continue
            t_row = part[part["시간강사명"] == t]
            rows.append({
                "보강교사": t, "유형": "시간강사", "우선순위": "시간강사", "_prio": 5,
                "담당과목": t_row["담당과목"].iloc[0] if not t_row.empty and "담당과목" in t_row.columns else "",
                "주당시수": 0, "누적보강": 0, "추천점수": 40
            })

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    return df.sort_values(["_prio", "추천점수"], ascending=[True, False]).drop(columns=["_prio"]).head(top_n).reset_index(drop=True)

def auto_assign_all(cid: str, on_date: str, day: str, rows: pd.DataFrame, include_pt: bool = False) -> list:
    log = []
    if rows.empty:
        return log
    norm_date = normalize_date_str(on_date)
    eff_tt = get_effective_timetable_for_date(norm_date)
    for r in rows.itertuples(index=False):
        p_val = safe_int(r.교시)
        cand = recommend_substitutes(day, p_val, r.과목, r.학급,
                                     r.교사명, norm_date, top_n=1, include_part_time=include_pt, eff_tt=eff_tt)
        if cand.empty:
            log.append((p_val, None, "배정 가능한 공강 교사가 없습니다."))
            continue
        pick = cand.iloc[0]
        add_substitute(cid, norm_date, day, p_val, r.학급, r.과목,
                       r.교사명, pick["보강교사"], "자동배정", pick["우선순위"], "")
        log.append((p_val, pick["보강교사"], pick["우선순위"]))
    return log

def add_substitute(cid, on_date, day, period, class_name, subject,
                   absent_teacher, sub_teacher, method, priority, memo):
    push_history(f"보강 배정 ({sub_teacher})")
    s = st.session_state.subs
    p_int = safe_int(period)
    norm_date = normalize_date_str(on_date)
    if not s.empty and "결강ID" in s.columns and "교시" in s.columns:
        s = s[~((s["결강ID"] == cid) & (s["교시"] == p_int))]
    new = {
        "결강ID": cid, "일자": norm_date, "요일": day, "교시": p_int,
        "학급": class_name, "과목": subject, "결강교사": absent_teacher,
        "보강교사": sub_teacher, "배정방식": method, "우선순위": priority,
        "비고": memo, "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    st.session_state.subs = pd.concat([s, pd.DataFrame([new])], ignore_index=True)
    save_work_data_to_gsheet()

def cancel_substitute(cid: str, period: int):
    p_int = safe_int(period)
    push_history(f"보강 배정 취소 (교시 {p_int})")
    s = st.session_state.subs
    if not s.empty and "결강ID" in s.columns and "교시" in s.columns:
        st.session_state.subs = s[~((s["결강ID"] == cid) & (s["교시"] == p_int))].reset_index(drop=True)
    save_work_data_to_gsheet()

# (이하 validate_swap, get_target_time_recommendations, do_swap, do_linked_swap, 
#  teacher_matrix, class_matrix, get_teacher_week_view, get_changed_teachers_for_week, 
#  build_report_html, to_excel_bytes 등은 원본과 동일하게 유지하되, 
#  is_free / absent_all_day가 복무를 반영하므로 자동으로 적용됨.
#  원본 코드의 해당 함수들을 그대로 붙여넣으시면 됩니다. 
#  여기서는 공간 관계상 핵심 변경 부분만 강조했습니다.)

# ==========================================================================================
# 사이드바 (일일 변경 내역서 버튼 추가)
# ==========================================================================================
with st.sidebar:
    st.header("데이터 (Google Sheets)")

    if st.button("🔄 원본 시간표 다시 불러오기", use_container_width=True):
        load_timetable_from_gsheet.clear()
        ti, tt = load_timetable_from_gsheet()
        st.session_state.teachers = ti
        st.session_state.timetable = tt
        _invalidate_all_caches()
        st.success("원본 시간표를 다시 불러왔습니다.")
        st.rerun()

    if st.button("🔄 작업 내역 다시 불러오기", use_container_width=True):
        load_work_data_from_gsheet.clear()
        absences, subs, swaps, part_time, cumulative, duties = load_work_data_from_gsheet()
        st.session_state.absences = absences
        st.session_state.subs = subs
        st.session_state.swaps = swaps
        st.session_state.part_time = part_time
        st.session_state.cumulative = cumulative
        st.session_state.duties = duties
        _invalidate_all_caches()
        st.success("작업 내역을 다시 불러왔습니다.")
        st.rerun()

    if st.button("💾 현재 작업 내역 저장", use_container_width=True, type="primary"):
        if save_work_data_to_gsheet():
            st.success("Google Sheets에 저장 완료! (작업내역 로그 포함)")

    st.divider()
    st.subheader("작업 되돌리기")
    c_undo, c_redo = st.columns(2)
    if c_undo.button("↩ 되돌리기 (Undo)", use_container_width=True):
        if undo():
            save_work_data_to_gsheet()
            st.success("이전 단계로 되돌렸습니다. (시트도 반영됨)")
            st.rerun()
        else:
            st.warning("더 이상 되돌릴 수 없습니다.")
    if c_redo.button("↪ 다시실행 (Redo)", use_container_width=True):
        if redo():
            save_work_data_to_gsheet()
            st.success("다시 실행했습니다. (시트도 반영됨)")
            st.rerun()
        else:
            st.warning("다시 실행할 단계가 없습니다.")

    if "history" in st.session_state and st.session_state.history:
        st.caption(f"히스토리: {st.session_state.history_index + 1} / {len(st.session_state.history)}")
        with st.expander("최근 작업 목록"):
            for h in reversed(st.session_state.history[-6:]):
                st.write(f"{h.get('time','')} - {h.get('action','')}")

    st.divider()
    st.caption(f"{SCHOOL_YEAR}학년도 · {SCHOOL_NAME}")
    st.metric("등록 교사", f"{len(st.session_state.teachers)} 명")
    st.metric("시간강사", f"{len(st.session_state.part_time)} 명")
    st.metric("주간 수업", f"{len(st.session_state.timetable)} 시수")
    st.metric("누적 보강", f"{len(st.session_state.subs)} 건")

    # ===== 일일 변경 내역서 버튼 (누적보강 UI 근처) =====
    st.divider()
    st.subheader("📄 일일 변경 내역서")
    report_date = st.date_input("출력 일자", value=date.today(), key="sidebar_report_date")
    rd = report_date.strftime("%Y-%m-%d")
    if st.button("📥 HTML/엑셀 다운로드 & 미리보기", use_container_width=True, type="primary"):
        html = build_report_html(rd)
        st.download_button("HTML 다운로드", data=html.encode("utf-8"),
                           file_name=f"결보강내역서_{rd}.html", mime="text/html")
        a = st.session_state.absences
        s = st.session_state.subs
        xls = to_excel_bytes({
            "결강": a[a["일자"] == rd] if not a.empty and "일자" in a.columns else pd.DataFrame(),
            "보강": s[s["일자"] == rd] if not s.empty and "일자" in s.columns else pd.DataFrame(),
            "시간표맞교환": st.session_state.swaps,
            "교사별시간표": teacher_matrix() if 'teacher_matrix' in globals() else pd.DataFrame(),
            "학급별시간표": class_matrix() if 'class_matrix' in globals() else pd.DataFrame(),
        })
        st.download_button("엑셀 백업 다운로드", data=xls,
                           file_name=f"결보강내역_{rd}.xlsx",
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
        st.components.v1.html(html, height=600, scrolling=True)

if st.session_state.timetable.empty:
    st.title("중학교 시간표·결보강 관리 프로그램")
    st.info("Google Sheets에서 시간표를 불러오는 중입니다...")
    st.stop()

# ==========================================================================================
# 6. 메인 화면 (탭 재구성)
# ==========================================================================================
st.title("전체 시간표 관리 · 결강/보강 자동 처리 (Google Sheets 연동)")

tabs = st.tabs([
    "시간표 조회",
    "시간강사 관리",
    "결강·보강",                     # ← 통합 탭
    "시간표 맞교환 & 변경 추천",
    "통계 (학기/월별)",
    "시간표 변경 테스트용",
    "변경된 교사 주간표",
    "🤖 복무 기반 판단 & 맞교환",
    "📋 복무 일정 관리"              # ← 신규 탭
])

# ------------------------------------------------------------------ 0. 시간표 조회 (기존과 동일)
with tabs[0]:
    view = st.radio("보기 방식", ["교사별 전체 매트릭스 (기본 주간)", "학급별 매트릭스 (기본 주간)", "교사 1인 주간표 (달력 – 주차 독립)"], horizontal=True)
    if view == "교사별 전체 매트릭스 (기본 주간)":
        st.dataframe(teacher_matrix(), use_container_width=True, height=760)
    elif view == "학급별 매트릭스 (기본 주간)":
        st.dataframe(class_matrix(), use_container_width=True, height=680)
    else:
        teacher_list = st.session_state.teachers["교사명"].tolist() if not st.session_state.teachers.empty else []
        if not teacher_list:
            st.warning("등록된 교사가 없습니다.")
        else:
            t = st.selectbox("교사 선택", teacher_list)
            ref_date = st.date_input("기준 날짜 (해당 주 월~금 표시)", value=date.today(), key="week_ref")
            grid, week_dates = get_teacher_week_view(t, ref_date)
            st.caption(f"주간: {week_dates[0].strftime('%Y-%m-%d')} ~ {week_dates[4].strftime('%Y-%m-%d')}")
            info_rows = st.session_state.teachers[st.session_state.teachers["교사명"] == t]
            info = info_rows.iloc[0] if not info_rows.empty else {}
            c = st.columns(4)
            c[0].metric("담당 과목", str(info.get("담당과목", "")) or "-")
            c[1].metric("담임 학급", str(info.get("담임학급", "")) or "-")
            tt_sub = st.session_state.timetable[st.session_state.timetable['교사명'] == t] if not st.session_state.timetable.empty else pd.DataFrame()
            c[2].metric("주당 시수", len(tt_sub))
            c[3].metric("누적 보강", cumulative_sub_count().get(t, 0))
            st.dataframe(grid, use_container_width=True)

# ------------------------------------------------------------------ 1. 시간강사 관리 (기존과 동일)
with tabs[1]:
    st.subheader("시간강사 등록 및 가능 시간표 (불가 = 불가능 / 가능 = 가능)")
    ed_pt = st.data_editor(st.session_state.part_time, num_rows="dynamic", use_container_width=True, height=400, key="ed_part")
    if st.button("시간강사 정보 저장", type="primary", use_container_width=True):
        st.session_state.part_time = ed_pt.reset_index(drop=True)
        push_history("시간강사 정보 수정")
        if save_work_data_to_gsheet():
            st.success("시간강사 정보가 Google Sheets에 저장되었습니다!")
            st.rerun()

# ------------------------------------------------------------------ 2. 결강·보강 통합 탭 (좌우 분할)
with tabs[2]:
    st.subheader("결강 등록 & 보강 배정 (한 화면)")
    left, right = st.columns(2)

    # ===== 왼쪽: 결강 등록 =====
    with left:
        st.markdown("### 📌 결강 등록")
        c = st.columns([1, 1, 1])
        d_sel = c[0].date_input("결강 일자", value=date.today(), key="abs_date")
        on_date = d_sel.strftime("%Y-%m-%d")
        day = WEEKDAY_KR[d_sel.weekday()]
        c[1].text_input("요일", value=day, disabled=True, key="abs_day")
        teacher_list = st.session_state.teachers["교사명"].tolist() if not st.session_state.teachers.empty else []
        who = c[2].selectbox("결강 교사", teacher_list if teacher_list else ["없음"], key="abs_teacher")
        reason = st.selectbox("사유", ABSENCE_REASONS, key="abs_reason")
        detail = st.text_input("상세 사유 / 비고 (선택)", key="abs_detail")

        if day not in DAYS:
            st.warning("주말은 시간표가 없습니다.")
        elif not teacher_list or who == "없음":
            st.warning("등록된 교사가 없습니다.")
        else:
            eff_tt_today = get_effective_timetable_for_date(on_date)
            todays = eff_tt_today[
                (eff_tt_today["교사명"] == who) &
                (eff_tt_today["요일"] == day)
            ].sort_values("교시") if not eff_tt_today.empty else pd.DataFrame()

            if todays.empty:
                st.info(f"{who} 교사는 {on_date} ({day}요일)에 예정된 수업이 없습니다. (교체 내역 반영됨)")
            else:
                opts = [f"{safe_int(r['교시'])}교시 · {r['학급']} · {r['과목']}" for _, r in todays.iterrows()]
                picked = st.multiselect("결강 교시 선택", opts, default=opts, key="abs_periods")

                if st.button("결강 등록", type="primary", key="btn_reg_abs"):
                    if not picked:
                        st.warning("교시를 선택하세요.")
                    else:
                        try:
                            push_history(f"결강 등록 ({who})")
                            cid = f"{on_date}-{who}"
                            sel_periods = [safe_int(o.split("교시")[0].strip()) for o in picked]
                            rows = todays[todays["교시"].apply(safe_int).isin(sel_periods)]
                            a = st.session_state.absences
                            if not a.empty and "일자" in a.columns:
                                a = a[~((a["일자"] == on_date) & (a["교사명"] == who))]
                            new_rows = []
                            for _, r in rows.iterrows():
                                new_rows.append({
                                    "결강ID": cid, "일자": on_date, "요일": day, "교사명": who,
                                    "사유": reason, "상세사유": detail, "교시": safe_int(r["교시"]),
                                    "학급": r["학급"], "과목": r["과목"],
                                    "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M")
                                })
                            if new_rows:
                                st.session_state.absences = pd.concat([a, pd.DataFrame(new_rows)], ignore_index=True)
                                save_work_data_to_gsheet()
                                st.success(f"{who} 교사 {len(new_rows)}시간 결강 등록 완료.")
                                st.rerun()
                        except Exception as e:
                            st.error(f"오류: {e}")

        st.divider()
        st.markdown("#### 등록된 결강 목록")
        ab = st.session_state.absences
        if not ab.empty:
            show = ab[ab["일자"] == on_date] if (st.checkbox("선택한 일자만 보기", value=True, key="abs_filter") and "일자" in ab.columns) else ab
            st.dataframe(show.sort_values(["일자", "교사명", "교시"]) if "일자" in show.columns else show, use_container_width=True, height=300)
        else:
            st.info("등록된 결강이 없습니다.")

    # ===== 오른쪽: 보강 배정 =====
    with right:
        st.markdown("### 📌 보강 배정")
        ab = st.session_state.absences
        if ab.empty or "결강ID" not in ab.columns:
            st.info("먼저 결강을 등록하세요.")
        else:
            unique_cids = sorted(ab["결강ID"].dropna().unique(), reverse=True)
            if not unique_cids:
                st.warning("유효한 결강 건이 없습니다.")
            else:
                cid = st.selectbox("결강 건 선택", unique_cids, key="sub_cid")
                rows = ab[ab["결강ID"] == cid].sort_values("교시")
                if rows.empty:
                    st.error("데이터가 없습니다.")
                else:
                    head = rows.iloc[0]
                    norm_head_date = normalize_date_str(head.get('일자',''))
                    st.markdown(f"**{norm_head_date} ({head.get('요일','')}) · {head.get('교사명','')} · {head.get('사유','')} · {len(rows)}시간**")
                    include_pt = st.checkbox("시간강사도 추천에 포함", value=False, key="sub_pt")

                    if st.button("전 교시 자동 배정", type="primary", key="btn_auto_sub"):
                        try:
                            push_history("자동 보강 배정")
                            log = auto_assign_all(cid, norm_head_date, head["요일"], rows, include_pt=include_pt)
                            for p, t, why in log:
                                if t:
                                    st.success(f"{p}교시 → {t} ({why})")
                                else:
                                    st.error(f"{p}교시 → {why}")
                            st.rerun()
                        except Exception as e:
                            st.error(f"오류: {e}")

                    st.divider()
                    eff_tt = get_effective_timetable_for_date(norm_head_date)
                    for _, r in rows.iterrows():
                        p = safe_int(r["교시"])
                        with st.expander(f"{p}교시 · {r.get('학급','')} · {r.get('과목','')}", expanded=True):
                            cur = st.session_state.subs
                            assigned = None
                            if not cur.empty and "결강ID" in cur.columns and "교시" in cur.columns:
                                m = cur[(cur["결강ID"] == cid) & (cur["교시"] == p)]
                                if not m.empty:
                                    assigned = m.iloc[0]["보강교사"]

                            if assigned:
                                st.success(f"현재 배정 : **{assigned}**")
                                if st.button("❌ 배정 취소", key=f"cancel_{cid}_{p}", type="secondary"):
                                    cancel_substitute(cid, p)
                                    st.success(f"{p}교시 배정이 취소되었습니다.")
                                    st.rerun()
                            else:
                                try:
                                    cand = recommend_substitutes(
                                        head["요일"], p, r["과목"], r["학급"],
                                        head["교사명"], norm_head_date, top_n=12, include_part_time=include_pt, eff_tt=eff_tt
                                    )
                                except Exception as e:
                                    st.error(f"추천 오류: {e}")
                                    cand = pd.DataFrame()

                                if cand.empty:
                                    st.error("가능한 공강 교사가 없습니다. (복무/결강/이미 배정된 교사 제외됨)")
                                else:
                                    st.dataframe(cand, use_container_width=True, hide_index=True, height=200)
                                    cc = st.columns([2, 2, 1])
                                    pick = cc[0].selectbox("보강 교사", cand["보강교사"].tolist(), key=f"pick_{cid}_{p}")
                                    memo = cc[1].text_input("비고", key=f"memo_{cid}_{p}")
                                    if cc[2].button("배정", key=f"btn_{cid}_{p}", use_container_width=True):
                                        try:
                                            pr_series = cand.loc[cand["보강교사"] == pick, "우선순위"]
                                            pr = pr_series.iloc[0] if not pr_series.empty else "수동배정"
                                            add_substitute(
                                                cid, norm_head_date, head["요일"], p, r["학급"], r["과목"],
                                                head["교사명"], pick, "수동배정", pr, memo
                                            )
                                            st.success(f"{p}교시 배정 완료.")
                                            st.rerun()
                                        except Exception as e:
                                            st.error(f"오류: {e}")

# ------------------------------------------------------------------ 3. 결강 등록
with tabs[3]:
    c = st.columns([1, 1, 1, 2])
    d_sel = c[0].date_input("결강 일자", value=date.today())
    on_date = d_sel.strftime("%Y-%m-%d")
    day = WEEKDAY_KR[d_sel.weekday()]
    c[1].text_input("요일", value=day, disabled=True)
    teacher_list = st.session_state.teachers["교사명"].tolist() if not st.session_state.teachers.empty else []
    who = c[2].selectbox("결강 교사", teacher_list) if teacher_list else c[2].selectbox("결강 교사", ["없음"])
    reason = c[3].selectbox("사유", ABSENCE_REASONS)
    detail = st.text_input("상세 사유 / 비고 (선택)")

    if day not in DAYS:
        st.warning("주말은 시간표가 없습니다.")
    elif not teacher_list:
        st.warning("등록된 교사가 없습니다.")
    else:
        eff_tt_today = get_effective_timetable_for_date(on_date)
        todays = eff_tt_today[
            (eff_tt_today["교사명"] == who) &
            (eff_tt_today["요일"] == day)
        ].sort_values("교시") if not eff_tt_today.empty else pd.DataFrame()

        if todays.empty:
            st.info(f"{who} 교사는 {on_date} ({day}요일)에 예정된 수업이 없습니다. (교체 내역 반영됨)")
        else:
            opts = [f"{safe_int(r['교시'])}교시 · {r['학급']} · {r['과목']}" for _, r in todays.iterrows()]
            picked = st.multiselect("결강 교시 선택", opts, default=opts)

            if st.button("결강 등록", type="primary"):
                if not picked:
                    st.warning("교시를 선택하세요.")
                else:
                    try:
                        push_history(f"결강 등록 ({who})")
                        cid = f"{on_date}-{who}"
                        sel_periods = [safe_int(o.split("교시")[0].strip()) for o in picked]
                        rows = todays[todays["교시"].apply(safe_int).isin(sel_periods)]
                        a = st.session_state.absences
                        if not a.empty and "일자" in a.columns:
                            a = a[~((a["일자"] == on_date) & (a["교사명"] == who))]
                        new_rows = []
                        for _, r in rows.iterrows():
                            new_rows.append({
                                "결강ID": cid, "일자": on_date, "요일": day, "교사명": who,
                                "사유": reason, "상세사유": detail, "교시": safe_int(r["교시"]),
                                "학급": r["학급"], "과목": r["과목"],
                                "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M")
                            })
                        if new_rows:
                            st.session_state.absences = pd.concat([a, pd.DataFrame(new_rows)], ignore_index=True)
                            save_work_data_to_gsheet()
                            st.success(f"{who} 교사 {len(new_rows)}시간 결강 등록 완료.")
                            st.rerun()
                    except Exception as e:
                        st.error(f"오류: {e}")

    st.divider()
    st.subheader("등록된 결강 목록")
    ab = st.session_state.absences
    if not ab.empty:
        show = ab[ab["일자"] == on_date] if (st.checkbox("선택한 일자만 보기", value=True) and "일자" in ab.columns) else ab
        st.dataframe(show.sort_values(["일자", "교사명", "교시"]) if "일자" in show.columns else show, use_container_width=True)
    else:
        st.info("등록된 결강이 없습니다.")

# ------------------------------------------------------------------ 4. 보강 배정
with tabs[4]:
    ab = st.session_state.absences
    if ab.empty or "결강ID" not in ab.columns:
        st.info("먼저 결강을 등록하세요.")
    else:
        unique_cids = sorted(ab["결강ID"].dropna().unique(), reverse=True)
        if not unique_cids:
            st.warning("유효한 결강 건이 없습니다.")
        else:
            cid = st.selectbox("결강 건 선택", unique_cids)
            rows = ab[ab["결강ID"] == cid].sort_values("교시")
            if rows.empty:
                st.error("데이터가 없습니다.")
            else:
                head = rows.iloc[0]
                norm_head_date = normalize_date_str(head.get('일자',''))
                st.markdown(f"### {norm_head_date} ({head.get('요일','')}) · **{head.get('교사명','')}** · {head.get('사유','')} · {len(rows)}시간")
                include_pt = st.checkbox("시간강사도 추천에 포함", value=False)

                if st.button("전 교시 자동 배정", type="primary"):
                    try:
                        push_history("자동 보강 배정")
                        log = auto_assign_all(cid, norm_head_date, head["요일"], rows, include_pt=include_pt)
                        for p, t, why in log:
                            if t:
                                st.success(f"{p}교시 → {t} ({why})")
                            else:
                                st.error(f"{p}교시 → {why}")
                        st.rerun()
                    except Exception as e:
                        st.error(f"오류: {e}")

                st.divider()
                eff_tt = get_effective_timetable_for_date(norm_head_date)
                for _, r in rows.iterrows():
                    p = safe_int(r["교시"])
                    with st.expander(f"{p}교시 · {r.get('학급','')} · {r.get('과목','')}", expanded=True):
                        cur = st.session_state.subs
                        assigned = None
                        if not cur.empty and "결강ID" in cur.columns and "교시" in cur.columns:
                            m = cur[(cur["결강ID"] == cid) & (cur["교시"] == p)]
                            if not m.empty:
                                assigned = m.iloc[0]["보강교사"]

                        if assigned:
                            st.success(f"현재 배정 : **{assigned}**")
                            if st.button("❌ 배정 취소 (결강 상태로 되돌리기)", key=f"cancel_{cid}_{p}", type="secondary"):
                                cancel_substitute(cid, p)
                                st.success(f"{p}교시 배정이 취소되었습니다.")
                                st.rerun()
                        else:
                            try:
                                cand = recommend_substitutes(
                                    head["요일"], p, r["과목"], r["학급"],
                                    head["교사명"], norm_head_date, top_n=12, include_part_time=include_pt, eff_tt=eff_tt
                                )
                            except Exception as e:
                                st.error(f"추천 오류: {e}")
                                cand = pd.DataFrame()

                            if cand.empty:
                                st.error("가능한 공강 교사가 없습니다.")
                            else:
                                st.dataframe(cand, use_container_width=True, hide_index=True)
                                cc = st.columns([2, 2, 1])
                                pick = cc[0].selectbox("보강 교사", cand["보강교사"].tolist(), key=f"pick_{cid}_{p}")
                                memo = cc[1].text_input("비고", key=f"memo_{cid}_{p}")
                                if cc[2].button("배정", key=f"btn_{cid}_{p}", use_container_width=True):
                                    try:
                                        pr_series = cand.loc[cand["보강교사"] == pick, "우선순위"]
                                        pr = pr_series.iloc[0] if not pr_series.empty else "수동배정"
                                        add_substitute(
                                            cid, norm_head_date, head["요일"], p, r["학급"], r["과목"],
                                            head["교사명"], pick, "수동배정", pr, memo
                                        )
                                        st.success(f"{p}교시 배정 완료.")
                                        st.rerun()
                                    except Exception as e:
                                        st.error(f"오류: {e}")

# ------------------------------------------------------------------ 5. 맞교환
with tabs[5]:
    st.markdown("### 🔄 스마트 시간표 변경 & 맞교환")
    st.caption("변경은 해당 날짜에만 적용됩니다. 이동된 수업은 주간표에서 원본 위치가 함께 표시됩니다.")

    col_a, col_b = st.columns(2)
    teacher_list = st.session_state.teachers["교사명"].tolist() if not st.session_state.teachers.empty else []

    with col_a:
        st.markdown("#### 1️⃣ 원본 수업 (교사 A)")
        date_a_input = st.date_input("원본 날짜", value=date.today(), key="date_a")
        date_a_str = date_a_input.strftime("%Y-%m-%d")
        day_a = WEEKDAY_KR[date_a_input.weekday()]

        if not teacher_list:
            st.warning("등록된 교사가 없습니다.")
            t_a = None
            pick_a = None
        else:
            t_a = st.selectbox("교사 A", teacher_list, key="teacher_a_select")
            eff_tt_a = get_effective_timetable_for_date(date_a_str)
            sub_a = eff_tt_a[(eff_tt_a["교사명"] == t_a) & (eff_tt_a["요일"] == day_a)].sort_values("교시") if not eff_tt_a.empty else pd.DataFrame()

            if sub_a.empty:
                st.warning(f"{t_a} 교사는 해당 요일에 수업이 없습니다.")
                pick_a = None
            else:
                opts_a = [f"{safe_int(r['교시'])}교시 · {r['학급']} · {r['과목']}" for _, r in sub_a.iterrows()]
                sel_a = st.selectbox("변경할 수업", opts_a, key="lesson_a_select")
                row_a = sub_a.iloc[opts_a.index(sel_a)]
                pick_a = {
                    "교사명": t_a, "일자": date_a_str, "요일": day_a,
                    "교시": safe_int(row_a["교시"]), "학급": row_a["학급"], "과목": row_a["과목"]
                }

    with col_b:
        st.markdown("#### 2️⃣ 이동 희망 시간")
        date_b_input = st.date_input("이동 희망 날짜", value=date.today(), key="date_b")
        date_b_str = date_b_input.strftime("%Y-%m-%d")
        day_b = WEEKDAY_KR[date_b_input.weekday()]

        p_b_input = st.selectbox("희망 교시", list(range(1, PERIODS_PER_DAY.get(day_b, MAX_PERIOD) + 1)), key="period_b_select")
        st.info(f"목표: [{date_b_str} ({day_b})] {p_b_input}교시")

    is_part_time_purpose = st.checkbox("☑ 시간강사 구인 목적", value=False)

    st.divider()

    if pick_a:
        st.subheader(f"📌 {pick_a['일자']} {pick_a['교시']}교시 {pick_a['학급']} {pick_a['과목']} → {date_b_str} {p_b_input}교시")

        df_swap_recs, df_linked_recs, df_sub_recs = get_target_time_recommendations(
            pick_a["교사명"], pick_a["일자"], pick_a["교시"], pick_a["학급"], pick_a["과목"],
            date_b_str, p_b_input
        )

        t_swap, t_linked, t_sub = st.tabs(["🔄 1:1 맞교환", "🔗 연계 교환", "➕ 대리/보강"])

        with t_swap:
            if df_swap_recs.empty:
                st.info("해당 시간에 수업이 있는 다른 교사가 없습니다.")
            else:
                for idx, row in df_swap_recs.iterrows():
                    b_item = row["b_info"]
                    with st.expander(f"{row['상태']} **{row['교사B']}** ({row['현재 수업']})", expanded=(idx==0)):
                        if not row["is_valid"]:
                            for e in row["errs"]:
                                st.error(e)
                        else:
                            st.success("맞교환 가능")
                            if st.button(f"'{row['교사B']}'와 맞교환 실행", key=f"exec_swap_{idx}", type="primary"):
                                if do_swap(pick_a, b_item, date_a_str, date_b_str, is_part_time_purpose):
                                    st.success("맞교환 기록 완료")
                                    st.rerun()

        with t_linked:
            if df_linked_recs.empty:
                st.info("연계 교환 가능한 교사가 없습니다.")
            else:
                for idx, row in df_linked_recs.iterrows():
                    with st.expander(f"{row['상태']} **{row['교사B']}**", expanded=(idx==0)):
                        st.write(row["설명"])
                        if st.button(f"'{row['교사B']}'와 연계 교환", key=f"exec_linked_{idx}", type="primary"):
                            if do_linked_swap(pick_a, row["교사B"], date_a_str, date_b_str, day_b, p_b_input, is_part_time_purpose):
                                st.success("연계 교환 기록 완료")
                                st.rerun()

        with t_sub:
            if not df_sub_recs.empty:
                st.dataframe(df_sub_recs[["보강/대리 교사", "우선순위", "담당과목", "누적보강"]], use_container_width=True, hide_index=True)
                c_s1, c_s2 = st.columns([3, 1])
                sub_pick = c_s1.selectbox("대리 교사", df_sub_recs["보강/대리 교사"], key="sel_sub")
                if c_s2.button("보강 등록", key="btn_reg_sub", type="primary"):
                    cid = f"{date_b_str}-{pick_a['교사명']}"
                    pr_series = df_sub_recs.loc[df_sub_recs["보강/대리 교사"] == sub_pick, "우선순위"]
                    pr_val = pr_series.iloc[0] if not pr_series.empty else "수동배정"
                    add_substitute(cid, date_b_str, day_b, p_b_input, pick_a["학급"], pick_a["과목"],
                                   pick_a["교사명"], sub_pick, "시간표변경_대리", pr_val, f"{pick_a['일자']} 이동")
                    st.success(f"'{sub_pick}' 등록 완료")
            else:
                st.info("추천 대리/보강 교사가 없습니다.")

    st.divider()
    st.subheader("변경 이력")
    st.dataframe(st.session_state.swaps, use_container_width=True)

# ------------------------------------------------------------------ 6. 내역서
with tabs[6]:
    rd = st.date_input("출력 일자", value=date.today(), key="rdate").strftime("%Y-%m-%d")
    html = build_report_html(rd)
    st.download_button("HTML 다운로드", data=html.encode("utf-8"),
                       file_name=f"결보강내역서_{rd}.html", mime="text/html", type="primary")
    a = st.session_state.absences
    s = st.session_state.subs
    xls = to_excel_bytes({
        "결강": a[a["일자"] == rd] if not a.empty and "일자" in a.columns else pd.DataFrame(),
        "보강": s[s["일자"] == rd] if not s.empty and "일자" in s.columns else pd.DataFrame(),
        "시간표맞교환": st.session_state.swaps,
        "교사별시간표": teacher_matrix(),
        "학급별시간표": class_matrix(),
    })
    st.download_button("엑셀 백업 다운로드", data=xls,
                       file_name=f"결보강내역_{rd}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.divider()
    st.components.v1.html(html, height=900, scrolling=True)

# ------------------------------------------------------------------ 7. 통계
with tabs[7]:
    st.subheader("보강 통계 – 기간 필터")
    col1, col2, col3 = st.columns(3)
    start_d = col1.date_input("시작일", value=date(2026, 3, 1))
    end_d = col2.date_input("종료일", value=date.today())
    period_label = col3.selectbox("빠른 선택", ["전체", "1학기 (3~7월)", "2학기 (8월~다음해 2월)", "이번 달"])

    if period_label == "1학기 (3~7월)":
        start_d, end_d = date(2026, 3, 1), date(2026, 7, 31)
    elif period_label == "2학기 (8월~다음해 2월)":
        start_d, end_d = date(2026, 8, 1), date(2027, 2, 28)
    elif period_label == "이번 달":
        start_d = date.today().replace(day=1)
        end_d = date.today()

    start_str = start_d.strftime("%Y-%m-%d")
    end_str = end_d.strftime("%Y-%m-%d")

    cum = cumulative_sub_count(start_str, end_str)
    df = pd.DataFrame({"교사명": list(cum.keys()), "누적보강": list(cum.values())})
    df["주당시수"] = df["교사명"].map(weekly_load()).fillna(0).astype(int)
    df = df.sort_values("누적보강", ascending=False)

    c = st.columns(4)
    c[0].metric("총 보강", int(df["누적보강"].sum()) if not df.empty else 0)
    c[1].metric("보강 참여 교사", int((df["누적보강"] > 0).sum()) if not df.empty else 0)
    c[2].metric("최다 보강", int(df["누적보강"].max()) if not df.empty else 0)
    c[3].metric("평균 보강", round(float(df["누적보강"].mean()), 2) if not df.empty else 0)

    st.caption(f"기간: {start_str} ~ {end_str}")
    if not df.empty:
        st.bar_chart(df.set_index("교사명")["누적보강"], height=340)
    st.dataframe(df, use_container_width=True, hide_index=True)

    st.divider()
    st.subheader("누적보강 시트에 저장")
    if st.button("📊 현재 통계를 누적보강 시트에 저장", type="primary"):
        save_df = df.copy()
        save_df["기간"] = f"{start_str} ~ {end_str}"
        save_df["저장시각"] = datetime.now().strftime("%Y-%m-%d %H:%M")
        if save_cumulative_to_gsheet(save_df):
            st.success("누적보강 시트에 저장 완료!")

# ------------------------------------------------------------------ 마지막: 복무 일정 관리 탭
with tabs[8]:
    st.subheader("📋 교사 복무(수업 불가능) 일정 관리")
    st.caption("여기에 등록된 교사는 해당 날짜에 보강 추천 대상에서 자동 제외됩니다. Google Sheets '복무' 시트에 저장됩니다.")

    col1, col2 = st.columns(2)
    with col1:
        duty_teacher = st.selectbox("교사 선택", st.session_state.teachers["교사명"].tolist() if not st.session_state.teachers.empty else [], key="duty_t")
        duty_date = st.date_input("복무 일자", value=date.today(), key="duty_d")
        duty_reason = st.selectbox("사유", ABSENCE_REASONS, key="duty_r")
        duty_detail = st.text_input("상세 사유 (선택)", key="duty_detail")

        if st.button("복무 등록", type="primary"):
            if duty_teacher:
                push_history(f"복무 등록 ({duty_teacher})")
                new_row = {
                    "교사명": duty_teacher,
                    "일자": duty_date.strftime("%Y-%m-%d"),
                    "사유": duty_reason,
                    "상세사유": duty_detail,
                    "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M")
                }
                d = st.session_state.duties
                # 동일 교사+일자 중복 제거
                if not d.empty:
                    d = d[~((d["교사명"] == duty_teacher) & (d["일자"] == new_row["일자"]))]
                st.session_state.duties = pd.concat([d, pd.DataFrame([new_row])], ignore_index=True)
                save_work_data_to_gsheet()
                st.success(f"{duty_teacher} 선생님 {new_row['일자']} 복무 등록 완료 (보강 추천에서 제외됨)")
                st.rerun()

    with col2:
        st.markdown("#### 등록된 복무 목록")
        duties = st.session_state.duties
        if not duties.empty:
            st.dataframe(duties.sort_values(["일자", "교사명"]), use_container_width=True, height=400)
            # 삭제 기능
            del_idx = st.number_input("삭제할 행 번호 (0부터)", min_value=0, max_value=max(0, len(duties)-1), value=0)
            if st.button("선택한 복무 삭제"):
                push_history("복무 삭제")
                st.session_state.duties = duties.drop(duties.index[del_idx]).reset_index(drop=True)
                save_work_data_to_gsheet()
                st.success("삭제 완료")
                st.rerun()
        else:
            st.info("등록된 복무 일정이 없습니다.")

# (나머지 탭 내용 – 맞교환, 통계, 테스트용, 변경된 교사 주간표, 복무기반 판단 – 은 
#  원본 코드의 해당 부분을 그대로 복사하여 tabs[3]~tabs[7]에 배치하시면 됩니다.
#  모든 is_free / recommend_substitutes 호출이 복무를 자동 반영합니다.)
