# -*- coding: utf-8 -*-
"""
==========================================================================================
 중학교 전체 시간표 관리 및 결강·보강 자동 처리 프로그램  (Streamlit Web App)
 2026 서라벌여자중학교용 – Google Sheets 연동 최종 버전
 (불가/가능 지원 + 문자열 강제 변환 + cache_data)
------------------------------------------------------------------------------------------
 실행 방법
   1) pip install streamlit pandas numpy openpyxl xlsxwriter gspread google-auth
   2) streamlit run timetable_app.py
==========================================================================================
"""

import io
import json
import os
from datetime import date, datetime, timedelta
from copy import deepcopy

import numpy as np
import pandas as pd
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials

# ==========================================================================================
# 0. 기본 설정 / 상수
# ==========================================================================================
st.set_page_config(page_title="시간표·결보강 관리", page_icon="📘", layout="wide")

SCHOOL_NAME = "서라벌여자중학교"
SCHOOL_YEAR = "2026"

DAYS = ["월", "화", "수", "목", "금"]
PERIODS_PER_DAY = {"월": 6, "화": 7, "수": 7, "목": 7, "금": 6}
MAX_PERIOD = 7
WEEKDAY_KR = {0: "월", 1: "화", 2: "수", 3: "목", 4: "금", 5: "토", 6: "일"}

# Google Sheets ID
TIMETABLE_SHEET_ID = "1jZhTHyJ8vKXn6tkoFXfY_f52-pj6eQTdVvRCo3cCmBA"   # 원본 시간표
WORK_SHEET_ID = "1g1B1cyZG_tfRn3AD1NZzr30YxYNYFewJeZYdos2obpU"        # 작업 내역

MAX_HISTORY = 30
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

def subject_group(subject: str) -> str:
    if not isinstance(subject, str):
        return ""
    s = subject.strip()
    if s in SUBJECT_GROUP:
        return SUBJECT_GROUP[s]
    return s.rstrip("0123456789")

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
    creds = Credentials.from_service_account_info(
        st.secrets["gcp_service_account"],
        scopes=scopes
    )
    return gspread.authorize(creds)

def get_worksheet(spreadsheet_id: str, sheet_name: str):
    client = get_gspread_client()
    sh = client.open_by_key(spreadsheet_id)
    try:
        return sh.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        return sh.add_worksheet(title=sheet_name, rows=2000, cols=40)

def df_from_worksheet(ws) -> pd.DataFrame:
    """모든 값을 안전하게 문자열로 강제 변환"""
    data = ws.get_all_values()
    if not data or len(data) < 2:
        return pd.DataFrame()

    headers = [str(h) for h in data[0]]
    
    cleaned_rows = []
    for row in data[1:]:
        cleaned_row = []
        for i in range(len(headers)):
            if i < len(row):
                val = row[i]
                cleaned_row.append("" if val is None else str(val))
            else:
                cleaned_row.append("")
        cleaned_rows.append(cleaned_row)

    df = pd.DataFrame(cleaned_rows, columns=headers)
    df = df.replace({"nan": "", "None": "", "NaN": "", "<NA>": ""})
    
    return df

def df_to_worksheet(ws, df: pd.DataFrame):
    ws.clear()
    if df is None or df.empty:
        return
    df_clean = df.fillna("").astype(str)
    values = [df_clean.columns.tolist()] + df_clean.values.tolist()
    ws.update(values, value_input_option="USER_ENTERED")

# ==========================================================================================
# 1. 히스토리 (Undo / Redo)
# ==========================================================================================
def push_history(action_name: str = "작업"):
    if "history" not in st.session_state:
        st.session_state.history = []
        st.session_state.history_index = -1

    st.session_state.history = st.session_state.history[:st.session_state.history_index + 1]

    snapshot = {
        "action": action_name,
        "time": datetime.now().strftime("%H:%M:%S"),
        "absences": st.session_state.absences.copy(deep=True),
        "subs": st.session_state.subs.copy(deep=True),
        "swaps": st.session_state.swaps.copy(deep=True),
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
    st.session_state.absences = snap["absences"].copy(deep=True)
    st.session_state.subs = snap["subs"].copy(deep=True)
    st.session_state.swaps = snap["swaps"].copy(deep=True)
    return True

def redo():
    if "history" not in st.session_state:
        return False
    if st.session_state.history_index >= len(st.session_state.history) - 1:
        return False
    st.session_state.history_index += 1
    snap = st.session_state.history[st.session_state.history_index]
    st.session_state.absences = snap["absences"].copy(deep=True)
    st.session_state.subs = snap["subs"].copy(deep=True)
    st.session_state.swaps = snap["swaps"].copy(deep=True)
    return True

def init_history_if_needed():
    if "history" not in st.session_state:
        st.session_state.history = []
        st.session_state.history_index = -1
        push_history("초기 상태")

# ==========================================================================================
# 2. 데이터 로드 / 저장 (Google Sheets)
# ==========================================================================================
REQUIRED_TT_COLS = ["교사명", "요일", "교시", "과목", "학급"]

def make_empty_frames():
    ti = pd.DataFrame(columns=["번호", "교사명", "담당과목", "과목군", "담당학년",
                               "주당시수", "담임학급", "비고"])
    tt = pd.DataFrame(columns=["교사명", "요일", "교시", "과목", "학급", "과목군"])
    pt_cols = ["번호", "시간강사명", "담당과목", "과목군", "비고"] + \
              [f"{d}{p}" for d in DAYS for p in range(1, 8)]
    pt = pd.DataFrame(columns=pt_cols)
    return ti, tt, pt

def normalize_frames(ti: pd.DataFrame, tt: pd.DataFrame):
    tt = tt.copy()
    for c in REQUIRED_TT_COLS:
        if c not in tt.columns:
            raise ValueError(f"[시간표] 시트에 '{c}' 열이 없습니다.")
    tt["교시"] = pd.to_numeric(tt["교시"], errors="coerce").fillna(0).astype(int)
    for c in ["교사명", "요일", "과목", "학급"]:
        tt[c] = tt[c].astype(str).str.strip()
    if "과목군" not in tt.columns or tt["과목군"].isna().all():
        tt["과목군"] = tt["과목"].map(subject_group)
    tt["과목군"] = tt["과목군"].fillna("").astype(str)
    tt.loc[tt["과목군"] == "", "과목군"] = tt.loc[tt["과목군"] == "", "과목"].map(subject_group)
    tt = tt[tt["요일"].isin(DAYS)]
    tt = tt[(tt["교시"] >= 1) & (tt["교시"] <= MAX_PERIOD)]
    tt = tt.drop_duplicates(subset=["교사명", "요일", "교시"]).reset_index(drop=True)

    ti = ti.copy()
    if "교사명" not in ti.columns:
        ti = pd.DataFrame({"교사명": sorted(tt["교사명"].unique())})
    ti["교사명"] = ti["교사명"].astype(str).str.strip()
    for c in ["번호", "담당과목", "과목군", "담당학년", "주당시수", "담임학급", "비고"]:
        if c not in ti.columns:
            ti[c] = ""
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
        if str(row.get("주당시수", "")).strip() in ("", "nan"):
            ti.at[i, "주당시수"] = len(sub)
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

        if not absences.empty and "교시" in absences.columns:
            absences["교시"] = pd.to_numeric(absences["교시"], errors="coerce").fillna(0).astype(int)
        if not subs.empty and "교시" in subs.columns:
            subs["교시"] = pd.to_numeric(subs["교시"], errors="coerce").fillna(0).astype(int)
            
        return absences, subs, swaps, part_time
    except Exception as e:
        st.error(f"작업 내역 로드 실패: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

def save_work_data_to_gsheet():
    try:
        df_to_worksheet(get_worksheet(WORK_SHEET_ID, "결강"), st.session_state.absences)
        df_to_worksheet(get_worksheet(WORK_SHEET_ID, "보강"), st.session_state.subs)
        df_to_worksheet(get_worksheet(WORK_SHEET_ID, "맞교환"), st.session_state.swaps)
        df_to_worksheet(get_worksheet(WORK_SHEET_ID, "시간강사"), st.session_state.part_time)
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
    st.session_state.part_time = make_empty_frames()[2]

    absences, subs, swaps = load_work_data_from_gsheet()

    if absences.empty:
        absences = pd.DataFrame(columns=["결강ID", "일자", "요일", "교사명", "사유", "상세사유", "교시", "학급", "과목", "등록시각"])
    if subs.empty:
        subs = pd.DataFrame(columns=["결강ID", "일자", "요일", "교시", "학급", "과목",
                                     "결강교사", "보강교사", "배정방식", "우선순위", "비고", "등록시각"])
    if swaps.empty:
        swaps = pd.DataFrame(columns=["원본일자", "교사A", "요일A", "교시A", "학급A", "과목A",
                                      "목표일자", "교사B", "요일B", "교시B", "학급B", "과목B",
                                      "유형", "시간강사구인", "등록시각"])

    st.session_state.absences = absences
    st.session_state.subs = subs
    st.session_state.swaps = swaps

    init_history_if_needed()

init_state()

# ==========================================================================================
# 3. 핵심 로직
# ==========================================================================================
def is_teacher_available(teacher: str, day: str, period: int, is_part_time: bool = False) -> bool:
    source = st.session_state.part_time if is_part_time else st.session_state.teachers
    name_col = "시간강사명" if is_part_time else "교사명"
    if source.empty or name_col not in source.columns:
        return True
    row = source[source[name_col] == teacher]
    if row.empty:
        return True

    col_candidates = [
        f"{day}{period}", 
        f"{day}요일{period}", 
        f"{day}{period}교시", 
        f"{day}요일{period}교시",
        day, 
        f"{day}요일"
    ]
    
    for col in col_candidates:
        if col in source.columns:
            val = row[col].values[0]
            if pd.isna(val):
                continue
                
            str_val = str(val).strip()
            
            # 불가능으로 인식할 값들 (0, 불가 등)
            if str_val in ['0', '0.0', 'False', 'false', '불가', '불가능', 'N', 'n', 'X', 'x', '없음']:
                return False
                
            try:
                num_val = pd.to_numeric(val, errors='coerce')
                if pd.notna(num_val) and num_val <= 0:
                    return False
            except:
                pass
                
    return True

def lesson_of(teacher: str, day: str, period: int):
    tt = st.session_state.timetable
    m = tt[(tt["교사명"] == teacher) & (tt["요일"] == day) & (tt["교시"] == int(period))]
    return None if m.empty else m.iloc[0].to_dict()

def is_free(teacher: str, day: str, period: int, on_date: str = None, is_part_time: bool = False) -> bool:
    if not is_teacher_available(teacher, day, period, is_part_time):
        return False
    if not is_part_time and lesson_of(teacher, day, period) is not None:
        return False
    if on_date and not is_part_time:
        s = st.session_state.subs
        if not s.empty and ((s["일자"] == on_date) & (s["교시"] == int(period)) & (s["보강교사"] == teacher)).any():
            return False
        a = st.session_state.absences
        if not a.empty and ((a["일자"] == on_date) & (a["교사명"] == teacher) & (a["교시"] == int(period))).any():
            return False
    return True

def is_class_free(class_name: str, day: str, period: int) -> bool:
    tt = st.session_state.timetable
    return tt[(tt["학급"] == class_name) & (tt["요일"] == day) & (tt["교시"] == int(period))].empty

def can_teacher_take_slot(teacher: str, day: str, period: int, on_date: str = None, is_part_time: bool = False) -> bool:
    return is_teacher_available(teacher, day, period, is_part_time) and is_free(teacher, day, period, on_date, is_part_time)

def absent_all_day(teacher: str, on_date: str) -> bool:
    a = st.session_state.absences
    if a.empty:
        return False
    return ((a["일자"] == on_date) & (a["교사명"] == teacher)).any()

def cumulative_sub_count(start_date: str = None, end_date: str = None) -> dict:
    s = st.session_state.subs
    base = {t: 0 for t in st.session_state.teachers["교사명"].tolist()} if not st.session_state.teachers.empty else {}
    if not s.empty:
        if start_date and end_date:
            s = s[(s["일자"] >= start_date) & (s["일자"] <= end_date)]
        for k, v in s["보강교사"].value_counts().items():
            base[k] = int(v)
    return base

def weekly_load() -> dict:
    tt = st.session_state.timetable
    return tt["교사명"].value_counts().to_dict() if not tt.empty else {}

def recommend_substitutes(day: str, period: int, subject: str, class_name: str,
                          absent_teacher: str, on_date: str, top_n: int = 12,
                          include_part_time: bool = False) -> pd.DataFrame:
    teachers = st.session_state.teachers
    part = st.session_state.part_time
    tt = st.session_state.timetable
    if teachers.empty:
        return pd.DataFrame()

    grp = subject_group(subject)
    grade = grade_of(class_name)
    cum = cumulative_sub_count()
    load = weekly_load()
    max_cum = max(cum.values()) if cum else 0
    rows = []

    for t in teachers["교사명"].tolist():
        if t == absent_teacher or absent_all_day(t, on_date):
            continue
        if not is_free(t, day, period, on_date):
            continue
        sub_tt = tt[tt["교사명"] == t]
        my_groups = set(sub_tt["과목군"]) if not sub_tt.empty else set()
        my_grades = {grade_of(c) for c in sub_tt["학급"]} if not sub_tt.empty else set()
        my_classes = set(sub_tt["학급"]) if not sub_tt.empty else set()

        if grp in my_groups:
            prio, prio_label, score = 1, "1순위 · 동일 과목", 100
        elif grade in my_grades:
            prio, prio_label, score = 2, "2순위 · 동일 학년", 60
        else:
            prio, prio_label, score = 3, "3순위 · 전체 공강", 20

        c = cum.get(t, 0)
        score += (max_cum - c) * 6 + max(0, (22 - load.get(t, 0))) * 0.8
        if class_name in my_classes:
            score += 5

        t_row = teachers[teachers["교사명"] == t]
        rows.append({
            "보강교사": t, "유형": "정규교사", "우선순위": prio_label, "_prio": prio,
            "담당과목": t_row["담당과목"].iloc[0] if not t_row.empty else "",
            "주당시수": int(load.get(t, 0)), "누적보강": c, "추천점수": round(score, 1)
        })

    if include_part_time and not part.empty and "시간강사명" in part.columns:
        for t in part["시간강사명"].dropna().tolist():
            if not is_free(t, day, period, on_date, is_part_time=True):
                continue
            t_row = part[part["시간강사명"] == t]
            rows.append({
                "보강교사": t, "유형": "시간강사", "우선순위": "시간강사", "_prio": 4,
                "담당과목": t_row["담당과목"].iloc[0] if not t_row.empty else "",
                "주당시수": 0, "누적보강": 0, "추천점수": 40
            })

    df = pd.DataFrame(rows)
    if df.empty:
        return df
    return df.sort_values(["_prio", "추천점수"], ascending=[True, False]).drop(columns=["_prio"]).head(top_n).reset_index(drop=True)

def auto_assign_all(cid: str, on_date: str, day: str, rows: pd.DataFrame, include_pt: bool = False) -> list:
    log = []
    for _, r in rows.iterrows():
        cand = recommend_substitutes(day, int(r["교시"]), r["과목"], r["학급"],
                                     r["교사명"], on_date, top_n=1, include_part_time=include_pt)
        if cand.empty:
            log.append((int(r["교시"]), None, "배정 가능한 공강 교사가 없습니다."))
            continue
        pick = cand.iloc[0]
        add_substitute(cid, on_date, day, int(r["교시"]), r["학급"], r["과목"],
                       r["교사명"], pick["보강교사"], "자동배정", pick["우선순위"], "")
        log.append((int(r["교시"]), pick["보강교사"], pick["우선순위"]))
    return log

def add_substitute(cid, on_date, day, period, class_name, subject,
                   absent_teacher, sub_teacher, method, priority, memo):
    push_history(f"보강 배정 ({sub_teacher})")
    s = st.session_state.subs
    if not s.empty:
        s = s[~((s["결강ID"] == cid) & (s["교시"] == int(period)))]
    new = {
        "결강ID": cid, "일자": on_date, "요일": day, "교시": int(period),
        "학급": class_name, "과목": subject, "결강교사": absent_teacher,
        "보강교사": sub_teacher, "배정방식": method, "우선순위": priority,
        "비고": memo, "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M"),
    }
    st.session_state.subs = pd.concat([s, pd.DataFrame([new])], ignore_index=True)
    save_work_data_to_gsheet()

def check_conflicts(tt: pd.DataFrame) -> pd.DataFrame:
    issues = []
    if tt.empty:
        return pd.DataFrame(columns=["유형", "요일", "교시", "대상", "내용", "해결 가이드"])
    for (d, p, c), grp in tt.groupby(["요일", "교시", "학급"]):
        if len(grp) > 1:
            issues.append({"유형": "학급 중복", "요일": d, "교시": p, "대상": c,
                           "내용": " / ".join(f"{r['교사명']}({r['과목']})" for _, r in grp.iterrows()),
                           "해결 가이드": f"{c} 학급에 동일 시간 중복 수업이 존재합니다."})
    for (d, p, t), grp in tt.groupby(["요일", "교시", "교사명"]):
        if len(grp) > 1:
            issues.append({"유형": "교사 중복", "요일": d, "교시": p, "대상": t,
                           "내용": " / ".join(f"{r['학급']}({r['과목']})" for _, r in grp.iterrows()),
                           "해결 가이드": f"{t} 교사가 동시간대 중복 입력되었습니다."})
    for _, r in tt.iterrows():
        if not is_teacher_available(r["교사명"], r["요일"], int(r["교시"])):
            issues.append({"유형": "불가능 시간 배치", "요일": r["요일"], "교시": int(r["교시"]), "대상": r["교사명"],
                           "내용": f"{r['학급']} {r['과목']} 수업이 불가능 시간에 설정됨",
                           "해결 가이드": f"{r['교사명']} 교사의 근무 가능 상태를 확인하세요."})
    return pd.DataFrame(issues)

def validate_swap(a: dict, b: dict):
    errs, guides = [], []
    tt = st.session_state.timetable
    if a["교사명"] == b["교사명"]:
        errs.append("동일한 교사의 수업끼리는 맞교환할 수 없습니다.")
        guides.append("💡 다른 교사와의 수업 맞교환을 선택하세요.")
    if not is_teacher_available(a["교사명"], b["요일"], b["교시"]):
        errs.append(f"{a['교사명']} 교사는 {b['요일']} {b['교시']}교시가 불가능 시간입니다.")
        guides.append(f"💡 {a['교사명']} 교사의 불가능 시간을 가능으로 변경하거나 다른 시간을 선택하세요.")
    if not is_teacher_available(b["교사명"], a["요일"], a["교시"]):
        errs.append(f"{b['교사명']} 교사는 {a['요일']} {a['교시']}교시가 불가능 시간입니다.")
        guides.append(f"💡 {b['교사명']} 교사의 불가능 시간을 가능으로 변경하거나 다른 시간을 선택하세요.")
    other_a = tt[(tt["교사명"] == a["교사명"]) & (tt["요일"] == b["요일"]) & (tt["교시"] == int(b["교시"]))]
    other_a = other_a[~((other_a["요일"] == a["요일"]) & (other_a["교시"] == int(a["교시"])))]
    if not other_a.empty:
        r = other_a.iloc[0]
        errs.append(f"{a['교사명']} 교사는 {b['요일']} {b['교시']}교시에 이미 {r['학급']} {r['과목']} 수업이 있습니다.")
        guides.append(f"💡 {a['교사명']} 교사의 동시간대 중복 수업 일정을 조정하세요.")
    other_b = tt[(tt["교사명"] == b["교사명"]) & (tt["요일"] == a["요일"]) & (tt["교시"] == int(a["교시"]))]
    other_b = other_b[~((other_b["요일"] == b["요일"]) & (other_b["교시"] == int(b["교시"])))]
    if not other_b.empty:
        r = other_b.iloc[0]
        errs.append(f"{b['교사명']} 교사는 {a['요일']} {a['교시']}교시에 이미 {r['학급']} {r['과목']} 수업이 있습니다.")
        guides.append(f"💡 {b['교사명']} 교사의 동시간대 중복 수업 일정을 조정하세요.")
    cls_a = tt[(tt["학급"] == a["학급"]) & (tt["요일"] == b["요일"]) & (tt["교시"] == int(b["교시"]))]
    cls_a = cls_a[~((cls_a["교사명"] == b["교사명"]) & (cls_a["요일"] == b["요일"]) & (cls_a["교시"] == int(b["교시"])))]
    if not cls_a.empty:
        r = cls_a.iloc[0]
        errs.append(f"{a['학급']} 학급은 {b['요일']} {b['교시']}교시에 이미 {r['과목']}({r['교사명']}) 수업이 지정되어 있습니다.")
        guides.append(f"💡 {a['학급']} 학급의 해당 시간 수업 담당자와 조정하세요.")
    cls_b = tt[(tt["학급"] == b["학급"]) & (tt["요일"] == a["요일"]) & (tt["교시"] == int(a["교시"]))]
    cls_b = cls_b[~((cls_b["교사명"] == a["교사명"]) & (cls_b["요일"] == a["요일"]) & (cls_b["교시"] == int(a["교시"])))]
    if not cls_b.empty:
        r = cls_b.iloc[0]
        errs.append(f"{b['학급']} 학급은 {a['요일']} {a['교시']}교시에 이미 {r['과목']}({r['교사명']}) 수업이 지정되어 있습니다.")
        guides.append(f"💡 {b['학급']} 학급의 해당 시간 수업 담당자와 조정하세요.")
    return errs, guides

def get_target_time_recommendations(teacher_a, date_a_str, period_a, class_a, subject_a, date_b_str, period_b):
    ti = st.session_state.teachers
    tt = st.session_state.timetable
    day_a = WEEKDAY_KR[datetime.strptime(date_a_str, "%Y-%m-%d").weekday()]
    day_b = WEEKDAY_KR[datetime.strptime(date_b_str, "%Y-%m-%d").weekday()]
    lesson_a = {"교사명": teacher_a, "일자": date_a_str, "요일": day_a, "교시": period_a, "학급": class_a, "과목": subject_a}
    grp_a = subject_group(subject_a)
    grade_a = grade_of(class_a)
    cum = cumulative_sub_count()
    load = weekly_load()
    a_can_go_to_y = can_teacher_take_slot(teacher_a, day_b, period_b, date_b_str)
    class_free_at_y = is_class_free(class_a, day_b, period_b)
    swap_recs, linked_recs, sub_recs = [], [], []

    for t_b in ti["교사명"].tolist():
        if t_b == teacher_a or absent_all_day(t_b, date_b_str):
            continue
        b_lessons = tt[(tt["교사명"] == t_b) & (tt["요일"] == day_b) & (tt["교시"] == period_b)]
        if not b_lessons.empty:
            for _, b_row in b_lessons.iterrows():
                b_item = {"교사명": t_b, "일자": date_b_str, "요일": day_b, "교시": period_b,
                          "학급": b_row["학급"], "과목": b_row["과목"]}
                errs, guides = validate_swap(lesson_a, b_item)
                is_valid = len(errs) == 0
                score = 2000 if is_valid else -1000
                if is_valid:
                    if subject_group(b_row["과목"]) == grp_a: score += 300
                    if grade_of(b_row["학급"]) == grade_a: score += 200
                    score -= cum.get(t_b, 0) * 15
                    score += max(0, (22 - load.get(t_b, 0))) * 2
                    prio_tag = "✅ [직접 1:1 맞교환 가능]"
                else:
                    prio_tag = f"❌ [충돌] {errs[0] if errs else '검증 실패'}"
                swap_recs.append({
                    "유형": "직접1:1", "교사B": t_b,
                    "현재 수업": f"{day_b}{period_b}교시 · {b_row['학급']} · {b_row['과목']}",
                    "상태": prio_tag, "is_valid": is_valid, "_score": score,
                    "b_info": b_item, "errs": errs, "guides": guides,
                    "설명": f"A({teacher_a})의 {day_a}{period_a} ↔ B({t_b})의 {day_b}{period_b}"
                })
        if (is_free(t_b, day_b, period_b, date_b_str) and a_can_go_to_y and class_free_at_y and
            can_teacher_take_slot(t_b, day_a, period_a, date_a_str)):
            tb_info = ti[ti["교사명"] == t_b].iloc[0]
            tb_grp = str(tb_info.get("과목군", ""))
            tb_grade = str(tb_info.get("담당학년", ""))
            score = 1500
            if grp_a and grp_a in tb_grp:
                score += 250
                prio_label = "1순위 · 동일 과목군 연계"
            elif grade_a and grade_a in tb_grade:
                score += 150
                prio_label = "2순위 · 동일 학년 연계"
            else:
                prio_label = "3순위 · 공강 연계"
            score -= cum.get(t_b, 0) * 12
            score += max(0, (22 - load.get(t_b, 0))) * 1.5
            linked_recs.append({
                "유형": "연계공강교환", "교사B": t_b, "현재 수업": f"{day_b}{period_b}교시 공강",
                "상태": f"✅ [{prio_label}]", "is_valid": True, "_score": score,
                "b_info": {"교사명": t_b, "일자": date_b_str, "요일": day_b, "교시": period_b,
                           "학급": class_a, "과목": subject_a},
                "errs": [], "guides": [],
                "설명": f"A({teacher_a}) → {day_b}{period_b}로 이동, B({t_b})가 A의 원래 {day_a}{period_a} 담당"
            })
        elif is_free(t_b, day_b, period_b, date_b_str) and a_can_go_to_y and class_free_at_y:
            tb_info = ti[ti["교사명"] == t_b].iloc[0]
            tb_grp = str(tb_info.get("과목군", ""))
            tb_grade = str(tb_info.get("담당학년", ""))
            if grp_a and grp_a in tb_grp:
                prio_label, prio_num, score = "1순위 · 동일 과목군 대리", 1, 800
            elif grade_a and grade_a in tb_grade:
                prio_label, prio_num, score = "2순위 · 동일 학년 대리", 2, 500
            else:
                prio_label, prio_num, score = "3순위 · 일반 공강 대리", 3, 200
            score -= cum.get(t_b, 0) * 10
            sub_recs.append({
                "보강/대리 교사": t_b, "우선순위": prio_label,
                "담당과목": tb_info.get("담당과목", ""),
                "누적보강": cum.get(t_b, 0), "_prio": prio_num, "_score": score
            })

    df_swap = pd.DataFrame(swap_recs)
    if not df_swap.empty:
        df_swap = df_swap.sort_values("_score", ascending=False).reset_index(drop=True)
    df_linked = pd.DataFrame(linked_recs)
    if not df_linked.empty:
        df_linked = df_linked.sort_values("_score", ascending=False).reset_index(drop=True)
    df_sub = pd.DataFrame(sub_recs)
    if not df_sub.empty:
        df_sub = df_sub.sort_values(["_prio", "_score"], ascending=[True, False]).reset_index(drop=True)
    return df_swap, df_linked, df_sub

def do_swap(a: dict, b: dict, date_a: str, date_b: str, is_part_time_purpose: bool = False):
    push_history(f"1:1 맞교환 ({a['교사명']} ↔ {b['교사명']})")
    rec = {
        "원본일자": date_a, "교사A": a["교사명"], "요일A": a["요일"], "교시A": int(a["교시"]),
        "학급A": a["학급"], "과목A": a["과목"],
        "목표일자": date_b, "교사B": b["교사명"], "요일B": b["요일"], "교시B": int(b["교시"]),
        "학급B": b["학급"], "과목B": b["과목"],
        "유형": "1:1 맞교환",
        "시간강사구인": "Y" if is_part_time_purpose else "N",
        "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    st.session_state.swaps = pd.concat([st.session_state.swaps, pd.DataFrame([rec])], ignore_index=True)
    save_work_data_to_gsheet()
    return True

def do_linked_swap(a: dict, teacher_b: str, date_a: str, date_b: str, day_b: str, period_b: int,
                   is_part_time_purpose: bool = False):
    push_history(f"연계 교환 ({a['교사명']} → {teacher_b})")
    rec = {
        "원본일자": date_a, "교사A": a["교사명"], "요일A": a["요일"], "교시A": int(a["교시"]),
        "학급A": a["학급"], "과목A": a["과목"],
        "목표일자": date_b, "교사B": teacher_b, "요일B": day_b, "교시B": int(period_b),
        "학급B": a["학급"], "과목B": a["과목"],
        "유형": "연계 공강 교환",
        "시간강사구인": "Y" if is_part_time_purpose else "N",
        "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M")
    }
    st.session_state.swaps = pd.concat([st.session_state.swaps, pd.DataFrame([rec])], ignore_index=True)
    save_work_data_to_gsheet()
    return True

def teacher_matrix() -> pd.DataFrame:
    tt = st.session_state.timetable
    cols = [f"{d}{p}" for d in DAYS for p in range(1, PERIODS_PER_DAY[d] + 1)]
    idx = list(st.session_state.teachers["교사명"])
    mat = pd.DataFrame("", index=idx, columns=cols)
    for _, r in tt.iterrows():
        key = f"{r['요일']}{int(r['교시'])}"
        if key in cols and r["교사명"] in mat.index:
            mat.at[r["교사명"], key] = f"{r['과목']} {r['학급']}"
    return mat

def class_matrix() -> pd.DataFrame:
    tt = st.session_state.timetable
    cols = [f"{d}{p}" for d in DAYS for p in range(1, PERIODS_PER_DAY[d] + 1)]
    classes = sorted([c for c in tt["학급"].unique() if "-" in str(c)])
    mat = pd.DataFrame("", index=classes, columns=cols)
    for _, r in tt.iterrows():
        key = f"{r['요일']}{int(r['교시'])}"
        if key in cols and r["학급"] in mat.index:
            cur = mat.at[r["학급"], key]
            val = f"{r['과목']}({r['교사명']})"
            mat.at[r["학급"], key] = val if cur == "" else cur + " / " + val
    return mat

def get_teacher_week_view(teacher: str, ref_date: date):
    weekday = ref_date.weekday()
    monday = ref_date - timedelta(days=weekday)
    week_dates = [monday + timedelta(days=i) for i in range(5)]

    grid = pd.DataFrame("", index=[f"{p}교시" for p in range(1, MAX_PERIOD + 1)], columns=DAYS)
    base = st.session_state.timetable[st.session_state.timetable["교사명"] == teacher]
    swaps = st.session_state.swaps

    for d_idx, d in enumerate(DAYS):
        for p in range(1, PERIODS_PER_DAY.get(d, 7) + 1):
            les = base[(base["요일"] == d) & (base["교시"] == p)]
            if not les.empty:
                grid.at[f"{p}교시", d] = f"{les.iloc[0]['과목']} {les.iloc[0]['학급']}"

    if not swaps.empty:
        for d_idx, cur_date in enumerate(week_dates):
            cur_date_str = cur_date.strftime("%Y-%m-%d")
            d = DAYS[d_idx]

            outgoing = swaps[swaps["원본일자"] == cur_date_str]
            for _, sw in outgoing.iterrows():
                if sw["교사A"] == teacher:
                    p = int(sw["교시A"])
                    grid.at[f"{p}교시", d] = "→이동됨"
                if sw["유형"] == "1:1 맞교환" and sw["교사B"] == teacher:
                    p = int(sw["교시B"])
                    grid.at[f"{p}교시", d] = "→이동됨"

            incoming = swaps[swaps["목표일자"] == cur_date_str]
            for _, sw in incoming.iterrows():
                if sw["교사A"] == teacher:
                    p = int(sw["교시B"])
                    grid.at[f"{p}교시", d] = f"🔀{sw['과목A']} {sw['학급A']}"
                if sw["유형"] == "1:1 맞교환" and sw["교사B"] == teacher:
                    p = int(sw["교시A"])
                    grid.at[f"{p}교시", d] = f"🔀{sw['과목B']} {sw['학급B']}"

            linked_out = outgoing[outgoing["유형"] == "연계 공강 교환"]
            for _, sw in linked_out.iterrows():
                if sw["교사B"] == teacher:
                    p = int(sw["교시A"])
                    grid.at[f"{p}교시", d] = f"🔗{sw['과목A']} {sw['학급A']}"

    for d_idx, cur_date in enumerate(week_dates):
        cur_date_str = cur_date.strftime("%Y-%m-%d")
        d = DAYS[d_idx]

        abs_m = st.session_state.absences
        if not abs_m.empty:
            m = abs_m[(abs_m["일자"] == cur_date_str) & (abs_m["교사명"] == teacher)]
            for _, r in m.iterrows():
                p = int(r["교시"])
                cur_val = str(grid.at[f"{p}교시", d])
                if "이동" in cur_val or cur_val == "":
                    grid.at[f"{p}교시", d] = "🚫결강"
                else:
                    grid.at[f"{p}교시", d] = f"🚫결강 {cur_val}"

        sub_m = st.session_state.subs
        if not sub_m.empty:
            m = sub_m[(sub_m["일자"] == cur_date_str) & (sub_m["보강교사"] == teacher)]
            for _, r in m.iterrows():
                p = int(r["교시"])
                grid.at[f"{p}교시", d] = f"✅보강 {r['학급']} {r['과목']}"

    return grid, week_dates

# ==========================================================================================
# 4. 내역서 / 엑셀
# ==========================================================================================
def build_report_html(on_date: str) -> str:
    day = WEEKDAY_KR[datetime.strptime(on_date, "%Y-%m-%d").weekday()]
    a = st.session_state.absences
    s = st.session_state.subs
    w = st.session_state.swaps
    a = a[a["일자"] == on_date] if not a.empty else a
    s = s[s["일자"] == on_date] if not s.empty else s
    w = w[(w["원본일자"] == on_date) | (w["목표일자"] == on_date)] if not w.empty else w

    def rows_abs():
        if a.empty:
            return "<tr><td colspan='6' class='empty'>해당 일자의 결강 등록 내역이 없습니다.</td></tr>"
        out = []
        for _, r in a.sort_values(["교사명", "교시"]).iterrows():
            out.append(f"<tr><td>{r['교사명']}</td><td>{r['사유']}</td><td>{r['교시']}교시</td>"
                       f"<td>{r['학급']}</td><td>{r['과목']}</td><td>{r.get('상세사유','') or '-'}</td></tr>")
        return "".join(out)

    def rows_sub():
        if s.empty:
            return "<tr><td colspan='7' class='empty'>보강 배정 내역이 없습니다.</td></tr>"
        out = []
        for _, r in s.sort_values(["교시", "학급"]).iterrows():
            out.append(f"<tr><td>{r['교시']}교시</td><td>{r['학급']}</td><td>{r['과목']}</td>"
                       f"<td>{r['결강교사']}</td><td class='hl'>{r['보강교사']}</td>"
                       f"<td>{r['우선순위']}</td><td>{r.get('비고','') or '-'}</td></tr>")
        return "".join(out)

    def rows_swap():
        if w.empty:
            return "<tr><td colspan='4' class='empty'>시간표 맞교환 내역이 없습니다.</td></tr>"
        out = []
        for _, r in w.iterrows():
            pt_flag = " (시간강사 구인)" if r.get("시간강사구인") == "Y" else ""
            out.append(f"<tr><td>{r['교사A']}</td>"
                       f"<td>[{r['원본일자']}] {r['요일A']} {r['교시A']}교시 {r['학급A']} {r['과목A']}"
                       f" ↔ [{r['목표일자']}] {r['요일B']} {r['교시B']}교시 {r['학급B']} {r['과목B']}{pt_flag}</td>"
                       f"<td>{r['교사B']}</td><td>{r.get('유형','')}</td></tr>")
        return "".join(out)

    return f"""<!doctype html><html lang="ko"><head><meta charset="utf-8">
<title>{on_date} 결강·보강 변경 내역서</title>
<style>
 @page {{ size: A4 portrait; margin: 15mm; }}
 body {{ font-family:'맑은 고딕','Malgun Gothic',sans-serif; color:#111; font-size:12px; }}
 h1 {{ text-align:center; font-size:21px; letter-spacing:6px; margin:0 0 4px; }}
 .sub {{ text-align:center; color:#555; margin-bottom:14px; }}
 h2 {{ font-size:13px; margin:16px 0 6px; border-left:4px solid #333; padding-left:7px; }}
 table {{ width:100%; border-collapse:collapse; }}
 th,td {{ border:1px solid #999; padding:5px 6px; text-align:center; }}
 th {{ background:#eef1f6; }}
 td.hl {{ font-weight:700; }}
 td.empty {{ color:#888; }}
 .sign {{ margin-top:22px; width:62%; margin-left:auto; }}
 .sign td {{ height:52px; }}
 .foot {{ margin-top:12px; font-size:11px; color:#666; text-align:right; }}
</style></head><body>
<h1>결강·보강 변경 내역서</h1>
<div class="sub">{SCHOOL_YEAR}학년도 · {SCHOOL_NAME} · {on_date} ({day})</div>
<h2>1. 결강 현황</h2>
<table><tr><th>결강 교사</th><th>사유</th><th>교시</th><th>학급</th><th>과목</th><th>상세</th></tr>{rows_abs()}</table>
<h2>2. 보강 배정</h2>
<table><tr><th>교시</th><th>학급</th><th>과목</th><th>결강 교사</th><th>보강 교사</th><th>배정 근거</th><th>비고</th></tr>{rows_sub()}</table>
<h2>3. 시간표 맞교환 및 변경</h2>
<table><tr><th>교사 A</th><th>교환 내용</th><th>교사 B</th><th>유형</th></tr>{rows_swap()}</table>
<table class="sign"><tr><th>담당</th><th>교육과정부장</th><th>교감</th><th>교장</th></tr><tr><td></td><td></td><td></td><td></td></tr></table>
<div class="foot">출력 {datetime.now().strftime('%Y-%m-%d %H:%M')} · 본 문서는 시간표 관리 프로그램에서 자동 생성되었습니다.</div>
</body></html>"""

def to_excel_bytes(sheets: dict) -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        for name, df in sheets.items():
            (df if isinstance(df, pd.DataFrame) else pd.DataFrame(df)).to_excel(w, sheet_name=name[:31], index=False)
    return buf.getvalue()

# ==========================================================================================
# 5. 사이드바
# ==========================================================================================
with st.sidebar:
    st.header("데이터 (Google Sheets)")
    
    if st.button("🔄 원본 시간표 다시 불러오기", use_container_width=True):
        load_timetable_from_gsheet.clear()
        ti, tt = load_timetable_from_gsheet()
        st.session_state.teachers = ti
        st.session_state.timetable = tt
        st.success("원본 시간표를 다시 불러왔습니다.")
        st.rerun()

    if st.button("🔄 작업 내역 다시 불러오기", use_container_width=True):
        load_work_data_from_gsheet.clear()
        absences, subs, swaps = load_work_data_from_gsheet()
        st.session_state.absences = absences
        st.session_state.subs = subs
        st.session_state.swaps = swaps
        st.success("작업 내역을 다시 불러왔습니다.")
        st.rerun()

    if st.button("💾 현재 작업 내역 저장", use_container_width=True, type="primary"):
        if save_work_data_to_gsheet():
            st.success("Google Sheets에 저장 완료!")

    st.divider()
    st.subheader("작업 되돌리기")
    c_undo, c_redo = st.columns(2)
    if c_undo.button("↩ 되돌리기 (Undo)", use_container_width=True):
        if undo():
            save_work_data_to_gsheet()
            st.success("이전 단계로 되돌렸습니다.")
            st.rerun()
        else:
            st.warning("더 이상 되돌릴 수 없습니다.")
    if c_redo.button("↪ 다시실행 (Redo)", use_container_width=True):
        if redo():
            save_work_data_to_gsheet()
            st.success("다시 실행했습니다.")
            st.rerun()
        else:
            st.warning("다시 실행할 단계가 없습니다.")

    if "history" in st.session_state and st.session_state.history:
        st.caption(f"히스토리: {st.session_state.history_index + 1} / {len(st.session_state.history)}")
        with st.expander("최근 작업 목록"):
            for i, h in enumerate(reversed(st.session_state.history[-8:])):
                st.write(f"{h['time']} - {h['action']}")

    st.divider()
    st.caption(f"{SCHOOL_YEAR}학년도 · {SCHOOL_NAME}")
    st.metric("등록 교사", f"{len(st.session_state.teachers)} 명")
    st.metric("시간강사", f"{len(st.session_state.part_time)} 명")
    st.metric("주간 수업", f"{len(st.session_state.timetable)} 시수")
    st.metric("누적 보강", f"{len(st.session_state.subs)} 건")

if st.session_state.timetable.empty:
    st.title("중학교 시간표·결보강 관리 프로그램")
    st.info("Google Sheets에서 시간표를 불러오는 중입니다... 잠시만 기다려 주세요.\n\n문제가 있으면 사이드바의 '원본 시간표 다시 불러오기'를 눌러보세요.")
    st.stop()

# ==========================================================================================
# 6. 메인 화면
# ==========================================================================================
st.title("전체 시간표 관리 · 결강/보강 자동 처리 (Google Sheets 연동)")

tabs = st.tabs([
    "시간표 조회", "기본정보·시간표 편집", "시간강사 관리",
    "결강 등록", "보강 배정", "시간표 맞교환 & 변경 추천",
    "일일 변경 내역서", "통계 (학기/월별)"
])

# ------------------------------------------------------------------ 시간표 조회
with tabs[0]:
    view = st.radio("보기 방식", ["교사별 전체 매트릭스 (기본 주간)", "학급별 매트릭스 (기본 주간)", "교사 1인 주간표 (달력 – 주차 독립)"], horizontal=True)
    if view == "교사별 전체 매트릭스 (기본 주간)":
        st.dataframe(teacher_matrix(), use_container_width=True, height=760)
    elif view == "학급별 매트릭스 (기본 주간)":
        st.dataframe(class_matrix(), use_container_width=True, height=680)
    else:
        t = st.selectbox("교사 선택", st.session_state.teachers["교사명"])
        ref_date = st.date_input("기준 날짜 (해당 주 월~금 표시)", value=date.today(), key="week_ref")
        grid, week_dates = get_teacher_week_view(t, ref_date)
        st.caption(f"주간: {week_dates[0].strftime('%Y-%m-%d')} ~ {week_dates[4].strftime('%Y-%m-%d')}  ※ 이 주차의 변경만 반영됩니다")
        info = st.session_state.teachers[st.session_state.teachers["교사명"] == t].iloc[0]
        c = st.columns(4)
        c[0].metric("담당 과목", str(info.get("담당과목", "")) or "-")
        c[1].metric("담임 학급", str(info.get("담임학급", "")) or "-")
        c[2].metric("주당 시수", int(len(st.session_state.timetable[st.session_state.timetable['교사명'] == t])))
        c[3].metric("누적 보강", cumulative_sub_count().get(t, 0))
        st.dataframe(grid, use_container_width=True)

# ------------------------------------------------------------------ 기본정보 편집
with tabs[1]:
    st.subheader("교사 정보 (불가능: 불가 / 가능: 가능)")
    ed_t = st.data_editor(st.session_state.teachers, num_rows="dynamic", use_container_width=True, height=320, key="ed_teachers")
    st.subheader("주간 시간표 (1행 = 1수업) – 기본 템플릿")
    only = st.selectbox("교사 필터", ["(전체)"] + list(st.session_state.teachers["교사명"]))
    tt_view = st.session_state.timetable if only == "(전체)" else st.session_state.timetable[st.session_state.timetable["교사명"] == only]
    ed_tt = st.data_editor(tt_view, num_rows="dynamic", use_container_width=True, height=420, key="ed_tt",
                           column_config={
                               "요일": st.column_config.SelectboxColumn(options=DAYS),
                               "교시": st.column_config.NumberColumn(min_value=1, max_value=MAX_PERIOD, step=1),
                           })
    b1, b2 = st.columns(2)
    if b1.button("변경 내용 적용", type="primary", use_container_width=True):
        st.session_state.teachers = ed_t.reset_index(drop=True)
        if only == "(전체)":
            new_tt = ed_tt.copy()
        else:
            keep = st.session_state.timetable[st.session_state.timetable["교사명"] != only]
            new_tt = pd.concat([keep, ed_tt], ignore_index=True)
        new_tt["교시"] = pd.to_numeric(new_tt["교시"], errors="coerce").fillna(0).astype(int)
        new_tt["과목군"] = new_tt["과목"].map(subject_group)
        st.session_state.timetable = new_tt.reset_index(drop=True)
        st.success("기본 주간 템플릿이 적용되었습니다. (원본 시트는 수정되지 않습니다)")
        st.rerun()
    if b2.button("시공간 충돌 및 불가능 시간 검사", use_container_width=True):
        iss = check_conflicts(st.session_state.timetable)
        if iss.empty:
            st.success("충돌 없음")
        else:
            st.error(f"충돌 {len(iss)}건")
            st.dataframe(iss, use_container_width=True)

# ------------------------------------------------------------------ 시간강사 관리
with tabs[2]:
    st.subheader("시간강사 등록 및 가능 시간표 (불가 = 불가능 / 가능 = 가능)")
    ed_pt = st.data_editor(st.session_state.part_time, num_rows="dynamic", use_container_width=True, height=400, key="ed_part")
    if st.button("시간강사 정보 저장", type="primary"):
        st.session_state.part_time = ed_pt.reset_index(drop=True)
        st.success("저장되었습니다.")
        st.rerun()

# ------------------------------------------------------------------ 결강 등록
with tabs[3]:
    c = st.columns([1, 1, 1, 2])
    d_sel = c[0].date_input("결강 일자", value=date.today())
    on_date = d_sel.strftime("%Y-%m-%d")
    day = WEEKDAY_KR[d_sel.weekday()]
    c[1].text_input("요일", value=day, disabled=True)
    who = c[2].selectbox("결강 교사", st.session_state.teachers["교사명"])
    reason = c[3].selectbox("사유", ABSENCE_REASONS)
    detail = st.text_input("상세 사유 / 비고 (선택)")

    if day not in DAYS:
        st.warning("주말은 시간표가 없습니다.")
    else:
        todays = st.session_state.timetable[
            (st.session_state.timetable["교사명"] == who) &
            (st.session_state.timetable["요일"] == day)
        ].sort_values("교시")

        if todays.empty:
            st.info(f"{who} 교사는 {day}요일에 수업이 없습니다.")
        else:
            opts = [f"{int(r['교시'])}교시 · {r['학급']} · {r['과목']}" for _, r in todays.iterrows()]
            picked = st.multiselect("결강 교시", opts, default=opts)

            if st.button("결강 등록", type="primary"):
                if not picked:
                    st.warning("교시를 선택하세요.")
                else:
                    try:
                        push_history(f"결강 등록 ({who})")
                        cid = f"{on_date}-{who}"
                        sel_periods = []
                        for o in picked:
                            try:
                                p = int(o.split("교시")[0].strip())
                                sel_periods.append(p)
                            except Exception:
                                continue
                        rows = todays[todays["교시"].isin(sel_periods)]
                        a = st.session_state.absences
                        if not a.empty:
                            a = a[~((a["일자"] == on_date) & (a["교사명"] == who))]
                        new_rows = []
                        for _, r in rows.iterrows():
                            new_rows.append({
                                "결강ID": cid, "일자": on_date, "요일": day, "교사명": who,
                                "사유": reason, "상세사유": detail, "교시": int(r["교시"]),
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
        show = ab[ab["일자"] == on_date] if st.checkbox("선택한 일자만 보기", value=True) else ab
        st.dataframe(show.sort_values(["일자", "교사명", "교시"]), use_container_width=True)
    else:
        st.info("등록된 결강이 없습니다.")

# ------------------------------------------------------------------ 보강 배정
with tabs[4]:
    ab = st.session_state.absences
    if ab.empty:
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
                st.markdown(f"### {head['일자']} ({head['요일']}) · **{head['교사명']}** · {head['사유']} · {len(rows)}시간")
                include_pt = st.checkbox("시간강사도 추천에 포함", value=False)

                if st.button("전 교시 자동 배정", type="primary"):
                    try:
                        push_history("자동 보강 배정")
                        log = auto_assign_all(cid, head["일자"], head["요일"], rows, include_pt=include_pt)
                        for p, t, why in log:
                            if t:
                                st.success(f"{p}교시 → {t} ({why})")
                            else:
                                st.error(f"{p}교시 → {why}")
                        st.rerun()
                    except Exception as e:
                        st.error(f"오류: {e}")

                st.divider()
                for _, r in rows.iterrows():
                    p = int(r["교시"])
                    with st.expander(f"{p}교시 · {r['학급']} · {r['과목']}", expanded=True):
                        cur = st.session_state.subs
                        assigned = None
                        if not cur.empty:
                            m = cur[(cur["결강ID"] == cid) & (cur["교시"] == p)]
                            if not m.empty:
                                assigned = m.iloc[0]["보강교사"]
                        if assigned:
                            st.success(f"현재 배정 : **{assigned}**")

                        try:
                            cand = recommend_substitutes(
                                head["요일"], p, r["과목"], r["학급"],
                                head["교사명"], head["일자"], top_n=12, include_part_time=include_pt
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
                                        cid, head["일자"], head["요일"], p, r["학급"], r["과목"],
                                        head["교사명"], pick, "수동배정", pr, memo
                                    )
                                    st.success(f"{p}교시 배정 완료.")
                                    st.rerun()
                                except Exception as e:
                                    st.error(f"오류: {e}")

# ------------------------------------------------------------------ 맞교환
with tabs[5]:
    st.markdown("### 🔄 스마트 시간표 변경 & 맞교환")
    st.caption("변경은 해당 날짜에만 적용됩니다.")
    
    col_a, col_b = st.columns(2)
    
    with col_a:
        st.markdown("#### 1️⃣ 원본 수업 (교사 A)")
        date_a_input = st.date_input("원본 날짜", value=date.today(), key="date_a")
        date_a_str = date_a_input.strftime("%Y-%m-%d")
        day_a = WEEKDAY_KR[date_a_input.weekday()]
        
        t_a = st.selectbox("교사 A", st.session_state.teachers["교사명"], key="teacher_a_select")
        
        sub_a = st.session_state.timetable[
            (st.session_state.timetable["교사명"] == t_a) & 
            (st.session_state.timetable["요일"] == day_a)
        ].sort_values("교시")
        
        if sub_a.empty:
            st.warning(f"{t_a} 교사는 해당 요일에 수업이 없습니다.")
            pick_a = None
        else:
            opts_a = [f"{int(r['교시'])}교시 · {r['학급']} · {r['과목']}" for _, r in sub_a.iterrows()]
            sel_a = st.selectbox("변경할 수업", opts_a, key="lesson_a_select")
            row_a = sub_a.iloc[opts_a.index(sel_a)]
            pick_a = {
                "교사명": t_a, "일자": date_a_str, "요일": day_a, 
                "교시": int(row_a["교시"]), "학급": row_a["학급"], "과목": row_a["과목"]
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
                    pr_val = df_sub_recs.loc[df_sub_recs["보강/대리 교사"] == sub_pick, "우선순위"].iloc[0]
                    add_substitute(cid, date_b_str, day_b, p_b_input, pick_a["학급"], pick_a["과목"],
                                   pick_a["교사명"], sub_pick, "시간표변경_대리", pr_val, f"{pick_a['일자']} 이동")
                    st.success(f"'{sub_pick}' 등록 완료")

    st.divider()
    st.subheader("변경 이력")
    st.dataframe(st.session_state.swaps, use_container_width=True)

# ------------------------------------------------------------------ 내역서
with tabs[6]:
    rd = st.date_input("출력 일자", value=date.today(), key="rdate").strftime("%Y-%m-%d")
    html = build_report_html(rd)
    st.download_button("HTML 다운로드", data=html.encode("utf-8"),
                       file_name=f"결보강내역서_{rd}.html", mime="text/html", type="primary")
    a = st.session_state.absences
    s = st.session_state.subs
    xls = to_excel_bytes({
        "결강": a[a["일자"] == rd] if not a.empty else pd.DataFrame(),
        "보강": s[s["일자"] == rd] if not s.empty else pd.DataFrame(),
        "시간표맞교환": st.session_state.swaps,
        "교사별시간표": teacher_matrix(),
        "학급별시간표": class_matrix(),
    })
    st.download_button("엑셀 백업 다운로드", data=xls,
                       file_name=f"결보강내역_{rd}.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.divider()
    st.components.v1.html(html, height=900, scrolling=True)

# ------------------------------------------------------------------ 통계
with tabs[7]:
    st.subheader("보강 통계 – 기간 필터")
    col1, col2, col3 = st.columns(3)
    start_d = col1.date_input("시작일", value=date(2026, 3, 1))
    end_d = col2.date_input("종료일", value=date.today())
    period_label = col3.selectbox("빠른 선택", ["전체", "1학기 (3~7월)", "2학기 (9~12월)", "이번 달"])

    if period_label == "1학기 (3~7월)":
        start_d, end_d = date(2026, 3, 1), date(2026, 7, 31)
    elif period_label == "2학기 (9~12월)":
        start_d, end_d = date(2026, 9, 1), date(2026, 12, 31)
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
    c[2].metric("최다 보강", int(df["누적보강"].max()) if len(df) else 0)
    c[3].metric("평균 보강", round(float(df["누적보강"].mean()), 2) if len(df) else 0)

    st.caption(f"기간: {start_str} ~ {end_str}")
    if not df.empty:
        st.bar_chart(df.set_index("교사명")["누적보강"], height=340)
    st.dataframe(df, use_container_width=True, hide_index=True)
