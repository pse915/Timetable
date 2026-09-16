# -*- coding: utf-8 -*-
"""
서라벌여중 시간표·결보강 관리 프로그램
2026 최종 완성판 (속도 최적화 + 버그 수정 버전)
- 결보강 계획서: 보강수업 칸에 '보강 배정된 교사'가 나오도록 수정
- 연계 순환 알고리즘 대폭 가속
- 모든 기존 기능 유지
"""

import io
import copy
from contextlib import contextmanager
import uuid
import time as _time
from datetime import date, datetime, timedelta, time
from zoneinfo import ZoneInfo
from collections import defaultdict
import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components
import gspread
from gspread.exceptions import APIError, WorksheetNotFound
from google.oauth2.service_account import Credentials
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.page import PageMargins

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
    .app-hero { padding: 18px 22px; border: 1px solid #dbe4ee; border-radius: 14px; background: linear-gradient(135deg, #f8fbff 0%, #f8fafc 100%); margin-bottom: 14px; }
    .app-hero-title { font-size: 1.35rem; font-weight: 800; margin-bottom: 4px; }
    .app-hero-sub { color: #64748b; font-size: .92rem; }
    .ux-chip { display:inline-block; padding: 4px 9px; margin: 2px 4px 2px 0; border-radius: 999px; background:#eef4fa; border:1px solid #d8e3ee; font-size:.78rem; }
    .weekly-help { padding: 8px 12px; border-radius: 9px; background:#f8fafc; border:1px solid #e2e8f0; color:#475569; font-size:.86rem; }
</style>
""", unsafe_allow_html=True)

SCHOOL_NAME = "서라벌여자중학교"
SCHOOL_YEAR = "2026"
APP_VERSION = "2.1.0"
DAYS = ["월", "화", "수", "목", "금"]
PERIODS_PER_DAY = {"월": 6, "화": 7, "수": 7, "목": 7, "금": 6}
MAX_PERIOD = 7
WEEKDAY_KR = {0: "월", 1: "화", 2: "수", 3: "목", 4: "금", 5: "토", 6: "일"}
SCHOOL_WEEKDAYS = (0, 1, 2, 3, 4)  # 학교 표시는 월~금만

# 주간표는 한국 학교 일정 기준으로 현재 주를 판정한다.
# 서버가 UTC여도 날짜가 하루 밀리지 않도록 Asia/Seoul을 명시한다.
KST = ZoneInfo("Asia/Seoul")

def _today_kst() -> date:
    return datetime.now(KST).date()

def _is_current_week(ref_date: date, today: date | None = None) -> bool:
    today = today or _today_kst()
    return (ref_date - timedelta(days=ref_date.weekday())) == (today - timedelta(days=today.weekday()))

def _hide_past_week_slots(matrix: pd.DataFrame, ref_date: date, *, hide_past=True) -> pd.DataFrame:
    """현재 주를 볼 때 이미 지나간 평일의 셀을 비운다.

    과거 주/미래 주를 조회할 때는 역사 조회를 방해하지 않도록 원본을 그대로 표시한다.
    현재 날짜 자체의 '몇 교시까지 지났는지'는 학교별 종/수업시간 정보가 코드에 없으므로
    임의의 시간을 추정하지 않고 오늘의 교시는 표시한다.
    """
    if not hide_past or matrix is None or matrix.empty or not _is_current_week(ref_date):
        return matrix
    today = _today_kst()
    monday = ref_date - timedelta(days=ref_date.weekday())
    out = matrix.copy()
    for i, day in enumerate(DAYS):
        day_date = monday + timedelta(days=i)
        if day_date >= today:
            continue
        for p in range(1, PERIODS_PER_DAY.get(day, MAX_PERIOD) + 1):
            col = f"{day}{p}"
            if col in out.columns:
                out[col] = ""
    return out


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


def _month_calendar_df(month_anchor: date, selected_dates=None, range_start=None, range_end=None):
    """월간 달력 매트릭스. 셀 선택으로 날짜를 고르게 하는 공통 UI용 데이터 생성기."""
    selected_dates = set(selected_dates or [])
    first = month_anchor.replace(day=1)
    next_month = (first.replace(day=28) + timedelta(days=4)).replace(day=1)
    last = next_month - timedelta(days=1)
    # 학교 시간표 기준으로 월~금만 표시한다. 토·일은 달력/기간 선택 표에서도 숨긴다.
    # 월요일 시작으로 주 단위를 맞추되, 실제 표의 열은 평일 5일만 사용한다.
    grid_start = first - timedelta(days=first.weekday())
    grid_end = last + timedelta(days=(6 - last.weekday()))
    weekdays = ["월", "화", "수", "목", "금"]
    rows = []
    cur = grid_start
    while cur <= grid_end:
        row = {}
        for i, wd in enumerate(weekdays):
            d = cur + timedelta(days=i)
            if d.month != first.month:
                row[wd] = ""
            else:
                mark = ""
                if d in selected_dates:
                    mark += "● "
                if range_start and d == range_start:
                    mark += "시작 "
                if range_end and d == range_end:
                    mark += "종료 "
                if range_start and range_end and range_start < d < range_end:
                    mark += "■ "
                row[wd] = f"{mark}{d.day:02d}"
        rows.append(row)
        cur += timedelta(days=7)
    return pd.DataFrame(rows, columns=weekdays)


def calendar_picker(label, value=None, key="calendar", help_text=None):
    """기본 날짜 입력을 월간 달력 매트릭스로 대체한다.

    달력 셀을 클릭하면 선택값이 저장되며, 좌우 버튼으로 월을 이동할 수 있다.
    반환값은 datetime.date이다.
    """
    value = value or _today_kst()
    # 학교 일정은 월~금만 사용한다. 주말이 기본값/이전 선택값으로 들어와도 금요일로 보정한다.
    if value.weekday() >= 5:
        value = value - timedelta(days=value.weekday() - 4)
    month_key = f"_{key}_month"
    selected_key = f"_{key}_selected"
    if month_key not in st.session_state:
        st.session_state[month_key] = value.replace(day=1)
    if selected_key not in st.session_state:
        st.session_state[selected_key] = value
    elif st.session_state[selected_key].weekday() >= 5:
        selected = st.session_state[selected_key]
        st.session_state[selected_key] = selected - timedelta(days=selected.weekday() - 4)

    st.markdown(f"**{label}**")
    nav1, nav2, nav3 = st.columns([1, 4, 1])
    with nav1:
        if st.button("◀", key=f"{key}_prev", use_container_width=True):
            m = st.session_state[month_key]
            st.session_state[month_key] = (m.replace(day=1) - timedelta(days=1)).replace(day=1)
            st.rerun()
    with nav2:
        st.markdown(f"<div style='text-align:center;font-weight:700;font-size:1.05rem'>{st.session_state[month_key].year}년 {st.session_state[month_key].month}월</div>", unsafe_allow_html=True)
    with nav3:
        if st.button("▶", key=f"{key}_next", use_container_width=True):
            m = st.session_state[month_key]
            nm = (m.replace(day=28) + timedelta(days=4)).replace(day=1)
            st.session_state[month_key] = nm
            st.rerun()

    quick1, quick2 = st.columns([1, 5])
    with quick1:
        if st.button("오늘", key=f"{key}_today", use_container_width=True):
            today = _today_kst()
            if today.weekday() >= 5:
                today = today - timedelta(days=today.weekday() - 4)
            st.session_state[month_key] = today.replace(day=1)
            st.session_state[selected_key] = today
            st.rerun()
    with quick2:
        st.caption(f"선택: **{st.session_state[selected_key]:%Y-%m-%d}** · 평일(월~금)만 표시됩니다. 날짜 셀을 클릭하세요.")

    cal = _month_calendar_df(st.session_state[month_key], selected_dates={st.session_state[selected_key]})
    event = st.dataframe(
        cal,
        hide_index=True,
        use_container_width=True,
        height=235,
        key=f"{key}_grid",
        on_select="rerun",
        selection_mode="single-cell",
        column_config={c: st.column_config.TextColumn(c, width="small") for c in cal.columns},
    )
    cells = getattr(getattr(event, "selection", None), "cells", []) or []
    if cells:
        row_idx, col_name = cells[0]
        text = str(cal.iloc[row_idx][col_name]).strip()
        import re
        m = re.search(r"(\d{1,2})$", text)
        if m:
            try:
                picked = date(st.session_state[month_key].year, st.session_state[month_key].month, int(m.group(1)))
                st.session_state[selected_key] = picked
            except ValueError:
                pass
    if help_text:
        st.caption(help_text)
    return st.session_state[selected_key]


def calendar_range_picker(start_value=None, end_value=None, key="calendar_range", help_text=None):
    """시작일/종료일을 각각 달력 매트릭스로 선택한다."""
    start_value = start_value or _today_kst()
    end_value = end_value or start_value
    c1, c2 = st.columns(2)
    with c1:
        start_date = calendar_picker("시작일", start_value, f"{key}_start")
    with c2:
        end_date = calendar_picker("종료일", end_value, f"{key}_end")
    if start_date > end_date:
        st.warning("시작일이 종료일보다 늦습니다. 종료일 달력에서 시작일 이후 날짜를 선택하세요.")
    else:
        st.caption(f"선택 기간: **{start_date:%Y-%m-%d} ~ {end_date:%Y-%m-%d}**")
    if help_text:
        st.caption(help_text)
    return start_date, end_date


def period_matrix_picker(label, key, selected=None, allow_all=True):
    """교시를 버튼형 매트릭스로 선택한다. 0은 하루 전체를 의미한다."""
    selected = set(safe_int(x) for x in (selected or []))
    st.markdown(f"**{label}**")
    cols = st.columns(7)
    picked = []
    for p, col in enumerate(cols, 1):
        with col:
            active = p in selected
            if st.button(f"{'✓ ' if active else ''}{p}교시", key=f"{key}_{p}", use_container_width=True, type="primary" if active else "secondary"):
                if active:
                    selected.remove(p)
                else:
                    selected.add(p)
                st.rerun()
    if allow_all:
        all_active = 0 in selected
        if st.button(f"{'✓ ' if all_active else ''}하루 전체", key=f"{key}_all", use_container_width=True, type="primary" if all_active else "secondary"):
            if all_active:
                selected.clear()
            else:
                selected = {0}
            st.rerun()
    if 0 in selected:
        return [0]
    return sorted(p for p in selected if 1 <= p <= 7)



def _daily_schedule_matrix(ref_date: date, *, teacher_filter=None, use_test=False, version=0):
    """선택한 하루를 교사×교시 매트릭스로 표시한다.
    셀에는 현재 적용 수업과 변경 유형을 함께 표시해, 날짜→셀 선택 흐름을 만든다.
    """
    norm = ref_date.strftime("%Y-%m-%d") if isinstance(ref_date, date) else normalize_date_str(ref_date)
    e = get_effective_timetable_for_date(norm, version, use_test=use_test)
    names = []
    if teacher_filter:
        names = [str(teacher_filter).strip()]
    else:
        base = st.session_state.get("teachers", pd.DataFrame())
        if not base.empty and "교사명" in base.columns:
            names = base["교사명"].dropna().astype(str).str.strip().tolist()
        if not e.empty:
            names += e["교사명"].dropna().astype(str).str.strip().tolist()
        names = sorted(set(x for x in names if x))
    idx = {}
    if not e.empty:
        for r in e.itertuples(index=False):
            idx[(str(r.교사명).strip(), safe_int(r.교시))] = r
    rows=[]
    for t in names:
        row={"교사명":t}
        for pno in range(1, MAX_PERIOD+1):
            r=idx.get((t,pno))
            if r is None:
                row[f"{pno}교시"]=""
                continue
            cell=f"{str(r.학급).strip()} {str(r.과목).strip()}".strip()
            typ=str(getattr(r,"변경유형","원본")).strip()
            if typ=="교환": cell += " 🔄"
            elif typ=="테스트교환": cell += " 🧪"
            elif typ=="보강": cell += " 🟢"
            elif typ=="시간강사": cell += f" 🟡 {str(getattr(r,'원본교사','')).strip()}→시간강사"
            row[f"{pno}교시"]=cell
        rows.append(row)
    return pd.DataFrame(rows, columns=["교사명"]+[f"{p}교시" for p in range(1,MAX_PERIOD+1)])


def daily_schedule_picker(ref_date=None, key="daily_schedule", *, teacher_filter=None, use_test=False,
                          multi=False, height=430, help_text=None):
    """달력 → 선택 날짜의 교사×교시 매트릭스. 날짜/대상별 위젯 상태를 분리한다."""
    ref_date = ref_date or _today_kst()
    picked_date = calendar_picker("날짜", ref_date, key=f"{key}_date")
    ver=st.session_state.get("_data_version",0)
    matrix=_daily_schedule_matrix(picked_date, teacher_filter=teacher_filter, use_test=use_test, version=ver)
    st.caption(f"{picked_date:%Y-%m-%d} ({WEEKDAY_KR[picked_date.weekday()]}) · 수업 셀을 클릭하면 작업 대상이 선택됩니다.")
    if matrix.empty:
        st.info("선택한 날짜의 시간표가 없습니다.")
        return picked_date, matrix, []
    teacher_key = str(teacher_filter or "all").strip().replace(" ", "_")
    matrix_key = f"{key}_matrix_{picked_date:%Y%m%d}_{teacher_key}_{'test' if use_test else 'live'}"
    state_key = f"_{key}_selected_cells_{picked_date:%Y%m%d}_{teacher_key}_{'test' if use_test else 'live'}"
    event=st.dataframe(matrix, hide_index=True, use_container_width=True, height=height,
                       key=matrix_key, on_select="rerun",
                       selection_mode="multi-cell" if multi else "single-cell")
    cells=getattr(getattr(event,"selection",None),"cells",[]) or []
    if cells:
        st.session_state[state_key]=list(cells)
    saved_cells=st.session_state.get(state_key,[])
    selections=[]
    for row_idx,col_name in saved_cells:
        if row_idx<0 or row_idx>=len(matrix) or col_name=="교사명": continue
        teacher=str(matrix.iloc[row_idx]["교사명"]).strip()
        period=safe_int(str(col_name).replace("교시",""))
        value=str(matrix.iloc[row_idx][col_name]).strip()
        if teacher and period>0 and value:
            selections.append({"교사명":teacher,"일자":picked_date.strftime("%Y-%m-%d"),
                               "요일":WEEKDAY_KR[picked_date.weekday()],"교시":period,"표시":value})
    if selections:
        st.success("선택: " + " · ".join(f"{x['교사명']} {x['교시']}교시" for x in selections[:8]))
    if help_text: st.caption(help_text)
    return picked_date, matrix, selections


def range_calendar_matrix_picker(start_value=None, end_value=None, key="range_matrix"):
    """두 날짜 달력과 시작/종료 교시 매트릭스를 하나의 선택 영역으로 제공한다."""
    start_value=start_value or _today_kst(); end_value=end_value or start_value
    st.markdown("#### 📅 기간 선택")
    start_date,end_date=calendar_range_picker(start_value,end_value,key=f"{key}_dates")
    st.markdown("#### 🕐 교시 범위 선택")
    c1,c2=st.columns(2)
    with c1:
        start_p=period_matrix_picker("시작일 적용 교시",f"{key}_start_period",st.session_state.get(f"{key}_start_periods",[]))
    with c2:
        end_p=period_matrix_picker("종료일 적용 교시",f"{key}_end_period",st.session_state.get(f"{key}_end_periods",[]))
    st.session_state[f"{key}_start_periods"]=start_p
    st.session_state[f"{key}_end_periods"]=end_p
    return start_date,end_date,start_p,end_p


def render_change_legend():
    st.caption("🟢 보강 · 🔄 실제 교환 · 🧪 테스트 교환 · 🟡 시간강사 · 빈칸=공강")

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
    """현재 상태를 히스토리에 저장한다. 호출은 변경 직전/직후 모두 가능하지만
    동일 상태의 연속 snapshot은 자동 제거한다."""
    if "history" not in st.session_state:
        st.session_state.history=[]; st.session_state.history_index=-1
    state = {
        "action": action_name,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "absences": st.session_state.get("absences", pd.DataFrame()).copy(deep=True),
        "subs": st.session_state.get("subs", pd.DataFrame()).copy(deep=True),
        "swaps": st.session_state.get("swaps", pd.DataFrame()).copy(deep=True),
        "part_time": st.session_state.get("part_time", pd.DataFrame()).copy(deep=True),
        "duties": st.session_state.get("duties", pd.DataFrame()).copy(deep=True),
        "budget_df": load_budget_df().copy(deep=True) if st.session_state.get("logged_in", False) else pd.DataFrame(),
    }
    def fingerprint(x):
        return tuple((k, tuple(v.fillna("").astype(str).tolist())) for k,v in sorted(x.items()) if isinstance(v,pd.DataFrame))
    if st.session_state.history:
        prev=st.session_state.history[st.session_state.history_index]
        if fingerprint(prev)==fingerprint(state):
            prev["action"]=action_name; prev["time"]=state["time"]; return
    st.session_state.history=st.session_state.history[:st.session_state.history_index+1]
    st.session_state.history.append(state)
    st.session_state.history_index=len(st.session_state.history)-1
    if len(st.session_state.history)>MAX_HISTORY:
        st.session_state.history.pop(0); st.session_state.history_index-=1


def _restore_history_snapshot(snap):
    for k in ["absences", "subs", "swaps", "part_time", "duties"]:
        st.session_state[k] = snap.get(k, pd.DataFrame()).copy(deep=True)
    budget_df = snap.get("budget_df")
    if isinstance(budget_df, pd.DataFrame) and not budget_df.empty:
        df_to_worksheet(get_worksheet(WORK_SHEET_ID, "예산"), budget_df)
        load_budget_df.clear()
    _invalidate_all_caches()


def undo():
    if st.session_state.get("history_index", 0) <= 0:
        return False
    st.session_state.history_index -= 1
    _restore_history_snapshot(st.session_state.history[st.session_state.history_index])
    return True



def redo():
    if st.session_state.get("history_index", -1) >= len(st.session_state.get("history", [])) - 1:
        return False
    st.session_state.history_index += 1
    _restore_history_snapshot(st.session_state.history[st.session_state.history_index])
    return True





def _invalidate_all_caches():
    st.session_state._data_version = st.session_state.get("_data_version", 0) + 1
    get_effective_timetable_for_date.clear()
    get_single_lesson_1to1_candidates.clear()
    get_single_lesson_linked_cycles.clear()
    effective_teacher_matrix.clear()
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

def save_work_data_to_gsheet(changed_sheets=None):
    """변경된 시트만 저장한다. None이면 수동 전체 저장이다."""
    if not can_full_data() and not is_teacher():
        st.warning("저장 권한이 없습니다.")
        return False
    try:
        if changed_sheets is None:
            changed_sheets = ["결강", "보강", "맞교환", "시간강사", "복무"]
        changed_sheets = set(changed_sheets)
        if "결강" in changed_sheets:
            df_to_worksheet(get_worksheet(WORK_SHEET_ID, "결강"), st.session_state.absences)
        if "보강" in changed_sheets:
            df_to_worksheet(get_worksheet(WORK_SHEET_ID, "보강"), st.session_state.subs)
        if "맞교환" in changed_sheets:
            df_to_worksheet(get_worksheet(WORK_SHEET_ID, "맞교환"), st.session_state.swaps)
        if "시간강사" in changed_sheets:
            st.session_state.part_time = ensure_part_time_columns(st.session_state.part_time)
            df_to_worksheet(get_worksheet(WORK_SHEET_ID, "시간강사"), st.session_state.part_time)
        if "복무" in changed_sheets:
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
# 핵심 로직
# ==========================================================================================

CHANGE_TYPES = {"원본", "교환", "테스트교환", "보강", "시간강사"}

def _new_change_id(prefix="CHG"):
    return f"{prefix}-{datetime.now().strftime('%Y%m%d%H%M%S')}-{uuid.uuid4().hex[:8]}"


def _lesson_record(teacher, day, period, subject, class_name, *, orig_teacher=None,
                   orig_date="", orig_period=0, change_type="원본", change_source="",
                   change_id="", change_detail=""):
    return {
        "교사명": str(teacher).strip(), "요일": day, "교시": safe_int(period),
        "과목": str(subject).strip(), "학급": str(class_name).strip(),
        "과목군": subject_group(str(subject).strip()),
        "원본교사": str(orig_teacher if orig_teacher is not None else teacher).strip(),
        "원본일자": normalize_date_str(orig_date), "원본교시": safe_int(orig_period or period),
        "변경유형": change_type, "변경출처": str(change_source).strip(),
        "변경ID": str(change_id).strip(), "변경상세": str(change_detail).strip(),
    }


def _truthy_availability(value):
    if value is None or pd.isna(value):
        return False
    s = str(value).strip().lower()
    return s not in {"", "nan", "none", "0", "x", "n", "no", "불가", "아니오", "×"}


def _part_time_available(prow, day, period):
    col = f"{day}{safe_int(period)}"
    return col in prow.index and _truthy_availability(prow.get(col, ""))


def validate_part_time_table(df):
    if df is None or df.empty: return True, ""
    seen = {}
    for _, r in df.iterrows():
        name=str(r.get("시간강사명","")).strip(); orig=str(r.get("대체교사","")).strip()
        start,end=normalize_date_str(r.get("시작일","")),normalize_date_str(r.get("종료일",""))
        if not name or not orig or not start or not end: continue
        if start>end: return False, f"시간강사 {name}: 시작일이 종료일보다 늦습니다."
        for d in DAYS:
            for p in range(1,8):
                if _part_time_available(r,d,p):
                    key=(orig,d,p)
                    if key in seen: return False, f"{orig}의 {d}{p}교시 시간강사 대체가 중복됩니다."
                    seen[key]=name
    return True, ""


def _effective_index(e_tt):
    idx = {}
    if e_tt is None or e_tt.empty:
        return idx
    for r in e_tt.itertuples(index=False):
        idx[(str(r.교사명).strip(), safe_int(r.교시))] = r
    return idx


def _class_slot_teachers(e_tt, class_name, period):
    if e_tt is None or e_tt.empty:
        return []
    m = e_tt[(e_tt["학급"].astype(str).str.strip() == str(class_name).strip()) & (e_tt["교시"].apply(safe_int) == safe_int(period))]
    return [str(x).strip() for x in m["교사명"].tolist()]


def validate_swap(a, b, date_a, date_b, *, is_test=False):
    da, db = normalize_date_str(date_a), normalize_date_str(date_b)
    ta, tb = str(a.get("교사명", "")).strip(), str(b.get("교사명", "")).strip()
    pa, pb = safe_int(a.get("교시", 0)), safe_int(b.get("교시", 0))
    if not da or not db or not ta or not tb or pa <= 0 or pb <= 0:
        return False, "교사·일자·교시 정보가 올바르지 않습니다."
    if ta == tb and da == db and pa == pb:
        return False, "동일한 교사·일자·교시는 교환할 수 없습니다."
    ver = st.session_state.get("_data_version", 0)
    e_a = get_effective_timetable_for_date(da, ver, use_test=is_test)
    e_b = get_effective_timetable_for_date(db, ver, use_test=is_test)
    ma = e_a[(e_a["교사명"] == ta) & (e_a["교시"].apply(safe_int) == pa)] if not e_a.empty else pd.DataFrame()
    mb = e_b[(e_b["교사명"] == tb) & (e_b["교시"].apply(safe_int) == pb)] if not e_b.empty else pd.DataFrame()
    if ma.empty:
        return False, f"현재 적용 시간표에서 {ta}의 {da} {pa}교시 수업을 찾을 수 없습니다."
    if mb.empty:
        return False, f"현재 적용 시간표에서 {tb}의 {db} {pb}교시 수업을 찾을 수 없습니다."
    for label, info, m in [("A", a, ma), ("B", b, mb)]:
        cls, subj = str(info.get("학급", "")).strip(), str(info.get("과목", "")).strip()
        if cls and str(m.iloc[0]["학급"]).strip() != cls:
            return False, f"{label} 수업의 학급이 현재 시간표와 달라졌습니다. 다시 검색하세요."
        if subj and str(m.iloc[0]["과목"]).strip() != subj:
            return False, f"{label} 수업의 과목이 현재 시간표와 달라졌습니다. 다시 검색하세요."
    # A의 목표는 B의 현재 슬롯, B의 목표는 A의 현재 슬롯이다. 다만 교차 날짜에는
    # 상대방이 아닌 제3의 수업이 같은 목표 교사 슬롯을 점유하고 있으면 거부한다.
    if db != da or pb != pa:
        a_at_target = e_b[(e_b["교사명"] == ta) & (e_b["교시"].apply(safe_int) == pb)] if not e_b.empty else pd.DataFrame()
        b_at_target = e_a[(e_a["교사명"] == tb) & (e_a["교시"].apply(safe_int) == pa)] if not e_a.empty else pd.DataFrame()
        if not a_at_target.empty:
            return False, f"A 교사의 목표 슬롯 {db} {pb}교시에 이미 수업이 있습니다."
        if not b_at_target.empty:
            return False, f"B 교사의 목표 슬롯 {da} {pa}교시에 이미 수업이 있습니다."
    if not is_test:
        used = get_actual_direct_swap_affected_slots()
        if (da, ta, pa) in used or (db, tb, pb) in used:
            return False, "이미 다른 맞교환에 사용된 슬롯입니다."
    else:
        if (da, ta, pa) in get_test_affected_slots() or (db, tb, pb) in get_test_affected_slots():
            return False, "테스트에서 이미 사용된 슬롯입니다."
    # 결강/보강/복무와 동시에 동일 슬롯을 조작하지 않도록 방어한다.
    if has_duty(ta, da, pa) or has_duty(tb, db, pb):
        return False, "복무가 등록된 슬롯은 맞교환할 수 없습니다."
    return True, ""

def validate_substitute(cid, on_date, period, class_name, subject, absent_teacher, sub_teacher, *, e_tt=None):
    norm = normalize_date_str(on_date); p = safe_int(period)
    absent_teacher, sub_teacher = str(absent_teacher).strip(), str(sub_teacher).strip()
    if not norm or p <= 0 or not absent_teacher or not sub_teacher:
        return False, "보강에 필요한 정보가 부족합니다."
    if absent_teacher == sub_teacher:
        return False, "결강교사와 보강교사가 같을 수 없습니다."
    ver = st.session_state.get("_data_version", 0)
    if e_tt is None:
        e_tt = get_effective_timetable_for_date(norm, ver)
    source = e_tt[(e_tt["교사명"] == absent_teacher) & (e_tt["교시"].apply(safe_int) == p)] if not e_tt.empty else pd.DataFrame()
    if source.empty:
        # 이미 다른 보강이 걸린 경우에는 현재 effective에서 결강 수업이 빠져 있을 수 있으므로 원본을 확인한다.
        base = st.session_state.timetable
        source = base[(base["교사명"] == absent_teacher) & (base["요일"] == WEEKDAY_KR[datetime.strptime(norm, "%Y-%m-%d").weekday()]) & (base["교시"].apply(safe_int) == p)] if not base.empty else pd.DataFrame()
    if source.empty:
        return False, f"{absent_teacher}의 해당 교시 원본 수업을 찾을 수 없습니다."
    if str(source.iloc[0].get("학급", "")).strip() != str(class_name).strip() or str(source.iloc[0].get("과목", "")).strip() != str(subject).strip():
        return False, "결강 수업 정보가 현재 시간표와 일치하지 않습니다."
    day = WEEKDAY_KR[datetime.strptime(norm, "%Y-%m-%d").weekday()]
    # 보강교사의 현재 effective 수업 여부
    if not is_free(sub_teacher, day, p, norm, e_tt):
        return False, f"{sub_teacher} 교사는 {day}{p}교시에 공강이 아닙니다."
    class_teachers = _class_slot_teachers(e_tt, class_name, p)
    others = [t for t in class_teachers if t != absent_teacher]
    if others:
        return False, f"{class_name}은 {p}교시에 이미 다른 교사({', '.join(others)})가 담당하고 있습니다."
    subs = st.session_state.get("subs", pd.DataFrame())
    if not subs.empty:
        same = subs[(subs["일자"].map(normalize_date_str) == norm) & (subs["교시"].apply(safe_int) == p)]
        same = same[(same["결강교사"].astype(str).str.strip() == absent_teacher) | (same["보강교사"].astype(str).str.strip() == sub_teacher)]
        if not same.empty:
            return False, "같은 시간대에 중복 보강 배정이 있습니다."
    return True, ""


def _apply_sub_to_effective(e_tt, class_name, subject, absent_teacher, sub_teacher, period, norm, change_id=""):
    if e_tt is None:
        return pd.DataFrame()
    current = [dict(x) for x in e_tt.to_dict("records")]
    current = [x for x in current if not (str(x.get("교사명", "")).strip() == absent_teacher and safe_int(x.get("교시")) == safe_int(period))]
    day = WEEKDAY_KR[datetime.strptime(norm, "%Y-%m-%d").weekday()]
    current.append(_lesson_record(sub_teacher, day, period, subject, class_name,
                                  orig_teacher=absent_teacher, orig_date=norm, orig_period=period,
                                  change_type="보강", change_source=f"{absent_teacher} 결강 → {sub_teacher} 보강",
                                  change_id=change_id or _new_change_id("SUB"),
                                  change_detail=f"{class_name} {subject}"))
    return pd.DataFrame(current)

def _is_direct_swap_type(typ: str) -> bool:
    return str(typ).strip() in ["1:1 맞교환", "1:1맞교환", "직접1:1"]


def get_test_affected_slots() -> set:
    """테스트 변경으로 이미 사용된 슬롯을 반환한다.

    반환 키: (일자, 교사명, 교시)
    - 1:1: A 원본, B 원본, A가 가는 목표, B가 가는 원본일 슬롯을 모두 잠근다.
    - 연계: A 원본 슬롯, 목표일의 B 기존 슬롯, 목표일의 A 도착 슬롯을 잠근다.
    """
    affected = set()
    test_swaps = st.session_state.get("test_swaps", pd.DataFrame())
    if test_swaps is None or test_swaps.empty:
        return affected

    for sw in test_swaps.itertuples(index=False):
        date_a = normalize_date_str(getattr(sw, "원본일자", ""))
        date_b = normalize_date_str(getattr(sw, "목표일자", ""))
        t_a = str(getattr(sw, "교사A", "")).strip()
        t_b = str(getattr(sw, "교사B", "")).strip()
        p_a = safe_int(getattr(sw, "교시A", 0))
        p_b = safe_int(getattr(sw, "교시B", 0))
        typ = str(getattr(sw, "유형", "")).strip()

        if date_a and t_a and p_a > 0:
            affected.add((date_a, t_a, p_a))
        if date_b and t_b and p_b > 0:
            affected.add((date_b, t_b, p_b))

        if _is_direct_swap_type(typ):
            if date_a and t_b and p_a > 0:
                affected.add((date_a, t_b, p_a))
            if date_b and t_a and p_b > 0:
                affected.add((date_b, t_a, p_b))
        elif "연계" in typ:
            if date_b and t_a and p_b > 0:
                affected.add((date_b, t_a, p_b))

    return affected


def _test_slot_is_affected(on_date: str, teacher: str, period: int) -> bool:
    norm = normalize_date_str(on_date)
    return (norm, str(teacher).strip(), safe_int(period)) in get_test_affected_slots()


def get_actual_direct_swap_affected_slots() -> set:
    """실제 저장된 1:1 맞교환으로 이미 사용된 슬롯을 모두 반환한다."""
    affected = set()
    swaps = st.session_state.get("swaps", pd.DataFrame())
    if swaps is None or swaps.empty:
        return affected
    for sw in swaps.itertuples(index=False):
        if not _is_direct_swap_type(str(getattr(sw, "유형", ""))):
            continue
        date_a = normalize_date_str(getattr(sw, "원본일자", ""))
        date_b = normalize_date_str(getattr(sw, "목표일자", ""))
        t_a = str(getattr(sw, "교사A", "")).strip()
        t_b = str(getattr(sw, "교사B", "")).strip()
        p_a = safe_int(getattr(sw, "교시A", 0))
        p_b = safe_int(getattr(sw, "교시B", 0))
        for item in [
            (date_a, t_a, p_a), (date_a, t_b, p_a),
            (date_b, t_b, p_b), (date_b, t_a, p_b),
        ]:
            if item[0] and item[1] and item[2] > 0:
                affected.add(item)
    return affected


def _actual_direct_slot_is_affected(on_date: str, teacher: str, period: int) -> bool:
    norm = normalize_date_str(on_date)
    return (norm, str(teacher).strip(), safe_int(period)) in get_actual_direct_swap_affected_slots()


def _effective_swap_origin_info(teacher: str, on_date: str, period: int, use_test: bool = False) -> str:
    """해당 셀이 기존/테스트 교환에 의해 다른 수업이 들어온 슬롯인지 반환한다."""
    norm = normalize_date_str(on_date)
    teacher = str(teacher).strip()
    period = safe_int(period)
    if not norm or not teacher or period <= 0:
        return ""

    tables = []
    actual = st.session_state.get("swaps", pd.DataFrame())
    if actual is not None and not actual.empty:
        tables.append(actual)
    if use_test:
        test = st.session_state.get("test_swaps", pd.DataFrame())
        if test is not None and not test.empty:
            tables.append(test)

    for swaps in tables:
        for sw in swaps.itertuples(index=False):
            typ = str(getattr(sw, "유형", "")).strip()
            date_a = normalize_date_str(getattr(sw, "원본일자", ""))
            date_b = normalize_date_str(getattr(sw, "목표일자", ""))
            teacher_a = str(getattr(sw, "교사A", "")).strip()
            teacher_b = str(getattr(sw, "교사B", "")).strip()
            period_a = safe_int(getattr(sw, "교시A", 0))
            period_b = safe_int(getattr(sw, "교시B", 0))

            if _is_direct_swap_type(typ):
                if date_a == norm and teacher_b == teacher and period_a == period:
                    return f"{date_b} {period_b}교시의 {teacher_a} 수업과 맞교환"
                if date_b == norm and teacher_a == teacher and period_b == period:
                    return f"{date_a} {period_a}교시의 {teacher_b} 수업과 맞교환"
            elif "연계" in typ:
                if date_b == norm and teacher_a == teacher and period_b == period:
                    return f"{date_a} {period_a}교시의 {teacher_b}와 연계교환"
    return ""

def _effective_sub_origin_info(teacher: str, on_date: str, period: int) -> str:
    """해당 날짜·교시에 이 교사가 보강으로 배정된 경우 보강 이력을 반환한다."""
    norm = normalize_date_str(on_date)
    teacher = str(teacher).strip()
    period = safe_int(period)
    if not norm or not teacher or period <= 0:
        return ""

    subs = st.session_state.get("subs", pd.DataFrame())
    if subs is None or subs.empty:
        return ""

    mask = (
        (subs["일자"].astype(str).map(normalize_date_str) == norm)
        & (subs["교시"].apply(safe_int) == period)
        & (subs["보강교사"].astype(str).str.strip() == teacher)
    )
    matches = subs[mask]
    if matches.empty:
        return ""

    row = matches.iloc[-1]
    absent_teacher = str(row.get("결강교사", "")).strip()
    method = str(row.get("배정방식", "")).strip()
    priority = str(row.get("우선순위", "")).strip()
    memo = str(row.get("비고", "")).strip()

    parts = []
    if absent_teacher:
        parts.append(f"{absent_teacher} 결강 → {teacher} 보강")
    else:
        parts.append(f"{teacher} 보강 배정")
    if method:
        parts.append(f"배정방식: {method}")
    if priority:
        parts.append(f"우선순위: {priority}")
    if memo:
        parts.append(f"비고: {memo}")
    return " / ".join(parts)


@st.cache_data(show_spinner=False, ttl=180)
def get_effective_timetable_for_date(on_date: str, version: int = 0, use_test: bool = False) -> pd.DataFrame:
    norm = normalize_date_str(on_date)
    columns = ["교사명","요일","교시","과목","학급","과목군","원본교사","원본일자","원본교시","변경유형","변경출처","변경ID","변경상세"]
    if not norm:
        base = st.session_state.timetable.copy()
        return base.assign(**{c: "" for c in columns if c not in base.columns})
    try:
        day = WEEKDAY_KR[datetime.strptime(norm, "%Y-%m-%d").weekday()]
    except Exception:
        return st.session_state.timetable.copy()
    tt = st.session_state.timetable
    if tt.empty:
        return pd.DataFrame(columns=columns)

    current = {}
    base = tt[tt["요일"] == day]
    for r in base.itertuples(index=False):
        p = safe_int(r.교시); t = str(r.교사명).strip()
        current[(t,p)] = _lesson_record(t, day, p, getattr(r,"과목", ""), getattr(r,"학급", ""),
                                         orig_teacher=t, orig_date=norm, orig_period=p, change_type="원본")

    def apply_swap_table(swaps, is_test_table=False):
        nonlocal current
        if swaps is None or swaps.empty:
            return
        mask = (swaps["원본일자"] == norm) | (swaps["목표일자"] == norm)
        for sw in swaps[mask].itertuples(index=False):
            ta, tb = str(getattr(sw,"교사A","")).strip(), str(getattr(sw,"교사B","")).strip()
            da, db = normalize_date_str(getattr(sw,"원본일자","")), normalize_date_str(getattr(sw,"목표일자",""))
            pa, pb = safe_int(getattr(sw,"교시A",0)), safe_int(getattr(sw,"교시B",0))
            typ = str(getattr(sw,"유형","")).strip(); cid = str(getattr(sw,"변경ID","")).strip() or ""
            sa, ca = str(getattr(sw,"과목A","")).strip(), str(getattr(sw,"학급A","")).strip()
            sb, cb = str(getattr(sw,"과목B","")).strip(), str(getattr(sw,"학급B","")).strip()
            if _is_direct_swap_type(typ):
                if da == db and pa == pb and da == norm:
                    current.pop((ta,pa), None); current.pop((tb,pb), None)
                    current[(tb,pa)] = _lesson_record(tb, day, pa, sa, ca,
                        orig_teacher=str(getattr(sw,"원본교사A",ta)).strip() or ta, orig_date=normalize_date_str(getattr(sw,"실제원본일자A",da)) or da, orig_period=safe_int(getattr(sw,"실제원본교시A",pa)) or pa,
                        change_type="테스트교환" if is_test_table else "교환", change_source=f"{ta} ↔ {tb}", change_id=cid,
                        change_detail=f"{da} {pa}교시의 {ta} 수업을 {tb}가 담당")
                    current[(ta,pb)] = _lesson_record(ta, day, pb, sb, cb,
                        orig_teacher=str(getattr(sw,"원본교사B",tb)).strip() or tb, orig_date=normalize_date_str(getattr(sw,"실제원본일자B",db)) or db, orig_period=safe_int(getattr(sw,"실제원본교시B",pb)) or pb,
                        change_type="테스트교환" if is_test_table else "교환", change_source=f"{ta} ↔ {tb}", change_id=cid,
                        change_detail=f"{db} {pb}교시의 {tb} 수업을 {ta}가 담당")
                else:
                    if da == norm:
                        current.pop((ta,pa), None)
                        if tb:
                            current[(tb,pa)] = _lesson_record(tb, day, pa, sa, ca,
                                orig_teacher=str(getattr(sw,"원본교사A",ta)).strip() or ta, orig_date=normalize_date_str(getattr(sw,"실제원본일자A",da)) or da, orig_period=safe_int(getattr(sw,"실제원본교시A",pa)) or pa,
                                change_type="테스트교환" if is_test_table else "교환", change_source=f"{ta} ↔ {tb}", change_id=cid,
                                change_detail=f"{da} {pa}교시의 {ta} 수업을 {tb}가 담당")
                    if db == norm:
                        current.pop((tb,pb), None)
                        if ta:
                            current[(ta,pb)] = _lesson_record(ta, day, pb, sb, cb,
                                orig_teacher=str(getattr(sw,"원본교사B",tb)).strip() or tb, orig_date=normalize_date_str(getattr(sw,"실제원본일자B",db)) or db, orig_period=safe_int(getattr(sw,"실제원본교시B",pb)) or pb,
                                change_type="테스트교환" if is_test_table else "교환", change_source=f"{ta} ↔ {tb}", change_id=cid,
                                change_detail=f"{db} {pb}교시의 {tb} 수업을 {ta}가 담당")
            elif "연계" in typ and db == norm and ta:
                # 연계교환은 목표 슬롯의 기존 수업을 제거하고 A의 수업을 이동시킨다.
                current.pop((tb,pb), None)
                current[(ta,pb)] = _lesson_record(ta, day, pb, sb, cb,
                    orig_teacher=ta, orig_date=da, orig_period=pa,
                    change_type="테스트교환" if is_test_table else "교환",
                    change_source=f"{ta} → {tb}", change_id=cid,
                    change_detail=f"{da} {pa}교시 수업의 연계 이동")
                if da == norm:
                    current.pop((ta,pa), None)

    apply_swap_table(st.session_state.get("swaps", pd.DataFrame()), False)
    if use_test:
        apply_swap_table(st.session_state.get("test_swaps", pd.DataFrame()), True)

    subs = st.session_state.get("subs", pd.DataFrame())
    if subs is not None and not subs.empty:
        for r in subs[subs["일자"] == norm].itertuples(index=False):
            p = safe_int(getattr(r,"교시",0)); abs_t = str(getattr(r,"결강교사","")).strip(); sub_t = str(getattr(r,"보강교사","")).strip()
            if p <= 0 or not sub_t: continue
            current.pop((abs_t,p), None)
            current.pop((sub_t,p), None)
            current[(sub_t,p)] = _lesson_record(sub_t, day, p, getattr(r,"과목",""), getattr(r,"학급",""),
                orig_teacher=abs_t, orig_date=norm, orig_period=p, change_type="보강",
                change_source=f"{abs_t} 결강 → {sub_t} 보강",
                change_id=str(getattr(r,"결강ID","")).strip() + f"-{p}",
                change_detail=str(getattr(r,"배정방식","")).strip())

    pt_df = st.session_state.get("part_time", pd.DataFrame())
    if pt_df is not None and not pt_df.empty:
        for _, prow in pt_df.iterrows():
            start, end = normalize_date_str(prow.get("시작일","")), normalize_date_str(prow.get("종료일",""))
            if not (start and end and start <= norm <= end): continue
            orig, pt_name = str(prow.get("대체교사","")).strip(), str(prow.get("시간강사명","")).strip()
            if not orig or not pt_name or orig == pt_name: continue
            for (teacher,p), lesson in list(current.items()):
                if teacher != orig or not _part_time_available(prow, day, p): continue
                current.pop((teacher,p), None)
                lesson = dict(lesson); lesson["교사명"] = pt_name; lesson["원본교사"] = orig
                lesson["변경유형"] = "시간강사"; lesson["변경출처"] = f"{orig} → {pt_name}"
                lesson["변경ID"] = lesson.get("변경ID") or _new_change_id("PT")
                lesson["변경상세"] = f"{day}{p} 가능시간에 따른 대체"
                current[(pt_name,p)] = lesson

    df = pd.DataFrame(list(current.values()))
    for c in columns:
        if c not in df.columns: df[c] = ""
    return df[columns].reset_index(drop=True)


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
    teachers = st.session_state.teachers
    norm = normalize_date_str(on_date)
    if e_tt is None:
        e_tt = get_effective_timetable_for_date(norm, st.session_state.get("_data_version", 0))
    grp, grade = subject_group(subject), grade_of(class_name)
    ver = st.session_state.get("_data_version", 0)
    cum, load = cumulative_sub_count(version=ver), weekly_load(version=ver)
    max_cum = max(cum.values()) if cum else 0
    rows=[]
    occupied = {(str(r.교사명).strip(), safe_int(r.교시)) for r in e_tt.itertuples(index=False)} if e_tt is not None and not e_tt.empty else set()
    duty_set = set()
    duties_df = st.session_state.get("duties", pd.DataFrame())
    if duties_df is not None and not duties_df.empty:
        dmask = (duties_df["교사명"].astype(str).str.strip() == str(absent_teacher).strip()) & (duties_df["일자"].astype(str).map(normalize_date_str) == norm)
        # 후보 교사별 복무 여부를 한 번에 인덱싱한다.
        duty_set = {(str(r.교사명).strip(), safe_int(r.교시)) for r in duties_df[duties_df["일자"].astype(str).map(normalize_date_str) == norm].itertuples(index=False)}
    teacher_groups = defaultdict(set); teacher_grades = defaultdict(set)
    if e_tt is not None and not e_tt.empty:
        for rr in e_tt.itertuples(index=False):
            tn = str(rr.교사명).strip()
            if not tn: continue
            teacher_groups[tn].add(str(getattr(rr, "과목군", "")).strip() or subject_group(getattr(rr, "과목", "")))
            teacher_grades[tn].add(grade_of(getattr(rr, "학급", "")))
    teacher_subject = {}
    if not teachers.empty and "교사명" in teachers.columns:
        for rr in teachers.itertuples(index=False):
            tn = str(getattr(rr, "교사명", "")).strip()
            if tn: teacher_subject[tn] = str(getattr(rr, "담당과목", "")).strip()
    for t in (teachers["교사명"].astype(str).str.strip().tolist() if not teachers.empty and "교사명" in teachers.columns else []):
        if t == absent_teacher or (t, safe_int(period)) in occupied or (t, safe_int(period)) in duty_set: continue
        groups=teacher_groups.get(t,set()); grades=teacher_grades.get(t,set())
        if grp in groups and grade in grades: prio,label,score=1,"1순위 · 동일 과목 & 동일 학년",100
        elif grp in groups: prio,label,score=2,"2순위 · 동일 과목",70
        elif grade in grades: prio,label,score=3,"3순위 · 동일 학년",45
        else: prio,label,score=4,"4순위 · 전체 공강",20
        score += (max_cum-cum.get(t,0))*2 + max(0,22-load.get(t,0))*0.3
        rows.append({"보강교사":t,"유형":"정규교사","우선순위":label,"_prio":prio,"담당과목":teacher_subject.get(t,""),"주당시수":load.get(t,0),"누적보강":cum.get(t,0),"추천점수":round(score,1)})
    if include_part_time:
        pt=st.session_state.get("part_time",pd.DataFrame())
        if not pt.empty:
            for _,prow in pt.iterrows():
                t=str(prow.get("시간강사명","")).strip()
                if not t or t==absent_teacher or not _part_time_available(prow,day,period): continue
                if not is_free(t,day,period,norm,e_tt): continue
                pgrp=subject_group(str(prow.get("담당과목","")).strip() or str(prow.get("과목군","")).strip())
                prio=1 if pgrp and pgrp==grp else 2 if pgrp else 4
                label="1순위 · 시간강사 동일 과목" if prio==1 else "2순위 · 시간강사" if prio==2 else "4순위 · 시간강사 공강"
                score=90 if prio==1 else 60 if prio==2 else 15
                rows.append({"보강교사":t,"유형":"시간강사","우선순위":label,"_prio":prio,"담당과목":prow.get("담당과목","") ,"주당시수":0,"누적보강":0,"추천점수":round(score,1)})
    if not rows: return pd.DataFrame()
    return pd.DataFrame(rows).sort_values(["_prio","추천점수"],ascending=[True,False]).drop(columns=["_prio"]).head(top_n).reset_index(drop=True)


@st.cache_data(show_spinner=False, ttl=120)
def get_cached_substitute_recommendations(day, period, subject, class_name, absent_teacher, on_date, top_n=10, include_part_time=True, version=0):
    """주간표 팝업용 보강 후보 캐시. 데이터 변경 버전이 바뀌면 자동으로 무효화된다."""
    return recommend_substitutes(
        day, period, subject, class_name, absent_teacher, on_date,
        top_n=top_n, include_part_time=include_part_time
    )

def add_substitute(cid, on_date, day, period, class_name, subject, absent_teacher, sub_teacher, method, priority, memo, *, save=True, history=True):
    ver = st.session_state.get("_data_version", 0)
    e_tt = get_effective_timetable_for_date(normalize_date_str(on_date), ver)
    ok, msg = validate_substitute(cid, on_date, period, class_name, subject, absent_teacher, sub_teacher, e_tt=e_tt)
    if not ok:
        st.warning(msg)
        return False
    p = safe_int(period); norm = normalize_date_str(on_date)
    s = st.session_state.subs.copy(deep=True)
    old = s[(s["결강ID"].astype(str) == str(cid)) & (s["교시"].apply(safe_int) == p)] if not s.empty else pd.DataFrame()
    if not old.empty:
        s = s.drop(old.index)
    new = pd.DataFrame([{
        "결강ID": cid, "일자": norm, "요일": day, "교시": p, "학급": class_name, "과목": subject,
        "결강교사": absent_teacher, "보강교사": sub_teacher, "배정방식": method, "우선순위": priority,
        "비고": memo, "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M"), "입력자": current_user()
    }])
    st.session_state.subs = pd.concat([s, new], ignore_index=True)
    if save:
        save_work_data_to_gsheet(["보강"])
        if old.empty:
            update_budget(-SUB_COST, f"보강 1건 ({sub_teacher} ← {absent_teacher})")
    _invalidate_all_caches()
    if history:
        push_history(f"보강 배정 ({sub_teacher})")
    return True


def add_substitutes_batch(assignments, action_name="자동 보강"):
    if not assignments: return 0, []
    original_subs = st.session_state.subs.copy(deep=True)
    working = original_subs.copy(deep=True)
    added_new = 0; accepted = 0; errors = []
    for rec in assignments:
        cid, norm, p = rec["cid"], normalize_date_str(rec["on_date"]), safe_int(rec["period"])
        # 검증은 현재 working 상태를 effective에 반영해 순차적으로 재검증한다.
        st.session_state.subs = working
        get_effective_timetable_for_date.clear()
        e_tt = get_effective_timetable_for_date(norm, st.session_state.get("_data_version",0))
        ok, msg = validate_substitute(cid, norm, p, rec["class_name"], rec["subject"], rec["absent_teacher"], rec["sub_teacher"], e_tt=e_tt)
        if not ok:
            errors.append(f"{p}교시: {msg}"); continue
        old = working[(working["결강ID"].astype(str) == str(cid)) & (working["교시"].apply(safe_int) == p)] if not working.empty else pd.DataFrame()
        if not old.empty: working = working.drop(old.index)
        else: added_new += 1
        working = pd.concat([working, pd.DataFrame([{
            "결강ID": cid, "일자": norm, "요일": rec["day"], "교시": p, "학급": rec["class_name"], "과목": rec["subject"],
            "결강교사": rec["absent_teacher"], "보강교사": rec["sub_teacher"], "배정방식": rec.get("method","자동"),
            "우선순위": rec.get("priority",""), "비고": rec.get("memo",""), "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M"), "입력자": current_user()
        }])], ignore_index=True)
        accepted += 1
    st.session_state.subs = working.reset_index(drop=True)
    if accepted:
        if save_work_data_to_gsheet(["보강"]):
            if added_new:
                update_budget(-(added_new * SUB_COST), f"{action_name} {added_new}건")
            push_history(action_name)
    else:
        # 작업이 하나도 없으면 history에 남긴 빈 snapshot을 제거
        if st.session_state.history and st.session_state.history_index == len(st.session_state.history)-1:
            st.session_state.history.pop(); st.session_state.history_index -= 1
    _invalidate_all_caches()
    return accepted, errors


def cancel_substitute(cid, period):
    if not can_full_data():
        s = st.session_state.subs; p = safe_int(period)
        m = s[(s["결강ID"] == cid) & (s["교시"] == p)]
        if not m.empty and m.iloc[0].get("입력자") != current_user():
            st.warning("본인이 입력한 데이터만 취소할 수 있습니다.")
            return False
    p = safe_int(period); s = st.session_state.subs.copy(deep=True)
    m = s[(s["결강ID"] == cid) & (s["교시"] == p)]
    if m.empty:
        return False
    st.session_state.subs = s.drop(m.index).reset_index(drop=True)
    save_work_data_to_gsheet(["보강"])
    update_budget(+SUB_COST, f"보강 취소 복구 ({period}교시)")
    _invalidate_all_caches()
    push_history(f"보강 취소 ({period}교시)")
    return True




def do_swap(a, b, date_a, date_b, is_part_time_purpose=False, is_test=False):
    ok, msg = validate_swap(a, b, date_a, date_b, is_test=is_test)
    if not ok:
        if not is_test: st.warning(msg)
        return False
    ver = st.session_state.get("_data_version", 0)
    e_a_now = get_effective_timetable_for_date(normalize_date_str(date_a), ver, use_test=is_test)
    e_b_now = get_effective_timetable_for_date(normalize_date_str(date_b), ver, use_test=is_test)
    ma_now = e_a_now[(e_a_now["교사명"] == a["교사명"]) & (e_a_now["교시"].apply(safe_int) == safe_int(a["교시"]))] if not e_a_now.empty else pd.DataFrame()
    mb_now = e_b_now[(e_b_now["교사명"] == b["교사명"]) & (e_b_now["교시"].apply(safe_int) == safe_int(b["교시"]))] if not e_b_now.empty else pd.DataFrame()
    oa = ma_now.iloc[0] if not ma_now.empty else a
    ob = mb_now.iloc[0] if not mb_now.empty else b
    rec = {
        "변경ID": _new_change_id("TESTSW" if is_test else "SWAP"),
        "원본교사A": str(oa.get("원본교사", a["교사명"])), "실제원본일자A": str(oa.get("원본일자", normalize_date_str(date_a))), "실제원본교시A": safe_int(oa.get("원본교시", a["교시"])),
        "원본교사B": str(ob.get("원본교사", b["교사명"])), "실제원본일자B": str(ob.get("원본일자", normalize_date_str(date_b))), "실제원본교시B": safe_int(ob.get("원본교시", b["교시"])),
        "원본일자": normalize_date_str(date_a), "교사A": a["교사명"], "요일A": a["요일"], "교시A": safe_int(a["교시"]),
        "학급A": a.get("학급", ""), "과목A": a.get("과목", ""), "목표일자": normalize_date_str(date_b),
        "교사B": b["교사명"], "요일B": b["요일"], "교시B": safe_int(b["교시"]), "학급B": b.get("학급", ""), "과목B": b.get("과목", ""),
        "유형": "1:1 맞교환", "시간강사구인": "Y" if is_part_time_purpose else "N",
        "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M"), "입력자": current_user()
    }
    if is_test:
        st.session_state.test_swaps = pd.concat([st.session_state.get("test_swaps", pd.DataFrame()), pd.DataFrame([rec])], ignore_index=True)
        get_effective_timetable_for_date.clear(); effective_teacher_matrix.clear(); get_single_lesson_1to1_candidates.clear(); get_single_lesson_linked_cycles.clear()
        return True
    st.session_state.swaps = pd.concat([st.session_state.swaps, pd.DataFrame([rec])], ignore_index=True)
    save_work_data_to_gsheet(["맞교환"])
    push_history(f"맞교환 ({a['교사명']} ↔ {b['교사명']})")
    return True



def do_linked_swap(a, teacher_b, date_a, date_b, day_b, period_b, is_part_time_purpose=False, is_test=False, subject_b=None, *, save=True, history=True):
    # 연계교환은 목표 슬롯을 실제로 비우고 A의 수업을 이동시킨다.
    b_probe = {"교사명": teacher_b, "일자": date_b, "요일": day_b, "교시": period_b, "학급": "", "과목": ""}
    ver = st.session_state.get("_data_version", 0)
    e_b = get_effective_timetable_for_date(normalize_date_str(date_b), ver, use_test=is_test)
    if not is_free(teacher_b, day_b, period_b, normalize_date_str(date_b), e_b):
        # 목표 교사가 가진 자신의 수업 슬롯을 연계 대상으로 허용한다.
        pass
    if is_test and _test_slot_is_affected(date_b, teacher_b, period_b):
        return False
    rec = {
        "변경ID": _new_change_id("TESTLINK" if is_test else "LINK"),
        "원본일자": normalize_date_str(date_a), "교사A": a["교사명"], "요일A": a["요일"], "교시A": safe_int(a["교시"]),
        "학급A": a.get("학급", ""), "과목A": a.get("과목", ""), "목표일자": normalize_date_str(date_b),
        "교사B": teacher_b, "요일B": day_b, "교시B": safe_int(period_b), "학급B": a.get("학급", ""),
        "과목B": subject_b if subject_b is not None else a.get("과목", ""), "유형": "연계 공강 교환",
        "시간강사구인": "Y" if is_part_time_purpose else "N", "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M"), "입력자": current_user()
    }
    if is_test:
        st.session_state.test_swaps = pd.concat([st.session_state.get("test_swaps", pd.DataFrame()), pd.DataFrame([rec])], ignore_index=True)
        get_effective_timetable_for_date.clear(); effective_teacher_matrix.clear(); get_single_lesson_1to1_candidates.clear(); get_single_lesson_linked_cycles.clear()
        return True
    st.session_state.swaps = pd.concat([st.session_state.swaps, pd.DataFrame([rec])], ignore_index=True)
    if save: save_work_data_to_gsheet(["맞교환"])
    if history: push_history(f"연계교환 ({a['교사명']} → {teacher_b})")
    return True


def apply_cycle_swaps(moves, is_test=False):
    if not moves: return False
    if is_test:
        for m in moves:
            a_info = {"교사명": m["teacher"], "요일": m.get("day_from", WEEKDAY_KR[datetime.strptime(m["from_date"], "%Y-%m-%d").weekday()]), "교시": m["from_period"], "학급": m["class"], "과목": m["subject"]}
            to_day = WEEKDAY_KR[datetime.strptime(m["to_date"], "%Y-%m-%d").weekday()]
            if not do_linked_swap(a_info, m.get("next_teacher", m["teacher"]), m["from_date"], m["to_date"], to_day, m["to_period"], is_test=True, subject_b=m.get("target_subject", m["subject"])):
                return False
        return True
    # 순환 전체를 하나의 atomic 작업으로 기록/저장한다.
    before = st.session_state.swaps.copy(deep=True)
    for m in moves:
        a_info = {"교사명": m["teacher"], "요일": m.get("day_from", WEEKDAY_KR[datetime.strptime(m["from_date"], "%Y-%m-%d").weekday()]), "교시": m["from_period"], "학급": m["class"], "과목": m["subject"]}
        to_day = WEEKDAY_KR[datetime.strptime(m["to_date"], "%Y-%m-%d").weekday()]
        if not do_linked_swap(a_info, m.get("next_teacher", m["teacher"]), m["from_date"], m["to_date"], to_day, m["to_period"], is_test=False, subject_b=m.get("target_subject", m["subject"]), save=False, history=False):
            st.session_state.swaps = before
            return False
    save_work_data_to_gsheet(["맞교환"])
    push_history(f"{len(moves)}인 연계교환")
    return True


# ==========================================================================================
# ★★★ 연계 공강 순환 알고리즘 (속도 대폭 최적화)
# ==========================================================================================
def find_cycle_linked_swaps(teacher_a, date_a_str, period_a, class_a, subject_a,
                           date_b_str, period_b, min_cycle=2, max_cycle=3, future_days=7, version=0, use_test=False):
    """
    학급 시수·담당·과목을 보존하는 지정 인원 범위의 순환.
    - 교사별 free-slot 사전 계산으로 O(n²) → O(n) 수준으로 개선
    - DFS depth + visited 제한 강화
    """
    original_slot = (normalize_date_str(date_a_str), safe_int(period_a))
    target_slot = (normalize_date_str(date_b_str), safe_int(period_b))

    try:
        da = datetime.strptime(date_a_str, "%Y-%m-%d")
        db = datetime.strptime(date_b_str, "%Y-%m-%d")
        base_min = min(da, db) - timedelta(days=2)
        base_max = max(da, db) + timedelta(days=future_days)
    except Exception:
        return [], "날짜 오류"

    # 후보 날짜 (평일만)
    candidates = []
    cur = base_min
    while cur <= base_max:
        if cur.weekday() < 5:
            candidates.append(cur.strftime("%Y-%m-%d"))
        cur += timedelta(days=1)
    candidates = sorted(set(candidates))

    # 캐시 한 번에 구축
    e_cache = {d: get_effective_timetable_for_date(d, version, use_test=use_test) for d in candidates}

    # 해당 학급의 슬롯 정보
    class_slots = {}          # (date, period) → {teacher, subject, day}
    teacher_occupied = defaultdict(set)  # teacher → set of (date, period)
    affected_test_slots = get_test_affected_slots() if use_test else set()
    weekday_cache = {d: WEEKDAY_KR[datetime.strptime(d, "%Y-%m-%d").weekday()] for d in candidates}

    for d in candidates:
        e = e_cache.get(d)
        if e is None or e.empty:
            continue
        day_kr = weekday_cache[d]
        # itertuples()은 iterrows()보다 훨씬 가볍고, 테스트 영향 슬롯도
        # 반복 호출하지 않고 한 번만 가져와 연계 순환 탐색 비용을 줄인다.
        for r in e.itertuples(index=False):
            t = str(getattr(r, "교사명", "")).strip()
            p = safe_int(getattr(r, "교시", 0))
            teacher_occupied[t].add((d, p))
            if str(getattr(r, "학급", "")).strip() == class_a:
                if use_test and (d, t, p) in affected_test_slots:
                    continue
                class_slots[(d, p)] = {
                    "teacher": t,
                    "subject": str(getattr(r, "과목", "")).strip(),
                    "day": day_kr
                }

    if original_slot not in class_slots or class_slots[original_slot]["teacher"] != teacher_a:
        return [], "원본 슬롯/교사 불일치"
    if target_slot not in class_slots:
        return [], "목표 슬롯에 학급 수업 없음 (공강 생성 금지)"

    day_b = WEEKDAY_KR[datetime.strptime(date_b_str, "%Y-%m-%d").weekday()]
    if not is_free(teacher_a, day_b, period_b, date_b_str, e_cache.get(date_b_str)):
        return [], "교사A 목표시간 수업 있음"

    # 모든 가능한 슬롯 리스트
    slots_list = list(class_slots.keys())

    # 교사별 free slot 사전 계산 (핵심 최적화)
    # free[t] = set of (date, period) where t is free
    free_of_teacher = {}
    all_teachers_in_class = {info["teacher"] for info in class_slots.values()}
    for t in all_teachers_in_class:
        occupied = teacher_occupied.get(t, set())
        free_set = set()
        for d in candidates:
            day_kr = weekday_cache[d]
            max_p = PERIODS_PER_DAY.get(day_kr, 7)
            for p in range(1, max_p + 1):
                slot = (d, p)
                if slot not in occupied and not has_duty(t, d, p):
                    free_set.add(slot)
        free_of_teacher[t] = free_set

    # free_moves: from_slot → list of to_slots that the teacher of from_slot can move to
    free_moves = defaultdict(list)
    for from_s, info in class_slots.items():
        t = info["teacher"]
        possible = free_of_teacher.get(t, set())
        for to_s in possible:
            if to_s != from_s and to_s in class_slots:  # only slots that exist for this class
                free_moves[from_s].append(to_s)

    # DFS (depth-limited + early stop)
    cycles = []
    max_found = 6

    def dfs(current, path, visited):
        if len(cycles) >= max_found:
            return
        if len(path) > max_cycle:
            return
        if current == original_slot:
            # 작은 순환을 먼저 종료해 확장 검색에 섞이지 않도록 한다.
            if len(path) < min_cycle:
                return
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
                    "subject": info["subject"],
                    "target_subject": tgt_info["subject"],
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
    return cycles[:max_found], f"{len(cycles)}개 순환 경로 발견" if cycles else "순환 경로 없음"

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
        date_b_str, period_b, max_cycle=3, future_days=7, version=ver
    )
    return df_swap, cycles, msg
def get_weekly_1to1_swap_table(teacher: str, ref_date: date, future_days: int = 0, version: int = 0) -> pd.DataFrame:
    """
    선택한 교사의 해당 주 + 미래 며칠 동안
    가능한 1:1 맞교환 후보를 모두 반환 (중복 완전 제거)
    """
    weekday = ref_date.weekday()
    monday = ref_date - timedelta(days=weekday)
    week_dates = [(monday + timedelta(days=i)).strftime("%Y-%m-%d") for i in range(5)]

    ver = version or st.session_state.get("_data_version", 0)
    results = []
    seen = set()   # 중복 방지용

    # 검색할 모든 날짜
    search_dates = week_dates[:]
    if future_days > 0:
        last = datetime.strptime(week_dates[-1], "%Y-%m-%d").date()
        for i in range(1, future_days + 1):
            nd = last + timedelta(days=i)
            if nd.weekday() < 5:
                search_dates.append(nd.strftime("%Y-%m-%d"))
    search_dates = sorted(set(search_dates))

    for d_str in week_dates:
        day_kr = WEEKDAY_KR[datetime.strptime(d_str, "%Y-%m-%d").weekday()]
        e_tt = get_effective_timetable_for_date(d_str, ver)
        if e_tt.empty:
            continue

        my_lessons = e_tt[(e_tt["교사명"] == teacher) & (e_tt["요일"] == day_kr)].sort_values("교시")
        for _, lesson in my_lessons.iterrows():
            p = safe_int(lesson["교시"])
            my_class = str(lesson["학급"]).strip()
            my_subj  = str(lesson["과목"]).strip()
            my_grade = grade_of(my_class)
            my_group = subject_group(my_subj)

            for td_str in search_dates:
                if td_str == d_str:
                    continue

                tday = WEEKDAY_KR[datetime.strptime(td_str, "%Y-%m-%d").weekday()]
                e_b = get_effective_timetable_for_date(td_str, ver)
                if e_b.empty:
                    continue

                others = e_b[(e_b["교시"] == p) & (e_b["교사명"] != teacher)]
                # 같은 교시·같은 교사 중복 제거
                others = others.drop_duplicates(subset=["교사명", "교시"])

                for _, o in others.iterrows():
                    other_teacher = str(o["교사명"]).strip()

                    # 중복 키
                    dup_key = (d_str, p, td_str, other_teacher)
                    if dup_key in seen:
                        continue

                    if not is_free(teacher, tday, p, td_str, e_b):
                        continue
                    if not is_free(other_teacher, day_kr, p, d_str, e_tt):
                        continue

                    other_class = str(o["학급"]).strip()
                    other_grade = grade_of(other_class)
                    same_class  = (other_class == my_class)
                    same_grade  = (other_grade == my_grade)

                    score = 0
                    if same_class: score += 200
                    elif same_grade: score += 100
                    if subject_group(str(o["과목"])) == my_group: score += 40
                    if td_str[:7] == d_str[:7]: score += 10

                    seen.add(dup_key)
                    results.append({
                        "원본일자": d_str,
                        "원본요일": day_kr,
                        "원본교시": p,
                        "원본학급": my_class,
                        "원본과목": my_subj,
                        "이동희망일": td_str,
                        "이동요일": tday,
                        "상대교사": other_teacher,
                        "상대학급": str(other_class), "상대과목": str(o["과목"]),
                        "상대수업": f"{other_class} {o['과목']}",
                        "동일학급": "🏆" if same_class else "",
                        "동학년": "⚠" if same_grade and not same_class else "",
                        "점수": score,
                        "_sort": (0 if same_class else 1, 0 if same_grade else 1, -score)
                    })

    if not results:
        return pd.DataFrame(columns=[
            "원본일자", "원본요일", "원본교시", "원본학급", "원본과목",
            "이동희망일", "이동요일", "상대교사", "상대수업",
            "동일학급", "동학년", "점수"
        ])

    df = pd.DataFrame(results)
    df = df.sort_values("_sort").drop(columns=["_sort"]).reset_index(drop=True)
    return df

@st.cache_data(show_spinner=False, ttl=180)
def get_single_lesson_1to1_candidates(
    teacher: str, orig_date_str: str, orig_period: int,
    orig_class: str, orig_subject: str, future_days: int = 0, version: int = 0, use_test: bool = False
) -> pd.DataFrame:
    """선택한 원본 수업 한 건만 대상으로 동일 학급 1:1 교환 후보를 찾는다.

    테스트 모드에서는
    - 원본 수업이 이미 테스트 변경에 사용되었으면 다시 후보를 만들지 않고,
    - 목표 수업도 이미 테스트 변경에 사용된 슬롯이면 후보에서 제외한다.
    """
    source_date = datetime.strptime(normalize_date_str(orig_date_str), "%Y-%m-%d").date()
    source_day = WEEKDAY_KR[source_date.weekday()]
    # 선택한 수업보다 이전 날짜는 교환 후보에서 제외한다.
    # 현재 주간을 유지하되, 선택일 이후의 평일만 후보로 검색한다.
    # 예: 수요일 수업을 선택하면 월/화 수업이 후보로 다시 나타나지 않는다.
    monday = source_date - timedelta(days=source_date.weekday())
    friday = monday + timedelta(days=4)
    search_dates = [
        source_date + timedelta(days=i)
        for i in range((friday - source_date).days + 1)
        if (source_date + timedelta(days=i)).weekday() < 5
    ]
    if future_days > 0:
        last_weekday = friday
        search_dates.extend(
            last_weekday + timedelta(days=i)
            for i in range(1, future_days + 1)
            if (last_weekday + timedelta(days=i)).weekday() < 5
        )

    ver = version or st.session_state.get("_data_version", 0)
    source_str = source_date.strftime("%Y-%m-%d")
    source_tt = get_effective_timetable_for_date(source_str, ver, use_test=use_test)
    if source_tt.empty:
        return pd.DataFrame()

    # 변경된 슬롯도 현재 적용 시간표의 실제 상태이므로 그대로 후보 계산에 포함한다.
    # 과거에 변경되었는지 여부가 아니라, 현재 시점의 교사/날짜/교시 상태를 기준으로 판정한다.
    source_match = source_tt[
        (source_tt["교사명"] == str(teacher).strip())
        & (source_tt["교시"] == safe_int(orig_period))
        & (source_tt["학급"] == str(orig_class).strip())
        & (source_tt["과목"] == str(orig_subject).strip())
    ]
    if source_match.empty:
        return pd.DataFrame()

    results = []
    seen = set()
    source_group = subject_group(orig_subject)

    for target_date in search_dates:
        target_str = target_date.strftime("%Y-%m-%d")
        # source_date 이전 날짜는 검색 대상이 아니다.
        if target_str <= source_str:
            continue
        target_day = WEEKDAY_KR[target_date.weekday()]
        target_tt = get_effective_timetable_for_date(target_str, ver, use_test=use_test)
        if target_tt.empty or not is_free(teacher, target_day, orig_period, target_str, target_tt):
            continue

        candidates = target_tt[
            (target_tt["교시"] == orig_period)
            & (target_tt["학급"] == orig_class)
            & (target_tt["교사명"] != teacher)
        ].drop_duplicates(subset=["교사명", "교시"])

        for _, candidate in candidates.iterrows():
            other_teacher = str(candidate["교사명"]).strip()
            target_period = safe_int(candidate["교시"])
            # 상대 수업도 현재 적용 시간표의 실제 수업이면 후보로 허용한다.
            key = (target_str, other_teacher, target_period)
            if key in seen or not is_free(other_teacher, source_day, orig_period, source_str, source_tt):
                continue
            seen.add(key)
            score = 200
            if subject_group(str(candidate["과목"])) == source_group:
                score += 40
            if target_str[:7] == source_str[:7]:
                score += 10
            results.append({
                "원본일자": source_str, "원본요일": source_day, "원본교시": orig_period,
                "원본학급": orig_class, "원본과목": orig_subject,
                "이동희망일": target_str, "이동요일": target_day,
                "상대교사": other_teacher,
                "상대학급": str(candidate["학급"]), "상대과목": str(candidate["과목"]),
                "상대수업": f"{candidate['학급']} {candidate['과목']}",
                "동일학급": "🏆", "동학년": "", "점수": score
            })

    return pd.DataFrame(results).sort_values(
        ["점수", "이동희망일", "상대교사"], ascending=[False, True, True]
    ).reset_index(drop=True) if results else pd.DataFrame()

@st.cache_data(show_spinner=False, ttl=180)
def get_single_lesson_linked_cycles(
    teacher: str, orig_date_str: str, orig_period: int,
    orig_class: str, orig_subject: str, future_days: int = 0, version: int = 0,
    min_cycle: int = 2, max_cycle: int = 3, use_test: bool = False
):
    """선택 수업용 연계 순환 후보를 지정 인원 범위에서 탐색한다."""
    source_date = datetime.strptime(normalize_date_str(orig_date_str), "%Y-%m-%d").date()
    source_str = source_date.strftime("%Y-%m-%d")
    monday = source_date - timedelta(days=source_date.weekday())
    search_dates = [monday + timedelta(days=i) for i in range(5)]
    if future_days > 0:
        friday = search_dates[-1]
        search_dates.extend(
            friday + timedelta(days=i)
            for i in range(1, future_days + 1)
            if (friday + timedelta(days=i)).weekday() < 5
        )

    ver = version or st.session_state.get("_data_version", 0)
    target_slots = []
    for target_date in search_dates:
        target_str = target_date.strftime("%Y-%m-%d")
        # source_date 이전 날짜는 검색 대상이 아니다.
        if target_str <= source_str:
            continue
        target_day = WEEKDAY_KR[target_date.weekday()]
        target_tt = get_effective_timetable_for_date(target_str, ver, use_test=use_test)
        if target_tt.empty:
            continue
        class_lessons = target_tt[target_tt["학급"] == orig_class]
        for _, row in class_lessons.iterrows():
            target_period = safe_int(row["교시"])
            if is_free(teacher, target_day, target_period, target_str, target_tt):
                if use_test:
                    # 목표 학급 슬롯 자체가 이미 테스트에 사용되었으면
                    # 다시 새로운 순환의 시작점으로 사용할 수 없다.
                    target_rows = target_tt[
                        (target_tt["교시"] == target_period)
                        & (target_tt["학급"] == orig_class)
                    ]
                    if any(
                        (target_str, str(r["교사명"]).strip(), target_period) in get_test_affected_slots()
                        for _, r in target_rows.iterrows()
                    ):
                        continue
                # 같은 교시·가까운 날짜·같은 과목군을 우선으로, 탐색 수를 제한한다.
                priority = (
                    0 if target_period == orig_period else 1,
                    abs((target_date - source_date).days),
                    0 if subject_group(str(row["과목"])) == subject_group(orig_subject) else 1
                )
                target_slots.append((priority, target_str, target_period))

    if not target_slots:
        return [], "연계 순환을 시작할 수 있는 빈 시간대가 없습니다."

    target_slots.sort()
    all_cycles, seen_paths = [], set()
    # 최대 10개 목표 슬롯만 검사해, 1:1 실패 시에도 화면 반응을 유지한다.
    for _, target_str, target_period in target_slots[:10]:
        cycles, _ = find_cycle_linked_swaps(
            teacher, source_str, orig_period, orig_class, orig_subject,
            target_str, target_period, min_cycle=min_cycle, max_cycle=max_cycle,
            future_days=future_days, version=ver, use_test=use_test
        )
        for cycle in cycles:
            cycle_key = tuple((m["teacher"], m["from_date"], m["from_period"], m["to_date"], m["to_period"]) for m in cycle["moves"])
            if cycle_key not in seen_paths:
                seen_paths.add(cycle_key)
                all_cycles.append(cycle)
        if len(all_cycles) >= 6:
            break

    all_cycles.sort(key=lambda c: (c["length"], -c["score"]))
    if not all_cycles:
        return [], "조건을 만족하는 연계 순환 경로가 없습니다."
    return all_cycles[:6], f"{len(all_cycles)}개 연계 순환 경로 발견"
# ==========================================================================================
# 뷰 헬퍼
# ==========================================================================================
@st.cache_data(show_spinner=False)
def teacher_matrix(version=0):
    tt = st.session_state.timetable
    if tt.empty: return pd.DataFrame()
    idx = {(str(r.교사명).strip(), str(r.요일).strip(), safe_int(r.교시)): f"{r.학급} {r.과목}" for r in tt.itertuples(index=False)}
    teachers = sorted(tt["교사명"].astype(str).str.strip().unique())
    rows=[]
    for t in teachers:
        row={"교사명":t}
        for d in DAYS:
            for p in range(1, PERIODS_PER_DAY.get(d,7)+1): row[f"{d}{p}"]=idx.get((t,d,p),"")
        rows.append(row)
    return pd.DataFrame(rows)


@st.cache_data(show_spinner=False, ttl=180)
def effective_teacher_matrix(ref_date: date, version: int = 0, use_test: bool = False) -> pd.DataFrame:
    monday = ref_date - timedelta(days=ref_date.weekday())
    daily_timetables={}; teacher_names=set()
    base_tt=st.session_state.get("timetable",pd.DataFrame())
    if not base_tt.empty: teacher_names.update(base_tt["교사명"].dropna().astype(str).str.strip())
    for i,d in enumerate(DAYS):
        day_date=monday+timedelta(days=i); ds=day_date.strftime("%Y-%m-%d")
        day_tt=get_effective_timetable_for_date(ds,version,use_test=use_test); daily_timetables[d]=(ds,day_tt)
        if not day_tt.empty: teacher_names.update(day_tt["교사명"].dropna().astype(str).str.strip())
    # 교사×교시 인덱스는 교사마다 다시 만들지 않고 요일별로 한 번만 만든다.
    # 이전 구현은 교사 수만큼 동일한 itertuples()/dict 생성을 반복해 주간표가 커질수록 느려졌다.
    daily_indexes = {}
    for d, (_, day_tt) in daily_timetables.items():
        daily_indexes[d] = (
            {(str(r.교사명).strip(), safe_int(r.교시)): r for r in day_tt.itertuples(index=False)}
            if not day_tt.empty else {}
        )

    rows=[]
    for t in sorted(x for x in teacher_names if x):
        row={"교사명":t}
        for d in DAYS:
            idx = daily_indexes[d]
            for p in range(1,PERIODS_PER_DAY.get(d,7)+1):
                r=idx.get((t,p))
                if r is None: row[f"{d}{p}"]=""; continue
                cell=f"{r.학급} {r.과목}".strip(); typ=str(getattr(r,"변경유형", "원본")).strip()
                if typ=="교환": cell += " 🔄 교환"
                elif typ=="테스트교환": cell += " 🧪 테스트교환"
                elif typ=="보강": cell += " 🟢 보강"
                elif typ=="시간강사": cell += f" 🟡 {r.원본교사}→시간강사"
                row[f"{d}{p}"]=cell
        rows.append(row)
    return pd.DataFrame(rows)


@st.cache_data(show_spinner=False)
def class_matrix(version=0, ref_date=None, use_test=False):
    """선택한 주의 실제 적용 학급 매트릭스. 날짜별 교환·보강·시간강사를 반영한다."""
    ref=ref_date or _today_kst(); monday=ref-timedelta(days=ref.weekday())
    daily={}
    classes=set()
    for i,d in enumerate(DAYS):
        ds=(monday+timedelta(days=i)).strftime("%Y-%m-%d")
        e=get_effective_timetable_for_date(ds,version,use_test=use_test); daily[d]=e
        if not e.empty: classes.update(e["학급"].dropna().astype(str).str.strip())
    # 요일별 (학급, 교시) 인덱스를 한 번만 만든다.
    daily_indexes = {}
    for d, e in daily.items():
        idx = {}
        if not e.empty:
            for r in e.itertuples(index=False):
                cls = str(getattr(r, "학급", "")).strip()
                if not cls:
                    continue
                idx[(cls, safe_int(getattr(r, "교시", 0)))] = r
        daily_indexes[d] = idx

    rows=[]
    for c in sorted(x for x in classes if x):
        row={"학급":c}
        for d in DAYS:
            idx = daily_indexes[d]
            for p in range(1,PERIODS_PER_DAY.get(d,7)+1):
                r=idx.get((c,p))
                if r is not None:
                    cell=f"{r.교사명} {r.과목}".strip(); typ=str(getattr(r,"변경유형","원본"))
                    if typ=="교환": cell += " 🔄"
                    elif typ=="테스트교환": cell += " 🧪"
                    elif typ=="보강": cell += " 🟢"
                    elif typ=="시간강사": cell += " 🟡"
                    row[f"{d}{p}"]=cell
                else:
                    row[f"{d}{p}"]=""
        rows.append(row)
    return pd.DataFrame(rows)



def _weekly_cell_parts(value):
    # 주간 매트릭스 셀을 짧고 안정적인 표시 단위로 분해한다.
    text = "" if value is None else str(value).strip()
    if not text:
        return "", "", "원본", ""
    marker = "원본"
    if "🧪" in text or "테스트교환" in text:
        marker = "테스트교환"
    elif "🔄" in text or "교환" in text:
        marker = "교환"
    elif "🟢" in text or "보강" in text:
        marker = "보강"
    elif "🟡" in text or "시간강사" in text:
        marker = "시간강사"
    clean = text.replace("🔄 교환", "").replace("🧪 테스트교환", "")
    clean = clean.replace("🟢 보강", "").replace("🟡 시간강사", "")
    clean = clean.replace(" 🔄", "").replace(" 🧪", "").replace(" 🟢", "")
    if "🟡" in clean:
        clean = clean.split("🟡", 1)[0].strip()
    if "[결강]" in clean:
        clean = clean.replace("[결강]", "").strip()
    parts = clean.split(None, 1)
    cls = parts[0] if parts else ""
    subject = parts[1] if len(parts) > 1 else ""
    icon = {"교환":"🔄", "테스트교환":"🧪", "보강":"🟢", "시간강사":"🟡"}.get(marker, "")
    return cls, subject, marker, icon



def _resolve_weekly_selection(selection, ref_date, use_test=False):
    if not selection:
        return None
    monday = ref_date - timedelta(days=ref_date.weekday())
    picked_date = monday + timedelta(days=selection["day_index"])
    ds = picked_date.strftime("%Y-%m-%d")
    ver = st.session_state.get("_data_version", 0)
    e = get_effective_timetable_for_date(ds, ver, use_test=use_test)
    if e.empty:
        return None
    p = selection["period"]
    if selection["row_label"] == "교사명":
        m = e[(e["교사명"].astype(str).str.strip() == selection["row_name"]) & (e["교시"].apply(safe_int) == p)]
    else:
        m = e[(e["학급"].astype(str).str.strip() == selection["row_name"]) & (e["교시"].apply(safe_int) == p)]
    if m.empty:
        return None
    r = m.iloc[0]
    return {
        "교사명": str(r.get("교사명", "")).strip(), "일자": ds, "요일": DAYS[selection["day_index"]],
        "교시": p, "학급": str(r.get("학급", "")).strip(), "과목": str(r.get("과목", "")).strip(),
        "변경유형": str(r.get("변경유형", "원본")).strip(), "변경출처": str(r.get("변경출처", "")).strip(),
        "변경ID": str(r.get("변경ID", "")).strip(), "원본교사": str(r.get("원본교사", r.get("교사명", ""))).strip(),
        "원본일자": str(r.get("원본일자", ds)), "원본교시": safe_int(r.get("원본교시", p)),
    }


def _register_absence_from_weekly(lesson, reason, detail=""):
    if not lesson:
        return False
    on_date, teacher, p = lesson["일자"], lesson["교사명"], safe_int(lesson["교시"])
    existing = st.session_state.absences
    if not existing.empty:
        dup = existing[(existing["일자"].astype(str) == on_date) & (existing["교사명"].astype(str).str.strip() == teacher) & (existing["교시"].apply(safe_int) == p)]
        if not dup.empty:
            st.warning("이미 등록된 결강입니다.")
            return False
    cid = f"{on_date}-{teacher}"
    new = pd.DataFrame([{
        "결강ID": cid, "일자": on_date, "요일": lesson["요일"], "교사명": teacher,
        "사유": reason, "상세사유": detail, "교시": p, "학급": lesson["학급"], "과목": lesson["과목"],
        "등록시각": datetime.now().strftime("%Y-%m-%d %H:%M"), "입력자": current_user()
    }])
    st.session_state.absences = pd.concat([existing.copy(deep=True), new], ignore_index=True)
    save_work_data_to_gsheet(["결강"])
    _invalidate_all_caches()
    push_history(f"주간표에서 결강 등록 ({teacher} {p}교시)")
    return True


def _weekly_display_matrix(matrix: pd.DataFrame) -> pd.DataFrame:
    """주간표 표시용 복사본. 원본 데이터는 건드리지 않고 셀을 짧게 표현한다."""
    if matrix is None or matrix.empty:
        return matrix.copy(deep=True) if isinstance(matrix, pd.DataFrame) else pd.DataFrame()
    out = matrix.copy(deep=True)
    for col in out.columns:
        if str(col) == "교사명" or str(col) == "학급":
            continue
        vals = []
        for raw in out[col].tolist():
            cls, subject, marker, icon = _weekly_cell_parts(raw)
            if not cls and not subject:
                vals.append("")
            else:
                # Streamlit dataframe은 줄바꿈을 지원하므로 3줄로 고정한다.
                vals.append("\n".join([x for x in (cls, subject, icon) if x]))
        out[col] = vals
    return out


def _weekly_styled_matrix(matrix: pd.DataFrame):
    """요일 그룹/변경 상태를 강조한 pandas Styler를 반환한다."""
    display = _weekly_display_matrix(matrix)
    if display is None or display.empty:
        return display
    styler = display.style
    # 셀 기본 가독성
    styler = styler.set_properties(**{
        "text-align": "center",
        "vertical-align": "middle",
        "white-space": "pre-wrap",
        "line-height": "1.15",
        "font-size": "12px",
    })
    if "교사명" in display.columns:
        styler = styler.set_properties(subset=["교사명"], **{
            "font-weight": "700", "text-align": "left", "white-space": "nowrap"
        })
    if "학급" in display.columns:
        styler = styler.set_properties(subset=["학급"], **{
            "font-weight": "700", "text-align": "left", "white-space": "nowrap"
        })

    # 요일별 아주 옅은 배경 + 요일 시작 열의 굵은 왼쪽 경계.
    day_rgba = [
        "rgba(59,130,246,.035)", "rgba(16,185,129,.035)", "rgba(245,158,11,.040)",
        "rgba(139,92,246,.035)", "rgba(236,72,153,.035)"
    ]
    for di, day in enumerate(DAYS):
        cols = [f"{day}{p}" for p in range(1, MAX_PERIOD + 1) if f"{day}{p}" in display.columns]
        if not cols:
            continue
        styler = styler.set_properties(subset=cols, **{"background-color": day_rgba[di]})
        first = cols[0]
        styler = styler.set_properties(subset=[first], **{"border-left": "3px solid rgba(71,85,105,.42)"})

    # 변경 상태 셀은 내용에 포함된 아이콘을 기준으로 강조한다.
    def _status_style(v):
        s = "" if v is None else str(v)
        if "🔄" in s:
            return "box-shadow: inset 0 0 0 2px rgba(239,68,68,.82); font-weight:700;"
        if "🧪" in s:
            return "box-shadow: inset 0 0 0 2px rgba(124,58,237,.82); font-weight:700;"
        if "🟢" in s:
            return "box-shadow: inset 0 0 0 2px rgba(22,163,74,.82); font-weight:700;"
        if "🟡" in s:
            return "box-shadow: inset 0 0 0 2px rgba(217,119,6,.82); font-weight:700;"
        return ""
    data_cols = [c for c in display.columns if str(c) not in ("교사명", "학급")]
    if data_cols:
        styler = styler.map(_status_style, subset=data_cols) if hasattr(styler, "map") else styler.applymap(_status_style, subset=data_cols)
    return styler


def _clear_weekly_selection():
    # dataframe selection은 위젯 내부 상태가 남을 수 있으므로 epoch를 증가시켜
    # 다음 실행에서 새로운 위젯 key를 사용한다. 같은 셀을 다시 눌러도 정상적으로
    # 새 선택 이벤트가 발생하며, 닫은 직후 팝업이 재오픈되는 현상도 막는다.
    st.session_state["weekly_matrix_epoch"] = int(st.session_state.get("weekly_matrix_epoch", 0) or 0) + 1
    st.session_state.pop("weekly_selected_lesson", None)
    st.session_state.pop("weekly_swap_source", None)
    st.session_state.pop("weekly_dialog_action_mode", None)
    st.session_state.pop("weekly_dialog_open", None)
    st.session_state.pop("weekly_swap_candidates", None)
    st.session_state.pop("weekly_swap_candidates_key", None)
    st.session_state.pop("weekly_cycle_candidates", None)
    st.session_state.pop("weekly_cycle_candidates_msg", None)
    st.session_state.pop("weekly_cycle_candidates_key", None)


def _validate_current_weekly_selection(lesson, *, use_test=False):
    """session_state에 남아 있는 선택이 현재 Effective Schedule과 일치하는지 확인한다.

    Streamlit의 dataframe selection은 rerun 사이에 잔존할 수 있고, 후보 캐시도
    과거 실행의 결과를 잠시 보유할 수 있다. 따라서 팝업을 그리기 직전에 선택 슬롯의
    현재 유효 수업을 다시 확인하여 "한 번 전 클릭한 수업"이나 오래된 수업을 차단한다.
    """
    if not isinstance(lesson, dict):
        return False
    try:
        ds = normalize_date_str(lesson.get("일자", ""))
        period = safe_int(lesson.get("교시", 0))
        teacher = str(lesson.get("교사명", "")).strip()
        klass = str(lesson.get("학급", "")).strip()
        subject = str(lesson.get("과목", "")).strip()
        if not ds or not teacher or period <= 0 or not klass or not subject:
            return False
        ver = st.session_state.get("_data_version", 0)
        e = get_effective_timetable_for_date(ds, ver, use_test=bool(use_test))
        if e is None or e.empty:
            return False
        m = e[(e["교사명"].astype(str).str.strip() == teacher)
               & (e["교시"].apply(safe_int) == period)
               & (e["학급"].astype(str).str.strip() == klass)
               & (e["과목"].astype(str).str.strip() == subject)]
        return not m.empty
    except Exception:
        return False


def _filter_current_swap_candidates(df, lesson, *, use_test=False):
    """후보 캐시에 오래된 수업이 섞여 있어도 현재 Effective Schedule과 대조해 제거한다."""
    if df is None or df.empty or not isinstance(lesson, dict):
        return pd.DataFrame(columns=list(df.columns) if isinstance(df, pd.DataFrame) else [])
    ver = st.session_state.get("_data_version", 0)
    source_date = normalize_date_str(lesson.get("일자", ""))
    source_period = safe_int(lesson.get("교시", 0))
    source_teacher = str(lesson.get("교사명", "")).strip()
    source_class = str(lesson.get("학급", "")).strip()
    source_subject = str(lesson.get("과목", "")).strip()
    source_e = get_effective_timetable_for_date(source_date, ver, use_test=bool(use_test))
    if source_e is None or source_e.empty:
        return df.iloc[0:0].copy()
    source_ok = not source_e[(source_e["교사명"].astype(str).str.strip() == source_teacher)
                              & (source_e["교시"].apply(safe_int) == source_period)
                              & (source_e["학급"].astype(str).str.strip() == source_class)
                              & (source_e["과목"].astype(str).str.strip() == source_subject)].empty
    if not source_ok:
        return df.iloc[0:0].copy()

    valid_rows = []
    target_cache = {}
    for idx, r in df.iterrows():
        td = normalize_date_str(r.get("이동희망일", ""))
        tp = safe_int(r.get("원본교시", source_period))
        tt = str(r.get("상대교사", "")).strip()
        tc = str(r.get("상대학급", "")).strip()
        ts = str(r.get("상대과목", "")).strip()
        # 선택 수업보다 과거인 날짜는 오래된 후보로 간주하여 제거한다.
        if (not td or not tt or tp <= 0 or not tc or not ts
                or td == source_date or td < source_date):
            continue
        if td not in target_cache:
            target_cache[td] = get_effective_timetable_for_date(td, ver, use_test=bool(use_test))
        te = target_cache[td]
        if te is None or te.empty:
            continue
        m = te[(te["교사명"].astype(str).str.strip() == tt)
               & (te["교시"].apply(safe_int) == tp)
               & (te["학급"].astype(str).str.strip() == tc)
               & (te["과목"].astype(str).str.strip() == ts)]
        if not m.empty:
            valid_rows.append(idx)
    return df.loc[valid_rows].reset_index(drop=True)


def _weekly_fragment_rerun():
    """Dialog 내부 상태 변경은 가능한 한 fragment rerun으로 처리해 팝업을 유지한다."""
    try:
        st.rerun(scope="fragment")
    except TypeError:
        # 구버전 Streamlit 호환: fragment scope를 지원하지 않으면 전체 rerun으로 폴백
        st.rerun()


@contextmanager
def _weekly_dialog_loading(label: str):
    """주간 작업 실행 중에만 보이는 초경량 로딩 표시.

    완료 후 별도의 상태 박스/안내 문구를 남기지 않는다. 실제 계산은
    호출부의 context 안에서 실행되므로 Streamlit의 spinner가 계산 중에만
    표시되고, 계산이 끝나면 자동으로 사라진다.
    """
    with st.spinner(f"🔄 로딩 중... {label}"):
        yield


@st.dialog("🎯 수업 작업", width="large")
def _weekly_action_dialog():
    """주간표 셀용 초경량 컨텍스트 팝업.

    초기 팝업에서는 무거운 후보 검색을 절대 실행하지 않는다.
    사용자가 실제 작업을 선택한 순간에만 필요한 검색을 지연 실행한다.
    검색 결과는 session_state + 하위 함수의 cache_data를 이용해 재사용한다.
    """
    lesson = st.session_state.get("weekly_selected_lesson")
    if not lesson:
        st.info("선택한 수업이 없습니다.")
        return

    use_test = bool(st.session_state.get("weekly_dialog_use_test", False))
    # 팝업이 열리기 직전 현재 Effective Schedule과 선택 수업을 재검증한다.
    # 이전 클릭의 session_state가 남아 있어도 오래된 수업을 표시하지 않는다.
    if not _validate_current_weekly_selection(lesson, use_test=use_test):
        _clear_weekly_selection()
        _weekly_fragment_rerun()

    title = st.session_state.get("weekly_dialog_title", "주간표 작업")
    status = lesson.get("변경유형") or "원본"
    ver = st.session_state.get("_data_version", 0)

    st.markdown(f"### {title}")
    st.info(
        f"**{lesson['일자']} ({lesson['요일']}) · {lesson['교사명']} · "
        f"{lesson['교시']}교시 · {lesson['학급']} · {lesson['과목']}**\n\n"
        f"현재 상태: **{status}**"
    )
    if st.session_state.get("weekly_dialog_result"):
        # 저장/테스트 완료 후에는 이전 수업의 유효성 재검증이나 무거운 후보 검색을
        # 다시 수행하지 않는다. 성공 결과를 팝업 안에 그대로 보여주고 사용자가 닫도록 한다.
        st.success(st.session_state.weekly_dialog_result)
        st.caption("시간표가 갱신되었습니다. 팝업을 닫으면 최신 주간표를 확인할 수 있습니다.")
        if st.button("✖ 닫기", key="dlg_result_close", use_container_width=True):
            st.session_state.pop("weekly_dialog_result", None)
            _clear_weekly_selection()
            _weekly_fragment_rerun()
        return

    # ----------------------------------------------------------------
    # 처음 팝업에서는 메뉴만 그린다. 후보/보강 계산은 하지 않는다.
    # ----------------------------------------------------------------
    # 1:1 맞교환을 팝업의 기본 작업으로 사용한다.
    # 다른 기능은 버튼을 눌렀을 때만 해당 화면으로 전환한다.
    action_mode = st.session_state.get("weekly_dialog_action_mode", "swap")
    if action_mode not in {"swap", "cycle", "absence", "substitute", "detail"}:
        action_mode = "swap"
        st.session_state.weekly_dialog_action_mode = "swap"

    if action_mode == "swap":
        st.markdown("#### 🔄 1:1 기본 맞교환")
        st.caption("가장 자주 사용하는 1:1 맞교환을 기본 화면으로 표시합니다. 다른 작업은 아래 버튼을 눌러 진행하세요.")
    else:
        nav_cols = st.columns(5)
        nav_items = [
            ("swap", "🔄 1:1 맞교환"),
            ("cycle", "🔗 연계 순환"),
            ("absence", "📌 결강"),
            ("substitute", "🟢 보강"),
            ("detail", "ℹ️ 상세"),
        ]
        for col, (mode, label) in zip(nav_cols, nav_items):
            with col:
                if st.button(label, type="primary" if mode == action_mode else "secondary",
                             key=f"dlg_action_{mode}", use_container_width=True):
                    st.session_state.weekly_dialog_action_mode = mode
                    _weekly_fragment_rerun()
        st.markdown(f"#### {dict(nav_items)[action_mode]}")

    # 기본 1:1 화면에서도 다른 작업으로 즉시 이동할 수 있게 작은 메뉴만 둔다.
    if action_mode == "swap":
        alt_cols = st.columns(4)
        alt_items = [("cycle", "🔗 연계 순환"), ("absence", "📌 결강"),
                     ("substitute", "🟢 보강"), ("detail", "ℹ️ 상세")]
        for col, (mode, label) in zip(alt_cols, alt_items):
            with col:
                if st.button(label, type="secondary", key=f"dlg_alt_{mode}", use_container_width=True):
                    st.session_state.weekly_dialog_action_mode = mode
                    _weekly_fragment_rerun()

    if status != "원본" and action_mode == "detail":
        st.caption(
            f"변경출처: {lesson.get('변경출처') or '-'} · 변경ID: {lesson.get('변경ID') or '-'} · "
            f"원본: {lesson.get('원본교사') or '-'} / {lesson.get('원본일자') or '-'} / "
            f"{lesson.get('원본교시') or '-'}교시"
        )

    # 검색 범위는 실제 검색 작업을 선택했을 때만 노출한다.
    # 팝업 속도를 위해 기본은 선택한 주간(월~금)만 검색한다.
    # 미래 날짜 검색은 사용자가 필요할 때만 확장한다.
    extra_days = int(st.session_state.get("weekly_dialog_extra_days", 7))
    if action_mode in ("swap", "cycle"):
        with st.expander("🔎 검색 범위 확장", expanded=False):
            extra_days = st.slider(
                "미래 추가 검색 일수", 0, 21, extra_days,
                key="weekly_dialog_extra_days_input",
                help="기본값은 미래 7일을 추가 검색합니다. 필요할 때 검색 범위를 조정할 수 있습니다.",
            )
            if extra_days != st.session_state.get("weekly_dialog_extra_days"):
                st.session_state.weekly_dialog_extra_days = extra_days
                # 범위가 바뀐 경우에만 후보 캐시를 무효화한다.
                st.session_state.pop("weekly_swap_candidates_key", None)
                st.session_state.pop("weekly_cycle_candidates_key", None)
                _weekly_fragment_rerun()

    # ----------------------------------------------------------------
    # 1:1 교환: 사용자가 버튼을 누른 뒤에만 후보 검색
    # ----------------------------------------------------------------
    if action_mode == "swap":
        cache_key = (
            str(lesson.get("교사명", "")), str(lesson.get("일자", "")), safe_int(lesson.get("교시", 0)),
            str(lesson.get("학급", "")), str(lesson.get("과목", "")), int(extra_days), int(ver), bool(use_test)
        )
        stored_key = st.session_state.get("weekly_swap_candidates_key")
        if stored_key != cache_key:
            with _weekly_dialog_loading("1:1 교환 후보 검색 중"):
                df_swap = get_single_lesson_1to1_candidates(
                    lesson["교사명"], lesson["일자"], safe_int(lesson["교시"]),
                    str(lesson["학급"]), str(lesson["과목"]),
                    future_days=extra_days, version=ver, use_test=use_test,
                )
            st.session_state.weekly_swap_candidates = df_swap
            st.session_state.weekly_swap_candidates_key = cache_key
        else:
            df_swap = st.session_state.get("weekly_swap_candidates", pd.DataFrame())

        # 캐시에 남은 후보도 표시 직전에 현재 Effective Schedule과 다시 대조한다.
        # 따라서 과거 날짜/과거 교사·과목 정보가 UI에 나타나지 않는다.
        df_swap = _filter_current_swap_candidates(df_swap, lesson, use_test=use_test)

        if df_swap.empty:
            st.info("현재 조건에서 가능한 1:1 맞교환 위치가 없습니다.")
        else:
            st.caption(f"가능한 1:1 교환 후보 {len(df_swap)}건 · 동일 학급을 우선 검색했습니다.")
            shortlist = df_swap.head(12).copy()
            labels = [
                f"{row['이동희망일']} ({row['이동요일']}) · {safe_int(row['원본교시'])}교시 · "
                f"{row['상대교사']} · {row['상대학급']} {row['상대과목']}"
                for _, row in shortlist.iterrows()
            ]
            dialog_instance = int(st.session_state.get("weekly_dialog_instance", 0) or 0)
            pick_label = st.selectbox("교환할 수업", labels, key=f"weekly_dialog_swap_pick_{dialog_instance}")
            picked = shortlist.iloc[labels.index(pick_label)]
            st.caption(
                f"상대 수업: **{picked['상대교사']} · {picked['이동희망일']} · "
                f"{safe_int(picked['원본교시'])}교시 · {picked['상대학급']} · {picked['상대과목']}**"
            )
            b_info = {
                "교사명": str(picked["상대교사"]), "일자": str(picked["이동희망일"]),
                "요일": str(picked["이동요일"]), "교시": safe_int(picked["원본교시"]),
                "학급": str(picked["상대학급"]), "과목": str(picked["상대과목"]),
            }
            button_label = "🧪 1:1 맞교환 테스트" if use_test else "✅ 1:1 맞교환 실행"
            if st.button(button_label, type="primary", key="dlg_direct_swap", use_container_width=True):
                try:
                    with _weekly_dialog_loading("1:1 맞교환 처리 중"):
                        ok = do_swap(lesson, b_info, lesson["일자"], b_info["일자"], is_test=use_test)
                except Exception as exc:
                    st.error(f"맞교환 처리 중 오류가 발생했습니다: {exc}")
                    ok = False
                if ok:
                    st.success("테스트 맞교환이 적용되었습니다." if use_test else "1:1 맞교환이 반영되었습니다.")
                    # 전체 앱 rerun 대신 dialog fragment만 갱신한다. 팝업을 유지한 채 결과를 보여준다.
                    st.session_state.weekly_dialog_result = (
                        "테스트 맞교환이 적용되었습니다." if use_test else "1:1 맞교환이 반영되었습니다."
                    )
                    _weekly_fragment_rerun()
                else:
                    st.error("현재 상태에서는 이 1:1 맞교환을 적용할 수 없습니다. 최신 시간표 상태를 다시 확인해 주세요.")

    # ----------------------------------------------------------------
    # 연계 순환: 사용자가 버튼을 누른 뒤에만 후보 검색
    # ----------------------------------------------------------------
    elif action_mode == "cycle":
        cache_key = (
            str(lesson.get("교사명", "")), str(lesson.get("일자", "")), safe_int(lesson.get("교시", 0)),
            str(lesson.get("학급", "")), str(lesson.get("과목", "")), int(extra_days), int(ver), bool(use_test)
        )
        stored_key = st.session_state.get("weekly_cycle_candidates_key")
        if stored_key != cache_key:
            with _weekly_dialog_loading("연계 순환 후보 계산 중"):
                cycles, cycle_msg = get_single_lesson_linked_cycles(
                    lesson["교사명"], lesson["일자"], safe_int(lesson["교시"]),
                    str(lesson["학급"]), str(lesson["과목"]),
                    future_days=extra_days, version=ver, min_cycle=2, max_cycle=3, use_test=use_test,
                )
            st.session_state.weekly_cycle_candidates = cycles
            st.session_state.weekly_cycle_candidates_msg = cycle_msg
            st.session_state.weekly_cycle_candidates_key = cache_key
        else:
            cycles = st.session_state.get("weekly_cycle_candidates", [])
            cycle_msg = st.session_state.get("weekly_cycle_candidates_msg", "")

        st.caption(cycle_msg or "선택한 수업을 시작점으로 연계 순환 가능성을 검사합니다.")
        if not cycles:
            st.info("현재 조건에서 가능한 2·3인 연계 순환 경로가 없습니다.")
        else:
            for idx, cyc in enumerate(cycles[:6]):
                with st.container(border=True):
                    st.markdown(
                        f"**{'🔗' if cyc['length'] > 2 else '↔️'} "
                        f"{cyc['length']}인 순환 · 점수 {cyc.get('score', '')}**"
                    )
                    st.caption(cyc.get("path_desc", ""))
                    st.caption("학급의 담당교사·과목·시수가 보존되는 순환 후보입니다.")
                    if st.button("🧪 이 연계 순환 테스트", key=f"dlg_cycle_test_{idx}", use_container_width=True):
                        try:
                            with _weekly_dialog_loading("연계 순환 테스트 중"):
                                ok = apply_cycle_swaps(cyc["moves"], is_test=True)
                        except Exception as exc:
                            st.error(f"연계 순환 테스트 중 오류가 발생했습니다: {exc}")
                            ok = False
                        if ok is not False:
                            st.session_state["test_has_cycle"] = True
                            st.session_state.weekly_dialog_result = f"테스트 {cyc['length']}인 연계 순환이 적용되었습니다. 실제 저장되지는 않습니다."
                            st.success(st.session_state.weekly_dialog_result)
                            _weekly_fragment_rerun()

    # ----------------------------------------------------------------
    # 결강: 입력 UI만 표시하고, 후보 검색은 하지 않는다.
    # ----------------------------------------------------------------
    elif action_mode == "absence":
        r1, r2 = st.columns([1, 2])
        with r1:
            reason = st.selectbox("사유", ABSENCE_REASONS, key="dlg_abs_reason")
        with r2:
            detail = st.text_input("상세사유", key="dlg_abs_detail")
        if st.button("📌 결강 등록", type="primary", key="dlg_abs_submit", use_container_width=True):
            try:
                with _weekly_dialog_loading("결강 정보 저장 중"):
                    ok = _register_absence_from_weekly(lesson, reason, detail)
            except Exception as exc:
                st.error(f"결강 등록 중 오류가 발생했습니다: {exc}")
                ok = False
            if ok:
                st.success("결강이 등록되었습니다.")
                st.session_state.weekly_dialog_result = "결강이 등록되었습니다."
                # 팝업은 유지하고 최신 상태만 fragment rerun으로 갱신
                _weekly_fragment_rerun()

    # ----------------------------------------------------------------
    # 보강: 사용자가 보강 메뉴를 선택한 경우에만 추천 계산
    # ----------------------------------------------------------------
    elif action_mode == "substitute":
        with _weekly_dialog_loading("보강 후보 확인 중"):
            cand = get_cached_substitute_recommendations(
                lesson["요일"], lesson["교시"], lesson["과목"], lesson["학급"], lesson["교사명"], lesson["일자"],
                top_n=10, include_part_time=True, version=ver
            )
        if cand.empty:
            st.warning("현재 조건에서 추천 가능한 보강 교사가 없습니다.")
        else:
            st.dataframe(cand, use_container_width=True, hide_index=True, height=240)
            abs_df = st.session_state.absences
            abs_match = (
                abs_df[(abs_df["일자"].astype(str) == lesson["일자"]) &
                       (abs_df["교사명"].astype(str).str.strip() == lesson["교사명"]) &
                       (abs_df["교시"].apply(safe_int) == lesson["교시"])]
                if not abs_df.empty else pd.DataFrame()
            )
            if not abs_match.empty:
                cid = str(abs_match.iloc[0]["결강ID"])
                labels = cand["보강교사"].astype(str).tolist()
                pick = st.selectbox("보강 교사", labels, key="dlg_sub_pick")
                picked = cand[cand["보강교사"].astype(str) == str(pick)].iloc[0]
                if st.button("🟢 선택 교사로 보강 배정", type="primary", key="dlg_sub_submit", use_container_width=True):
                    try:
                        with _weekly_dialog_loading("보강 배정 처리 중"):
                            ok = add_substitute(
                                cid, lesson["일자"], lesson["요일"], lesson["교시"], lesson["학급"], lesson["과목"],
                                lesson["교사명"], str(picked["보강교사"]), "주간표", picked.get("우선순위", ""),
                                "주간 시간표 셀에서 배정"
                            )
                    except Exception as exc:
                        st.error(f"보강 배정 중 오류가 발생했습니다: {exc}")
                        ok = False
                    if ok:
                        st.success("보강이 배정되었습니다.")
                        st.session_state.weekly_dialog_result = "보강이 배정되었습니다."
                        _weekly_fragment_rerun()
            else:
                st.caption("이 수업에 등록된 결강이 없습니다. 결강 등록 후 바로 보강을 배정할 수 있습니다.")

    # ----------------------------------------------------------------
    # 상세
    # ----------------------------------------------------------------
    elif action_mode == "detail":
        detail_rows = [
            ("교사", lesson.get("교사명", "")), ("일자", lesson.get("일자", "")),
            ("요일", lesson.get("요일", "")), ("교시", lesson.get("교시", "")),
            ("학급", lesson.get("학급", "")), ("과목", lesson.get("과목", "")),
            ("변경유형", lesson.get("변경유형", "원본")), ("변경출처", lesson.get("변경출처", "")),
            ("변경ID", lesson.get("변경ID", "")), ("원본교사", lesson.get("원본교사", "")),
            ("원본일자", lesson.get("원본일자", "")), ("원본교시", lesson.get("원본교시", "")),
        ]
        st.dataframe(pd.DataFrame(detail_rows, columns=["항목", "내용"]), use_container_width=True, hide_index=True)

    if st.button("✖ 닫기", key="dlg_close", use_container_width=True):
        st.session_state.pop("weekly_dialog_result", None)
        _clear_weekly_selection()
        _weekly_fragment_rerun()

def _resolve_matrix_cell_selection(matrix, ref_date, row_label, selected_cells, *, use_test=False):
    """st.dataframe의 단일 셀 선택을 현재 유효 시간표의 실제 수업으로 해석한다."""
    if not selected_cells:
        return None
    try:
        row_idx, column_name = selected_cells[0]
        row_idx = int(row_idx)
        column_name = str(column_name)
    except Exception:
        return None
    if row_idx < 0 or row_idx >= len(matrix) or column_name in ("교사명", "학급"):
        return None
    day = column_name[:1]
    period = safe_int(column_name[1:])
    if day not in DAYS or not (1 <= period <= MAX_PERIOD):
        return None
    row_name = str(matrix.iloc[row_idx].get(row_label, "")).strip()
    if not row_name or not str(matrix.iloc[row_idx].get(column_name, "")).strip():
        return None
    monday = ref_date - timedelta(days=ref_date.weekday())
    picked_date = monday + timedelta(days=DAYS.index(day))
    ds = picked_date.strftime("%Y-%m-%d")
    ver = st.session_state.get("_data_version", 0)
    e = get_effective_timetable_for_date(ds, ver, use_test=use_test)
    if e.empty:
        return None
    if row_label == "교사명":
        m = e[(e["교사명"].astype(str).str.strip() == row_name) & (e["교시"].apply(safe_int) == period)]
    else:
        m = e[(e["학급"].astype(str).str.strip() == row_name) & (e["교시"].apply(safe_int) == period)]
    if m.empty:
        return None
    r = m.iloc[0]
    return {
        "교사명": str(r.get("교사명", "")).strip(), "일자": ds, "요일": day, "교시": period,
        "학급": str(r.get("학급", "")).strip(), "과목": str(r.get("과목", "")).strip(),
        "변경유형": str(r.get("변경유형", "원본")).strip(), "변경출처": str(r.get("변경출처", "")).strip(),
        "변경ID": str(r.get("변경ID", "")).strip(), "원본교사": str(r.get("원본교사", r.get("교사명", ""))).strip(),
        "원본일자": str(r.get("원본일자", ds)), "원본교시": safe_int(r.get("원본교시", period)),
    }


def render_weekly_selection_panel(ref_date, *, use_test=False, title="선택 수업 작업"):
    """하위 호환용. 팝업을 직접 렌더링하지 않고 중앙 렌더러에 위임한다."""
    lesson = st.session_state.get("weekly_selected_lesson")
    if lesson:
        st.session_state.weekly_dialog_use_test = use_test
        st.session_state.weekly_dialog_title = title
        st.session_state.weekly_dialog_open = True
    else:
        st.caption("주간표의 수업 셀을 클릭하면 작은 팝업에서 결강·맞교환·보강 작업을 시작할 수 있습니다.")


def render_standard_weekly_matrix(matrix: pd.DataFrame, ref_date: date, *, row_label="교사명", key="weekly_matrix", title=None, use_test=False):
    """모든 탭이 동일한 주간 매트릭스 렌더러 설정을 사용하도록 하는 표준 래퍼."""
    return render_weekly_matrix(
        matrix, ref_date, row_label=row_label, height=650, key=key,
        title=title, show_week_dates=True, use_test=use_test, open_dialog=True
    )


def _weekly_selection_signature(selected_cells):
    """주간표 dataframe 선택값을 안정적으로 비교하기 위한 불변 signature."""
    try:
        return tuple((int(r), str(c)) for r, c in selected_cells)
    except Exception:
        return tuple()


def render_weekly_matrix(matrix: pd.DataFrame, ref_date: date, *, row_label="교사명", height=700,
                         key="weekly_matrix", title=None, show_week_dates=True, use_test=False, open_dialog=True):
    """주간 5일×7교시 인터랙티브 렌더러.

    URL/query parameter를 전혀 사용하지 않는다.
    사용자가 주간표의 수업 셀을 클릭하면 Streamlit의 selection 이벤트로
    현재 실행 상태에 선택을 저장하고, native dialog를 띄워 작업한다.

    결과적으로 동작은 바탕화면의 '셀 선택 → 컨텍스트 메뉴'와 비슷하지만,
    브라우저 URL을 바꾸거나 초기 페이지로 이동하지 않는다.
    """
    if matrix is None or matrix.empty:
        st.info("표시할 주간 시간표가 없습니다.")
        return None
    monday = ref_date - timedelta(days=ref_date.weekday())
    if title:
        st.markdown(f"#### {title}")
    if show_week_dates:
        dates = [monday + timedelta(days=i) for i in range(5)]
        st.caption(" · ".join(f"{DAYS[i]} {dates[i]:%Y.%m.%d}" for i in range(5)))
    if _is_current_week(ref_date):
        st.caption("🕒 현재 주: 이미 지난 평일의 수업은 자동으로 숨기고, 오늘부터 남은 시간표를 표시합니다.")
    st.caption("💡 수업 셀을 한 번 클릭하면 페이지 이동 없이 작은 작업 팝업이 열립니다. URL은 변경하지 않습니다.")

    # 현재 주라면 이미 지나간 평일의 수업 셀은 숨긴다.
    # 단, 기준일을 과거/미래 주로 선택한 경우에는 역사 조회를 위해 그대로 보여준다.
    visible_matrix = _hide_past_week_slots(matrix, ref_date, hide_past=True)
    display = _weekly_styled_matrix(visible_matrix)
    column_config = {}
    # 1500px급 브라우저에서 좌우 여백까지 고려해 월~금 전체가 들어오도록 폭을 고정한다.
    # 내부 표는 약 1455px(행 이름 90px + 교시 39px × 최대 35칸)를 목표로 한다.
    # 35칸보다 적은 실제 요일 교시를 가진 학교에서도 같은 규칙을 유지한다.
    compact_period_width = 39
    row_name_width = 90
    if row_label in display.columns:
        column_config[row_label] = st.column_config.TextColumn(row_label, width=row_name_width)
    monday = ref_date - timedelta(days=ref_date.weekday())
    week_dates = [monday + timedelta(days=i) for i in range(5)]
    for day_idx, day in enumerate(DAYS):
        day_date = week_dates[day_idx]
        for p in range(1, MAX_PERIOD + 1):
            col = f"{day}{p}"
            if col in display.columns:
                # 첫 교시에 요일+날짜를 표시하고, 나머지는 요일+교시로 표시해 헤더를 압축한다.
                label = f"{day} {day_date.day}" if p == 1 else f"{day}{p}"
                column_config[col] = st.column_config.TextColumn(label, width=compact_period_width)

    event = None
    selected_cells = []
    # 모든 주간표 위젯은 선택 epoch를 포함한 안정적인 key를 사용한다.
    # 팝업 닫기 후 기존 dataframe의 selection이 재전달되어 팝업이 즉시 재오픈되는
    # Streamlit 특유의 상태 잔존 문제를 방지한다.
    matrix_epoch = int(st.session_state.get("weekly_matrix_epoch", 0) or 0)
    widget_key = f"{key}__{ref_date:%Y%m%d}__{row_label}__{'test' if use_test else 'live'}__sel{matrix_epoch}"
    try:
        event = st.dataframe(
            display,
            hide_index=True,
            use_container_width=False,
            height=height,
            key=widget_key,
            on_select="rerun",
            selection_mode="single-cell",
            column_config=column_config,
            width=1450,
        )
        try:
            selected_cells = list(event.selection.cells)
        except Exception:
            selected_cells = []

        # st.tabs() 안에서는 모든 탭의 dataframe이 같은 실행에서 렌더링된다.
        # 따라서 각 dataframe의 selection이 session에 남아 있으면, 클릭하지 않은
        # 다른 표가 '마지막 클릭'처럼 다시 전달되어 이전 수업으로 팝업이 바뀌는
        # 문제가 생길 수 있다. 각 widget key별 마지막 selection을 기억하고,
        # 실제로 selection이 변경된 widget만 새 선택으로 인정한다.
        selection_seen_key = f"_weekly_selection_seen__{widget_key}"
        current_signature = _weekly_selection_signature(selected_cells)
        previous_signature = st.session_state.get(selection_seen_key, tuple())
        selection_changed = current_signature != previous_signature
        st.session_state[selection_seen_key] = current_signature
        if not selection_changed:
            selected_cells = []
    except TypeError:
        # 구버전 Streamlit에서는 selection API가 없을 수 있다.
        # 이 경우 표 자체는 정상 표시하고 URL 이동 방식으로 대체하지 않는다.
        st.dataframe(
            display,
            hide_index=True,
            use_container_width=False,
            height=height,
            key=f"{key}_legacy__sel{matrix_epoch}",
            column_config=column_config,
            width=1450,
        )
        st.warning("현재 Streamlit 버전에서는 주간표 셀 클릭 기능을 지원하지 않습니다. Streamlit을 최신 버전으로 업데이트하면 셀 클릭 팝업을 사용할 수 있습니다.")
        return None

    lesson = _resolve_matrix_cell_selection(matrix, ref_date, row_label, selected_cells, use_test=use_test)
    if lesson:
        # 선택 상태는 URL이 아니라 session_state에만 저장한다.
        st.session_state.weekly_selected_lesson = lesson
        # 셀을 새로 선택하면 항상 1:1 맞교환을 기본 작업으로 연다.
        st.session_state.weekly_dialog_action_mode = "swap"
        # 이전 셀에서 확장했던 미래 검색 범위를 새 셀에 그대로 물려주지 않는다.
        st.session_state.weekly_dialog_extra_days = 7
        st.session_state.pop("weekly_dialog_extra_days_input", None)
        st.session_state.weekly_dialog_use_test = bool(use_test)
        st.session_state.weekly_dialog_title = title or "주간표 작업"
        # 새 수업을 클릭할 때 팝업 내부의 후보 selectbox도 새 위젯으로 만든다.
        # 이전 수업의 선택 후보가 새 수업에 그대로 남는 것을 방지한다.
        st.session_state["weekly_dialog_instance"] = int(st.session_state.get("weekly_dialog_instance", 0) or 0) + 1
        st.session_state.weekly_dialog_open = bool(open_dialog)
        # 중요: dialog는 이 함수에서 직접 렌더링하지 않는다.
    # 주간표 렌더러는 선택 상태만 기록하고 Dialog는 앱 마지막에서 단 한 번 중앙 렌더링한다.
    return lesson


def get_teacher_week_view(teacher: str, ref_date: date, use_test=False):
    monday=ref_date-timedelta(days=ref_date.weekday()); week_dates=[monday+timedelta(days=i) for i in range(5)]
    ver=st.session_state.get("_data_version",0); grid=[]
    for p in range(1,MAX_PERIOD+1):
        row={"교시":p}
        for i,d in enumerate(DAYS):
            ds=week_dates[i].strftime("%Y-%m-%d"); e=get_effective_timetable_for_date(ds,ver,use_test=use_test)
            m=e[(e["교사명"]==teacher)&(e["교시"].apply(safe_int)==p)] if not e.empty else pd.DataFrame()
            if m.empty: row[d]=""; continue
            r=m.iloc[0]; cell=f"{r['학급']} {r['과목']}".strip(); typ=str(r.get("변경유형","원본"))
            if typ=="교환": cell += f" 🔄 {r.get('변경출처','교환')}"
            elif typ=="테스트교환": cell += f" 🧪 {r.get('변경출처','테스트교환')}"
            elif typ=="보강": cell += f" 🟢 {r.get('변경출처','보강')}"
            elif typ=="시간강사": cell += f" 🟡 {r.get('원본교사','')}→시간강사"
            if not st.session_state.absences.empty and ((st.session_state.absences["일자"]==ds)&(st.session_state.absences["교사명"]==teacher)&(st.session_state.absences["교시"]==p)).any(): cell=f"[결강] {cell}"
            row[d]=cell
        grid.append(row)
    return pd.DataFrame(grid),week_dates


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


def build_weekly_schedule_excel_bytes(ref_date: date, *, use_test=False, title="전체 교사 시간표") -> bytes:
    """업로드된 '체육과 보강 시간표.xlsx' 양식을 기준으로 만든 주간 시간표 Excel."""
    monday = ref_date - timedelta(days=ref_date.weekday())
    ver = st.session_state.get("_data_version", 0)
    wb = Workbook(); ws = wb.active; ws.title = "전체 교사 시간표"
    # 업로드 양식의 핵심 스타일: 돋움, 2행/교사, 번호·교사 병합, 요일 병합 헤더.
    title_font = Font(name="돋움", size=20, bold=True)
    small_font = Font(name="돋움", size=8)
    head_font = Font(name="돋움", size=9, bold=True)
    body_font = Font(name="돋움", size=8)
    thin = Side(style="hair", color="B7B7B7"); med = Side(style="medium", color="808080")
    fill_head = PatternFill("solid", fgColor="D9EAF7")
    fill_day = [PatternFill("solid", fgColor=x) for x in ("DDEBF7","E2F0D9","FFF2CC","E4DFEC","FCE4D6")]
    days = DAYS; periods = {d: list(range(1, PERIODS_PER_DAY.get(d,7)+1)) for d in days}
    col = 3
    day_ranges = {}
    for i,d in enumerate(days):
        start=col; end=col+len(periods[d])-1; day_ranges[d]=(start,end)
        col=end+1
    last_col=col-1
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=last_col)
    ws["A1"] = title; ws["A1"].font=title_font; ws["A1"].alignment=Alignment(horizontal="center",vertical="center")
    ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=2); ws["A2"]=f"{SCHOOL_YEAR} 학년도"; ws["A2"].font=small_font; ws["A2"].alignment=Alignment(horizontal="left",vertical="center")
    ws.merge_cells(start_row=2,start_column=last_col-4,end_row=2,end_column=last_col); ws.cell(2,last_col-4).value=SCHOOL_NAME; ws.cell(2,last_col-4).font=small_font; ws.cell(2,last_col-4).alignment=Alignment(horizontal="right",vertical="center")
    ws["A3"]="번호"; ws["B3"]="교사"
    for c in (1,2): ws.cell(3,c).font=head_font; ws.cell(3,c).alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); ws.cell(3,c).fill=fill_head
    ws.merge_cells("A3:A4"); ws.merge_cells("B3:B4")
    for i,d in enumerate(days):
        start,end=day_ranges[d]
        ws.merge_cells(start_row=3,start_column=start,end_row=3,end_column=end)
        cell=ws.cell(3,start); cell.value=f"{d}({(monday + timedelta(days=i)):%m/%d})"; cell.font=head_font; cell.alignment=Alignment(horizontal="center",vertical="center"); cell.fill=fill_day[i]
        for pno in periods[d]:
            cc=start+pno-1; pc=ws.cell(4,cc); pc.value=pno; pc.font=head_font; pc.alignment=Alignment(horizontal="center",vertical="center"); pc.fill=fill_day[i]
    teachers=sorted(st.session_state.timetable["교사명"].dropna().astype(str).str.strip().unique()) if not st.session_state.timetable.empty else []
    # 날짜별 effective timetable을 한 번만 만들고 (교사, 교시) 인덱스를 재사용한다.
    daily_indexes = {}
    for i, d in enumerate(days):
        ds=(monday + timedelta(days=i)).strftime("%Y-%m-%d")
        e=get_effective_timetable_for_date(ds,ver,use_test=use_test)
        daily_indexes[d] = {(str(r.교사명).strip(), safe_int(r.교시)): r for r in e.itertuples(index=False)} if not e.empty else {}
    row=5
    for n,t in enumerate(teachers,1):
        ws.merge_cells(start_row=row,start_column=1,end_row=row+1,end_column=1); ws.merge_cells(start_row=row,start_column=2,end_row=row+1,end_column=2)
        ws.cell(row,1).value=n; ws.cell(row,2).value=t
        for rr in (row,row+1):
            for cc in range(1,last_col+1):
                c=ws.cell(rr,cc); c.font=body_font; c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=Border(left=thin,right=thin,top=thin,bottom=thin)
        ws.cell(row,1).border=Border(left=med,right=thin,top=thin,bottom=thin); ws.cell(row,2).border=Border(left=thin,right=med,top=thin,bottom=thin)
        for i,d in enumerate(days):
            idx = daily_indexes[d]
            for pno in periods[d]:
                cc=day_ranges[d][0]+pno-1
                r=idx.get((t,pno))
                if r is None: continue
                subj=str(getattr(r,"과목","")).strip(); cls=str(getattr(r,"학급","")).strip(); typ=str(getattr(r,"변경유형","원본")).strip()
                ws.cell(row,cc).value=subj
                ws.cell(row+1,cc).value=cls
                if typ=="교환": fill=PatternFill("solid",fgColor="F4CCCC"); mark="🔄"
                elif typ=="보강": fill=PatternFill("solid",fgColor="D9EAD3"); mark="🟢"
                elif typ=="테스트교환": fill=PatternFill("solid",fgColor="EADCF8"); mark="🧪"
                elif typ=="시간강사": fill=PatternFill("solid",fgColor="FCE5CD"); mark="🟡"
                else: fill=fill_day[i]; mark=""
                ws.cell(row,cc).fill=fill; ws.cell(row+1,cc).fill=fill
                if mark: ws.cell(row,cc).value=f"{subj} {mark}".strip()
        row += 2
    ws.row_dimensions[1].height=22.5; ws.row_dimensions[2].height=14.25; ws.row_dimensions[3].height=18; ws.row_dimensions[4].height=16
    for rr in range(5,row): ws.row_dimensions[rr].height=18 if rr%2==1 else 16
    ws.column_dimensions["A"].width=5; ws.column_dimensions["B"].width=12
    for cc in range(3,last_col+1): ws.column_dimensions[get_column_letter(cc)].width=11
    ws.freeze_panes="C5"; ws.sheet_view.showGridLines=False
    ws.page_setup.orientation="landscape"; ws.page_setup.paperSize=ws.PAPERSIZE_A4; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0; ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_margins=PageMargins(left=0.25,right=0.25,top=0.4,bottom=0.4,header=0.2,footer=0.2)
    ws.print_title_rows="1:4"; ws.print_area=f"A1:{get_column_letter(last_col)}{row-1}"
    return _workbook_bytes(wb)



def build_weekly_class_schedule_excel_bytes(ref_date: date, *, use_test=False, title="전체 학급 시간표") -> bytes:
    """교사 주간표와 동일한 양식으로 학급 기준 주간표를 만든다.
    각 학급은 과목(첫 줄)·담당교사(둘째 줄)로 표시해 엑셀에서도 화면과 동일한 맥락을 유지한다.
    """
    monday = ref_date - timedelta(days=ref_date.weekday())
    ver = st.session_state.get("_data_version", 0)
    wb = Workbook(); ws = wb.active; ws.title = "전체 학급 시간표"
    title_font=Font(name="돋움",size=20,bold=True); small_font=Font(name="돋움",size=8); head_font=Font(name="돋움",size=9,bold=True); body_font=Font(name="돋움",size=8)
    thin=Side(style="hair",color="B7B7B7"); med=Side(style="medium",color="808080"); fill_head=PatternFill("solid",fgColor="D9EAF7")
    fill_day=[PatternFill("solid",fgColor=x) for x in ("DDEBF7","E2F0D9","FFF2CC","E4DFEC","FCE4D6")]
    periods={d:list(range(1,PERIODS_PER_DAY.get(d,7)+1)) for d in DAYS}; col=3; day_ranges={}
    for i,d in enumerate(DAYS):
        start=col; end=col+len(periods[d])-1; day_ranges[d]=(start,end); col=end+1
    last_col=col-1
    ws.merge_cells(start_row=1,start_column=1,end_row=1,end_column=last_col); ws["A1"]=title; ws["A1"].font=title_font; ws["A1"].alignment=Alignment(horizontal="center",vertical="center")
    ws.merge_cells(start_row=2,start_column=1,end_row=2,end_column=2); ws["A2"]=f"{SCHOOL_YEAR} 학년도"; ws["A2"].font=small_font
    ws.merge_cells(start_row=2,start_column=last_col-4,end_row=2,end_column=last_col); ws.cell(2,last_col-4).value=SCHOOL_NAME; ws.cell(2,last_col-4).font=small_font; ws.cell(2,last_col-4).alignment=Alignment(horizontal="right",vertical="center")
    ws["A3"]="번호"; ws["B3"]="학급"; ws.merge_cells("A3:A4"); ws.merge_cells("B3:B4")
    for c in (1,2): ws.cell(3,c).font=head_font; ws.cell(3,c).alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); ws.cell(3,c).fill=fill_head
    for i,d in enumerate(DAYS):
        start,end=day_ranges[d]; ws.merge_cells(start_row=3,start_column=start,end_row=3,end_column=end); cell=ws.cell(3,start); cell.value=f"{d}({(monday+timedelta(days=i)):%m/%d})"; cell.font=head_font; cell.alignment=Alignment(horizontal="center",vertical="center"); cell.fill=fill_day[i]
        for pno in periods[d]:
            pc=ws.cell(4,start+pno-1); pc.value=pno; pc.font=head_font; pc.alignment=Alignment(horizontal="center",vertical="center"); pc.fill=fill_day[i]
    daily_indexes={}
    classes=set()
    for i,d in enumerate(DAYS):
        ds=(monday+timedelta(days=i)).strftime("%Y-%m-%d"); e=get_effective_timetable_for_date(ds,ver,use_test=use_test); daily_indexes[d]={(str(r.학급).strip(),safe_int(r.교시)):r for r in e.itertuples(index=False) if str(getattr(r,"학급","")).strip()} if not e.empty else {}
        classes.update(k[0] for k in daily_indexes[d])
    row=5
    for n,cls in enumerate(sorted(classes),1):
        ws.merge_cells(start_row=row,start_column=1,end_row=row+1,end_column=1); ws.merge_cells(start_row=row,start_column=2,end_row=row+1,end_column=2); ws.cell(row,1).value=n; ws.cell(row,2).value=cls
        for rr in (row,row+1):
            for cc in range(1,last_col+1):
                c=ws.cell(rr,cc); c.font=body_font; c.alignment=Alignment(horizontal="center",vertical="center",wrap_text=True); c.border=Border(left=thin,right=thin,top=thin,bottom=thin)
        for i,d in enumerate(DAYS):
            for pno in periods[d]:
                r=daily_indexes[d].get((cls,pno)); cc=day_ranges[d][0]+pno-1
                if r is None: continue
                subj=str(getattr(r,"과목","")).strip(); teacher=str(getattr(r,"교사명","")).strip(); typ=str(getattr(r,"변경유형","원본")).strip(); ws.cell(row,cc).value=subj; ws.cell(row+1,cc).value=teacher
                fill=fill_day[i]; mark=""
                if typ=="교환": fill=PatternFill("solid",fgColor="F4CCCC"); mark=" 🔄"
                elif typ=="보강": fill=PatternFill("solid",fgColor="D9EAD3"); mark=" 🟢"
                elif typ=="테스트교환": fill=PatternFill("solid",fgColor="EADCF8"); mark=" 🧪"
                elif typ=="시간강사": fill=PatternFill("solid",fgColor="FCE5CD"); mark=" 🟡"
                ws.cell(row,cc).value=f"{subj}{mark}".strip(); ws.cell(row,cc).fill=fill; ws.cell(row+1,cc).fill=fill
        row+=2
    ws.row_dimensions[1].height=22.5; ws.row_dimensions[2].height=14.25; ws.row_dimensions[3].height=18; ws.row_dimensions[4].height=16
    for rr in range(5,row): ws.row_dimensions[rr].height=18 if rr%2==1 else 16
    ws.column_dimensions["A"].width=5; ws.column_dimensions["B"].width=12
    for cc in range(3,last_col+1): ws.column_dimensions[get_column_letter(cc)].width=11
    ws.freeze_panes="C5"; ws.sheet_view.showGridLines=False; ws.page_setup.orientation="landscape"; ws.page_setup.paperSize=ws.PAPERSIZE_A4; ws.page_setup.fitToWidth=1; ws.page_setup.fitToHeight=0; ws.sheet_properties.pageSetUpPr.fitToPage=True
    ws.page_margins=PageMargins(left=0.25,right=0.25,top=0.4,bottom=0.4,header=0.2,footer=0.2); ws.print_title_rows="1:4"; ws.print_area=f"A1:{get_column_letter(last_col)}{row-1}"
    return _workbook_bytes(wb)


def _workbook_bytes(wb):
    buf=io.BytesIO(); wb.save(buf); return buf.getvalue()

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

def _week_anchor(ref_date=None):
    ref = ref_date or _today_kst()
    if not isinstance(ref, date):
        try:
            ref = datetime.strptime(normalize_date_str(ref), "%Y-%m-%d").date()
        except Exception:
            ref = _today_kst()
    return ref - timedelta(days=ref.weekday())


def render_app_overview():
    """교사가 첫 화면에서 현재 업무 상태를 3초 안에 파악할 수 있는 경량 대시보드."""
    today = _today_kst().strftime("%Y-%m-%d")
    abs_df = st.session_state.get("absences", pd.DataFrame())
    sub_df = st.session_state.get("subs", pd.DataFrame())
    swap_df = st.session_state.get("swaps", pd.DataFrame())
    today_abs = 0 if abs_df.empty else int((abs_df["일자"].astype(str).map(normalize_date_str) == today).sum())
    today_sub = 0 if sub_df.empty else int((sub_df["일자"].astype(str).map(normalize_date_str) == today).sum())
    week_start = _week_anchor()
    week_end = week_start + timedelta(days=4)
    if swap_df.empty:
        week_swap = 0
    else:
        date_min, date_max = week_start.strftime("%Y-%m-%d"), week_end.strftime("%Y-%m-%d")
        dates = swap_df["원본일자"].astype(str).map(normalize_date_str) if "원본일자" in swap_df.columns else pd.Series("", index=swap_df.index)
        targets = swap_df["목표일자"].astype(str).map(normalize_date_str) if "목표일자" in swap_df.columns else pd.Series("", index=swap_df.index)
        week_swap = int((dates.between(date_min, date_max) | targets.between(date_min, date_max)).sum())
    changed_today = today_abs + today_sub
    st.markdown(f"""<div class='app-hero'><div class='app-hero-title'>📘 {SCHOOL_NAME} · {SCHOOL_YEAR}학년도</div>
    <div class='app-hero-sub'>{current_name()} · {current_role()} · 오늘 {today} · 주간표는 월~금 기준</div></div>""", unsafe_allow_html=True)
    cols = st.columns(5)
    metrics = [("📌 오늘 결강", today_abs), ("🟢 오늘 보강", today_sub), ("🔄 이번 주 교환", week_swap), ("⚠️ 오늘 처리대상", changed_today), ("👩‍🏫 등록 교사", len(st.session_state.get("teachers", pd.DataFrame())))]
    for c, (label, value) in zip(cols, metrics):
        with c:
            st.metric(label, value)
    st.markdown("<div class='ux-chip'>📅 달력으로 날짜 선택</div><div class='ux-chip'>🗓️ 주간표에서 수업 클릭</div><div class='ux-chip'>🔄 1:1 맞교환이 기본</div><div class='ux-chip'>🔗 연계 순환은 필요할 때만 계산</div>", unsafe_allow_html=True)


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

        plan_date = calendar_picker("계획서 기준일", _today_kst(), key="plan_date")
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
            rd = calendar_picker("전체 내역서 일자", _today_kst(), key="sidebar_rd")
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
render_app_overview()

allowed_tabs = get_user_allowed_tabs()
visible_tabs = [t for t in ALL_TABS if t in allowed_tabs]

if can_manage_ids():
    for t in ["🔑 아이디·권한 관리", "📑 회원별 탭 권한 관리", "🛠️ 다중 출장·전체 조정 추천"]:
        if t not in visible_tabs:
            visible_tabs.append(t)

if not visible_tabs:
    st.warning("접근 가능한 탭이 없습니다.")
    st.stop()

# Streamlit의 st.tabs()는 보이지 않는 탭의 본문도 같은 실행에서 모두 계산한다.
# 이 앱은 각 탭에 주간 매트릭스·Google Sheets 데이터·추천 계산이 많기 때문에,
# 메뉴 하나만 실제로 렌더링하는 방식으로 전환한다. UI는 탭과 같은 역할을 하되
# 비활성 화면을 계산하지 않아 클릭/팝업 반응성과 초기 로딩을 크게 줄인다.
if "active_tab" not in st.session_state or st.session_state.active_tab not in visible_tabs:
    st.session_state.active_tab = visible_tabs[0]
active_tab = st.selectbox(
    "업무 메뉴", visible_tabs,
    index=visible_tabs.index(st.session_state.active_tab),
    key="main_active_tab",
    help="한 번에 한 업무 화면만 렌더링합니다. 주간표와 후보 검색의 불필요한 백그라운드 계산을 줄입니다.",
)
st.session_state.active_tab = active_tab
tab_map = {active_tab: st.container()}
st.caption(f"현재 업무: **{active_tab}** · 다른 메뉴는 위에서 선택하세요.")

# 주간표 팝업은 현재 활성 화면의 주간표가 선택 상태를 기록한 뒤
# 스크립트 마지막에서 단 한 번 렌더링한다. 비활성 메뉴는 렌더링하지 않으므로
# 다른 메뉴의 dataframe selection이 현재 선택을 덮어쓰는 문제도 줄어든다.

# ------------------------------------------------------------------ 시간표 조회
if "시간표 조회" in tab_map:
    with tab_map["시간표 조회"]:
        st.subheader("📅 시간표 조회 — 달력 → 시간표 매트릭스")
        render_change_legend()
        view = st.radio("보기 방식", ["선택 날짜 매트릭스", "교사별 주간 매트릭스", "학급별 주간 매트릭스", "교사 1인 주간표"], horizontal=True, key="view_mode")
        ver = st.session_state.get("_data_version", 0)
        if view == "선택 날짜 매트릭스":
            picked, matrix, selections = daily_schedule_picker(_today_kst(), key="view_daily", height=560,
                help_text="날짜를 달력에서 선택한 후 교사×교시 셀을 클릭하세요.")
            if selections:
                st.markdown("#### 선택 수업 상세")
                e=get_effective_timetable_for_date(picked.strftime("%Y-%m-%d"),ver)
                details=[]
                for sel in selections:
                    m=e[(e["교사명"]==sel["교사명"]) & (e["교시"].apply(safe_int)==sel["교시"])]
                    if not m.empty:
                        r=m.iloc[0]
                        details.append({"교사":r["교사명"],"교시":sel["교시"],"학급":r["학급"],"과목":r["과목"],"변경유형":r.get("변경유형","원본"),"변경상세":r.get("변경상세","")})
                st.dataframe(pd.DataFrame(details),use_container_width=True,hide_index=True)
        elif view == "교사별 주간 매트릭스":
            ref=calendar_picker("주간 기준일",_today_kst(),key="view_week_ref")
            render_standard_weekly_matrix(effective_teacher_matrix(ref,ver,use_test=False), ref, row_label='교사명', key='view_teacher_week_matrix', title='교사별 주간 시간표', use_test=False)
            xlsx = build_weekly_schedule_excel_bytes(ref, use_test=False)
            st.download_button('📥 이 주간표 Excel 다운로드', xlsx, file_name=f'전체교사_주간시간표_{ref:%Y%m%d}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key='weekly_xlsx_teacher')
        elif view == "학급별 주간 매트릭스":
            ref=calendar_picker("주간 기준일",_today_kst(),key="view_class_ref")
            render_standard_weekly_matrix(class_matrix(ver, ref_date=ref), ref, row_label='학급', key='view_class_week_matrix', title='학급별 주간 시간표', use_test=False)
            xlsx = build_weekly_class_schedule_excel_bytes(ref, use_test=False)
            st.download_button('📥 이 주간표 Excel 다운로드', xlsx, file_name=f'전체학급_주간시간표_{ref:%Y%m%d}.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', key='weekly_xlsx_class')
        else:
            tlist=get_all_teacher_names()
            t=st.selectbox("교사 선택",tlist,key="view_t")
            ref=calendar_picker("주간 기준일",_today_kst(),key="view_ref")
            teacher_week = effective_teacher_matrix(ref, ver, use_test=False)
            if not teacher_week.empty and t:
                teacher_week = teacher_week[teacher_week["교사명"].astype(str).str.strip() == str(t).strip()].reset_index(drop=True)
            render_standard_weekly_matrix(teacher_week, ref, row_label="교사명", key="view_single_teacher_week_matrix", title=f"{t} 주간 시간표", use_test=False)

# ------------------------------------------------------------------ 시간강사 관리
if "시간강사 관리" in tab_map:
    with tab_map["시간강사 관리"]:
        st.subheader("시간강사 등록 및 가능 시간표")
        st.info("기본 입력은 달력과 가능 시간 매트릭스를 사용합니다. 아래 표는 기존 자료를 일괄 수정할 때 사용하는 고급 편집 영역입니다.")
        if can_full_data() or is_teacher():
            st.markdown("#### 📅 시간강사 빠른 등록")
            pt_names = st.session_state.part_time.get("시간강사명", pd.Series(dtype=str)).dropna().astype(str).str.strip().tolist() if not st.session_state.part_time.empty else []
            default_pt = pt_names[0] if pt_names else ""
            pt_name = st.text_input("시간강사명", value=default_pt, key="quick_pt_name")
            subjects = st.session_state.timetable.get("과목", pd.Series(dtype=str)).dropna().astype(str).unique().tolist()
            pt_subject = st.selectbox("담당과목", [""]+sorted(subjects), key="quick_pt_subject")
            orig_teachers = st.session_state.teachers["교사명"].dropna().astype(str).tolist() if not st.session_state.teachers.empty else []
            pt_orig = st.selectbox("대체교사", orig_teachers, key="quick_pt_orig")
            qstart,qend=calendar_range_picker(_today_kst(),_today_kst(),key="quick_pt_range",help_text="달력에서 대체 기간을 선택합니다.")
            st.markdown("가능한 교시를 클릭하세요")
            qcols=st.columns(5); qdays=DAYS
            avail={}
            for di,dayname in enumerate(qdays):
                with qcols[di]:
                    st.markdown(f"**{dayname}**")
                    for pp in range(1,8):
                        k=f"quick_pt_{dayname}_{pp}"
                        avail[(dayname,pp)]=st.checkbox(f"{pp}",key=k)
            if st.button("달력·매트릭스로 시간강사 추가",key="quick_pt_add",type="primary"):
                if not pt_name.strip() or not pt_orig.strip() or qstart>qend:
                    st.error("시간강사명·대체교사·기간을 확인하세요.")
                elif not any(avail.values()):
                    st.error("가능한 요일·교시를 하나 이상 선택하세요.")
                else:
                    row={c:"" for c in ensure_part_time_columns(pd.DataFrame()).columns}
                    row["시간강사명"]=pt_name.strip(); row["담당과목"]=pt_subject; row["대체교사"]=pt_orig; row["시작일"]=qstart.strftime("%Y-%m-%d"); row["종료일"]=qend.strftime("%Y-%m-%d")
                    for (dd,pp),flag in avail.items():
                        if flag: row[f"{dd}{pp}"]=1
                    candidate=ensure_part_time_columns(pd.concat([st.session_state.part_time,pd.DataFrame([row])],ignore_index=True))
                    ok,msg=validate_part_time_table(candidate)
                    if not ok: st.error(msg)
                    else:
                        st.session_state.part_time=candidate; save_work_data_to_gsheet(["시간강사"]); push_history("시간강사 추가"); st.success("추가 완료"); st.rerun()
            st.markdown("#### 🛠️ 기존 자료 고급 편집")
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
                candidate_pt = ensure_part_time_columns(ed)
                ok, msg = validate_part_time_table(candidate_pt)
                if not ok:
                    st.error(msg)
                else:
                    st.session_state.part_time = candidate_pt
                    save_work_data_to_gsheet(["시간강사"])
                    push_history("시간강사 수정")
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
            d_sel, abs_matrix, abs_cells = daily_schedule_picker(_today_kst(), key="abs_daily", multi=True, height=430,
                help_text="달력에서 날짜를 선택하고 결강할 수업 셀을 하나 이상 클릭하세요. 교사와 교시가 자동으로 결정됩니다.")
            on_date=d_sel.strftime("%Y-%m-%d"); day=WEEKDAY_KR[d_sel.weekday()]
            reason=st.selectbox("사유",ABSENCE_REASONS,key="abs_reason")
            detail=st.text_input("상세 사유",key="abs_detail")
            ver=st.session_state.get("_data_version",0)
            selected_abs={}
            for x in abs_cells:
                selected_abs.setdefault(x["교사명"],[]).append(x["교시"])
            if selected_abs:
                st.markdown("#### 선택한 결강 수업")
                st.write(" · ".join(f"{t}: {', '.join(map(str,sorted(ps)))}교시" for t,ps in selected_abs.items()))
            if st.button("결강 등록",type="primary",key="btn_abs"):
                if not selected_abs:
                    st.error("시간표 매트릭스에서 결강 수업을 하나 이상 선택하세요.")
                else:
                    a=st.session_state.absences
                    new_rows=[]
                    for who,sel_p in selected_abs.items():
                        cid=f"{on_date}-{who}"
                        if not a.empty:
                            a=a[~((a["일자"]==on_date)&(a["교사명"]==who))]
                        e_tt=get_effective_timetable_for_date(on_date,ver)
                        todays=e_tt[(e_tt["교사명"]==who)&(e_tt["요일"]==day)].sort_values("교시")
                        for r in todays[todays["교시"].apply(safe_int).isin(sel_p)].itertuples():
                            new_rows.append({"결강ID":cid,"일자":on_date,"요일":day,"교사명":who,"사유":reason,"상세사유":detail,"교시":safe_int(r.교시),"학급":r.학급,"과목":r.과목,"등록시각":datetime.now().strftime("%Y-%m-%d %H:%M"),"입력자":current_user()})
                    if new_rows:
                        st.session_state.absences=pd.concat([a,pd.DataFrame(new_rows)],ignore_index=True)
                        save_work_data_to_gsheet(["결강"]); push_history("결강 등록")
                        st.success(f"{len(new_rows)}건 결강 등록 완료"); st.rerun()
            show_abs=filter_by_owner(st.session_state.absences)
            st.dataframe(show_abs[show_abs["일자"]==on_date] if not show_abs.empty else pd.DataFrame(),height=250,hide_index=True)

        with right:
            st.markdown("### 📌 보강 배정")
            ab = filter_by_owner(st.session_state.absences)
            if ab.empty:
                st.info("먼저 결강을 등록하세요." if can_full_data() else "본인이 등록한 결강이 없습니다.")
            else:
                st.markdown("#### 📅 보강 대상 선택 — 달력 → 결강 매트릭스")
                sub_ref=calendar_picker("결강 날짜", _today_kst(), key="sub_ref")
                sub_date=sub_ref.strftime("%Y-%m-%d")
                day_abs=ab[ab["일자"].map(normalize_date_str)==sub_date].copy()
                if day_abs.empty:
                    st.info("선택한 날짜에 등록된 결강이 없습니다. 달력에서 다른 날짜를 선택하세요.")
                    cid=None; rows=pd.DataFrame(); head=None
                else:
                    teachers_for_day=sorted(day_abs["교사명"].dropna().astype(str).str.strip().unique())
                    idx={(str(r["교사명"]).strip(),safe_int(r["교시"])):r for _,r in day_abs.iterrows()}
                    sm_rows=[]
                    for teacher in teachers_for_day:
                        rr={"교사명":teacher}
                        for pp in range(1,MAX_PERIOD+1):
                            r=idx.get((teacher,pp))
                            rr[f"{pp}교시"]=(f"{r.get('학급','')} {r.get('과목','')} ❗".strip() if r is not None else "")
                        sm_rows.append(rr)
                    sub_matrix=pd.DataFrame(sm_rows,columns=["교사명"]+[f"{p}교시" for p in range(1,MAX_PERIOD+1)])
                    sub_event=st.dataframe(sub_matrix,hide_index=True,use_container_width=True,height=300,key=f"sub_abs_matrix_{sub_date}",on_select="rerun",selection_mode="single-cell")
                    sub_cells=getattr(getattr(sub_event,"selection",None),"cells",[]) or []
                    cid=None; rows=pd.DataFrame(); head=None
                    if sub_cells:
                        ri,cc=sub_cells[0]
                        if 0<=ri<len(sub_matrix) and cc!="교사명":
                            chosen_teacher=str(sub_matrix.iloc[ri]["교사명"]).strip(); chosen_period=safe_int(str(cc).replace("교시",""))
                            hit=day_abs[(day_abs["교사명"].astype(str).str.strip()==chosen_teacher)&(day_abs["교시"].apply(safe_int)==chosen_period)]
                            if not hit.empty:
                                cid=str(hit.iloc[0]["결강ID"]); rows=ab[ab["결강ID"]==cid].sort_values("교시"); head=rows.iloc[0]
                    if head is not None:
                        st.success(f"선택: **{head['일자']} · {head['교사명']} · {len(rows)}시간**")
                if head is not None:
                    include_pt = st.checkbox("시간강사 포함", key="sub_pt")
                    show_all = st.checkbox("전체 공강 교사 보기 (최대 20명)", value=True, key="show_all_free")

                    if st.button("전 교시 자동 배정", type="primary", key="auto_all"):
                        assignments = []
                        for r in rows.itertuples():
                            cand = recommend_substitutes(head["요일"], r.교시, r.과목, r.학급, head["교사명"], head["일자"], top_n=1, include_part_time=include_pt)
                            if not cand.empty:
                                assignments.append({"cid":cid,"on_date":head["일자"],"day":head["요일"],"period":r.교시,"class_name":r.학급,"subject":r.과목,"absent_teacher":head["교사명"],"sub_teacher":cand.iloc[0]["보강교사"],"method":"자동","priority":cand.iloc[0]["우선순위"],"memo":""})
                        accepted, errors = add_substitutes_batch(assignments, "자동 보강")
                        st.success(f"자동 보강 {accepted}건 처리 완료")
                        if errors: st.warning(" / ".join(errors[:5]))
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
        st.caption(
            "이제 별도의 출발점 지정이나 후보표를 먼저 열 필요가 없습니다. "
            "주간 시간표에서 수업 셀을 클릭하면 바로 작업 팝업이 열립니다."
        )

        week_anchor = calendar_picker(
            "교환 검색 기준 주", _today_kst(),
            key="exchange_week_anchor",
            help_text="선택한 날짜가 포함된 평일 주간 시간표를 표시합니다. 토·일은 표시하지 않습니다."
        )
        ver = st.session_state.get("_data_version", 0)

        # 실제 적용 시간표를 그대로 주간 매트릭스로 보여준다.
        # 셀 클릭은 render_weekly_matrix의 네이티브 selection → dialog 흐름을 사용한다.
        lesson_matrix = effective_teacher_matrix(week_anchor, ver, use_test=False)
        render_standard_weekly_matrix(
            lesson_matrix, week_anchor, row_label="교사명",
            key="exchange_weekly_matrix",
            title="📅 현재 적용 주간 시간표 — 수업을 클릭해서 바로 작업",
            use_test=False
        )

        st.divider()
        st.markdown("#### 📜 실제 맞교환 이력")
        show_swaps = filter_by_owner(st.session_state.swaps)
        if show_swaps.empty:
            st.info("등록된 실제 맞교환 이력이 없습니다.")
        else:
            st.dataframe(show_swaps, use_container_width=True, hide_index=True)

# ------------------------------------------------------------------ 통계
if "통계" in tab_map:
    with tab_map["통계"]:
        st.subheader("보강 통계")
        period = st.selectbox("빠른 기간", ["직접 선택", "전체", "1학기", "2학기", "이번 달"], key="st_p")
        if period == "직접 선택":
            start, end = calendar_range_picker(date(2026, 3, 1), _today_kst(), key="stats_range", help_text="달력에서 시작일과 종료일을 선택합니다.")
        else:
            start, end = date(2026, 3, 1), _today_kst()
        if period == "1학기":
            start, end = date(2026, 3, 1), date(2026, 7, 31)
        elif period == "2학기":
            start, end = date(2026, 8, 1), date(2027, 2, 28)
        elif period == "이번 달":
            start = _today_kst().replace(day=1)
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

        st.markdown("#### 수업 선택 — **현재 적용 + 테스트 변경 결과**에서 수업 셀 하나를 클릭")
        st.caption("실제 변경과 현재까지의 테스트 변경을 모두 반영합니다. 선택한 현재 상태를 기준으로 다음 1:1 가능 위치를 계산합니다.")
        test_week_anchor = calendar_picker("테스트 검색 기준 주", _today_kst(), key="test_week_anchor", help_text="선택한 날짜가 포함된 주간 시간표를 테스트 기준으로 사용합니다.")
        ver = st.session_state.get("_data_version", 0)
        test_matrix = effective_teacher_matrix(test_week_anchor, ver, use_test=True)
        st.caption("표시 기준: 🔄 교환 변경 이력 · 🟢/ [보강] 보강 처리 이력 · 테스트 변경도 함께 반영")
        # 테스트 화면도 별도의 선택표를 두지 않고, 위 주간 매트릭스 자체를 작업 시작점으로 사용한다.
        # 셀을 클릭하면 동일한 네이티브 팝업에서 1:1 맞교환 테스트와 연계 순환 테스트를 바로 선택한다.
        render_standard_weekly_matrix(
            test_matrix, test_week_anchor, row_label="교사명",
            key="test_week_preview", title="테스트 적용 주간표", use_test=True
        )

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
        ref = calendar_picker("기준 날짜", _today_kst(), key="chg_ref")
        changed = get_changed_teachers_for_week(ref)
        if not changed:
            st.success("이번 주차 변경 교사 없음")
        else:
            st.info(f"변경 교사 {len(changed)}명: {', '.join(changed)}")
            changed_week = effective_teacher_matrix(ref, st.session_state.get("_data_version", 0), use_test=False)
            for idx, t in enumerate(changed):
                with st.expander(f"👤 {t}", expanded=False):
                    st.caption("🔄 교환 · 🟢 보강 · 🟡 시간강사 이력이 셀에 표시됩니다.")
                    one = changed_week[changed_week["교사명"].astype(str).str.strip() == str(t).strip()].reset_index(drop=True) if not changed_week.empty else pd.DataFrame()
                    render_standard_weekly_matrix(one, ref, row_label="교사명", key=f"changed_teacher_week_{idx}", title=f"{t} 주간 시간표", use_test=False)

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

            t = st.selectbox("교사", teacher_options, key="duty_t")
            d, duty_matrix, duty_cells = daily_schedule_picker(_today_kst(), key="duty_daily", teacher_filter=t, multi=True, height=360,
                help_text="달력에서 날짜를 선택하고 해당 교사의 수업 셀을 클릭하세요. 빈 셀도 복무 시간 선택에 사용할 수 있도록 아래 교시 버튼을 제공합니다.")
            reason = st.selectbox("사유", ABSENCE_REASONS, key="duty_r")
            detail = st.text_input("상세", key="duty_det")
            all_day = st.checkbox("하루 전체", key="duty_all")
            if all_day:
                periods=[]
            else:
                periods=sorted({x["교시"] for x in duty_cells if x["교사명"]==t})
                if periods:
                    st.info("선택된 교시: " + ", ".join(f"{p}교시" for p in periods))
                else:
                    periods=period_matrix_picker("복무 교시", "duty_p_matrix", st.session_state.get("duty_p_matrix_sel", []))
                    st.session_state["duty_p_matrix_sel"]=periods

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
                    save_work_data_to_gsheet(["복무"])
                    push_history(f"복무 ({t})")
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
                                            "teacher": o.교사명, "class": str(other_class), "subject": str(o.과목),
                                            "lesson": f"{other_class} {o.과목}",
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
                                        if do_swap(a_info, b_info, d_str, c["date"]):
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

        start_date, end_date, start_periods, end_periods = range_calendar_matrix_picker(
            _today_kst(), _today_kst(), key="multi_range",
        )
        start_all = 0 in start_periods
        end_all = 0 in end_periods
        st.markdown("#### 🗓️ 선택 범위 미리보기")
        st.caption(f"{start_date:%Y-%m-%d} ({WEEKDAY_KR[start_date.weekday()]}) → {end_date:%Y-%m-%d} ({WEEKDAY_KR[end_date.weekday()]}) · 시작 {format_periods(start_periods)} / 종료 {format_periods(end_periods)}")

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

# ------------------------------------------------------------------ 주간표 공통 팝업 중앙 렌더러
# 모든 탭의 주간표가 먼저 렌더링되어 이번 실행의 최신 셀 선택을 session_state에 기록한 뒤
# 여기서 native dialog를 정확히 한 번만 호출한다. 따라서 클릭 직전의 이전 선택이 아니라
# 방금 클릭한 셀이 팝업에 표시된다. 또한 주간표 렌더러 안에서 dialog가 중복 생성되지 않는다.
if st.session_state.get("weekly_dialog_open") and st.session_state.get("weekly_selected_lesson"):
    _weekly_action_dialog()

st.caption(f"서라벌여중 시간표 관리 시스템 20260916 v2.1.0 · {current_name()} ({current_user()}) · {current_role()}")
