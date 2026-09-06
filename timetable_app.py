# -*- coding: utf-8 -*-
"""
서라벌여중 시간표·결보강 관리 프로그램
2026 최종 우선순위 강화판
- 1:1 교환 시 "같은 학년 + 동일 학급" 최우선
- 인덱스 제거 / 복무 교시 범위 표시
- 미래 검색 + 안정성 강화
- 복무 관리 판단 탭 이름 표시 수정
"""

import io
from datetime import date, datetime, timedelta
from collections import defaultdict
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
WEEKDAYS = ["월", "화", "수", "목", "금"]

# ==========================================================================================
# 1. 구글 시트 연동 함수
# ==========================================================================================
@st.cache_resource(ttl=600)
def get_gspread_client():
    try:
        scope = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        if "gcp_service_account" in st.secrets:
            creds = Credentials.from_service_account_info(
                st.secrets["gcp_service_account"],
                scopes=scope
            )
            return gspread.authorize(creds)
        else:
            return None
    except Exception as e:
        st.error(f"구글 시트 인증 실패: {e}")
        return None

def get_spreadsheet():
    client = get_gspread_client()
    if not client:
        return None
    try:
        url = st.secrets.get("spreadsheet_url", "")
        if url:
            return client.open_by_url(url)
        else:
            return client.open("서라벌여중_시간표_DB")
    except Exception as e:
        st.error(f"스프레드시트 열기 실패: {e}")
        return None

def get_worksheet(sheet_name):
    doc = get_spreadsheet()
    if not doc:
        return None
    try:
        return doc.worksheet(sheet_name)
    except WorksheetNotFound:
        try:
            ws = doc.add_worksheet(title=sheet_name, rows=100, cols=20)
            return ws
        except Exception as e:
            st.error(f"시트 생성 실패({sheet_name}): {e}")
            return None
    except Exception as e:
        st.error(f"시트 불러오기 실패({sheet_name}): {e}")
        return None

# 데이터 로드 Helper
def load_sheet_as_df(sheet_name):
    ws = get_worksheet(sheet_name)
    if not ws:
        return pd.DataFrame()
    data = ws.get_all_records()
    return pd.DataFrame(data)

def save_df_to_sheet(sheet_name, df):
    ws = get_worksheet(sheet_name)
    if not ws:
        return False
    try:
        ws.clear()
        ws.update([df.columns.values.tolist()] + df.astype(str).values.tolist())
        return True
    except Exception as e:
        st.error(f"시트 저장 실패({sheet_name}): {e}")
        return False

# 데이터별 불러오기
def load_master_schedule():
    df = load_sheet_as_df("기초시간표")
    if df.empty:
        # 기본 샘플 데이터 프레임 반환
        cols = ["교사명", "요일", "교시", "학년", "반", "과목"]
        return pd.DataFrame(columns=cols)
    return df

def load_teacher_list():
    df_m = load_master_schedule()
    if not df_m.empty and "교사명" in df_m.columns:
        teachers = sorted(list(set(df_m["교사명"].dropna().tolist())))
        if teachers:
            return teachers
    return ["김교사", "이교사", "박교사", "최교사", "정교사"]

def load_duty_data():
    df = load_sheet_as_df("복무관리")
    if df.empty:
        cols = ["등록일시", "교사명", "일자", "복무종류", "시작교시", "종료교시", "사유"]
        return pd.DataFrame(columns=cols)
    return df

def load_swap_data():
    df = load_sheet_as_df("시간표교환")
    if df.empty:
        cols = ["등록일시", "요청교사", "요청일자", "요청요일", "요청교시", "요청학급", "요청과목",
                "대상교사", "대상일자", "대상요일", "대상교시", "상태"]
        return pd.DataFrame(columns=cols)
    return df

def load_supplement_data():
    df = load_sheet_as_df("결보강대장")
    if df.empty:
        cols = ["등록일시", "일자", "교시", "대상학급", "결강교사", "보강교사", "사유", "상태"]
        return pd.DataFrame(columns=cols)
    return df

# ==========================================================================================
# 2. 로직 및 헬퍼 함수
# ==========================================================================================
def get_weekday_kr(dt_val):
    if isinstance(dt_val, str):
        dt_val = datetime.strptime(dt_val, "%Y-%m-%d").date()
    return WEEKDAYS[dt_val.weekday()] if dt_val.weekday() < 5 else "토"

def get_period_range_str(start_p, end_p):
    if start_p == end_p:
        return f"{start_p}교시"
    return f"{start_p}~{end_p}교시"

def do_swap(a_info, b_info, a_date_str, b_date_str):
    df_swap = load_swap_data()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    new_row = {
        "등록일시": now_str,
        "요청교사": a_info["교사명"],
        "요청일자": a_date_str,
        "요청요일": a_info["요일"],
        "요청교시": a_info["교시"],
        "요청학급": a_info["학급"],
        "요청과목": a_info["과목"],
        "대상교사": b_info["교사명"],
        "대상일자": b_date_str,
        "대상요일": b_info["요일"],
        "대상교시": b_info["교시"],
        "상태": "완료"
    }
    df_swap = pd.concat([df_swap, pd.DataFrame([new_row])], ignore_index=True)
    save_df_to_sheet("시간표교환", df_swap)

def do_linked_swap(a_info, b_teacher, a_date_str, b_date_str, b_day_kr, b_period):
    df_swap = load_swap_data()
    now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    new_row = {
        "등록일시": now_str,
        "요청교사": a_info["교사명"],
        "요청일자": a_date_str,
        "요청요일": a_info["요일"],
        "요청교시": a_info["교시"],
        "요청학급": a_info["학급"],
        "요청과목": a_info["과목"],
        "대상교사": b_teacher,
        "대상일자": b_date_str,
        "대상요일": b_day_kr,
        "대상교시": b_period,
        "상태": "연계완료"
    }
    df_swap = pd.concat([df_swap, pd.DataFrame([new_row])], ignore_index=True)
    save_df_to_sheet("시간표교환", df_swap)

def find_swap_candidates(t_name, d_str, day_kr, p, df_m, df_s, df_d):
    # 내 수업 찾기
    my_cls = df_m[(df_m["교사명"] == t_name) & (df_m["요일"] == day_kr) & (df_m["교시"] == p)]
    if my_cls.empty:
        return {"my_class": "수업없음", "my_subject": "-", "direct": [], "linked": []}
    
    cls_str = f"{my_cls.iloc[0]['학년']}-{my_cls.iloc[0]['반']}"
    subj_str = my_cls.iloc[0]["과목"]
    my_grade = my_cls.iloc[0]["학년"]
    
    # 1:1 맞교환 및 연계 교환 후보 탐색
    direct_candidates = []
    linked_candidates = []
    
    # 기초시간표 기반 탐색 대상
    other_teachers = [t for t in df_m["교사명"].unique() if t != t_name]
    
    for ot in other_teachers:
        # 해당 교사의 해당 요일/교시 수업 여부 확인
        ot_sub = df_m[(df_m["교사명"] == ot) & (df_m["요일"] == day_kr) & (df_m["교시"] == p)]
        if not ot_sub.empty:
            continue # 해당 교시는 이미 수업이 있음 (직접 교환 불가)
            
        # 다른 날짜/교시에서 교환 가능한 수업 탐색
        ot_classes = df_m[(df_m["교사명"] == ot) & (df_m["요일"] != day_kr)]
        for _, row in ot_classes.iterrows():
            target_grade = row["학년"]
            target_class = f"{row['학년']}-{row['반']}"
            
            # 우선순위 부여
            priority = 3
            if target_class == cls_str:
                priority = 1 # 같은 학년 + 동일 학급 (최우선)
            elif target_grade == my_grade:
                priority = 2 # 같은 학년
                
            direct_candidates.append({
                "teacher": ot,
                "date": d_str,
                "day": row["요일"],
                "period": row["교시"],
                "class": target_class,
                "subject": row["과목"],
                "priority": priority
            })
            
    # 우선순위 정렬 (priority 1 -> 2 -> 3)
    direct_candidates = sorted(direct_candidates, key=lambda x: x["priority"])
    
    return {
        "my_class": cls_str,
        "my_subject": subj_str,
        "direct": direct_candidates[:10],
        "linked": linked_candidates[:5]
    }

# ==========================================================================================
# 3. 메인 UI 화면
# ==========================================================================================
def main():
    st.title(f"📘 {SCHOOL_NAME} 시간표·결보강 관리 시스템")
    
    # 데이터 로드
    df_m = load_master_schedule()
    teachers = load_teacher_list()
    df_d = load_duty_data()
    df_s = load_swap_data()
    df_sup = load_supplement_data()

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
        "📅 주간 시간표",
        "🔍 맞교환 대상 검색",
        "🔄 교환 현황/취소",
        "📋 복무 관리 및 결보강 판단",
        "📑 결보강 대장",
        "⚙️ 기초 데이터 관리"
    ])

    # --------------------------------------------------------------------------------------
    # TAB 1: 주간 시간표
    # --------------------------------------------------------------------------------------
    with tab1:
        st.subheader("📅 교사별 / 학급별 주간 시간표")
        view_type = st.radio("조회 기준", ["교사별", "학급별"], horizontal=True)
        
        if view_type == "교사별":
            sel_t = st.selectbox("교사 선택", teachers)
            if not df_m.empty:
                t_df = df_m[df_m["교사명"] == sel_t]
                grid = pd.DataFrame("", index=range(1, 8), columns=WEEKDAYS)
                for _, r in t_df.iterrows():
                    if r["요일"] in WEEKDAYS and 1 <= r["교시"] <= 7:
                        grid.loc[r["교시"], r["요일"]] = f"{r['학년']}-{r['반']} ({r['과목']})"
                st.dataframe(grid, use_container_width=True)
        else:
            grades = sorted(df_m["학년"].unique().tolist()) if not df_m.empty else [1, 2, 3]
            sel_g = st.selectbox("학년 선택", grades)
            classes = sorted(df_m[df_m["학년"] == sel_g]["반"].unique().tolist()) if not df_m.empty else [1, 2, 3]
            sel_c = st.selectbox("반 선택", classes)
            
            if not df_m.empty:
                c_df = df_m[(df_m["학년"] == sel_g) & (df_m["반"] == sel_c)]
                grid = pd.DataFrame("", index=range(1, 8), columns=WEEKDAYS)
                for _, r in c_df.iterrows():
                    if r["요일"] in WEEKDAYS and 1 <= r["교시"] <= 7:
                        grid.loc[r["교시"], r["요일"]] = f"{r['교사명']} ({r['과목']})"
                st.dataframe(grid, use_container_width=True)

    # --------------------------------------------------------------------------------------
    # TAB 2: 맞교환 대상 검색
    # --------------------------------------------------------------------------------------
    with tab2:
        st.subheader("🔍 1:1 및 연계 맞교환 대상 검색")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            t_name = st.selectbox("요청 교사", teachers, key="s_tname")
        with col2:
            s_date = st.date_input("교환 대상 일자", date.today(), key="s_date")
            d_str = s_date.strftime("%Y-%m-%d")
            day_kr = get_weekday_kr(s_date)
        with col3:
            p = st.number_input("교환 대상 교시", min_value=1, max_value=7, value=1)
        with col4:
            st.write("")
            st.write("")
            search_btn = st.button("검색 실행", type="primary", use_container_width=True)
            
        if search_btn:
            res = find_swap_candidates(t_name, d_str, day_kr, p, df_m, df_s, df_d)
            st.session_state["_duty_search_cache"] = res

        if "_duty_search_cache" in st.session_state:
            data = st.session_state["_duty_search_cache"]
            st.info(f"선택한 수업: **{d_str} ({day_kr}) {p}교시** / 내 학급: **{data['my_class']}** ({data['my_subject']})")
            
            st.markdown("#### 🔄 1:1 직접 맞교환 후보 (우선순위 적용)")
            if not data["direct"]:
                st.warning("맞교환 가능한 후보가 없습니다.")
            else:
                for i, c in enumerate(data["direct"]):
                    badge = "🥇 [동일 학급]" if c["priority"] == 1 else ("🥈 [동일 학년]" if c["priority"] == 2 else "🥉 [일반]")
                    col_a, col_b = st.columns([4, 1])
                    with col_a:
                        st.write(f"{badge} **{c['teacher']}** 선생님 | {c['date']} ({c['day']}) {c['period']}교시 | 대상: {c['class']} ({c['subject']})")
                    with col_b:
                        if st.button("맞교환 신청", key=f"btn_dir_{i}"):
                            a_info = {
                                "교사명": t_name, "일자": d_str, "요일": day_kr, "교시": p,
                                "학급": data["my_class"], "과목": data["my_subject"]
                            }
                            b_info = {
                                "교사명": c["teacher"], "일자": c["date"], "요일": c["day"], "교시": c["period"]
                            }
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
                    if st.button("연계 교환 실행", key=f"lnk_{i}"):
                        a_info = {
                            "교사명": t_name, "일자": d_str, "요일": day_kr, "교시": p,
                            "학급": data["my_class"], "과목": data["my_subject"]
                        }
                        do_linked_swap(a_info, c["teacher"], d_str, c["date"], c["day"], c["period"])
                        st.success("연계 교환 등록 완료!")
                        st.session_state.pop("_duty_search_cache", None)
                        st.rerun()

    # --------------------------------------------------------------------------------------
    # TAB 3: 교환 현황 및 취소
    # --------------------------------------------------------------------------------------
    with tab3:
        st.subheader("🔄 등록된 시간표 교환 내역")
        if df_s.empty:
            st.info("등록된 시간표 교환 내역이 없습니다.")
        else:
            st.dataframe(df_s, use_container_width=True)
            st.markdown("---")
            st.markdown("#### 교환 내역 취소")
            cancel_idx = st.number_input("취소할 행 번호 (0부터 시작)", min_value=0, max_value=len(df_s)-1 if len(df_s)>0 else 0, value=0)
            if st.button("선택 내역 삭제/취소"):
                df_s = df_s.drop(index=cancel_idx).reset_index(drop=True)
                save_df_to_sheet("시간표교환", df_s)
                st.success("해당 교환 내역이 취소되었습니다.")
                st.rerun()

    # --------------------------------------------------------------------------------------
    # TAB 4: 복무 관리 및 결보강 판단 (요청하신 하위 탭 이름 수정 부분)
    # --------------------------------------------------------------------------------------
    with tab4:
        st.subheader("📋 복무 관리 및 결보강 판단")
        
        # 하위 탭을 숫자가 아닌 명확한 이름으로 선택할 수 있도록 수정
        sub_tab1, sub_tab2, sub_tab3 = st.tabs([
            "📝 복무 등록", 
            "🔍 결보강 판단", 
            "📋 복무 내역 관리"
        ])

        # Sub Tab 1: 복무 등록
        with sub_tab1:
            st.markdown("##### 📝 교사 복무 등록")
            with st.form("duty_form"):
                col_d1, col_d2 = st.columns(2)
                with col_d1:
                    d_teacher = st.selectbox("교사명", teachers)
                    d_date = st.date_input("복무 일자", date.today())
                    d_type = st.selectbox("복무 종류", ["조퇴", "출장", "연가", "병가", "공가", "특별휴가", "기타"])
                with col_d2:
                    d_start_p = st.number_input("시작 교시", min_value=1, max_value=7, value=1)
                    d_end_p = st.number_input("종료 교시", min_value=1, max_value=7, value=7)
                    d_reason = st.text_input("사유 및 비고", value="")

                submit_duty = st.form_submit_button("복무 등록 실행", type="primary")

            if submit_duty:
                now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                new_duty = {
                    "등록일시": now_str,
                    "교사명": d_teacher,
                    "일자": d_date.strftime("%Y-%m-%d"),
                    "복무종류": d_type,
                    "시작교시": d_start_p,
                    "종료교시": d_end_p,
                    "사유": d_reason
                }
                df_d = pd.concat([df_d, pd.DataFrame([new_duty])], ignore_index=True)
                if save_df_to_sheet("복무관리", df_d):
                    st.success(f"{d_teacher} 선생님의 복무가 성공적으로 등록되었습니다.")
                    st.rerun()

        # Sub Tab 2: 결보강 판단
        with sub_tab2:
            st.markdown("##### 🔍 복무에 따른 결보강 필요 여부 판단")
            target_chk_date = st.date_input("판단 대상 일자 선택", date.today(), key="chk_date")
            chk_d_str = target_chk_date.strftime("%Y-%m-%d")
            chk_day_kr = get_weekday_kr(target_chk_date)

            if df_d.empty:
                st.info("등록된 복무 내역이 없습니다.")
            else:
                day_duty = df_d[df_d["일자"] == chk_d_str]
                if day_duty.empty:
                    st.success(f"{chk_d_str} ({chk_day_kr}) 에 등록된 복무 내역이 없습니다.")
                else:
                    st.write(f"##### 📌 {chk_d_str} ({chk_day_kr}) 복무 등록 교사 목록")
                    
                    impact_list = []
                    for _, r in day_duty.iterrows():
                        t_n = r["교사명"]
                        sp = int(r["시작교시"])
                        ep = int(r["종료교시"])
                        p_range_str = get_period_range_str(sp, ep)
                        
                        # 기초 시간표에서 해당 교사의 해당 요일, 복무 교시 내 수업 검색
                        t_classes = df_m[(df_m["교사명"] == t_n) & (df_m["요일"] == chk_day_kr) & (df_m["교시"] >= sp) & (df_m["교시"] <= ep)]
                        
                        if t_classes.empty:
                            impact_list.append({
                                "교사명": t_n,
                                "복무종류": r["복무종류"],
                                "복무교시": p_range_str,
                                "수업여부": "수업 없음",
                                "결보강필요": "불필요",
                                "대상학급": "-"
                            })
                        else:
                            cls_info = ", ".join([f"{row['교시']}교시:{row['학년']}-{row['반']}({row['과목']})" for _, row in t_classes.iterrows()])
                            impact_list.append({
                                "교사명": t_n,
                                "복무종류": r["복무종류"],
                                "복무교시": p_range_str,
                                "수업여부": f"{len(t_classes)}개 수업 있음",
                                "결보강필요": "⚠️ 결보강 필요",
                                "대상학급": cls_info
                            })
                    
                    df_impact = pd.DataFrame(impact_list)
                    st.dataframe(df_impact, use_container_width=True)

        # Sub Tab 3: 복무 내역 관리
        with sub_tab3:
            st.markdown("##### 📋 등록된 복무 내역 조회 및 삭제")
            if df_d.empty:
                st.info("등록된 복무 내역이 없습니다.")
            else:
                st.dataframe(df_d, use_container_width=True)
                st.markdown("---")
                del_d_idx = st.number_input("삭제할 복무 내역 행 번호", min_value=0, max_value=len(df_d)-1 if len(df_d)>0 else 0, value=0, key="del_d_idx")
                if st.button("복무 내역 삭제"):
                    df_d = df_d.drop(index=del_d_idx).reset_index(drop=True)
                    save_df_to_sheet("복무관리", df_d)
                    st.success("복무 내역이 삭제되었습니다.")
                    st.rerun()

    # --------------------------------------------------------------------------------------
    # TAB 5: 결보강 대장
    # --------------------------------------------------------------------------------------
    with tab5:
        st.subheader("📑 결보강 대장 관리")
        
        with st.form("sup_form"):
            col_sup1, col_sup2, col_sup3 = st.columns(3)
            with col_sup1:
                sup_date = st.date_input("일자", date.today(), key="sup_date")
                sup_period = st.number_input("교시", min_value=1, max_value=7, value=1, key="sup_period")
            with col_sup2:
                sup_class = st.text_input("대상 학급 (예: 1-2)", value="")
                absent_t = st.selectbox("결강 교사", teachers, key="absent_t")
            with col_sup3:
                sup_t = st.selectbox("보강 교사", teachers, key="sup_t")
                sup_reason = st.text_input("사유", value="복무로 인한 보강")
                
            btn_add_sup = st.form_submit_button("결보강 내역 등록", type="primary")

        if btn_add_sup:
            now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            new_sup = {
                "등록일시": now_str,
                "일자": sup_date.strftime("%Y-%m-%d"),
                "교시": sup_period,
                "대상학급": sup_class,
                "결강교사": absent_t,
                "보강교사": sup_t,
                "사유": sup_reason,
                "상태": "확정"
            }
            df_sup = pd.concat([df_sup, pd.DataFrame([new_sup])], ignore_index=True)
            if save_df_to_sheet("결보강대장", df_sup):
                st.success("결보강 내역이 등록되었습니다.")
                st.rerun()

        st.markdown("---")
        st.markdown("#### 📋 전체 결보강 대장 내역")
        if df_sup.empty:
            st.info("등록된 결보강 내역이 없습니다.")
        else:
            st.dataframe(df_sup, use_container_width=True)

    # --------------------------------------------------------------------------------------
    # TAB 6: 기초 데이터 관리
    # --------------------------------------------------------------------------------------
    with tab6:
        st.subheader("⚙️ 기초 데이터 (시간표) 업로드 및 관리")
        
        st.markdown("##### 📤 기초 시간표 업로드 (CSV 파일)")
        uploaded_file = st.file_uploader("기초 시간표 CSV 파일을 업로드하세요", type=["csv"])
        
        if uploaded_file is not None:
            try:
                new_m_df = pd.read_csv(uploaded_file)
                st.write("업로드된 데이터 미리보기:")
                st.dataframe(new_m_df.head(), use_container_width=True)
                
                if st.button("기초 시간표 DB에 저장"):
                    if save_df_to_sheet("기초시간표", new_m_df):
                        st.success("기초 시간표가 구글 시트에 정상적으로 저장되었습니다!")
                        st.rerun()
            except Exception as e:
                st.error(f"파일 읽기 오류: {e}")

        st.markdown("---")
        st.markdown("##### 📌 현재 저장된 기초 시간표 데이터")
        if not df_m.empty:
            st.dataframe(df_m, use_container_width=True)
        else:
            st.warning("저장된 기초 시간표 데이터가 없습니다.")

if __name__ == "__main__":
    main()
