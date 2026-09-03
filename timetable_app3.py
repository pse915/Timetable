import streamlit as st
import pandas as pd
from datetime import datetime, date
import time

# ==========================================================================================
# 1. 페이지 설정 및 데이터 캐싱 (API 429 Quota Exceeded 오류 방지)
# ==========================================================================================
st.set_page_config(page_title="스마트 시간표 및 결보강 관리 시스템", layout="wide")

WEEKDAY_KR = {0: "월", 1: "화", 2: "수", 3: "목", 4: "금", 5: "토", 6: "일"}
PERIODS_PER_DAY = {"월": 7, "화": 7, "수": 6, "목": 7, "금": 6}
MAX_PERIOD = 7

# Google Sheets API 429 오류 방지: 데이터를 5분간 캐싱하여 반복적인 API 호출 제한
@st.cache_data(ttl=300)
def load_initial_teachers():
    return pd.DataFrame([
        {"교사명": "김철수", "담당과목": "수학", "담당학년": "3", "담임학급": "3-1", "과목군": "수학"},
        {"교사명": "이영희", "담당과목": "국어", "담당학년": "3", "담임학급": "3-2", "과목군": "국어"},
        {"교사명": "박민수", "담당과목": "영어", "담당학년": "1,3", "담임학급": "1-1", "과목군": "영어"},
        {"교사명": "정수진", "담당과목": "수학", "담당학년": "2", "담임학급": "2-1", "과목군": "수학"},
        {"교사명": "최동현", "담당과목": "체육", "담당학년": "3", "담임학급": "3-3", "과목군": "체육"},
        {"교사명": "한지민", "담당과목": "국어", "담당학년": "1,2", "담임학급": "1-2", "과목군": "국어"},
    ])

@st.cache_data(ttl=300)
def load_initial_timetable():
    return pd.DataFrame([
        {"교사명": "김철수", "요일": "월", "교시": 1, "과목": "수학", "학급": "3-1"},
        {"교사명": "김철수", "요일": "월", "교시": 3, "과목": "수학", "학급": "3-2"},
        {"교사명": "김철수", "요일": "금", "교시": 1, "과목": "수학", "학급": "3-1"},
        {"교사명": "이영희", "요일": "월", "교시": 2, "과목": "국어", "학급": "3-2"},
        {"교사명": "이영희", "요일": "금", "교시": 3, "과목": "국어", "학급": "3-2"},
        {"교사명": "박민수", "요일": "금", "교시": 3, "과목": "영어", "학급": "1-2"},
        {"교사명": "박민수", "요일": "월", "교시": 4, "과목": "영어", "학급": "3-1"},
        {"교사명": "정수진", "요일": "금", "교시": 3, "과목": "수학", "학급": "2-1"},
        {"교사명": "최동현", "요일": "금", "교시": 2, "과목": "체육", "학급": "3-3"},
        {"교사명": "한지민", "요일": "월", "교시": 1, "과목": "국어", "학급": "1-1"},
    ])

def init_session_state():
    if "teachers" not in st.session_state:
        st.session_state.teachers = load_initial_teachers().copy()
    if "timetable" not in st.session_state:
        st.session_state.timetable = load_initial_timetable().copy()
    if "substitutions" not in st.session_state:
        st.session_state.substitutions = pd.DataFrame(
            columns=["ID", "날짜", "요일", "교시", "학급", "원과목", "원교사", "대리교사", "유형", "상태", "비고"]
        )
    if "swaps" not in st.session_state:
        st.session_state.swaps = pd.DataFrame(
            columns=["변경일시", "교사A", "수업A_일시", "수업A_내용", "교사B", "수업B_일시", "수업B_내용", "상태"]
        )

init_session_state()

# ==========================================================================================
# 2. 헬퍼 함수
# ==========================================================================================
def grade_of(class_name: str) -> str:
    if isinstance(class_name, str) and "-" in class_name:
        return class_name.split("-")[0]
    return ""

def subject_group(subject_name: str) -> str:
    if not isinstance(subject_name, str):
        return ""
    if any(k in subject_name for k in ["수학", "수학A", "수학B"]): return "수학"
    if any(k in subject_name for k in ["국어", "문학", "독서"]): return "국어"
    if any(k in subject_name for k in ["영어", "영어I", "영어II"]): return "영어"
    if any(k in subject_name for k in ["체육", "운동"]): return "체육"
    return subject_name

def is_free(teacher: str, day: str, period: int, date_str: str) -> bool:
    tt = st.session_state.timetable
    has_tt = not tt[(tt["교사명"] == teacher) & (tt["요일"] == day) & (tt["교시"] == period)].empty
    if has_tt:
        return False
    
    subs = st.session_state.substitutions
    if not subs.empty:
        has_sub = not subs[(subs["대리교사"] == teacher) & (subs["날짜"] == date_str) & (subs["교시"] == period)].empty
        if has_sub:
            return False
    return True

def cumulative_sub_count() -> dict:
    subs = st.session_state.substitutions
    if subs.empty:
        return {}
    return subs["대리교사"].value_counts().to_dict()

def validate_swap(lesson_a: dict, lesson_b: dict):
    errors = []
    guides = []
    tt = st.session_state.timetable
    
    a_conflict = tt[
        (tt["교사명"] == lesson_a["교사명"]) & 
        (tt["요일"] == lesson_b["요일"]) & 
        (tt["교시"] == lesson_b["교시"]) & 
        (tt["학급"] != lesson_a["학급"])
    ]
    if not a_conflict.empty:
        errors.append(f"{lesson_a['교사명']} 교사가 이동 희망 시간({lesson_b['요일']} {lesson_b['교시']}교시)에 이미 수업({a_conflict.iloc[0]['학급']})이 있습니다.")
        guides.append("교사 A의 동시간대 기존 수업을 먼저 조정하세요.")

    b_conflict = tt[
        (tt["교사명"] == lesson_b["교사명"]) & 
        (tt["요일"] == lesson_a["요일"]) & 
        (tt["교시"] == lesson_a["교시"]) & 
        (tt["학급"] != lesson_b["학급"])
    ]
    if not b_conflict.empty:
        errors.append(f"{lesson_b['교사명']} 교사가 원본 시간({lesson_a['요일']} {lesson_a['교시']}교시)에 이미 수업({b_conflict.iloc[0]['학급']})이 있습니다.")
        guides.append("교사 B의 해당 시간대 일정을 확인하세요.")

    return errors, guides

def do_swap(lesson_a: dict, lesson_b: dict, date_a_str: str, date_b_str: str):
    tt = st.session_state.timetable
    
    idx_a = tt[(tt["교사명"] == lesson_a["교사명"]) & (tt["요일"] == lesson_a["요일"]) & (tt["교시"] == lesson_a["교시"])].index
    idx_b = tt[(tt["교사명"] == lesson_b["교사명"]) & (tt["요일"] == lesson_b["요일"]) & (tt["교시"] == lesson_b["교시"])].index
    
    if len(idx_a) > 0 and len(idx_b) > 0:
        tt.loc[idx_a[0], "요일"] = lesson_b["요일"]
        tt.loc[idx_a[0], "교시"] = lesson_b["교시"]
        
        tt.loc[idx_b[0], "요일"] = lesson_a["요일"]
        tt.loc[idx_b[0], "교시"] = lesson_a["교시"]
        
        new_swap = {
            "변경일시": datetime.now().strftime("%Y-%m-%d %H:%M"),
            "교사A": lesson_a["교사명"],
            "수업A_일시": f"{date_a_str}({lesson_a['요일']}) {lesson_a['교시']}교시 ➔ {date_b_str}({lesson_b['요일']}) {lesson_b['교시']}교시",
            "수업A_내용": f"{lesson_a['학급']} {lesson_a['과목']}",
            "교사B": lesson_b["교사명"],
            "수업B_일시": f"{date_b_str}({lesson_b['요일']}) {lesson_b['교시']}교시 ➔ {date_a_str}({lesson_a['요일']}) {lesson_a['교시']}교시",
            "수업B_내용": f"{lesson_b['학급']} {lesson_b['과목']}",
            "상태": "완료"
        }
        st.session_state.swaps = pd.concat([st.session_state.swaps, pd.DataFrame([new_swap])], ignore_index=True)
        return True
    return False

def add_substitute(cid, date_str, day_str, period, cls_str, subj_str, orig_t, sub_t, stype, status, note):
    new_sub = {
        "ID": cid, "날짜": date_str, "요일": day_str, "교시": period,
        "학급": cls_str, "원과목": subj_str, "원교사": orig_t,
        "대리교사": sub_t, "유형": stype, "상태": status, "비고": note
    }
    st.session_state.substitutions = pd.concat([st.session_state.substitutions, pd.DataFrame([new_sub])], ignore_index=True)

# ==========================================================================================
# 3. 맞교환 & 공강 추천 로직
# ==========================================================================================
def get_target_time_recommendations_v2(
    teacher_a: str, date_a_str: str, period_a: int, class_a: str, subject_a: str,
    date_b_str: str, period_b: int
):
    ti = st.session_state.teachers
    tt = st.session_state.timetable
    day_a = WEEKDAY_KR[datetime.strptime(date_a_str, "%Y-%m-%d").weekday()]
    day_b = WEEKDAY_KR[datetime.strptime(date_b_str, "%Y-%m-%d").weekday()]
    
    lesson_a = {
        "교사명": teacher_a, "일자": date_a_str, "요일": day_a, "교시": period_a,
        "학급": class_a, "과목": subject_a
    }
    
    grp_a = subject_group(subject_a)
    grade_a = grade_of(class_a)
    cum = cumulative_sub_count()
    
    all_recommendations = []
    
    for _, t_row in ti.iterrows():
        t_b = t_row["교사명"]
        if t_b == teacher_a:
            continue
            
        tb_grades = str(t_row.get("담당학년", ""))
        tb_subject = str(t_row.get("담당과목", ""))
        is_same_grade = (grade_a != "") and (grade_a in tb_grades)
        
        b_lessons = tt[(tt["교사명"] == t_b) & (tt["요일"] == day_b) & (tt["교시"] == period_b)]
        
        if not b_lessons.empty:
            for _, b_row in b_lessons.iterrows():
                b_item = {
                    "교사명": t_b, "일자": date_b_str, "요일": day_b, "교시": period_b,
                    "학급": b_row["학급"], "과목": b_row["과목"]
                }
                errs, guides = validate_swap(lesson_a, b_item)
                is_valid = (len(errs) == 0)
                
                score = 1000 if is_valid else -500
                if is_same_grade:
                    score += 300
                if subject_group(b_row["과목"]) == grp_a:
                    score += 200
                score -= cum.get(t_b, 0) * 10
                
                grade_tag = f"🌟 {grade_a}학년 담당" if is_same_grade else "타 학년"
                
                all_recommendations.append({
                    "교사명": t_b,
                    "구분": "🔄 1:1 맞교환 대상",
                    "학년구분": grade_tag,
                    "담당과목": tb_subject,
                    "상대수업": f"{b_row['학급']} ({b_row['과목']})",
                    "is_valid": is_valid,
                    "is_same_grade": is_same_grade,
                    "_score": score,
                    "b_info": b_item,
                    "errs": errs,
                    "guides": guides,
                    "type": "SWAP"
                })
        else:
            if is_free(t_b, day_b, period_b, date_b_str):
                score = 500
                if is_same_grade:
                    score += 400
                    prio_label = f"🌟 {grade_a}학년 공강"
                elif grp_a and grp_a in str(t_row.get("과목군", "")):
                    score += 200
                    prio_label = "동일 과목군"
                else:
                    prio_label = "일반 공강"
                score -= cum.get(t_b, 0) * 10
                
                grade_tag = f"🌟 {grade_a}학년 담당" if is_same_grade else "타 학년"
                
                all_recommendations.append({
                    "교사명": t_b,
                    "구분": "➕ 공강 (보강/대리 가능)",
                    "학년구분": grade_tag,
                    "담당과목": tb_subject,
                    "상대수업": f"공강 ({prio_label})",
                    "is_valid": True,
                    "is_same_grade": is_same_grade,
                    "_score": score,
                    "b_info": None,
                    "errs": [],
                    "guides": [],
                    "type": "FREE"
                })

    df_all = pd.DataFrame(all_recommendations)
    if not df_all.empty:
        df_all = df_all.sort_values("_score", ascending=False).reset_index(drop=True)
    return df_all, grade_a

# ==========================================================================================
# 4. Streamlit UI
# ==========================================================================================
st.title("🏫 스마트 시간표 및 결보강 관리 시스템")

# 상단 데이터 새로고침 (캐시 초기화 버튼)
top_col1, top_col2 = st.columns([5, 1])
with top_col2:
    if st.button("🔄 데이터 새로고침", help="API 데이터 캐시를 초기화합니다."):
        st.cache_data.clear()
        st.rerun()

tabs = st.tabs([
    "📋 전체 시간표", 
    "👤 교사별 시간표", 
    "🏫 학급별 시간표", 
    "🚨 결강/보강 관리", 
    "🔄 스마트 맞교환 & 공강 추천"
])

# ------------------------------------------------------------------ TAB 1
with tabs[0]:
    st.subheader("📋 전체 교사 시간표")
    st.dataframe(st.session_state.timetable, use_container_width=True)

# ------------------------------------------------------------------ TAB 2
with tabs[1]:
    st.subheader("👤 교사별 개인 시간표")
    sel_teacher = st.selectbox("교사 선택", st.session_state.teachers["교사명"].unique(), key="t_view_sel")
    t_df = st.session_state.timetable[st.session_state.timetable["교사명"] == sel_teacher]
    
    matrix = pd.DataFrame("", index=[f"{p}교시" for p in range(1, MAX_PERIOD + 1)], columns=["월", "화", "수", "목", "금"])
    for _, r in t_df.iterrows():
        matrix.loc[f"{r['교시']}교시", r["요일"]] = f"{r['학급']}\n({r['과목']})"
    st.table(matrix)

# ------------------------------------------------------------------ TAB 3
with tabs[2]:
    st.subheader("🏫 학급별 시간표")
    all_classes = sorted(list(st.session_state.timetable["학급"].unique()))
    sel_class = st.selectbox("학급 선택", all_classes, key="c_view_sel")
    c_df = st.session_state.timetable[st.session_state.timetable["학급"] == sel_class]
    
    c_matrix = pd.DataFrame("", index=[f"{p}교시" for p in range(1, MAX_PERIOD + 1)], columns=["월", "화", "수", "목", "금"])
    for _, r in c_df.iterrows():
        c_matrix.loc[f"{r['교시']}교시", r["요일"]] = f"{r['과목']}\n({r['교사명']})"
    st.table(c_matrix)

# ------------------------------------------------------------------ TAB 4
with tabs[3]:
    st.subheader("🚨 결강 등록 및 대리 교사 배정 이력")
    st.dataframe(st.session_state.substitutions, use_container_width=True)

# ------------------------------------------------------------------ TAB 5: 직관적인 카드형 UI 적용
with tabs[4]:
    st.markdown("### 🔄 스마트 시간표 변경 & 학년별 맞교환/보강 통합 추천")
    st.caption("변경할 수업과 이동 희망 시간을 선택하면 교체 가능한 수업과 맞교환/보강 내역이 한눈에 표시됩니다.")
    
    col_a, col_b = st.columns(2)
    
    with col_a:
        st.markdown("#### 1️⃣ 원본 수업 선택 (교사 A)")
        date_a_input = st.date_input("원본 날짜 (Date A)", value=date.today(), key="date_a")
        date_a_str = date_a_input.strftime("%Y-%m-%d")
        day_a = WEEKDAY_KR[date_a_input.weekday()]
        
        t_a = st.selectbox("교사 A 선택", st.session_state.teachers["교사명"], key="teacher_a_select")
        
        sub_a = st.session_state.timetable[
            (st.session_state.timetable["교사명"] == t_a) & 
            (st.session_state.timetable["요일"] == day_a)
        ].sort_values("교시")
        
        if sub_a.empty:
            st.warning(f"⚠️ {t_a} 교사는 {date_a_str} ({day_a}요일)에 수업이 없습니다.")
            pick_a = None
        else:
            opts_a = [f"{int(r['교시'])}교시 · {r['학급']} · {r['과목']}" for _, r in sub_a.iterrows()]
            sel_a = st.selectbox(f"변경할 수업 선택 ({day_a}요일 수업)", opts_a, key="lesson_a_select")
            row_a = sub_a.iloc[opts_a.index(sel_a)]
            pick_a = {
                "교사명": t_a, "일자": date_a_str, "요일": day_a, 
                "교시": int(row_a["교시"]), "학급": row_a["학급"], "과목": row_a["과목"]
            }

    with col_b:
        st.markdown("#### 2️⃣ 이동 희망 날짜 & 시간 선택 (Target Time)")
        date_b_input = st.date_input("이동 희망 날짜 (Date B)", value=date.today(), key="date_b")
        date_b_str = date_b_input.strftime("%Y-%m-%d")
        day_b = WEEKDAY_KR[date_b_input.weekday()]
        
        p_b_input = st.selectbox("희망 교시 (Target Period)", list(range(1, PERIODS_PER_DAY.get(day_b, MAX_PERIOD) + 1)), key="period_b_select")
        st.info(f"🎯 **목표 이동 대상 시간**: [{date_b_str} ({day_b})] {p_b_input}교시")

    st.divider()

    if pick_a:
        df_recs, target_grade = get_target_time_recommendations_v2(
            pick_a["교사명"], pick_a["일자"], pick_a["교시"], pick_a["학급"], pick_a["과목"],
            date_b_str, p_b_input
        )
        
        st.markdown(f"#### 💡 추천 대상 교사 및 교체 가능 수업 목록 ({target_grade}학년 기준)")

        c_flt1, c_flt2 = st.columns([2, 2])
        only_same_grade = c_flt1.checkbox(f"🌟 {target_grade}학년 담당 교사만 보기", value=False)
        filter_type = c_flt2.radio("구분 필터", ["전체 보기", "🔄 맞교환(수업 중)", "➕ 공강 교사"], horizontal=True)

        view_df = df_recs.copy()
        if only_same_grade:
            view_df = view_df[view_df["is_same_grade"] == True]
        if filter_type == "🔄 맞교환(수업 중)":
            view_df = view_df[view_df["type"] == "SWAP"]
        elif filter_type == "➕ 공강 교사":
            view_df = view_df[view_df["type"] == "FREE"]

        if view_df.empty:
            st.warning("조건에 해당하는 교사가 없습니다.")
        else:
            # 💡 가독성을 높인 직관적인 카드형 표시
            for idx, row in view_df.iterrows():
                with st.container(border=True):
                    c1, c2, c3 = st.columns([2, 4, 2])
                    
                    with c1:
                        st.markdown(f"### **{row['교사명']}** 교사")
                        st.caption(f"{row['구분']} | {row['학년구분']}")
                        st.write(f"**담당 과목**: {row['담당과목']}")
                    
                    with c2:
                        st.markdown("  **🔄 시간표 교체 내역**")
                        if row["type"] == "SWAP":
                            st.write(f"• **{pick_a['교사명']}**: {pick_a['일자']}({pick_a['요일']}) {pick_a['교시']}교시 [{pick_a['학급']} {pick_a['과목']}] ➔ **{date_b_str}({day_b}) {p_b_input}교시**")
                            st.write(f"• **{row['교사명']}**: {date_b_str}({day_b}) {p_b_input}교시 [{row['상대수업']}] ➔ **{pick_a['일자']}({pick_a['요일']}) {pick_a['교시']}교시**")
                        else:
                            st.write(f"• **{pick_a['교사명']} 수업**: {pick_a['일자']}({pick_a['요일']}) {pick_a['교시']}교시 ➔ {date_b_str}({day_b}) {p_b_input}교시로 이동")
                            st.write(f"• **{row['교사명']} 역할**: 해당 시간 공강으로 **수업 대리/보강 진행** 가능")
                    
                    with c3:
                        if row["type"] == "SWAP":
                            if not row["is_valid"]:
                                st.error("⚠️ 맞교환 불가")
                                st.caption(row['errs'][0] if row['errs'] else "")
                            else:
                                st.success("✅ 맞교환 가능")
                                if st.button("1:1 맞교환 실행", key=f"btn_swap_{idx}", type="primary", use_container_width=True):
                                    if do_swap(pick_a, row["b_info"], date_a_str, date_b_str):
                                        st.success("시간표 맞교환이 완료되었습니다.")
                                        st.rerun()
                        else:
                            st.info("✅ 공강/보강 가능")
                            if st.button("보강/대리 지정", key=f"btn_free_{idx}", type="primary", use_container_width=True):
                                cid = f"{date_b_str}-{pick_a['교사명']}"
                                add_substitute(
                                    cid, date_b_str, day_b, p_b_input, pick_a["학급"], pick_a["과목"],
                                    pick_a["교사명"], row["교사명"], "시간표변경_대리", "완료", f"{pick_a['일자']} {pick_a['교시']}교시 수업 이동"
                                )
                                st.success(f"대리 교사로 '{row['교사명']}' 교사가 지정되었습니다.")
                                st.rerun()

    st.divider()
    st.subheader("📜 시간표 변경 및 맞교환 이력")
    st.dataframe(st.session_state.swaps, use_container_width=True)
