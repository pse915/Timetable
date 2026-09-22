"""
[중1 정보] 활동지 : 데이터의 구조화 (Streamlit 앱)
- GitHub 저장소에 이 파일(app.py)과 requirements.txt를 업로드한 후,
  Streamlit Community Cloud (share.streamlit.io)에서 1클릭으로 무료 배포할 수 있습니다.
- 로컬 실행: streamlit run app.py
"""

import streamlit as st
import pandas as pd

# 1. 페이지 기본 설정
st.set_page_config(
    page_title="활동지 - 데이터의 구조화 (중1 정보)",
    page_icon="📋",
    layout="wide",
    initial_sidebar_state="expanded"
)

# 2. 커스텀 CSS (깔끔한 교육용 카드 스타일)
st.markdown("""
<style>
    .main-title {
        font-size: 26px;
        font-weight: 800;
        color: #1e293b;
        margin-bottom: 2px;
    }
    .sub-title {
        font-size: 14px;
        color: #64748b;
        margin-bottom: 16px;
    }
    .section-card {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 18px;
        margin-bottom: 20px;
    }
    .badge {
        background-color: #dbeafe;
        color: #1d4ed8;
        padding: 4px 10px;
        border-radius: 6px;
        font-size: 12px;
        font-weight: bold;
    }
</style>
""", unsafe_allow_html=True)

# 3. 사이드바: 학생 정보 및 단어 보관함, 채점 컨트롤
with st.sidebar:
    st.image("https://api.iconify.design/lucide:book-open.svg?color=%232563eb", width=40)
    st.markdown("### 👨‍🎓 학생 정보 입력")
    col_sb1, col_sb2 = st.columns(2)
    with col_sb1:
        grade = st.text_input("학년", value="1학년")
        class_num = st.text_input("반", value="1반")
    with col_sb2:
        num = st.text_input("번호", value="1번")
        name = st.text_input("이름", placeholder="홍길동")
    
    st.markdown("---")
    st.markdown("### 💡 단어 보관함 (힌트)")
    st.info("""
    - 데이터
    - 특성
    - 정리 및 배열
    - 통일된 모양
    - 쉽게 찾을
    - 관계
    - 효율적으로 관리
    - 기준
    - 소프트웨어 개발 전문가
    - 시스템 SW 개발자
    - 응용 SW 개발자
    """)
    
    st.markdown("---")
    show_answer_key = st.toggle("🔍 교사용 모범답안 보기", value=False)

# 4. 상단 메인 헤더
st.markdown('<span class="badge">중학교 1학년 정보 Ⅱ.데이터</span>', unsafe_allow_html=True)
st.markdown(f'<div class="main-title">활동지 - 데이터의 구조화</div>', unsafe_allow_html=True)
st.markdown(f'<div class="sub-title">작성자: {grade} {class_num} {num} <strong>{name if name else "(이름을 입력하세요)"}</strong></div>', unsafe_allow_html=True)

# 5. 본문 섹션 1: 데이터 구조화의 뜻
st.markdown("### 1. 데이터 구조화의 뜻과 왜 해야 할까?")
with st.container():
    st.markdown("""
    > **[데이터 구조화의 정의]**  
    > 전달하려고 하는 **( ① )**의 내용 요소들을 **( ② )**에 맞게 **( ③ )**하여 **( ④ )**으로 표현하는 것.
    """)
    
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        q1 = st.text_input("① 빈칸", value="데이터" if show_answer_key else "", placeholder="① 빈칸 입력", key="q1")
    with col2:
        q2 = st.text_input("② 빈칸", value="특성" if show_answer_key else "", placeholder="② 빈칸 입력", key="q2")
    with col3:
        q3 = st.text_input("③ 빈칸", value="정리 및 배열" if show_answer_key else "", placeholder="③ 빈칸 입력", key="q3")
    with col4:
        q4 = st.text_input("④ 빈칸", value="통일된 모양" if show_answer_key else "", placeholder="④ 빈칸 입력", key="q4")

# 본문 섹션 1-2: 4가지 이유와 실생활 예시
st.markdown("#### 📌 데이터 구조화를 하는 4가지 주요 이유")
col_r1, col_r2 = st.columns(2)
with col_r1:
    q5 = st.text_input("이유 ①: 데이터를 [ ? ] 수 있다.", value="쉽게 찾을" if show_answer_key else "", placeholder="예: 쉽게 찾을", key="q5")
    st.caption("실생활 예시: 학번순 학생 명부, 도서관 십진분류 청구기호")
    
    q6 = st.text_input("이유 ②: 데이터 사이의 [ ? ]을 쉽게 이해할 수 있다.", value="관계" if show_answer_key else "", placeholder="예: 관계", key="q6")
    st.caption("실생활 예시: 가족 가계도, 학교 및 회사 조직도")

with col_r2:
    st.text_input("이유 ③: 빠지거나 잘못된 내용(오류/결측치)을 쉽게 파악", value="오류 및 누락 파악", disabled=True)
    st.caption("실생활 예시: 월별 지출 누락 확인, 시간표 중복 체크")
    
    q7 = st.text_input("이유 ④: 데이터를 [ ? ] 할 수 있다.", value="효율적으로 관리" if show_answer_key else "", placeholder="예: 효율적으로 관리", key="q7")
    st.caption("실생활 예시: 연간 통계 분석 및 데이터베이스 저장")

st.markdown("---")

# 6. 본문 섹션 2: 3가지 구조화 방법 (목록 ➔ 표 ➔ 다이어그램)
st.markdown("### 2. 데이터는 어떤 방법으로 구조화할 수 있을까? (순서: 목록 ➔ 표 ➔ 다이어그램)")
col_m1, col_m2, col_m3 = st.columns(3)

with col_m1:
    st.success("##### 1. 목록 (List)")
    q8 = st.text_input("일정한 [ ? ]에 맞추어 나열", value="기준" if show_answer_key else "", placeholder="기준", key="q8")
    st.caption("예: 준비물 체크리스트, 레시피 순서")

with col_m2:
    st.info("##### 2. 표 (Table)")
    st.markdown("**행(가로줄)**과 **열(세로줄)**의 격자 구조로 구성")
    st.caption("예: 학급 시간표, 친구 주소록")

with col_m3:
    st.warning("##### 3. 다이어그램 (Diagram)")
    st.markdown("점, 선, 도형, 화살표 등으로 시각화")
    st.caption("예: 지하철 노선도, 계층 트리, 막대그래프")

st.markdown("---")

# 7. [활동 1] 목록 만들기 실습
st.markdown("### [활동 1] 목록(List) 만들기 실습: 우리 모둠 체크리스트")
st.write("아래 표를 클릭하여 직접 우리 모둠의 할 일과 담당자를 수정해 보세요.")
default_checklist = pd.DataFrame([
    {"순번": 1, "할 일 / 점검 항목": "정보 교과서 및 필기도구 챙기기", "담당자": "나", "완료": True},
    {"순번": 2, "할 일 / 점검 항목": "모둠 태블릿 PC 배터리 충전 확인", "담당자": "친구1", "완료": False},
    {"순번": 3, "할 일 / 점검 항목": "발표 자료 취합 및 정리", "담당자": "친구2", "완료": False},
])
edited_checklist = st.data_editor(default_checklist, use_container_width=True, num_rows="dynamic")

# 8. [활동 2] 표 만들기 실습
st.markdown("### [활동 2] 표(Table) 만들기 실습: 우리 모둠 친구 프로필 데이터")
st.write("행과 열로 구성된 격자 표에 친구들의 데이터를 입력해 보세요.")
default_profiles = pd.DataFrame([
    {"번호": 1, "친구 이름": "김민수", "생일(월/일)": "3월 15일", "취미/특기": "축구, 코딩", "모둠 내 역할": "자료 조사 및 발표"},
    {"번호": 2, "친구 이름": "이영희", "생일(월/일)": "7월 20일", "취미/특기": "독서, 그리기", "모둠 내 역할": "보고서 서기"},
    {"번호": 3, "친구 이름": "박지훈", "생일(월/일)": "11월 5일", "취미/특기": "게임, 음악 감상", "모둠 내 역할": "자료 제작 및 검토"},
])
edited_profiles = st.data_editor(default_profiles, use_container_width=True, num_rows="dynamic")

st.markdown("---")

# 9. [활동 3] 다이어그램 만들기 실습
st.markdown("### [활동 3] 다이어그램 만들기: 계층형 다이어그램 빈칸 채우기")
st.info("""
**(가) 제시문:**  
"소프트웨어 개발 전문가는 시스템 SW 개발자와 응용 SW 개발자로 나뉜다. 시스템 SW 개발자는 운영체제 프로그래머와 임베디드 프로그래머가 있다. 응용 SW 개발자는 응용 SW 프로그래머, 네트워크 프로그래머, 컴퓨터 및 모바일 게임 프로그래머가 있다."
""")

st.markdown("#### (나) 계층형 다이어그램")
col_tree_center = st.columns([1, 2, 1])[1]
with col_tree_center:
    q9 = st.text_input("🏢 [최상위 직무 분류]", value="소프트웨어 개발 전문가" if show_answer_key else "", placeholder="최상위 직무 입력", key="q9")
    st.markdown("<div style='text-align: center; font-size: 20px; color: #0284c7;'>│<br>▼</div>", unsafe_allow_html=True)

col_sub1, col_sub2 = st.columns(2)
with col_sub1:
    q10 = st.text_input("↙ 하위 분류 1", value="시스템 SW 개발자" if show_answer_key else "", placeholder="하위 분류 1 입력", key="q10")
    st.markdown("- 운영체제 프로그래머\n- 임베디드 프로그래머")

with col_sub2:
    q11 = st.text_input("↘ 하위 분류 2", value="응용 SW 개발자" if show_answer_key else "", placeholder="하위 분류 2 입력", key="q11")
    st.markdown("- 응용 SW 프로그래머\n- 네트워크 프로그래머\n- 모바일 게임 프로그래머")

st.markdown("---")

# 10. [활동 4] 체험학습 계획 및 통계 막대그래프 다이어그램
st.markdown("### [활동 4] 민호의 체험학습 계획 구조화 및 통계 막대그래프")
st.markdown("""
> "4월에는 봄꽃을 보러 경복궁, 창덕궁, 종묘 등을 방문할 예정이다. 서울 지하철 3호선을 탄다.  
> 5월에는 가족과 함께 지하철 2호선을 타고 미술관이 있는 덕수궁에 간다.  
> 6월에는 선조들의 얼을 기리고자 전쟁기념관과 국립중앙박물관을 방문하며 지하철 4호선을 탄다."
""")

chart_data = pd.DataFrame({
    "지하철 호선 (월별)": ["3호선 (4월)", "2호선 (5월)", "4호선 (6월)"],
    "방문 장소 수": [3, 1, 2],
    "방문 장소 명단": ["경복궁, 창덕궁, 종묘", "덕수궁", "전쟁기념관, 국립중앙박물관"]
})

st.markdown("#### 📊 호선별 방문 장소 수 비교 (막대그래프 다이어그램)")
st.bar_chart(chart_data, x="지하철 호선 (월별)", y="방문 장소 수", color="#3b82f6")

reflection = st.text_area("자신이 정리한 구조화 내용(목록/표/노선 흐름)을 자유롭게 기록해 보세요:", placeholder="체험학습 계획을 목록이나 표로 요약하여 기록해 보세요...")

st.markdown("---")

# 11. 스스로 배움 점검하기 & 자동 채점 결과
st.markdown("### 🎯 스스로 배움 점검하기")
c1 = st.checkbox("1. 데이터 구조화의 뜻과 필요한 이유 4가지를 설명할 수 있다.")
c2 = st.checkbox("2. 데이터의 특성에 맞춰 목록, 표, 다이어그램을 올바르게 선택할 수 있다.")
c3 = st.checkbox("3. 줄글 데이터를 표나 계층형/차트 다이어그램으로 직접 표현할 수 있다.")

st.markdown("---")

# 채점 로직
if st.button("📝 자동 채점 및 결과 확인", type="primary", use_container_width=True):
    answers = {
        "q1": ("데이터", q1),
        "q2": ("특성", q2),
        "q3": ("정리 및 배열", q3),
        "q4": ("통일된 모양", q4),
        "q5": ("쉽게 찾을", q5),
        "q6": ("관계", q6),
        "q7": ("효율적으로 관리", q7),
        "q8": ("기준", q8),
        "q9": ("소프트웨어 개발 전문가", q9),
        "q10": ("시스템 SW 개발자", q10),
        "q11": ("응용 SW 개발자", q11),
    }
    
    score = 0
    total = len(answers)
    
    for key, (correct_val, user_val) in answers.items():
        if user_val.strip().replace(" ", "") == correct_val.strip().replace(" ", ""):
            score += 1
            
    percent = int((score / total) * 100)
    
    if percent >= 80:
        st.balloons()
        st.success(f"🎉 훌륭합니다! 총 {total}개 빈칸 중 {score}개를 맞혔습니다. (점수: {percent}점)")
    elif percent >= 50:
        st.warning(f"👍 잘했습니다! 총 {total}개 빈칸 중 {score}개를 맞혔습니다. (점수: {percent}점) 틀린 부분을 다시 검토해 보세요.")
    else:
        st.error(f"총 {total}개 빈칸 중 {score}개를 맞혔습니다. (점수: {percent}점) 사이드바의 힌트 단어 보관함을 참고하여 다시 도전해 보세요!")
