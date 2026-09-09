
import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# 페이지 기본 설정
st.set_page_config(
    page_title="가족 의사소통 & 갈등 해결 마스터",
    page_icon="💖",
    layout="centered"
)

# ================= gspread 직접 연동 구글 시트 저장 함수 =================
def submit_to_google_sheet(std_id, name, score):
    sheet_url = "https://docs.google.com/spreadsheets/d/1IiG4q_CY6yUPUqnrvb0O-YWteqmti3qxLdoovxMLXYI/edit?usp=sharing"
    
    try:
        # Streamlit Secrets에서 GCP 서비스 계정 정보 로드
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], 
            scopes=scopes
        )
        client = gspread.authorize(creds)
        
        # 구글 시트 열기 (첫 번째 시트)
        sheet = client.open_by_url(sheet_url).sheet1
        
        # 제출 일시, 학번, 이름, 점수 행 추가
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        sheet.append_row([now_str, std_id, name, score])
        return True
        
    except Exception as e:
        st.error(f"구글 시트 전송 실패: {e}")
        return False

# ================= 질문 데이터베이스 =================
QUESTIONS = [
    {
        "cat": "1단계: 나-전달법 (행동 묘사)",
        "dialogue": "동생이 허락 없이 학용품을 가져가 잃어버린 상황!",
        "q": "비난이나 평가 없이 상대방의 '행동'만 객관적으로 표현한 것은?",
        "options": [
            ("1. 너는 왜 매번 허락도 없이 남의 물건을 마음대로 건드리니?", False),
            ("2. 내 허락 없이 필통에서 학용품을 가져가서 잃어버렸어.", True),
            ("3. 넌 항상 정리정돈도 안 하고 무책임한 태도를 보이더라.", False),
            ("4. 남의 물건을 몰래 가져가는 건 정말 나쁜 행동이야.", False),
            ("5. 네가 자꾸 내 물건을 건드리니까 내가 항상 화가 나는 거야.", False)
        ]
    },
    {
        "cat": "2단계: 나-전달법 (영향 및 감정)",
        "dialogue": "형(누나)이 약속 시간에 30분 넘게 연락도 없이 늦게 온 상황!",
        "q": "나에게 미친 '영향과 솔직한 감정'을 올바르게 표현한 것은?",
        "options": [
            ("1. 오랫동안 혼자 기다리면서 걱정되고 내 시간이 허비되어 속상했어.", True),
            ("2. 너 진짜 시간 개념이 없구나. 약속을 왜 하자고 했니?", False),
            ("3. 너 때문에 오늘 내 하루 기분을 완전히 다 망쳐버렸어.", False),
            ("4. 다음부터 늦으면 나도 똑같이 늦게 나갈 테니 알아서 해.", False),
            ("5. 약속 하나 제대로 못 지키면서 무슨 중요한 일을 하겠다는 거니?", False)
        ]
    },
    {
        "cat": "3단계: 나-전달법 (바라는 사항)",
        "dialogue": "부모님이 내 의견을 묻지 않고 주말 일정을 일방적으로 정하셨을 때!",
        "q": "상대방에게 올바르게 '바라는 사항'을 요청하는 문장은?",
        "options": [
            ("1. 부모님 마음대로 하실 거면 앞으로 저한테 아무것도 묻지 마세요.", False),
            ("2. 제 일정과 의견도 먼저 물어봐 주시고 함께 결정해 주셨으면 좋겠어요.", True),
            ("3. 이번 결정 취소 안 해주시면 저 주말에 집 나가서 안 들어올 거예요.", False),
            ("4. 제발 저 좀 그만 괴롭히시고 그냥 제 일에 신경 꺼주세요.", False),
            ("5. 부모님의 결정 방식은 완전히 잘못되었으니 당장 수정해 주세요.", False)
        ]
    },
    {
        "cat": "4단계: 너-전달법 → 나-전달법 변환",
        "dialogue": "너-전달법: \"너는 왜 내가 말할 때마다 폰만 보고 딴청이니?\"",
        "q": "위 '너-전달법'을 올바른 '나-전달법'으로 바꾼 것은?",
        "options": [
            ("1. 폰 좀 그만 보고 사람 얼굴을 보며 대화하는 예의를 갖춰라.", False),
            ("2. 내가 말할 때 스마트폰을 보면 내 말을 경청받지 못하는 느낌이 들어 서운해.", True),
            ("3. 너는 스마트폰 중독이라 사람과 제대로 된 대화가 불가능하구나.", False),
            ("4. 앞으로 대화할 때는 네 스마트폰을 전부 압수해야겠어.", False),
            ("5. 내가 말하는데 폰을 계속 보면 너랑 다시는 대화 안 할 거야.", False)
        ]
    },
    {
        "cat": "5단계: 나-전달법 3요소 완성",
        "dialogue": "\"네가 연락 없이 약속에 늦어서(행동), 기다리며 걱정되고 속상했어(영향/감정).\"",
        "q": "이 문장 뒤에 이어질 마지막 '바라는 사항'으로 가장 적절한 것은?",
        "options": [
            ("1. 앞으로는 늦을 것 같으면 미리 나에게 연락을 해주면 좋겠어.", True),
            ("2. 다음부터 한 번만 더 늦으면 너랑 다시는 안 놀아.", False),
            ("3. 너도 똑같이 30분 동안 길거리에서 기다려 봐야 정신 차리지?", False),
            ("4. 앞으로 너와의 모든 일정은 내가 전부 취소하도록 할게.", False),
            ("5. 늦은 시간만큼 네가 맛있는 걸 사서 정식으로 사과하면 좋겠어.", False)
        ]
    },
    {
        "cat": "6단계: 경청과 공감",
        "dialogue": "가족 구성원이 시험이나 일로 마음처럼 되지 않아 우울하다고 고민할 때!",
        "q": "경청과 공감의 바람직한 대화 태도는 무엇일까요?",
        "options": [
            ("1. 상대방의 평소 생활 습관과 잘못된 점을 즉시 지적해 준다.", False),
            ("2. 말하는 중간에 개입하여 나의 더 안 좋았던 경험담을 이야기한다.", False),
            ("3. 비판이나 성급한 조언 전에 상대방이 느꼈을 좌절감에 먼저 공감해 준다.", True),
            ("4. 별일 아니라는 듯 대수롭지 않게 넘기며 빠르게 주제를 바꾼다.", False),
            ("5. 해결책을 제시하기 위해 상대방의 말을 끊고 논리적으로 질문한다.", False)
        ]
    },
    {
        "cat": "7단계: 나-전달법과 비언어적 표현",
        "dialogue": "나-전달법으로 말하지만 표정은 찌푸리고 팔짱을 끼고 있는 상황!",
        "q": "나-전달법을 사용할 때 비언어적 표현(표정, 말투, 시선)의 올바른 태도는?",
        "options": [
            ("1. 말의 내용보다 상대를 제압하는 강한 눈빛과 억양이 중요하다.", False),
            ("2. 말만 나-전달법으로 한다면 비꼬는 말투나 표정은 상관없다.", False),
            ("3. 진정성 전달을 위해 언어적 메시지와 비언어적 표현을 일치시켜야 한다.", True),
            ("4. 감정을 숨기기 위해 무표정한 얼굴과 기계적인 목소리를 유지한다.", False),
            ("5. 상대방이 미안함을 느끼도록 가벼운 한숨을 쉬며 말하는 것이 좋다.", False)
        ]
    },
    {
        "cat": "8단계: 성격 유형별 대화 (사고형 T)",
        "dialogue": "원칙과 논리적 사실 관계를 중시하는 '사고형(T)' 아빠와의 대화!",
        "q": "T형 가족 구성원과 갈등을 해결할 때 가장 효과적인 대화법은?",
        "options": [
            ("1. 감정적으로 눈물을 흘리며 내 기분만 알아달라고 호소한다.", False),
            ("2. 객관적인 사실과 이유를 차분하고 논리적으로 설명한다.", True),
            ("3. 상대방의 논리적 오류를 계속 지적하며 언쟁에서 이기려 한다.", False),
            ("4. 논리적인 대화는 무의미하므로 대화를 완전히 포기한다.", False),
            ("5. 상대방의 서운한 점을 지적하며 감정적인 대답을 강요한다.", False)
        ]
    },
    {
        "cat": "9단계: 성격 유형별 대화 (감정형 F)",
        "dialogue": "관계와 공감, 마음의 공유를 중시하는 '감정형(F)' 동생과의 대화!",
        "q": "F형 가족 구성원의 마음을 열 수 있는 바람직한 대화법은?",
        "options": [
            ("1. 옳고 그름을 따지기 전에 상대방이 느꼈을 감정과 입장을 인정해 준다.", True),
            ("2. 상대방의 감정은 비이성적이라며 차갑게 사실만 지적한다.", False),
            ("3. 감정적인 이야기에는 응하지 않고 빠른 해결책만 제시한다.", False),
            ("4. 동조해 주는 척하면서 은근히 상대방의 잘못을 깨닫게 한다.", False),
            ("5. 상대방의 감정 표현을 장난으로 넘기며 분위기를 전환한다.", False)
        ]
    },
    {
        "cat": "10단계: 가족 갈등 해결 4단계",
        "dialogue": "가족 회의에서 갈등을 올바르게 해결하는 체계적인 4단계 프로세스!",
        "q": "가족 갈등을 해결하는 올바른 순서로 가장 적절한 것은?",
        "options": [
            ("1. 갈등 확인 → 해결 방법 탐색 → 해결 방법 결정 → 실행 및 평가", True),
            ("2. 해결 방법 결정 → 갈등 확인 → 실행 및 평가 → 해결 방법 탐색", False),
            ("3. 갈등 확인 → 실행 및 평가 → 해결 방법 탐색 → 해결 방법 결정", False),
            ("4. 해결 방법 탐색 → 갈등 확인 → 해결 방법 결정 → 실행 및 평가", False),
            ("5. 갈등 확인 → 해결 방법 결정 → 해결 방법 탐색 → 실행 및 평가", False)
        ]
    }
]

# ================= 세션 상태 초기화 =================
if "q_idx" not in st.session_state:
    st.session_state.q_idx = 0
if "score" not in st.session_state:
    st.session_state.score = 0
if "combo" not in st.session_state:
    st.session_state.combo = 0
if "feedback" not in st.session_state:
    st.session_state.feedback = None
if "submitted" not in st.session_state:
    st.session_state.submitted = False

# ================= UI 레이아웃 =================
st.title("👨‍👩‍👧‍👦 가족 의사소통 & 갈등 해결 마스터")

# 상단 진행률 및 게이지
col1, col2 = st.columns([3, 1])
with col1:
    st.progress(st.session_state.score / 100)
with col2:
    st.metric("가족 화목도", f"{st.session_state.score}%")

st.divider()

# 퀴즈 진행 중 (0~9번 문제)
if st.session_state.q_idx < len(QUESTIONS):
    q_data = QUESTIONS[st.session_state.q_idx]

    col_cat, col_combo = st.columns([3, 1])
    with col_cat:
        st.caption(f"📍 STAGE {st.session_state.q_idx + 1}. {q_data['cat']}")
    with col_combo:
        if st.session_state.combo > 1:
            st.write(f"🔥 **{st.session_state.combo} COMBO!**")

    st.info(f"💬 **[상황]** {q_data['dialogue']}")

    if st.session_state.feedback:
        fb_type, fb_text = st.session_state.feedback
        if fb_type == "success":
            st.success(fb_text)
        else:
            st.error(fb_text)

    st.subheader(q_data["q"])

    for idx, (opt_text, is_correct) in enumerate(q_data["options"]):
        if st.button(opt_text, key=f"q_{st.session_state.q_idx}_opt_{idx}", use_container_width=True):
            if is_correct:
                st.session_state.score += 10
                st.session_state.combo += 1
                st.session_state.feedback = ("success", f"🎉 정답! 가족 화목도 UP! ({st.session_state.combo}연속 성공!)")
            else:
                st.session_state.combo = 0
                st.session_state.feedback = ("error", "💔 오답! 상처주는 대화법입니다.")

            st.session_state.q_idx += 1
            st.rerun()

# 결과 제출 화면
else:
    if st.session_state.feedback:
        fb_type, fb_text = st.session_state.feedback
        if fb_type == "success":
            st.success(fb_text)
        else:
            st.error(fb_text)

    st.balloons()
    st.header("🏆 학습 완료! 결과를 제출하세요")
    st.subheader(f"최종 가족 화목도: {st.session_state.score}점 / 100점")

    if not st.session_state.submitted:
        with st.form("submit_form"):
            std_id = st.text_input("학번", placeholder="예: 10101")
            name = st.text_input("이름", placeholder="예: 홍길동")
            submit_btn = st.form_submit_button("구글 시트에 제출", use_container_width=True)

            if submit_btn:
                if std_id and name:
                    with st.spinner("구글 시트에 전송 중..."):
                        success = submit_to_google_sheet(std_id, name, st.session_state.score)
                    if success:
                        st.session_state.submitted = True
                        st.rerun()
                else:
                    st.warning("학번과 이름을 모두 입력해 주세요!")
    else:
        st.success("✨ 구글 시트에 성공적으로 기록되었습니다!")

        if st.button("🔄 처음부터 다시 풀기", use_container_width=True):
            st.session_state.q_idx = 0
            st.session_state.score = 0
            st.session_state.combo = 0
            st.session_state.feedback = None
            st.session_state.submitted = False
            st.rerun()
