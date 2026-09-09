import streamlit as st
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# 1. 페이지 및 레이아웃 기본 설정
st.set_page_config(
    page_title="나 전달법 마스터: 대화의 신",
    page_icon="💬",
    layout="centered"
)

# 2. 커스텀 CSS (캐릭터 카드 & 말풍선 UI 디자인)
st.markdown("""
<style>
    .char-card {
        background-color: #f8fafc;
        border: 2px solid #e2e8f0;
        border-radius: 16px;
        padding: 16px;
        text-align: center;
        margin-bottom: 12px;
    }
    .char-avatar {
        font-size: 50px;
        margin-bottom: -10px;
    }
    .char-name {
        font-weight: bold;
        color: #334155;
        font-size: 16px;
    }
    .speech-bubble {
        position: relative;
        background: #e0f2fe;
        border-radius: 12px;
        padding: 16px;
        color: #0369a1;
        font-weight: 600;
        font-size: 15px;
        border: 1px solid #bae6fd;
        margin-bottom: 16px;
    }
    .stButton>button {
        border-radius: 10px;
        font-size: 14px;
        padding: 10px 14px;
        transition: all 0.2s ease;
    }
</style>
""", unsafe_allow_html=True)

# 3. 구글 시트 연동 함수
def submit_to_google_sheet(std_id, name, score, title):
    sheet_url = "https://docs.google.com/spreadsheets/d/1IiG4q_CY6yUPUqnrvb0O-YWteqmti3qxLdoovxMLXYI/edit?usp=sharing"
    
    try:
        scopes = ["https://www.googleapis.com/auth/spreadsheets"]
        creds = Credentials.from_service_account_info(
            st.secrets["gcp_service_account"], 
            scopes=scopes
        )
        client = gspread.authorize(creds)
        sheet = client.open_by_url(sheet_url).sheet1
        
        now_str = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        sheet.append_row([now_str, std_id, name, score, title])
        return True
    except Exception as e:
        st.error(f"구글 시트 저장 실패: {e}")
        return False

# 4. 검토 완료된 5지 선다형 10개 퀴즈 데이터베이스
QUESTIONS = [
    {
        "stage": "STAGE 1",
        "cat": "1단계: 비난 없는 '객관적 행동 묘사'",
        "char_name": "동생 민지 (15세)",
        "avatar": "👧",
        "dialogue": "동생이 내 옷을 허락 없이 입고 나갔다가 떡볶이 국물을 묻혀서 돌아온 상황!",
        "q": "비난이나 평가 없이 상대방의 '행동 사실'만 객관적으로 표현한 문장은?",
        "options": [
            {"text": "1. 너는 왜 매번 허락도 없이 남의 옷을 마음대로 입고 나가니?", "correct": False, "exp": "❌ '매번~', '마음대로'는 상대의 태도를 비난하는 너-전달법입니다."},
            {"text": "2. 내 허락 없이 옷을 입고 나가서 떡볶이 국물을 묻혀왔어.", "correct": True, "exp": "⭕ 정답! 비판이나 감정을 배제하고 일어난 '행동 사실'만 명확히 기술했습니다."},
            {"text": "3. 너처럼 무책임하고 남 생각 안 하는 애는 정말 처음 본다.", "correct": False, "exp": "❌ 상대방의 인격을 공격하는 바람직하지 못한 대화법입니다."},
            {"text": "4. 남의 옷을 망가뜨리는 건 정말 예의가 아닌 행동이야.", "correct": False, "exp": "❌ 도덕적 훈계와 비판이 앞서면 상대는 방어 태세를 취하게 됩니다."},
            {"text": "5. 자꾸 내 옷에 국물을 묻혀오니까 내가 항상 화가 나는 거잖아.", "correct": False, "exp": "❌ 감정의 원인을 전적으로 상대 탓으로 돌리는 표현입니다."}
        ]
    },
    {
        "stage": "STAGE 2",
        "cat": "2단계: 구체적 '영향 및 감정 표현'",
        "char_name": "친구 태양 (16세)",
        "avatar": "👦",
        "dialogue": "중요한 조별 과제 모임에 친구가 연락도 없이 40분이나 늦게 도착한 상황!",
        "q": "친구에게 나에게 미친 '구체적 영향'과 '나의 솔직한 감정'을 올바르게 전달한 문장은?",
        "options": [
            {"text": "1. 40분 동안 연락이 안 돼서 걱정됐고, 발표 준비 시간이 부족해져서 마음이 초조했어.", "correct": True, "exp": "⭕ 정답! 내가 겪은 불이익(영향)과 솔직하게 느낀 기분(감정)을 정확히 전달했습니다."},
            {"text": "2. 너 진짜 시간 개념이 없구나. 약속을 왜 하자고 했니?", "correct": False, "exp": "❌ 상대의 성격과 태도를 단정 짓고 비난하는 '너-전달법'입니다."},
            {"text": "3. 너 때문에 우리 조 발표 준비 완전히 망쳤으니까 책임져.", "correct": False, "exp": "❌ 책임을 상대에게 전가하면 친구는 반발심만 느끼게 됩니다."},
            {"text": "4. 다음부터 늦으면 나도 약속 장소에 안 나갈 테니까 알아서 해.", "correct": False, "exp": "❌ 감정적 협박이나 통보는 갈등을 더 악화시킵니다."},
            {"text": "5. 약속 하나 제대로 못 지키면서 무슨 중요한 일을 하겠다는 거니?", "correct": False, "exp": "❌ 상대방의 능력과 인격을 비하하는 공격적 대화입니다."}
        ]
    },
    {
        "stage": "STAGE 3",
        "cat": "3단계: 긍정적 '바라는 사항(요청)'",
        "char_name": "아버지 (48세)",
        "avatar": "👨",
        "dialogue": "부모님이 내 의사는 묻지도 않고 주말 친척 집 방문 일정을 일방적으로 정하신 상황!",
        "q": "부모님께 예의를 지키면서 내가 '원하는 바'를 올바르게 요청하는 문장은?",
        "options": [
            {"text": "1. 제 일정은 왜 안 물어보세요? 저 주말에 절대로 안 갈 거예요.", "correct": False, "exp": "❌ 일방적인 거부 표현은 부모님과의 대화를 단절시킵니다."},
            {"text": "2. 제 계획도 미리 물어봐 주시고, 일정을 정할 때 함께 이야기해 주셨으면 좋겠어요.", "correct": True, "exp": "⭕ 정답! 부모님을 존중하면서도 내가 바라는 변화(수용 가능한 요청)를 명확히 표현했습니다."},
            {"text": "3. 부모님 마음대로 하실 거면 앞으로 저한테 아무것도 요구하지 마세요.", "correct": False, "exp": "❌ 비꼬는 태도는 갈등 해결에 아무런 도움이 되지 않습니다."},
            {"text": "4. 친척 집 가는 것보다 제 수행평가가 훨씬 중요하거든요?", "correct": False, "exp": "❌ 상대의 가치를 폄하하는 말투는 감정싸움을 유발합니다."},
            {"text": "5. 제발 제 일에 신경 좀 끄시고 부모님 일이나 잘 신경 쓰세요.", "correct": False, "exp": "❌ 반발심과 무례함이 섞인 바람직하지 않은 언어 표현입니다."}
        ]
    },
    {
        "stage": "STAGE 4",
        "cat": "4단계: '너-전달법'을 '나-전달법'으로 완벽 변환",
        "char_name": "어머니 (46세)",
        "avatar": "👩",
        "dialogue": "공격적인 너-전달법: \"너는 왜 내가 말할 때마다 폰만 보고 딴청이니?\"",
        "q": "위 문장을 나-전달법의 3요소(행동-영향/감정-바람)를 모두 갖춘 문장으로 바꾼 것은?",
        "options": [
            {"text": "1. 말할 때 스마트폰을 계속 보면 내 말이 무시당하는 느낌이 들어서 서운해. 나를 바라보며 대화해 주면 좋겠어.", "correct": True, "exp": "⭕ 정답! 행동(폰 봄), 영향/감정(무시당하는 느낌/서운함), 바람(바라보며 대화)이 완벽히 갖춰졌습니다."},
            {"text": "2. 스마트폰 좀 치우고 사람 얼굴 보고 대화하는 예의를 갖춰라.", "correct": False, "exp": "❌ 지적과 명령조는 상대의 방어 기제를 자극합니다."},
            {"text": "3. 너는 스마트폰 중독이라 사람과 제대로 된 대화가 불가능하구나.", "correct": False, "exp": "❌ 진단과 인격 비난이 섞여 상대의 기분을 해치게 됩니다."},
            {"text": "4. 앞으로 대화할 때는 네 스마트폰을 전부 압수해야겠어.", "correct": False, "exp": "❌ 일방적인 통제와 처벌 언급은 소통을 차단합니다."},
            {"text": "5. 내가 말하는데 폰을 계속 보면 너랑 다시는 대화 안 할 거야.", "correct": False, "exp": "❌ 감정적 협박은 진정한 행동 변화를 이끌어낼 수 없습니다."}
        ]
    },
    {
        "stage": "STAGE 5",
        "cat": "5단계: 공감과 나-전달법의 조화",
        "char_name": "남동생 현우 (13세)",
        "avatar": "👦",
        "dialogue": "내가 공부하는 동안 옆에서 남동생이 계속 시끄럽게 소리를 지르며 장난치는 상황!",
        "q": "동생과의 갈등을 평화롭게 해결하기 위한 가장 바람직한 나-전달법 표현은?",
        "options": [
            {"text": "1. 네가 신나서 기분 좋은 건 알겠는데, 큰 소리를 내면 집중하기가 힘들어서 스트레스를 받아. 조금만 조용히 놀아줄래?", "correct": True, "exp": "⭕ 정답! 상대의 기분에 공감해 준 뒤 나의 영향과 바람을 말하면 요청의 수용률이 높아집니다."},
            {"text": "2. 야! 너 진짜 산만하다. 당장 네 방으로 나가!", "correct": False, "exp": "❌ 일방적인 명령과 인격 비하는 동생의 반발을 부릅니다."},
            {"text": "3. 네가 자꾸 시끄럽게 구니까 내가 공부를 집중해서 못 하잖아!", "correct": False, "exp": "❌ 자신의 실패 원인을 모두 남 탓으로 돌리는 대화 방식입니다."},
            {"text": "4. 너 한 번만 더 소리 지르면 컴퓨터 전원 꺼버린다.", "correct": False, "exp": "❌ 위협과 협박은 갈등만 더욱 악화시킵니다."},
            {"text": "5. 남을 배려할 줄 모르는 행동은 정말 나쁜 습관이야.", "correct": False, "exp": "❌ 도덕적 훈계와 지적은 상대방의 마음을 닫히게 만듭니다."}
        ]
    },
    {
        "stage": "STAGE 6",
        "cat": "6단계: 비언어적 표현의 일치",
        "char_name": "짝꿍 지유 (16세)",
        "avatar": "👧",
        "dialogue": "말로는 \"네 생각도 이해해~\"라고 하지만, 팔짱을 끼고 인상을 찌푸리며 한숨을 쉬는 상황!",
        "q": "나-전달법을 사용할 때 언어적 메시지만큼 중요한 '비언어적 태도'의 바람직한 모습은?",
        "options": [
            {"text": "1. 상대방을 제압할 수 있도록 강렬한 눈빛과 비꼬는 말투를 유지한다.", "correct": False, "exp": "❌ 공격적인 태도는 언어 표현과 상관없이 갈등을 증폭시킵니다."},
            {"text": "2. 진정성 있는 전달을 위해 언어적 메시지와 표정, 억양, 시선 등의 비언어적 표현을 일치시킨다.", "correct": True, "exp": "⭕ 정답! 언어 표현과 비언어적 표현이 일치해야 상대방이 진심으로 받아들입니다."},
            {"text": "3. 내 감정을 숨기기 위해 로봇처럼 무표정과 기계적인 목소리로 대화한다.", "correct": False, "exp": "❌ 영혼 없는 대화는 오히려 상대를 거부하거나 무시하는 느낌을 줍니다."},
            {"text": "4. 말만 나-전달법 공식을 지키면 한숨이나 팔짱 같은 행동은 상관없다.", "correct": False, "exp": "❌ 대화에서 비언어적 요소가 차지하는 비중은 약 70%에 달합니다."},
            {"text": "5. 상대방이 미안함을 느끼도록 가벼운 한숨을 쉬며 눈을 피한다.", "correct": False, "exp": "❌ 수동공격적인 행동은 진정한 소통을 방해합니다."}
        ]
    },
    {
        "stage": "STAGE 7",
        "cat": "7단계: '과장된 단어' 배제하기",
        "char_name": "형 준호 (18세)",
        "avatar": "👦",
        "dialogue": "형이 다 쓴 휴지갑을 책상 위에 그대로 두고 방을 나간 상황!",
        "q": "나-전달법에서 상대를 억울하게 만드는 '과장·단정하는 단어(항상, 절대로 등)'를 빼고 올바르게 표현한 문장은?",
        "options": [
            {"text": "1. 형은 '항상', '단 한 번도' 자기 쓰레기를 제때 치운 적이 없잖아.", "correct": False, "exp": "❌ '항상', '단 한 번도' 같은 일반화 단어는 상대를 억울하게 만들어 반발을 부릅니다."},
            {"text": "2. 책상 위에 다 쓴 휴지갑이 있어 치우느라 내 흐름이 깨졌어. 다 쓴 휴지갑은 휴지통에 버려주면 좋겠어.", "correct": True, "exp": "⭕ 정답! '항상'이라는 과장 없이, 이번에 일어난 특정 상황과 영향만 객관적으로 말했습니다."},
            {"text": "3. 형은 도대체 정리정돈 개념이라는 게 있기는 해?", "correct": False, "exp": "❌ 상대의 습관과 인성을 단정 짓는 감정적 공격입니다."},
            {"text": "4. 내가 형 하인이야? 형 쓰레기는 형이 알아서 버려.", "correct": False, "exp": "❌ 도발적인 질문은 건설적인 갈등 해결을 방해합니다."},
            {"text": "5. 형이 자꾸 어지르니까 이 방 전체가 항상 난장판이 되는 거잖아.", "correct": False, "exp": "❌ '항상 난장판'이라는 과장적 표현으로 상대에게 죄책감을 강요하는 대화입니다."}
        ]
    },
    {
        "stage": "STAGE 8",
        "cat": "8단계: 감정형(F) 상대와의 대화법",
        "char_name": "친구 유진 (17세)",
        "avatar": "👩",
        "dialogue": "시험을 망쳐서 너무 속상하고 눈물이 난다며 마음을 털어놓는 친구!",
        "q": "공감과 관계를 중시하는 감정형(F) 친구의 마음을 열어주는 올바른 대화법은?",
        "options": [
            {"text": "1. 네가 열심히 준비했는데 결과가 안 나와서 얼마나 속상할지 느끼니 나도 마음이 아프다.", "correct": True, "exp": "⭕ 정답! 옳고 그름이나 논리적 조언보다 상대의 슬픈 감정에 먼저 깊이 공감해 주었습니다."},
            {"text": "2. 울어봤자 성적이 올라가지 않아. 공부 방법을 분석해서 해결책을 찾자.", "correct": False, "exp": "❌ 감정이 격해진 상태에서의 조기 해결책 제시나 논리적 분석은 상처를 줄 수 있습니다."},
            {"text": "3. 평소에 딴짓할 때 알아봤어. 다음부터는 오답 노트 작성 잘해라.", "correct": False, "exp": "❌ 상대의 약점을 지적하고 지적하는 평가적 태도입니다."},
            {"text": "4. 겨우 시험 하나 가지고 왜 그래? 다음 시험 준비나 해.", "correct": False, "exp": "❌ 상대의 감정을 가볍게 여기고 축소하는 대화는 소통을 차단합니다."},
            {"text": "5. 시험 결과에 일희일비하는 건 이성적이지 않은 행동이야.", "correct": False, "exp": "❌ 상대의 감정 표현을 평가절하하는 냉담한 대화 방식입니다."}
        ]
    },
    {
        "stage": "STAGE 9",
        "cat": "9단계: 사고형(T) 상대와의 대화법",
        "char_name": "삼촌 (35세)",
        "avatar": "👨",
        "dialogue": "원칙과 논리, 객관적 사실관계를 가장 중요하게 생각하는 삼촌과의 대화!",
        "q": "이성적이고 논리적인 사고형(T) 상대와 갈등을 해결할 때 가장 효과적인 대화법은?",
        "options": [
            {"text": "1. 무작정 서운하다고 떼를 쓰며 내 감정만 무조건 이해해달라고 호소한다.", "correct": False, "exp": "❌ 논리적 근거 없는 감정적 호소는 사고형 상대에게 설득력이 떨어집니다."},
            {"text": "2. 구체적인 사실관계와 원인을 차분하고 논리적으로 설명하며 바라는 점을 전달한다.", "correct": True, "exp": "⭕ 정답! 객관적 사실과 타당한 이유를 명확히 제시할 때 설득과 수용이 잘 이뤄집니다."},
            {"text": "3. 상대방의 논리적 오류를 하나하나 지적하며 말싸움에서 이기려 한다.", "correct": False, "exp": "❌ 논쟁에서 이기려 들면 갈등 해결이 아닌 자존심 싸움으로 변질됩니다."},
            {"text": "4. 논리적인 대화는 무의미하므로 대화를 완전히 포기하고 입을 닫는다.", "correct": False, "exp": "❌ 대화 포기는 갈등을 방치하고 관계를 악화시킵니다."},
            {"text": "5. 상대방의 냉정한 태도를 지적하며 감정적인 사과를 강요한다.", "correct": False, "exp": "❌ 사고형 상대의 성향을 이해하지 못하고 감정을 강요하는 행동입니다."}
        ]
    },
    {
        "stage": "STAGE 10",
        "cat": "10단계: 가족 갈등 해결 4단계 프로세스",
        "char_name": "가족 전체 (4인)",
        "avatar": "👨‍👩‍👧‍👦",
        "dialogue": "가족 회의에서 집안일 분담으로 생긴 갈등을 민주적으로 해결하려는 상황!",
        "q": "기술가정 시간에 배운 '가족 갈등 해결 4단계'의 올바른 순서는?",
        "options": [
            {"text": "1. 갈등 확인 ➔ 해결 방안 탐색 ➔ 최선의 방안 결정 ➔ 실행 및 평가", "correct": True, "exp": "⭕ 정답! 문제를 명확히 한 뒤 대안을 찾고, 합의하여 실행한 후 평가하는 것이 정석입니다."},
            {"text": "2. 최선의 방안 결정 ➔ 갈등 확인 ➔ 실행 및 평가 ➔ 해결 방안 탐색", "correct": False, "exp": "❌ 문제를 확인하기도 전에 결정부터 내리는 오류입니다."},
            {"text": "3. 해결 방안 탐색 ➔ 갈등 확인 ➔ 실행 및 평가 ➔ 최선의 방안 결정", "correct": False, "exp": "❌ 순서가 엉켜 합리적인 의사결정이 불가능합니다."},
            {"text": "4. 갈등 확인 ➔ 실행 및 평가 ➔ 해결 방안 탐색 ➔ 최선의 방안 결정", "correct": False, "exp": "❌ 실행을 대안 탐색보다 먼저 할 수 없습니다."},
            {"text": "5. 해결 방안 결정 ➔ 해결 방안 탐색 ➔ 갈등 확인 ➔ 실행 및 평가", "correct": False, "exp": "❌ 의사결정의 절차가 완전히 거꾸로 진행된 잘못된 순서입니다."}
        ]
    }
]

# 5. 세션 상태 관리
if "q_idx" not in st.session_state:
    st.session_state.q_idx = 0
if "score" not in st.session_state:
    st.session_state.score = 0
if "combo" not in st.session_state:
    st.session_state.combo = 0
if "selected_exp" not in st.session_state:
    st.session_state.selected_exp = None
if "is_correct" not in st.session_state:
    st.session_state.is_correct = None
if "submitted" not in st.session_state:
    st.session_state.submitted = False

# 6. 메인 UI 헤더 및 상태창
st.title("💬 나 전달법 마스터: 대화의 신")
st.caption("🏫 기술가정 대화법 프로젝트 | 상처 주지 않고 내 마음을 지혜롭게 전달하기")

# 화목도 게이지 및 스코어
score_percent = st.session_state.score
col_g1, col_g2 = st.columns([3, 1])
with col_g1:
    st.progress(score_percent / 100)
with col_g2:
    st.metric("가족/친구 화목도", f"{score_percent}%")

st.divider()

# 7. 게임 진행 화면 (0~9번 문제)
if st.session_state.q_idx < len(QUESTIONS):
    q = QUESTIONS[st.session_state.q_idx]

    # 스테이지 타이틀 & 콤보
    col_st1, col_st2 = st.columns([3, 1])
    with col_st1:
        st.subheader(f"🚩 {q['stage']}: {q['cat']}")
    with col_st2:
        if st.session_state.combo > 1:
            st.markdown(f"🔥 **{st.session_state.combo} COMBO!**")

    # 캐릭터 아바타 & 상황 말풍선
    st.markdown(f"""
    <div class="char-card">
        <div class="char-avatar">{q['avatar']}</div>
        <div class="char-name">{q['char_name']}</div>
    </div>
    <div class="speech-bubble">
        "{q['dialogue']}"
    </div>
    """, unsafe_allow_html=True)

    # 발문
    st.write(f"**❓ 질문:** {q['q']}")

    # 5개 선택지 및 해설 영역
    if st.session_state.selected_exp is None:
        for idx, opt in enumerate(q["options"]):
            if st.button(opt["text"], key=f"btn_{st.session_state.q_idx}_{idx}", use_container_width=True):
                st.session_state.selected_exp = opt["exp"]
                st.session_state.is_correct = opt["correct"]
                if opt["correct"]:
                    st.session_state.score += 10
                    st.session_state.combo += 1
                else:
                    st.session_state.combo = 0
                st.rerun()
    else:
        # 정답/오답 피드백
        if st.session_state.is_correct:
            st.success(st.session_state.selected_exp)
        else:
            st.error(st.session_state.selected_exp)

        if st.button("다음 문제로 이동 ➔", type="primary", use_container_width=True):
            st.session_state.q_idx += 1
            st.session_state.selected_exp = None
            st.session_state.is_correct = None
            st.rerun()

# 8. 최종 결과 및 구글 시트 저장 화면
else:
    st.balloons()
    st.header("🏆 학습 완료! 성적표 및 소통 칭호")
    
    # 점수별 칭호 부여
    final_score = st.session_state.score
    if final_score == 100:
        badge = "🥇 소통의 신 (마스터)"
    elif final_score >= 80:
        badge = "🥈 따뜻한 대화가 (전문가)"
    elif final_score >= 60:
        badge = "🥉 공감 노력파 (수련생)"
    else:
        badge = "🌱 대화 초보자 (재도전 필요)"

    col_res1, col_res2 = st.columns(2)
    with col_res1:
        st.metric("최종 점수", f"{final_score}점 / 100점")
    with col_res2:
        st.metric("획득 칭호", badge)

    st.divider()

    # 데이터 제출 폼
    if not st.session_state.submitted:
        st.subheader("📝 수행평가 결과 구글 시트 제출")
        with st.form("result_form"):
            std_id = st.text_input("학번 (예: 10101)", placeholder="학번 5자리를 입력하세요")
            name = st.text_input("이름", placeholder="이름을 입력하세요")
            submit_btn = st.form_submit_button("구글 시트에 제출하기", use_container_width=True)

            if submit_btn:
                if std_id and name:
                    with st.spinner("선생님 구글 시트로 안전하게 전송 중..."):
                        ok = submit_to_google_sheet(std_id, name, final_score, badge)
                    if ok:
                        st.session_state.submitted = True
                        st.rerun()
                else:
                    st.warning("학번과 이름을 모두 정확히 입력해 주세요!")
    else:
        st.success("🎉 선생님 구글 시트에 제출되었습니다!")
        if st.button("🔄 처음부터 다시 풀기", use_container_width=True):
            st.session_state.q_idx = 0
            st.session_state.score = 0
            st.session_state.combo = 0
            st.session_state.selected_exp = None
            st.session_state.is_correct = None
            st.session_state.submitted = False
            st.rerun()
