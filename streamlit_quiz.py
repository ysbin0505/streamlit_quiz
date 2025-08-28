#streamlit_quiz.py
import streamlit as st
import os
import json
import random

# ===== 유틸리티 =====
def reset_to_home():
    st.session_state.app_mode = 'setup'
    st.session_state.show_answer = False
    st.session_state.submitted = False
    st.session_state.step = 0
    st.session_state.score = 0
    st.session_state.finished = False

# ===== "처음으로" 버튼 (항상 상단에) =====
st.sidebar.markdown("## 🚀 메뉴")
if st.sidebar.button("🏠 처음으로", key="btn_home_sidebar"):
    reset_to_home()
    st.rerun()

# ===== 데이터 폴더 및 파일 리스트 =====
DATA_DIR = './quiz_data'
files = [f for f in os.listdir(DATA_DIR) if f.endswith('.json')]
subjects = [os.path.splitext(f)[0] for f in files]
if not subjects:
    st.error("❌ quiz_data 폴더에 json 파일이 없습니다.")
    st.stop()

if 'app_mode' not in st.session_state:
    st.session_state.app_mode = 'setup'
if 'selected_subject' not in st.session_state:
    st.session_state.selected_subject = subjects[0]
if 'order_mode' not in st.session_state:
    st.session_state.order_mode = "랜덤"
if 'solve_mode' not in st.session_state:
    st.session_state.solve_mode = "한 문제씩(즉시 채점)"
if 'last_judged' not in st.session_state:
    st.session_state.last_judged = None  # None, 'correct', 'wrong'

# ===== 1. 첫 화면: 모든 옵션 선택 =====
if st.session_state.app_mode == 'setup':
    st.markdown("<h1 style='color:#0066CC'>📝 영양교육 객관식 퀴즈</h1>", unsafe_allow_html=True)
    st.markdown("---")

    st.markdown("#### 1. 과목(파일) 선택")
    subject = st.selectbox("과목(파일)", subjects, index=subjects.index(st.session_state.selected_subject))
    st.session_state.selected_subject = subject

    st.markdown("#### 2. 문제 순서")
    order_mode = st.radio("문제 순서", ["랜덤", "순차"], index=0 if st.session_state.order_mode == "랜덤" else 1, horizontal=True)
    st.session_state.order_mode = order_mode

    st.markdown("#### 3. 풀이 모드")
    solve_mode = st.radio("풀이 모드", ["한 문제씩(즉시 채점)", "모의고사(최종 제출)"], index=0 if st.session_state.solve_mode.startswith("한 문제씩") else 1, horizontal=True)
    st.session_state.solve_mode = solve_mode

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("🚩 문제풀이 시작", use_container_width=True):
        filepath = os.path.join(DATA_DIR, subject + '.json')
        with open(filepath, encoding='utf-8') as f:
            questions = json.load(f)
        st.session_state.questions = questions
        indices = list(range(len(questions)))
        if order_mode == "랜덤":
            random.shuffle(indices)
        st.session_state.quiz_order = indices
        st.session_state.step = 0
        st.session_state.score = 0
        st.session_state.last_input = ""
        st.session_state.show_answer = False
        st.session_state.inputs = [None] * len(questions)
        st.session_state.answered = [False] * len(questions)
        st.session_state.finished = False
        st.session_state.submitted = False
        st.session_state.last_judged = None
        st.session_state.app_mode = 'quiz'
        st.rerun()
    st.stop()

# ===== 2. 퀴즈풀이 화면 (모드별) =====
questions = st.session_state.questions
order = st.session_state.quiz_order
step = st.session_state.step
score = st.session_state.score
solve_mode = st.session_state.solve_mode
inputs = st.session_state.inputs

if st.session_state.app_mode == 'quiz':
    st.markdown(f"<h2 style='color:#0066CC'>[{st.session_state.selected_subject}] 퀴즈풀이</h2>", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown(f"<div style='padding:8px 0 0 0; color:#333'><b>문제 {step+1} / {len(questions)}</b></div>", unsafe_allow_html=True)
    st.progress((step + 1) / len(questions) if step < len(questions) else 1.0)
    st.markdown(f"<span style='color: #16a34a; font-weight:700;'>현재 점수: {score}</span>", unsafe_allow_html=True)

# ===== 한 문제씩(즉시 채점) 모드 =====
if solve_mode == "한 문제씩(즉시 채점)" and not st.session_state.finished:
    if step >= len(order):
        st.session_state.finished = True
        st.rerun()

    idx = order[step]
    q = questions[idx]
    st.markdown("-----")
    st.markdown(f"<b style='font-size:1.1em;'>{q['question']}</b>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)

    # 객관식 라디오버튼 입력
    # 입력 내역이 있으면 해당 선택지로 index 지정
    index_val = q["choices"].index(inputs[idx]) if inputs[idx] in q["choices"] else None
    choice = st.radio("정답을 선택하세요", q["choices"], key=f"choice_{step}", index=index_val)
    inputs[idx] = choice

    answer_btn_col, submit_col = st.columns([1, 1])
    with answer_btn_col:
        # 보기/숨기기 기능은 필요하다면 여기에 추가
        pass
    with submit_col:
        if st.button("✅ 제출", key=f"submit_{step}", use_container_width=True):
            if not choice:
                st.warning("정답을 선택해 주세요!")
            else:
                if choice == q["answer"]:
                    st.success("🎉 정답입니다! 대단해요!")
                    st.balloons()
                    st.session_state.last_judged = 'correct'
                    st.session_state.score += 1
                else:
                    st.error("😥 오답입니다... 조금만 더 힘내요!")
                    st.snow()
                    st.session_state.last_judged = 'wrong'
                st.session_state.answered[idx] = True
                st.session_state.submitted = True
                st.session_state.show_answer = True  # 제출 시 정답 보기
                st.rerun()

    if st.session_state.show_answer:
        st.info(f"정답: {q['answer']}")

    # 최근 맞음/틀림 이모지
    if st.session_state.get('last_judged') == 'correct':
        st.markdown("<div style='font-size:2em;'>🎊 🎉 🥳</div>", unsafe_allow_html=True)
    elif st.session_state.get('last_judged') == 'wrong':
        st.markdown("<div style='font-size:2em;'>😭 ❄️ 💧</div>", unsafe_allow_html=True)

    if st.session_state.get('submitted', False):
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("➡️ 다음 문제", use_container_width=True):
            st.session_state.step += 1
            st.session_state.submitted = False
            st.session_state.last_judged = None
            st.session_state.show_answer = False
            st.rerun()

    if step >= len(questions) - 1 and st.session_state.answered[idx]:
        st.session_state.finished = True
        st.rerun()

# ===== 모의고사(전체 제출) 모드 =====
elif solve_mode == "모의고사(최종 제출)" and not st.session_state.finished:
    st.markdown("-----")
    st.write("모든 문제에 답을 입력한 뒤 **최종 제출** 버튼을 누르세요.")
    for i, idx in enumerate(order):
        q = questions[idx]
        with st.expander(f"문제 {i+1}: {q['question']}"):
            # 입력 내역이 있으면 해당 선택지로 index 지정
            index_val = q["choices"].index(inputs[idx]) if inputs[idx] in q["choices"] else None
            choice = st.radio("정답을 선택하세요", q["choices"], key=f"mock_choice_{i}", index=index_val)
            inputs[idx] = choice
            # 정답보기 기능이 필요하다면 아래처럼 사용
            # st.info(f"정답: {q['answer']}")

    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("📊 최종 제출", use_container_width=True):
        score = 0
        for i, idx in enumerate(order):
            user_choice = inputs[idx]
            q = questions[idx]
            if user_choice == q["answer"]:
                score += 1
        st.session_state.score = score
        st.session_state.finished = True
        st.rerun()

# ===== 3. 결과 및 재시작 =====
if st.session_state.finished:
    st.balloons()
    st.markdown(f"<h2 style='color:#0066CC;'>🥳 퀴즈 완료!</h2>", unsafe_allow_html=True)
    st.markdown(f"<div style='font-size:1.5em; color:#333;'>총 <span style='color:#2563eb;font-weight:700;'>{len(questions)}</span>문제 중 <span style='color:#16a34a;font-weight:700;'>{st.session_state.score}</span>개 맞추셨습니다!</div>", unsafe_allow_html=True)
    st.markdown("<br>", unsafe_allow_html=True)
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("🔄 다시 시작", use_container_width=True):
            reset_to_home()
            st.rerun()
    with col2:
        if st.button("🏠 처음으로", key="btn_home_final", use_container_width=True):
            reset_to_home()
            st.rerun()
