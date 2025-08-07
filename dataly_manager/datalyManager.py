import streamlit as st
import zipfile
import tempfile
import os
from newspaper_eval_merged import json_to_excel_stacked
from newspaper_eval_json import merge_newspaper_eval


# --- 커스텀 CSS (예: 헤더, 버튼, 공통 배경 등) ---
st.markdown("""
    <style>
    .main-title {font-size:2.1rem; font-weight:bold; color:#174B99; margin-bottom:0;}
    .sub-desc {font-size:1.1rem; color:#222;}
    .logo {height:60px; margin-bottom:15px;}
    .footer {color:#999; font-size:0.9rem; margin-top:40px;}
    /* 버튼 개선 */
    div.stButton > button:first-child {background:#174B99; color:white; font-weight:bold; border-radius:8px;}
    /* 탭 선택 강조 */
    .stTabs [data-baseweb="tab-list"] {background:#F6FAFD;}
    </style>
""", unsafe_allow_html=True)

# --- 상단 브랜드/로고 (이미지 경로는 직접 지정) ---
col1, col2 = st.columns([0.15, 0.85])
with col1:
    st.image("https://static.streamlit.io/examples/cat.jpg", width=55)   # 로고 URL or 파일경로
with col2:
    st.markdown('<div class="main-title">Dataly Manager</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-desc">업무 자동화, 평가 데이터 변환, 관리, 수합 웹앱</div>', unsafe_allow_html=True)

# --- 상단 탭 메뉴 ---
tabs = st.tabs(["🏠 홈", "📰 신문평가 수합", "💬 대화평가 병합", "📦 신문평가 병합"])

# --- 각 탭별 컨텐츠 ---
with tabs[0]:  # 홈
    st.markdown("#### 👋 환영합니다!<br>아래 탭에서 기능을 선택해 주세요.", unsafe_allow_html=True)
    st.markdown("""
    - **📰 신문평가 수합**: 신문 JSON을 엑셀로 변환
    - **💬 대화평가 병합**: 대화 평가 병합 (추가예정)
    """)

with tabs[1]:  # 신문평가 수합
    st.header("📰 신문평가 JSON → 엑셀 자동 수합기")
    st.info("아래 순서대로 업로드 및 실행을 진행하세요.")
    uploaded_zip = st.file_uploader("1. 평가 데이터 ZIP 업로드 (폴더를 압축)", type=["zip"])
    week_num = st.number_input("2. 수합할 주차 (예: 1)", min_value=1, step=1, value=1)
    storage_folder = st.selectbox("3. storage 폴더명 선택", ["storage0", "storage1"])
    run_btn = st.button("실행 (엑셀 변환)", key="run_newspaper")

    if uploaded_zip and run_btn:
        with tempfile.TemporaryDirectory() as temp_dir:
            zip_path = os.path.join(temp_dir, "data.zip")
            with open(zip_path, "wb") as f:
                f.write(uploaded_zip.read())
            with zipfile.ZipFile(zip_path, "r") as zip_ref:
                zip_ref.extractall(temp_dir)
            folder_list = [f for f in os.listdir(temp_dir) if os.path.isdir(os.path.join(temp_dir, f))]
            if not folder_list:
                st.error("압축파일 내부에 폴더가 없습니다. 폴더째 압축한 zip만 지원합니다.")
            else:
                root_path = os.path.join(temp_dir, folder_list[0])
                st.info("엑셀 변환을 시작합니다. (수초~수십초 소요)")
                json_to_excel_stacked(root_path, week_num, storage_folder)
                excel_path = os.path.join(root_path, "summary_eval_all.xlsx")
                if os.path.exists(excel_path):
                    with open(excel_path, "rb") as f:
                        st.success("엑셀 변환 완료! 아래 버튼으로 다운로드하세요.")
                        st.download_button(
                            label="summary_eval_all.xlsx 다운로드",
                            data=f,
                            file_name="summary_eval_all.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                        )
                else:
                    st.error("엑셀 파일 생성 실패. 내부 오류를 확인하세요.")
    else:
        st.info("ZIP, 주차, 폴더명 입력 후 [실행]을 눌러주세요.")

with tabs[2]:  # 대화평가 병합 (예시, 추후 구현)
    st.header("💬 대화평가 병합 (준비중)")
    st.info("이 기능은 곧 추가됩니다. 원하시는 기능이 있다면 문의해 주세요.")

with tabs[3]:  # 신문평가 병합
    st.header("📦 신문평가 JSON 병합")
    st.info("""
    이 기능은 지정 경로에 있는 신문 JSON 데이터를 주차별로 병합하여 새로운 JSON을 생성합니다.
    (폴더 구조/경로는 코드에 명시된 대로 /Users/data.ly/Desktop/말뭉치배포/신문 내 A팀/B팀/merged가 필요합니다.)
    """)

    week_num = st.number_input("병합할 주차 (예: 1)", min_value=1, step=1, value=1)
    files_per_week = st.number_input("병합할 파일 수 (보통 102)", min_value=1, step=1, value=102)
    run_merge_btn = st.button("신문평가 병합 실행", key="run_newspaper_merge")

    if run_merge_btn:
        with st.spinner("병합 중입니다... (경로/입력값 확인)"):
            msg = merge_newspaper_eval(week_num=int(week_num), files_per_week=int(files_per_week))
        st.success(f"병합 결과: {msg}")
        # 필요시 병합된 폴더 링크 안내 등 추가


# --- 하단 푸터/상태바 ---
st.markdown("""
<hr>
<div class="footer">
문의: 검증 엔지니어 | Powered by Streamlit<br>
Copyright &copy; 2025. All rights reserved.
</div>
""", unsafe_allow_html=True)
