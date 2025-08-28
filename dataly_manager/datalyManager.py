# datalyManager.py
import streamlit as st
import zipfile
import tempfile
import os
import json

# 외부 기능 모듈(기존)
from newspaper_eval_merged import json_to_excel_stacked
from newspaper_eval_json import merge_newspaper_eval

# 새로 분리한 변환 모듈
from dataly_tools.table_to_excel import table_json_to_xlsx_bytes
from dataly_tools.photo_to_excel import photo_json_to_xlsx_bytes

st.markdown("""
    <style>
    .main-title {font-size:2.1rem; font-weight:bold; color:#174B99; margin-bottom:0;}
    .sub-desc {font-size:1.1rem; color:#222;}
    .logo {height:60px; margin-bottom:15px;}
    .footer {color:#999; font-size:0.9rem; margin-top:40px;}
    div.stButton > button:first-child {background:#174B99; color:white; font-weight:bold; border-radius:8px;}
    .stTabs [data-baseweb="tab-list"] {background:#F6FAFD;}
    </style>
""", unsafe_allow_html=True)

col1, col2 = st.columns([0.15, 0.85])
with col1:
    st.image("https://static.streamlit.io/examples/cat.jpg", width=55)
with col2:
    st.markdown('<div class="main-title">Dataly Manager</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-desc">업무 자동화, 평가 데이터 변환, 관리, 수합 웹앱</div>', unsafe_allow_html=True)

tabs = st.tabs([
    "🏠 홈",
    "📰 신문평가 수합",
    "💬 대화평가 병합",
    "📦 신문평가 병합",
    "📊 표 변환 (JSON→Excel)",
    "🖼️ 사진 변환 (JSON→Excel)"
])

# 홈
with tabs[0]:
    st.markdown("#### 👋 환영합니다!<br>아래 탭에서 기능을 선택해 주세요.", unsafe_allow_html=True)
    st.markdown("""
    - 📰 신문평가 수합: ZIP→엑셀
    - 💬 대화평가 병합: 준비중
    - 📦 신문평가 병합: A/B팀 JSON ZIP 병합
    - 📊 표 변환: 단일 JSON→엑셀
    - 🖼️ 사진 변환: 단일 JSON→엑셀(이미지용 스키마)
    """)

# 신문평가 수합
with tabs[1]:
    st.header("📰 신문평가 JSON → 엑셀 자동 수합기")
    st.info("아래 순서대로 업로드 및 실행을 진행하세요.")
    uploaded_zip = st.file_uploader("1. 평가 데이터 ZIP 업로드 (폴더를 압축)", type=["zip"], key="zip_sum")
    sum_week_num = st.number_input("2. 수합할 주차 (예: 1)", min_value=1, step=1, value=1, key="week_sum")
    storage_folder = st.selectbox("3. storage 폴더명 선택", ["storage0", "storage1"], key="storage_sum")
    if st.button("실행 (엑셀 변환)", key="btn_sum"):
        if not uploaded_zip:
            st.error("ZIP 파일을 업로드하세요.")
        else:
            with tempfile.TemporaryDirectory() as temp_dir:
                zip_path = os.path.join(temp_dir, "data.zip")
                with open(zip_path, "wb") as f:
                    f.write(uploaded_zip.read())
                with zipfile.ZipFile(zip_path, "r") as zip_ref:
                    zip_ref.extractall(temp_dir)
                folder_list = [f for f in os.listdir(temp_dir) if os.path.isdir(os.path.join(temp_dir, f))]
                if not folder_list:
                    st.error("압축 내부에 폴더가 없습니다. 폴더째 압축한 zip만 지원합니다.")
                else:
                    root_path = os.path.join(temp_dir, folder_list[0])
                    st.info("엑셀 변환 중입니다…")
                    json_to_excel_stacked(root_path, sum_week_num, storage_folder)
                    excel_path = os.path.join(root_path, "summary_eval_all.xlsx")
                    if os.path.exists(excel_path):
                        with open(excel_path, "rb") as f:
                            st.success("엑셀 변환 완료!")
                            st.download_button(
                                label="summary_eval_all.xlsx 다운로드",
                                data=f,
                                file_name="summary_eval_all.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        st.error("엑셀 파일 생성 실패. 내부 오류를 확인하세요.")

# 대화평가 병합 (준비중)
with tabs[2]:
    st.header("💬 대화평가 병합 (준비중)")
    st.info("원하시는 기능이 있다면 요청해 주세요.")

# 신문평가 병합
with tabs[3]:
    st.header("📦 신문평가 JSON 병합")
    st.info("ZIP 내 'A/A팀', 'B/B팀' 폴더와 JSON 파일이 있어야 합니다.")
    uploaded_zip = st.file_uploader("병합할 신문 원본 ZIP 업로드 (A/B팀 포함 폴더)", type=["zip"], key="zip_merge")
    merge_week_num = st.number_input("병합할 주차 (예: 1)", min_value=1, step=1, value=1, key="week_merge")
    files_per_week = st.number_input("병합할 파일 수 (보통 102)", min_value=1, step=1, value=102, key="files_per_week")
    if st.button("신문평가 병합 실행", key="btn_merge"):
        if not uploaded_zip:
            st.error("ZIP 파일을 업로드하세요.")
        else:
            with tempfile.TemporaryDirectory() as temp_dir:
                zip_path = os.path.join(temp_dir, "src.zip")
                with open(zip_path, "wb") as f:
                    f.write(uploaded_zip.read())
                with zipfile.ZipFile(zip_path, "r") as zip_ref:
                    zip_ref.extractall(temp_dir)

                candidate_dirs = [os.path.join(temp_dir, d) for d in os.listdir(temp_dir)
                                  if os.path.isdir(os.path.join(temp_dir, d))]
                if not candidate_dirs:
                    st.error("압축 내부 폴더를 찾을 수 없습니다. ZIP 폴더 구조를 확인하세요.")
                else:
                    base_dir = candidate_dirs[0]
                    with st.spinner("병합 중입니다..."):
                        msg, output_dir, out_zip_path = merge_newspaper_eval(
                            week_num=int(merge_week_num),
                            files_per_week=int(files_per_week),
                            base_dir=base_dir
                        )
                    st.success(f"병합 결과: {msg}")
                    with open(out_zip_path, "rb") as f:
                        st.download_button(
                            label=f"{merge_week_num}주차 병합 JSON ZIP 다운로드",
                            data=f,
                            file_name=f"merged_{merge_week_num}주차.zip",
                            mime="application/zip"
                        )

# 표 변환 (JSON→Excel) — table_to_excel.py 사용
with tabs[4]:
    st.header("📊 표 변환 (단일 JSON → Excel)")
    st.info("project_*.json 1개를 업로드하면 표 형태 엑셀로 변환합니다.")
    uploaded_json = st.file_uploader("JSON 업로드 (project_*.json)", type=["json"], key="json_table")
    if st.button("엑셀 변환 실행", key="btn_table"):
        if not uploaded_json:
            st.error("JSON 파일을 업로드하세요.")
        else:
            try:
                raw = uploaded_json.getvalue()
                data = json.loads(raw)
            except Exception as e:
                st.error(f"JSON 파싱 실패: {e}")
            else:
                with st.spinner("엑셀 생성 중..."):
                    xlsx_bytes = table_json_to_xlsx_bytes(data)
                st.success("엑셀 생성 완료!")
                st.download_button(
                    label="표_변환.xlsx 다운로드",
                    data=xlsx_bytes,
                    file_name="표_변환.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

# 사진 변환 (JSON→Excel) — photo_to_excel.py 사용
with tabs[5]:
    st.header("🖼️ 사진 변환 (단일 JSON → Excel)")
    st.info("project_*.json 1개를 업로드하면 이미지 전용 스키마를 엑셀로 변환합니다.")
    uploaded_json_img = st.file_uploader("JSON 업로드 (project_*.json)", type=["json"], key="json_photo")
    if st.button("엑셀 변환 실행", key="btn_photo"):
        if not uploaded_json_img:
            st.error("JSON 파일을 업로드하세요.")
        else:
            try:
                raw = uploaded_json_img.getvalue()
                data = json.loads(raw)
            except Exception as e:
                st.error(f"JSON 파싱 실패: {e}")
            else:
                with st.spinner("엑셀 생성 중..."):
                    xlsx_bytes = photo_json_to_xlsx_bytes(data)
                st.success("엑셀 생성 완료!")
                st.download_button(
                    label="사진_변환.xlsx 다운로드",
                    data=xlsx_bytes,
                    file_name="사진_변환.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

st.markdown("""
<hr>
<div class="footer">
문의: 검증 엔지니어 | Powered by Streamlit<br>
Copyright &copy; 2025. All rights reserved.
</div>
""", unsafe_allow_html=True)
