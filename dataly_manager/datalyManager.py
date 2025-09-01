# datalyManager.py
import streamlit as st
import zipfile
import tempfile
import os
import sys
import json
import importlib  # ← 추가

# (안전) 현재 디렉토리를 import 경로에 추가
APP_DIR = os.path.dirname(__file__)
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)

# 외부 기능 모듈(기존)
from newspaper_eval_merged import json_to_excel_stacked
from newspaper_eval_json import merge_newspaper_eval

# ─────────────────────────────────────────────────────────────
# 새로 분리한 변환 모듈: 직접 심볼 import 대신 모듈 import로 통일
import dataly_tools.table_to_excel as t2e
import dataly_tools.photo_to_excel as p2e
# ─────────────────────────────────────────────────────────────

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
                    importlib.reload(t2e)  # ← 최신 코드 보장
                    xlsx_bytes = t2e.table_json_to_xlsx_bytes(data)
                st.success("엑셀 생성 완료!")
                st.download_button(
                    label="표_변환.xlsx 다운로드",
                    data=xlsx_bytes,
                    file_name="표_변환.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    st.divider()
    st.subheader("🔁 엑셀의 ‘설명 문장’ → JSON 반영 (ZIP)")
    st.caption("ZIP 안에 .xlsx 1개와 project_*.json 1개가 있어야 합니다. 시트명을 비우면 첫 시트를 사용합니다.")
    apply_zip = st.file_uploader("ZIP 업로드 (Excel + JSON)", type=["zip"], key="zip_apply_desc_tab4")
    sheet_name = st.text_input("엑셀 시트명(선택)", value="", key="sheet_apply_desc_tab4")

    if st.button("적용 실행", key="btn_apply_desc_tab4"):
        if not apply_zip:
            st.error("ZIP 파일을 업로드하세요.")
        else:
            try:
                zip_bytes = apply_zip.getvalue()
                sheet_arg = sheet_name.strip() or None

                importlib.reload(t2e)  # ← 최신 코드 보장
                if not hasattr(t2e, "apply_excel_desc_to_json_from_zip"):
                    st.error("table_to_excel 모듈에 apply_excel_desc_to_json_from_zip가 없습니다.")
                    st.caption(f"loaded from: {t2e.__file__}")
                else:
                    updated_bytes, suggested_name = t2e.apply_excel_desc_to_json_from_zip(zip_bytes, sheet_arg)
            except Exception as e:
                st.error(f"적용 중 오류: {e}")
            else:
                st.success("JSON 업데이트 완료!")
                st.download_button(
                    label=f"{suggested_name} 다운로드",
                    data=updated_bytes,
                    file_name=suggested_name,
                    mime="application/json"
                )

# 사진 변환 (JSON→Excel) — photo_to_excel.py 사용
with tabs[5]:
    st.header("🖼️ 사진 변환 (단일 JSON → Excel)")
    st.info("project_*.json 1개를 업로드하면 엑셀로 변환합니다.")
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
                    importlib.reload(p2e)  # ← 최신 코드 보장
                    xlsx_bytes = p2e.photo_json_to_xlsx_bytes(data)
                st.success("엑셀 생성 완료!")
                st.download_button(
                    label="사진_변환.xlsx 다운로드",
                    data=xlsx_bytes,
                    file_name="사진_변환.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )

    st.divider()
    st.subheader("🔁 엑셀의 ‘설명 문장’ → JSON 반영 (ZIP)")
    st.caption("ZIP 안에 .xlsx 1개와 project_*.json 1개가 있어야 합니다. 시트명을 비우면 첫 시트를 사용합니다.")
    apply_zip_img = st.file_uploader("ZIP 업로드 (Excel + JSON)", type=["zip"], key="zip_apply_desc_tab5")
    sheet_name_img = st.text_input("엑셀 시트명(선택)", value="", key="sheet_apply_desc_tab5")

    if st.button("적용 실행 (사진)", key="btn_apply_desc_tab5"):
        if not apply_zip_img:
            st.error("ZIP 파일을 업로드하세요.")
        else:
            try:
                import importlib; importlib.reload(p2e)  # 최신 코드 보장
                zip_bytes = apply_zip_img.getvalue()
                sheet_arg = sheet_name_img.strip() or None
                updated_bytes, suggested_name = p2e.apply_excel_desc_to_json_from_zip(zip_bytes, sheet_arg)
            except Exception as e:
                st.error(f"적용 중 오류: {e}")
            else:
                st.success("JSON 업데이트 완료!")
                st.download_button(
                    label=f"{suggested_name} 다운로드",
                    data=updated_bytes,
                    file_name=suggested_name,
                    mime="application/json"
                )


st.markdown("""
<hr>
<div class="footer">
문의: 검증 엔지니어 | Powered by Streamlit<br>
Copyright &copy; 2025. All rights reserved.
</div>
""", unsafe_allow_html=True)
