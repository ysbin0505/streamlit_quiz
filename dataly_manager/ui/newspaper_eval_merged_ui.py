# ui/newspaper_eval_merged_ui.py
import streamlit as st
import zipfile
import tempfile
import os
from dataly_manager.dataly_tools.newspaper_eval_merged import json_to_excel_stacked

def render_sum_eval_tab():
    st.header("📰 신문평가 JSON → 엑셀 자동 수합기")
    st.info("아래 순서대로 업로드 및 실행을 진행하세요.")

    uploaded_zip = st.file_uploader("1. 평가 데이터 ZIP 업로드 (폴더를 압축)", type=["zip"], key="zip_sum")
    sum_week_num = st.number_input("2. 수합할 주차 (예: 1)", min_value=1, step=1, value=1, key="week_sum")
    storage_folder = st.selectbox("3. storage 폴더명 선택", ["storage0", "storage1"], key="storage_sum")

    if st.button("실행 (엑셀 변환)", key="btn_sum"):
        if not uploaded_zip:
            st.error("ZIP 파일을 업로드하세요.")
            return

        with tempfile.TemporaryDirectory() as temp_dir:
            zip_path = os.path.join(temp_dir, "data.zip")
            with open(zip_path, "wb") as f:
                f.write(uploaded_zip.read())

            with zipfile.ZipFile(zip_path, "r") as zip_ref:
                zip_ref.extractall(temp_dir)

            folder_list = [f for f in os.listdir(temp_dir) if os.path.isdir(os.path.join(temp_dir, f))]
            if not folder_list:
                st.error("압축 내부에 폴더가 없습니다. 폴더째 압축한 zip만 지원합니다.")
                return

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
