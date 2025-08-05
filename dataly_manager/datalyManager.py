import streamlit as st
import zipfile
import tempfile
import os
from newspaper_eval_merged import json_to_excel_stacked

st.set_page_config(page_title="신문평가 자동 병합기", layout="centered")
st.title("📰 신문평가 JSON → 엑셀 자동 병합기")

st.markdown("""
#### 사용법 안내
1. **평가 데이터가 들어있는 폴더**를 ZIP 압축해서 업로드하세요.
2. **주차(week_num)**와 **스토리지 폴더명**을 입력한 후 실행을 누르세요.
3. 변환이 끝나면 아래에서 **엑셀 파일을 다운로드** 할 수 있습니다.
""")

# --- 1. ZIP 파일 업로드
uploaded_zip = st.file_uploader("1. 평가 데이터 ZIP 업로드 (폴더를 압축)", type=["zip"])
week_num = st.number_input("2. 병합할 주차 (예: 1)", min_value=1, step=1, value=1)
storage_folder = st.selectbox("3. storage 폴더명 선택", ["storage0", "storage1"])

run_btn = st.button("실행 (엑셀 변환)")

output_ready = False
output_excel = None

if uploaded_zip and run_btn:
    # 임시 디렉토리 생성 및 ZIP 해제
    with tempfile.TemporaryDirectory() as temp_dir:
        zip_path = os.path.join(temp_dir, "data.zip")
        with open(zip_path, "wb") as f:
            f.write(uploaded_zip.read())
        with zipfile.ZipFile(zip_path, "r") as zip_ref:
            zip_ref.extractall(temp_dir)
        # 루트 폴더 탐색 (압축 내 최상위 폴더)
        folder_list = [f for f in os.listdir(temp_dir) if os.path.isdir(os.path.join(temp_dir, f))]
        if not folder_list:
            st.error("압축파일 내부에 폴더가 없습니다. 폴더째 압축한 zip만 지원합니다.")
        else:
            root_path = os.path.join(temp_dir, folder_list[0])
            # 실제 변환 실행
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
    st.info("좌측/위 입력값을 모두 지정하고 [실행]을 눌러주세요.")

st.markdown("---")
st.markdown("문의: 검증 엔지니어 | Powered by Streamlit")

