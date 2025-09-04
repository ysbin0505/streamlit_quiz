# datalyManager.py
import streamlit as st
import zipfile
import tempfile
import os
import sys
import json
import importlib

# (안전) 현재 디렉토리를 import 경로에 추가
APP_DIR = os.path.dirname(__file__)
if APP_DIR not in sys.path:
    sys.path.insert(0, APP_DIR)

from newspaper_eval_merged_ui import render_sum_eval_tab
import dataly_tools.table_to_excel as t2e
import dataly_tools.photo_to_excel as p2e

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
    st.image("https://static.streamlit.io/examples/dog.jpg", width=55)
with col2:
    st.markdown('<div class="main-title">Dataly Manager</div>', unsafe_allow_html=True)
    st.markdown('<div class="sub-desc">업무 자동화, 데이터 변환 및 관리 웹앱</div>', unsafe_allow_html=True)

tabs = st.tabs([
    "🏠 홈",
    "📰 신문평가 수합",
    "📊 표 변환 (JSON→Excel)",
    "🖼️ 사진 변환 (JSON→Excel)",
    "🧪 정합성 검사"
])

# 홈
with tabs[0]:
    st.header("🏠 홈")
    st.markdown("""
    왼쪽 탭에서 원하는 기능을 선택해 사용하세요.
    """)
    st.markdown("### 빠른 안내")
    st.markdown("""
    - **📰 신문평가 수합**: ZIP을 업로드하면 주차별로 엑셀을 생성합니다.
    - **📊 표 변환**: `project_*.json` → 엑셀 변환 및 엑셀의 설명을 JSON에 반영.  
    - **🖼️ 사진 변환**: 사진용 `project_*.json` → 엑셀 변환 및 설명 반영.
    """)

    st.divider()
    st.markdown("### 빠른 점검")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.metric("Python", sys.version.split()[0])
    with col_b:
        st.metric("작업 디렉토리", os.path.basename(APP_DIR))
    with col_c:
        st.metric("캐시", "스트림릿 런타임")

    st.caption("상단 탭에서 기능을 선택해 주세요.")

# 신문평가 수합
with tabs[1]:
    render_sum_eval_tab()
# 표 변환 (JSON→Excel) — table_to_excel.py 사용
with tabs[2]:
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
with tabs[3]:
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
