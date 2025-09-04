# ui/table_to_excel_ui.py
import streamlit as st
import json, importlib
from dataly_manager.dataly_tools import table_to_excel as t2e


def render_table_to_excel():
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
                    importlib.reload(t2e)  # 최신 코드 보장
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

                importlib.reload(t2e)  # 최신 코드 보장
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
