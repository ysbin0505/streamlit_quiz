# dataly_manager/ui/wsd_to_excel_ui.py
import os
import sys
import streamlit as st

# 패키지 루트(= dataly_manager의 부모) 경로 세팅 - 다른 UI 파일과 동일 패턴
APP_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(APP_DIR)
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from dataly_manager.dataly_tools import jsons_to_wsd_excel

def render_wsd_to_excel_ui():
    st.header("📄 WSD/DP/SRL/ZA → 엑셀 변환")

    with st.expander("도움말", expanded=False):
        st.markdown("""
        - 폴더 안의 `*.json`을 스캔해 **WSD 시트**(단어 단위)와 선택 시 **Memos 시트**를 생성합니다.  
        - SRL/ZA 정보는 다음 컬럼으로 추출됩니다.  
          - **SRL**: `SRL Span`, `SRL Label`, `SRL Predicate Lamma`  
          - **ZA**: `ant_sen_id`, `ant_word_id`, `ant_form`, `restored_form`, `restored_type`  
        - `SRL Span`은 **argument의 word_id**,  
          `SRL Predicate Lamma`는 **predicate의 `word_id/lemma`** 형식입니다.
        """)

    col1, col2 = st.columns([2, 1], gap="large")
    with col1:
        base_dir = st.text_input("변환할 JSON 폴더 경로", value="", placeholder="/path/to/json/dir")
        excel_name = st.text_input("저장 파일명", value="WSD_sense_tagging_simple.xlsx")

    with col2:
        include_memo_sheet = st.checkbox("Memos 시트 포함", value=True)
        memo_placement = st.selectbox(
            "메모 배치 방식",
            options=["by_row", "first", "repeat"],
            index=0,
            help="- by_row: 메모의 row == word_id 인 행만 기입\n- first: 문장 첫 단어 행만 기입\n- repeat: 문장 내 모든 단어 행에 반복"
        )
        memo_sep = st.text_input("메모 구분자", value=" | ")

    run = st.button("🚀 변환 실행", type="primary", use_container_width=True)

    if run:
        if not base_dir or not os.path.isdir(base_dir):
            st.error("유효한 폴더 경로를 입력해주세요.")
            return

        with st.status("변환 중입니다...", expanded=True) as status:
            try:
                out_path = jsons_to_wsd_excel(
                    base_dir=base_dir,
                    excel_name=excel_name,
                    include_memo_sheet=include_memo_sheet,
                    memo_placement=memo_placement,
                    memo_sep=memo_sep,
                )
                status.update(label="완료!", state="complete")
            except Exception as e:
                status.update(label="에러 발생", state="error")
                st.exception(e)
                return

        if os.path.exists(out_path):
            st.success(f"엑셀 파일 생성: {out_path}")
            with open(out_path, "rb") as f:
                st.download_button(
                    label="⬇️ 엑셀 다운로드",
                    data=f.read(),
                    file_name=os.path.basename(out_path),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

            # 미리보기(상위 100행)
            try:
                import pandas as pd
                df_preview = pd.read_excel(out_path, sheet_name="WSD", nrows=100)
                st.subheader("미리보기 (WSD 시트 상위 100행)")
                st.dataframe(df_preview, use_container_width=True, height=400)
            except Exception:
                st.info("미리보기를 열 수 없습니다. 파일을 직접 확인해주세요.")
