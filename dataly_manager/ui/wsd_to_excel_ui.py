# dataly_manager/ui/wsd_to_excel_ui.py
import os
import sys
import io
import tempfile
import zipfile
import streamlit as st

# 패키지 루트(= dataly_manager의 부모) 경로 세팅
APP_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(APP_DIR)
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from dataly_manager.dataly_tools.wsd_to_excel import jsons_to_wsd_excel

def render_wsd_to_excel_ui():
    st.header("📄 WSD/DP/SRL/ZA → 엑셀 변환")

    with st.expander("도움말", expanded=False):
        st.markdown("""
        - **ZIP 업로드** 또는 **로컬 폴더 경로** 중 하나로 입력하세요. (ZIP이 있으면 ZIP이 우선됩니다)
        - 폴더/ZIP 안의 모든 하위 폴더까지 재귀적으로 `*.json`을 스캔합니다.
        - 생성 시트
          - **WSD**: 단어 단위 테이블 (+ DP, SRL, ZA 컬럼 포함)
          - **Memos**(옵션): 문장/문서 메모 목록
        - SRL/ZA 컬럼
          - **SRL**: `SRL Span`, `SRL Label`, `SRL Predicate Lamma`
          - **ZA**: `ant_sen_id`, `ant_word_id`, `ant_form`, `restored_form`, `restored_type`
        - `SRL Span`은 *argument의 word_id*, `SRL Predicate Lamma`는 *predicate의 `word_id/lemma`* 형식입니다.
        """)

    col1, col2 = st.columns([2, 1], gap="large")
    with col1:
        uploaded_zip = st.file_uploader("JSON ZIP 업로드", type=["zip"])
        base_dir = st.text_input("또는 변환할 JSON **폴더 경로**", value="", placeholder="/path/to/json/dir")
        excel_name = st.text_input("저장 파일명", value="SRL_ZA.xlsx")

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
        # 입력 검증
        if not uploaded_zip and not (base_dir and os.path.isdir(base_dir)):
            st.error("ZIP을 업로드하거나, 유효한 폴더 경로를 입력해 주세요.")
            return

        excel_bytes = None
        out_path_display = None

        with st.status("변환 중입니다...", expanded=True) as status:
            try:
                if uploaded_zip:
                    # ZIP → 임시 폴더로 해제 후 그 폴더를 대상으로 변환
                    with tempfile.TemporaryDirectory() as tmpdir:
                        zpath = os.path.join(tmpdir, "input.zip")
                        with open(zpath, "wb") as f:
                            f.write(uploaded_zip.getbuffer())

                        with zipfile.ZipFile(zpath) as zf:
                            zf.extractall(tmpdir)

                        # ZIP 파일명으로 기본 결과 이름 제안
                        if excel_name.strip() == "SRL_ZA.xlsx" and uploaded_zip.name:
                            base_name = os.path.splitext(os.path.basename(uploaded_zip.name))[0]
                            excel_out_name = f"{base_name}_SRL_ZA.xlsx"
                        else:
                            excel_out_name = excel_name

                        out_path = jsons_to_wsd_excel(
                            base_dir=tmpdir,
                            excel_name=excel_out_name,
                            include_memo_sheet=include_memo_sheet,
                            memo_placement=memo_placement,
                            memo_sep=memo_sep,
                        )
                        out_path_display = out_path  # 표시용
                        with open(out_path, "rb") as f:
                            excel_bytes = f.read()
                else:
                    # 폴더 직접 처리
                    out_path = jsons_to_wsd_excel(
                        base_dir=base_dir,
                        excel_name=excel_name,
                        include_memo_sheet=include_memo_sheet,
                        memo_placement=memo_placement,
                        memo_sep=memo_sep,
                    )
                    out_path_display = out_path
                    with open(out_path, "rb") as f:
                        excel_bytes = f.read()

                status.update(label="완료!", state="complete")
            except Exception as e:
                status.update(label="에러 발생", state="error")
                st.exception(e)
                return

        if excel_bytes:
            st.success(f"엑셀 파일 생성: {out_path_display}")
            st.download_button(
                label="⬇️ 엑셀 다운로드",
                data=excel_bytes,
                file_name=os.path.basename(out_path_display),
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )

            # 미리보기(상위 100행) — 메모리 바이트로 로드
            try:
                import pandas as pd
                xbio = io.BytesIO(excel_bytes)
                df_preview = pd.read_excel(xbio, sheet_name="SRL_ZA", nrows=100)
                st.subheader("미리보기 (SRL_ZA 시트 상위 100행)")
                st.dataframe(df_preview, use_container_width=True, height=400)
            except Exception:
                st.info("미리보기를 열 수 없습니다. 파일을 직접 확인해주세요.")
