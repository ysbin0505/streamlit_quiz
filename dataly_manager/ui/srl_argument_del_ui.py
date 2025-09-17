# dataly_manager/ui/srl_argument_del_ui.py
# -*- coding: utf-8 -*-
from __future__ import annotations

"""
ZIP 업로드 → 임시폴더 해제 → SRL 정리(write_back=True) →
결과 엑셀까지 포함한 통합 ZIP( clean JSON + report.xlsx )을 '단일' 다운로드로 제공.
재실행에도 버튼이 없어지지 않도록 st.session_state로 바이트를 보존.
"""

import io
import os
import zipfile
import tempfile
from pathlib import Path
import streamlit as st

# 패키지 루트(= dataly_manager의 부모 폴더)를 sys.path에 추가
import sys
APP_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(APP_DIR)
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from dataly_manager.dataly_tools.srl_argument_del import (
    srl_argument_cleanup,
    make_excel_report,
)

# ---------------- Session State 초기화 ----------------
if "srl_bundle_zip" not in st.session_state:
    st.session_state["srl_bundle_zip"] = None         # bytes
    st.session_state["srl_bundle_name"] = "srl_cleaned_json_and_report.zip"
    st.session_state["srl_metrics"] = None            # dict
    st.session_state["srl_log_preview"] = None        # str


def _zip_jsons_and_excel(dir_path: Path, excel_bytes: bytes, excel_name: str = "srl_cleanup_result.xlsx") -> bytes:
    """
    dir_path 아래의 모든 *.json 파일과 엑셀 바이트를 하나의 ZIP으로 묶어 반환.
    ZIP 루트:
      - cleaned_jsons/...(원래 폴더 구조 유지)
      - srl_cleanup_result.xlsx
    """
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        # 엑셀 리포트 파일
        zf.writestr(excel_name, excel_bytes)
        # JSON들
        for p in dir_path.rglob("*.json"):
            if p.is_file():
                arc = Path("cleaned_jsons") / p.relative_to(dir_path)
                zf.write(p, arcname=str(arc))
    mem.seek(0)
    return mem.getvalue()


def render_srl_argument_del_ui():
    st.markdown("### 🧹 SRL 인자 정리 (ZIP 업로드 → 통합 ZIP: JSON + Excel)")
    st.caption("규칙: argument.label이 비어 있고 해당 영역에 VX 형태소가 포함되면 argument 삭제, 모든 argument가 사라지면 SRL 항목 삭제합니다.")

    # 업로더와 실행/초기화 UI
    up = st.file_uploader("JSON 파일들이 들어있는 ZIP을 업로드하세요", type=["zip"], key="srl_zip_uploader")
    col_run, col_reset = st.columns([0.6, 0.4])
    run = col_run.button("실행", type="primary", use_container_width=True)
    reset = col_reset.button("초기화", use_container_width=True)

    if reset:
        st.session_state["srl_bundle_zip"] = None
        st.session_state["srl_metrics"] = None
        st.session_state["srl_log_preview"] = None
        st.success("상태를 초기화했습니다.")

    if run:
        if not up:
            st.error("ZIP 파일을 업로드해 주세요.")
            st.stop()

        with tempfile.TemporaryDirectory() as td:
            tdir = Path(td)

            # 1) ZIP 해제
            try:
                with zipfile.ZipFile(up) as zf:
                    zf.extractall(tdir)
            except Exception as e:
                st.error(f"ZIP 해제 실패: {e}")
                st.stop()

            # 2) 정리 수행 (임시폴더에 바로 적용)
            prog = st.progress(0, text="처리 시작…")

            def _cb(cur, total, path):
                denom = max(total, 1)
                prog.progress(min(cur / denom, 1.0), text=f"[{cur}/{total}] {path.name} 처리 중")

            result = srl_argument_cleanup(in_path=tdir, write_back=True, progress_cb=_cb)
            prog.progress(1.0, text="완료")

            # 3) 결과 엑셀 생성
            xlsx_bytes = make_excel_report(result)

            # 4) 통합 ZIP(정리된 JSON + 결과 엑셀) 생성
            bundle_zip = _zip_jsons_and_excel(tdir, xlsx_bytes, excel_name="srl_cleanup_result.xlsx")

            # 5) 세션에 저장(재실행에도 유지)
            st.session_state["srl_bundle_zip"] = bundle_zip
            st.session_state["srl_bundle_name"] = "srl_cleaned_json_and_report.zip"
            st.session_state["srl_metrics"] = {
                "total_files": result["total_files"],
                "changed_files": result["changed_files"],
                "skipped_files": result["skipped_files"],
            }
            # 로그 미리보기 문자열 저장
            rows = result.get("log_rows") or []
            head = rows[:51]  # header + 50
            preview = "\n".join([",".join(map(str, r)) for r in head]) if head else "(로그 없음)"
            st.session_state["srl_log_preview"] = preview

            st.success("처리가 완료되었습니다. 아래에서 통합 ZIP을 다운로드하세요.")

    # ---------------- 결과 표시(세션 기반, 항상 렌더) ----------------
    if st.session_state["srl_bundle_zip"] is not None:
        st.download_button(
            label="통합 ZIP 다운로드 (정리된 JSON + 결과 엑셀)",
            data=st.session_state["srl_bundle_zip"],
            file_name=st.session_state["srl_bundle_name"],
            mime="application/zip",
            use_container_width=True,
        )

        # 메트릭
        m = st.session_state["srl_metrics"] or {}
        col1, col2, col3 = st.columns(3)
        col1.metric("총 파일", m.get("total_files", 0))
        col2.metric("변경된 파일", m.get("changed_files", 0))
        col3.metric("변경 없음/스킵", m.get("skipped_files", 0))

        with st.expander("로그 미리보기 (상위 50행)"):
            st.code(st.session_state["srl_log_preview"] or "(로그 없음)", language="text")
