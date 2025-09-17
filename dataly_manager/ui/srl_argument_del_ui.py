# -*- coding: utf-8 -*-
from __future__ import annotations

"""
ZIP 업로드 → 임시폴더에 해제 → SRL 정리(write_back=True) → 
1) 정리된 JSON ZIP 다운로드
2) 요약 엑셀(xlsx) 다운로드
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


def _zip_jsons(dir_path: Path) -> bytes:
    """
    dir_path 아래의 모든 *.json 파일만 보존 경로로 ZIP으로 묶어 메모리 바이트로 반환.
    """
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for p in dir_path.rglob("*.json"):
            if p.is_file():
                zf.write(p, arcname=str(p.relative_to(dir_path)))
    mem.seek(0)
    return mem.getvalue()


def render_srl_argument_del_ui():
    st.markdown("### 🧹 SRL 인자 정리 (ZIP 업로드 → JSON ZIP + Excel)")
    st.caption("규칙: argument.label이 비어 있고 해당 영역에 VX 형태소가 포함되면 해당 argument를 삭제, 모든 argument가 사라지면 SRL 항목을 삭제합니다.")

    up = st.file_uploader("JSON 파일들이 들어있는 ZIP을 업로드하세요", type=["zip"])
    run = st.button("실행", type="primary", use_container_width=True)

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

            # 2) 정리 수행 (write_back=True → 해제된 JSON에 바로 적용)
            prog = st.progress(0, text="처리 시작…")

            def _cb(cur, total, path):
                denom = max(total, 1)
                prog.progress(min(cur / denom, 1.0), text=f"[{cur}/{total}] {path.name} 처리 중")

            result = srl_argument_cleanup(in_path=tdir, write_back=True, progress_cb=_cb)
            prog.progress(1.0, text="완료")

            # 3) 결과 메트릭
            c1, c2, c3 = st.columns(3)
            c1.metric("총 파일", result["total_files"])
            c2.metric("변경된 파일", result["changed_files"])
            c3.metric("변경 없음/스킵", result["skipped_files"])

            # 4) 정리된 JSON ZIP 다운로드
            cleaned_zip_bytes = _zip_jsons(tdir)
            st.download_button(
                label="정리된 JSON ZIP 다운로드 (srl_cleaned_json.zip)",
                data=cleaned_zip_bytes,
                file_name="srl_cleaned_json.zip",
                mime="application/zip",
                use_container_width=True,
            )

            # 5) 결과 엑셀 다운로드
            xlsx_bytes = make_excel_report(result)
            st.download_button(
                label="결과 엑셀 다운로드 (srl_cleanup_result.xlsx)",
                data=xlsx_bytes,
                file_name="srl_cleanup_result.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

            # 6) 로그 미리보기(상위 50행)
            with st.expander("로그 미리보기 (상위 50행)"):
                rows = result.get("log_rows") or []
                head = rows[:51]  # header + 50
                preview = "\n".join([",".join(map(str, r)) for r in head]) if head else "(로그 없음)"
                st.code(preview, language="text")
