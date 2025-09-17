# dataly_manager/ui/srl_argument_del_ui.py
# -*- coding: utf-8 -*-
from __future__ import annotations

"""
ZIP 업로드 → 임시폴더에 해제 → SRL 정리(write_back=True) →
- 업로드 ZIP의 폴더 구조를 그대로 보존하여 '적용된 JSON만' ZIP으로 단일 다운로드 제공
- 엑셀 파일은 생성/포함하지 않음
- 세션 키 안전 초기화로 KeyError 방지
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
    srl_argument_cleanup,   # 엑셀은 사용하지 않음
)


def _zip_jsons_keep_structure(dir_path: Path) -> bytes:
    """
    dir_path 아래의 모든 *.json을 원래 상대 경로(= dir_path 기준) 그대로 ZIP에 담아 반환.
    다른 파일 형식은 포함하지 않음.
    """
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", compression=zipfile.ZIP_DEFLATED) as zf:
        for p in dir_path.rglob("*.json"):
            if p.is_file():
                zf.write(p, arcname=str(p.relative_to(dir_path)))
    mem.seek(0)
    return mem.getvalue()


def render_srl_argument_del_ui():
    st.markdown("### 🧹 SRL 불필요 값 삭제")
    st.caption("규칙: argument.label이 비어 있고 해당 영역에 VX 형태소가 포함되면 argument 삭제, 모든 argument가 사라지면 SRL 항목 삭제합니다. 엑셀은 생성/포함하지 않습니다.")

    # ---------------- 세션 키 안전 초기화 ----------------
    st.session_state.setdefault("srl_json_zip_bytes", None)   # bytes
    st.session_state.setdefault("srl_json_zip_name", "srl_cleaned_json.zip")
    st.session_state.setdefault("srl_metrics", None)          # dict
    st.session_state.setdefault("srl_log_preview", None)      # str

    # 업로더 + 실행/초기화
    up = st.file_uploader("JSON 파일들이 들어있는 ZIP을 업로드하세요", type=["zip"], key="srl_zip_uploader")
    col_run, col_reset = st.columns([0.6, 0.4])
    run = col_run.button("실행", type="primary", use_container_width=True)
    reset = col_reset.button("초기화", use_container_width=True)

    if reset:
        st.session_state["srl_json_zip_bytes"] = None
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

            # 3) 적용된 JSON만 폴더 구조 그대로 ZIP으로 패키징
            cleaned_zip = _zip_jsons_keep_structure(tdir)

            # 4) 세션에 저장(재실행에도 다운로드 버튼 유지)
            st.session_state["srl_json_zip_bytes"] = cleaned_zip
            st.session_state["srl_json_zip_name"] = "srl_cleaned_json.zip"
            st.session_state["srl_metrics"] = {
                "total_files": result["total_files"],
                "changed_files": result["changed_files"],
                "skipped_files": result["skipped_files"],
            }
            rows = result.get("log_rows") or []
            head = rows[:51]  # header + 50
            preview = "\n".join([",".join(map(str, r)) for r in head]) if head else "(로그 없음)"
            st.session_state["srl_log_preview"] = preview

            st.success("처리가 완료되었습니다. 아래에서 ZIP을 다운로드하세요.")

    # ---------------- 결과 표시(세션 기반, 항상 렌더) ----------------
    zip_bytes = st.session_state.get("srl_json_zip_bytes")
    if zip_bytes is not None:
        st.download_button(
            label="정리된 JSON ZIP 다운로드",
            data=zip_bytes,
            file_name=st.session_state.get("srl_json_zip_name", "srl_cleaned_json.zip"),
            mime="application/zip",
            use_container_width=True,
        )

        m = st.session_state.get("srl_metrics") or {}
        col1, col2, col3 = st.columns(3)
        col1.metric("총 파일", m.get("total_files", 0))
        col2.metric("변경된 파일", m.get("changed_files", 0))
        col3.metric("변경 없음/스킵", m.get("skipped_files", 0))

        with st.expander("로그 미리보기 (상위 50행)"):
            st.code(st.session_state.get("srl_log_preview") or "(로그 없음)", language="text")
