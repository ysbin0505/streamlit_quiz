# -*- coding: utf-8 -*-
from __future__ import annotations

import os
from pathlib import Path
import streamlit as st

# 패키지 루트(= dataly_manager의 부모 폴더)를 sys.path에 추가 (메인과 동일 전략)
import sys
APP_DIR = os.path.dirname(os.path.abspath(__file__))
ROOT_DIR = os.path.dirname(APP_DIR)
if ROOT_DIR not in sys.path:
    sys.path.insert(0, ROOT_DIR)

from dataly_manager.dataly_tools.srl_argument_del import srl_argument_cleanup


def render_srl_argument_del_ui():
    st.markdown("### 🧹 SRL 인자 정리 (빈 label + VX 포함 제거)")
    st.caption("조건: argument.label이 비어 있고, 해당 argument가 커버하는 단어들 중 morph.label == 'VX'가 하나라도 있으면 해당 argument를 삭제합니다. argument가 모두 사라지면 SRL 항목 자체를 삭제합니다.")

    with st.container(border=True):
        col1, col2 = st.columns([0.55, 0.45])
        with col1:
            in_path = st.text_input(
                "입력 경로 (파일 또는 폴더)",
                value="",
                placeholder="/Users/you/data or /Users/you/file.json"
            )
            use_outdir = st.checkbox("별도 출력 디렉터리에 저장", value=False)
            out_dir = st.text_input(
                "출력 디렉터리 (선택)",
                value="",
                placeholder="/Users/you/output",
                disabled=not use_outdir
            )
        with col2:
            make_csv = st.checkbox("보고용 CSV 로그 생성", value=True)
            report_csv = st.text_input(
                "CSV 경로 (선택)",
                value="srl_cleanup_VX_log.csv",
                disabled=not make_csv
            )

        run = st.button("실행", type="primary", use_container_width=True)

    if run:
        if not in_path.strip():
            st.error("입력 경로를 입력해 주세요.")
            st.stop()

        p_in = Path(in_path.strip())
        p_out = Path(out_dir.strip()) if (use_outdir and out_dir.strip()) else None
        p_csv = Path(report_csv.strip()) if (make_csv and report_csv.strip()) else None

        prog = st.progress(0, text="처리 시작…")
        last_total = 1

        def _cb(cur: int, total: int, path: Path):
            nonlocal last_total
            last_total = total
            prog.progress(min(cur / max(total, 1), 1.0), text=f"[{cur}/{total}] 처리 중: {path.name}")

        try:
            result = srl_argument_cleanup(
                in_path=p_in,
                out_dir=p_out,
                report_csv=p_csv,
                progress_cb=_cb
            )
        except Exception as e:
            prog.empty()
            st.error(f"에러: {e}")
            st.stop()

        prog.progress(1.0, text="완료")
        st.success("SRL 인자 정리가 완료되었습니다.")

        colA, colB, colC = st.columns(3)
        colA.metric("총 파일", result["total_files"])
        colB.metric("변경된 파일", result["changed_files"])
        colC.metric("변경 없음/스킵", result["skipped_files"])

        if result.get("report_csv"):
            st.info(f"로그 CSV: {result['report_csv']}")

        with st.expander("변경된 파일 목록 보기"):
            if result["outputs"]:
                for item in result["outputs"]:
                    src = Path(item["src"])
                    dst = Path(item["dst"])
                    st.write(f"- {src.name} → {dst}")
            else:
                st.write("변경된 파일이 없습니다.")

        with st.expander("세부 로그 미리보기 (상위 50행)"):
            rows = result["log_rows"][:51]  # header + 50
            preview = "\n".join([",".join(map(str, r)) for r in rows])
            st.code(preview, language="text")
