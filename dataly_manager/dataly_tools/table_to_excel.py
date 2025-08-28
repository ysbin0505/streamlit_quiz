#dataly_tools/table_to_excel.py

# -*- coding: utf-8 -*-
"""
표(JSON) -> Excel 변환기 (bytes 반환)
- 입력: dict(JSON 파싱 결과)
- 출력: bytes(XLSX)
- 같은 id 블록 기준으로 [A:id, B:worker_id_cnst, E:metadata, F:mdfcn_infos] 병합
- metadata 첫 행에만 URL 하이퍼링크
"""
import json
from collections import defaultdict
from io import BytesIO
from typing import Dict, Any, List

import pandas as pd
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side


# ===== 유형 매핑 =====
REF_MAP = {
    "table_ref": "표 설명 문장",
    "row_ref": "행 설명 문장",
    "col_ref": "열 설명 문장",
    "cell_ref": "불연속 영역 설명 문장",
}
TYPE_TAGS = {"table_ref", "row_ref", "col_ref", "cell_ref"}

# --- 추가: 설명문장 안전 추출 헬퍼들 ---
def _iter_exp_items(ex_obj):
    """ex_obj.get('exp_sentence')가 list/dict/str 어떤 형태든 리스트로 정규화"""
    raw = ex_obj.get("exp_sentence", [])
    if isinstance(raw, list):
        return raw
    if isinstance(raw, dict):
        return [raw]
    if isinstance(raw, str):
        return [{"설명문장": [raw]}]
    return []

def _pick_sentence(exp_item) -> str:
    """
    exp_item에서 설명문장 1개를 뽑아 반환.
    - '설명문장' / '설명 문장' / 비슷한 변형(공백 제거 후 비교) 우선
    - 그래도 없으면 dict 값들 중 list[str] 또는 str을 첫 번째로 사용
    """
    if isinstance(exp_item, str):
        return exp_item.strip()

    if isinstance(exp_item, dict):
        # 1) 우선 '설명문장' / '설명 문장' 같이 보이는 키를 탐색(공백 제거 후 비교)
        for k, v in exp_item.items():
            kn = str(k).replace(" ", "")
            if kn in ("설명문장", "설명문장들", "설명"):
                if isinstance(v, list) and v:
                    return str(v[0]).strip()
                return str(v).strip() if v is not None else ""

        # 2) fallback: 값들 중 list[str] 또는 str을 사용
        for v in exp_item.values():
            if isinstance(v, list) and v and isinstance(v[0], str):
                return v[0].strip()
            if isinstance(v, str):
                return v.strip()

    return ""  # 못 찾으면 빈 문자열


def extract_mdfcn_values(obj, sep: str = "\n") -> str:
    """mdfcn_infos에서 value만 추출(중복 제거, 순서 유지) 후 sep로 결합"""
    values: List[str] = []

    def _dedup_keep_order(items: List[str]) -> List[str]:
        seen, out = set(), []
        for it in items:
            if it and it not in seen:
                seen.add(it)
                out.append(it)
        return out

    def _walk(x):
        if x is None:
            return
        if isinstance(x, str):
            s = x.strip()
            if not s:
                return
            if s[:1] in ("[", "{"):
                try:
                    _walk(json.loads(s))
                    return
                except Exception:
                    if s not in TYPE_TAGS:
                        values.append(s)
                    return
            if s not in TYPE_TAGS:
                values.append(s)
            return
        if isinstance(x, dict):
            v = x.get("value")
            if isinstance(v, str):
                v = v.strip()
                if v:
                    values.append(v)
            mm = x.get("mdfcn_memo")
            if isinstance(mm, str):
                mm_s = mm.strip()
                if mm_s:
                    try:
                        _walk(json.loads(mm_s))
                    except Exception:
                        pass
            for k, sub in x.items():
                if k in ("value", "mdfcn_memo"):
                    continue
                if isinstance(sub, (list, dict)):
                    _walk(sub)
            return
        if isinstance(x, (list, tuple)):
            for it in x:
                _walk(it)
            return

    _walk(obj)
    return sep.join(_dedup_keep_order([v for v in values if isinstance(v, str)]))


def extract_url(meta: Any) -> str:
    """metadata.url을 안전하게 추출"""
    if isinstance(meta, dict):
        u = meta.get("url", "")
        if isinstance(u, list):
            return u[0] if u else ""
        return str(u or "")
    return ""


def table_json_to_xlsx_bytes(data: Dict[str, Any]) -> bytes:
    """datalyManager에서 호출하는 공개 API"""
    rows: List[Dict[str, Any]] = []
    group_counts = defaultdict(int)

    for doc in data.get("document", []) or []:
        doc_id = doc.get("id", "")
        worker = str(doc.get("worker_id_cnst") or "").strip()
        metadata = doc.get("metadata", {}) or {}

        # mdfcn_infos 원본 수집(키 변동 대비)
        mdfcn_raw = doc.get("mdfcn_infos", doc.get("mdfcn_memo", []))
        mdfcn_text = extract_mdfcn_values(mdfcn_raw, sep="\n")

        for ex in doc.get("EX", []) or []:
            ref_type = ex.get("reference", {}).get("reference_type", "")
            for exp_item in _iter_exp_items(ex):
                sentence = _pick_sentence(exp_item)  # ← 키 변형 안전 처리
                rows.append({
                    "id": doc_id,
                    "worker_id_cnst": worker,
                    "유형": REF_MAP.get(ref_type, ref_type),
                    "설명 문장": sentence,
                    "metadata": json.dumps(metadata, ensure_ascii=False, indent=2),
                    "mdfcn_infos": mdfcn_text,
                    "url": extract_url(metadata),
                })
                group_counts[doc_id] += 1

    # 빈 데이터여도 헤더만 있는 파일 생성
    df = pd.DataFrame(
        rows,
        columns=["id", "worker_id_cnst", "유형", "설명 문장", "metadata", "mdfcn_infos"]
    ).fillna("")

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="sheet1")
        ws = writer.sheets["sheet1"]

        # 열 너비
        widths = {"A": 18, "B": 16, "C": 13, "D": 80, "E": 50, "F": 50}
        for col, w in widths.items():
            ws.column_dimensions[col].width = w

        # 머리글 스타일
        header_fill = PatternFill("solid", fgColor="D9E1F2")
        header_font = Font(bold=True)
        for cell in ws[1]:
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        # metadata/검수이력 헤더 강조
        ws["E1"].fill = PatternFill("solid", fgColor="BDD7EE")
        ws["F1"].fill = PatternFill("solid", fgColor="FCE4D6")

        # 데이터 영역: 줄바꿈 + 상단 정렬 + 테두리
        thin = Side(style="thin", color="999999")
        border = Border(left=thin, right=thin, top=thin, bottom=thin)

        max_row = ws.max_row
        max_col = ws.max_column

        for r in range(2, max_row + 1):
            for c in range(1, max_col + 1):
                cell = ws.cell(row=r, column=c)
                cell.alignment = Alignment(wrap_text=True, vertical="top") if c >= 4 else Alignment(vertical="top")
                cell.border = border

        # id 블록 병합: A(id), B(worker), E(metadata), F(mdfcn_infos)
        cur_row = 2
        for doc_id, count in group_counts.items():
            if count > 1:
                for col in (1, 2, 5, 6):
                    ws.merge_cells(
                        start_row=cur_row, start_column=col,
                        end_row=cur_row + count - 1, end_column=col
                    )
                    top_cell = ws.cell(row=cur_row, column=col)
                    top_cell.alignment = Alignment(vertical="top", wrap_text=True)
                    for rr in range(cur_row, cur_row + count):
                        ws.cell(row=rr, column=col).border = border
            cur_row += count

        # 하이퍼링크: 같은 id의 첫 행만 E열에 설정
        first_row_for_id: Dict[str, int] = {}
        for idx, row in enumerate(rows, start=2):
            first_row_for_id.setdefault(row["id"], idx)

        for idx, row in enumerate(rows, start=2):
            url = row.get("url", "")
            if not url:
                continue
            if idx != first_row_for_id[row["id"]]:
                continue
            ws.cell(row=idx, column=5).hyperlink = url  # E열

        # 행 높이: D/E/F의 개행 수 기준 근사 조절
        for r in range(2, max_row + 1):
            max_lines = 1
            for c in (4, 5, 6):  # D,E,F
                val = ws.cell(row=r, column=c).value
                if isinstance(val, str) and "\n" in val:
                    lines = val.count("\n") + 1
                    if lines > max_lines:
                        max_lines = lines
            ws.row_dimensions[r].height = min(15 + (max_lines - 1) * 12, 200)

        # 틀 고정
        ws.freeze_panes = "A2"

    output.seek(0)
    return output.getvalue()
