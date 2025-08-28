#dataly_tools/photo_to_excel.py
"""
사진(이미지) JSON -> Excel 변환기
- 입력: dict(JSON 파싱 결과)
- 출력: bytes(XLSX), datalyManager에서 st.download_button으로 바로 다운로드
- 같은 document.id 묶음에서 [id, worker_id_cnst, Medium_category, metadata, mdfcn_memo] 세로 병합
- metadata: 멀티라인 텍스트 + 같은 id 첫 행에만 URL 하이퍼링크(파란색, 밑줄 없음)
"""

import json
import math
from typing import Any, Dict, Iterable, List, Tuple
from io import BytesIO

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill, Border, Side
from openpyxl.utils import get_column_letter

# 표시 순서(메타 키)
META_ORDER = [
    "note", "image", "copyright", "term_id", "Major_category",
    "title", "url", "Medium_category", "domain", "media_id",
    "publisher", "term", "source_id",
]

THIN_BORDER = Border(
    left=Side(style="thin"), right=Side(style="thin"),
    top=Side(style="thin"), bottom=Side(style="thin"),
)
HEADER_FILL = PatternFill(start_color="EEECE1", end_color="EEECE1", fill_type="solid")
LINK_BLUE = "0563C1"

# [타입] 문장 형태 파싱용 ([Type] 내용)
import re
TYPE_BRACKET_RE = re.compile(r"^\s*\[(.+?)\]\s*(.*)$")


def extract_sentences(doc: Dict[str, Any]) -> List[Tuple[str, str]]:
    """
    EX.exp_sentence 내부를 탐색해 ([Type] sentence) 또는 그냥 sentence를
    (type, sentence) 튜플 리스트로 반환
    """
    out: List[Tuple[str, str]] = []
    for ex in doc.get("EX", []):
        for item in ex.get("exp_sentence", []) or []:
            if not isinstance(item, dict):
                continue
            for _k, v in item.items():
                seq = v if isinstance(v, list) else [v]
                for s in seq:
                    if not s:
                        continue
                    text = str(s).strip()
                    m = TYPE_BRACKET_RE.match(text)
                    if m:
                        out.append((m.group(1).strip(), m.group(2).strip()))
                    else:
                        out.append(("", text))
    return out


def _clean_url(u: str) -> str:
    if not u:
        return ""
    u = str(u).strip()
    if (u.startswith('"') and u.endswith('"')) or (u.startswith("'") and u.endswith("'")):
        u = u[1:-1].strip()
    return u


def format_metadata_and_url(meta: Dict[str, Any]) -> Tuple[str, str]:
    """
    metadata를 멀티라인 문자열로 정리하고, url만 분리해서 반환
    """
    url_only = _clean_url(meta.get("url", ""))
    lines = ['metadata : {']
    for k in META_ORDER:
        v = meta.get(k, "")
        if k == "url":
            v = url_only or meta.get("url", "")
        lines.append(f'  "{k}": "{v}",' if k != META_ORDER[-1] else f'  "{k}": "{v}"')
    lines.append("}")
    return "\n".join(lines), url_only


def extract_mdfcn_memo(mdfcn_infos):
    """
    mdfcn_infos[*].mdfcn_memo 가 JSON 문자열이면 value만 추출해
    "작업 목록 1 : ..." 형태로 번호 매겨 반환.
    항목 간에 빈 줄 한 줄 추가.
    """
    if not mdfcn_infos:
        return ""
    out, idx = [], 1
    for info in mdfcn_infos:
        raw = info.get("mdfcn_memo", "")
        if not raw:
            continue
        try:
            arr = json.loads(raw)
            if isinstance(arr, list):
                for obj in arr:
                    val = str((obj or {}).get("value", "")).strip()
                    if val:
                        out.append(f"작업 목록 {idx} : {val}")
                        idx += 1
        except Exception:
            val = str(raw).strip()
            if val:
                out.append(f"작업 목록 {idx} : {val}")
                idx += 1
    return "\n\n".join(out)


def to_rows(data: Dict[str, Any]) -> List[Dict[str, Any]]:
    """
    JSON(dict) -> 행 리스트
    """
    rows: List[Dict[str, Any]] = []
    docs = data.get("document", [])
    if not isinstance(docs, list):
        return rows

    for doc in docs:
        img_id = str(doc.get("id", ""))
        meta = doc.get("metadata", {}) or {}
        medium_category = str(meta.get("Medium_category", "") or "")
        worker_id_cnst = str(doc.get("worker_id_cnst", "") or "")

        md_text, md_url = format_metadata_and_url(meta)
        memo_text = extract_mdfcn_memo(doc.get("mdfcn_infos", []) or [])

        pairs = extract_sentences(doc) or [("", "")]
        for typ, sent in pairs:
            rows.append({
                "id": img_id,
                "worker_id_cnst": worker_id_cnst,
                "Medium_category": medium_category,
                "유형": typ,
                "설명 문장": sent,
                "metadata": md_text,
                "meta_url": md_url,  # 하이퍼링크용(엑셀 열에는 포함 안 함)
                "mdfcn_memo(검수자 수정 이력)": memo_text,
            })
    return rows


def estimate_wrapped_lines(text: str, col_chars: int) -> int:
    if not text:
        return 1
    total = 0
    width = max(col_chars, 5)
    for para in str(text).split("\n"):
        total += max(1, math.ceil(len(para) / (width * 1.08)))
    return max(1, total)


def _write_excel_to_bytes(all_rows: List[Dict[str, Any]]) -> bytes:
    """
    행 리스트 -> Excel bytes
    """
    wb = Workbook()
    ws = wb.active
    ws.title = "result"

    headers = [
        "id", "worker_id_cnst", "Medium_category",
        "유형", "설명 문장", "metadata", "mdfcn_memo\n(검수자 수정 이력)"
    ]
    ws.append(headers)

    # 헤더 스타일
    for c in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = THIN_BORDER
        cell.fill = HEADER_FILL

    # 열 너비(문자폭 기준 추정)
    widths = {1: 12, 2: 16, 3: 14, 4: 16, 5: 80, 6: 60, 7: 50}
    for col_idx, w in widths.items():
        ws.column_dimensions[get_column_letter(col_idx)].width = w

    # 그룹 시작/개수 추적
    start_row_by_group: Dict[Tuple[str], int] = {}
    count_by_group: Dict[Tuple[str], int] = {}

    current_row = 2
    for row in all_rows:
        ws.append([
            row.get("id",""),
            row.get("worker_id_cnst",""),
            row.get("Medium_category",""),
            row.get("유형",""),
            row.get("설명 문장",""),
            row.get("metadata",""),
            row.get("mdfcn_memo(검수자 수정 이력)",""),
        ])
        for c in range(1, len(headers) + 1):
            ws.cell(row=current_row, column=c).alignment = Alignment(
                vertical="top", wrap_text=(c in (5, 6, 7))
            )
            ws.cell(row=current_row, column=c).border = THIN_BORDER

        key = (row.get("id",""),)
        if key not in start_row_by_group:
            start_row_by_group[key] = current_row
            count_by_group[key] = 0
        count_by_group[key] += 1
        current_row += 1

    # 병합: 같은 id 블록에서 [id, worker, Medium_category, metadata, mdfcn_memo] 병합
    merge_cols = [1, 2, 3, 6, 7]
    for key, start in start_row_by_group.items():
        cnt = count_by_group[key]
        if cnt > 1:
            end = start + cnt - 1
            for col in merge_cols:
                ws.merge_cells(start_row=start, start_column=col, end_row=end, end_column=col)
                ws.cell(row=start, column=col).alignment = Alignment(vertical="top", wrap_text=True)

    # metadata 하이퍼링크(같은 id 첫 행만)
    first_url_by_id: Dict[str, str] = {}
    for _r, row in enumerate(all_rows, start=2):
        rid = row.get("id", "")
        if rid and rid not in first_url_by_id:
            first_url_by_id[rid] = row.get("meta_url", "") or ""

    for key, start in start_row_by_group.items():
        doc_id = key[0]
        url = first_url_by_id.get(doc_id, "")
        if url and url.startswith(("http://", "https://")):
            c = ws.cell(row=start, column=6)
            c.hyperlink = url
            # 파란색, 밑줄 없음
            c.font = Font(color=LINK_BLUE, underline=None)
            c.alignment = Alignment(vertical="top", wrap_text=True)
            c.border = (THIN_BORDER)

    # 행 높이 대략 조정
    LINE_HEIGHT_PT = 18
    group_starts = set(start_row_by_group.values())
    for r in range(2, current_row):
        desc = ws.cell(row=r, column=5).value or ""
        desc_lines = estimate_wrapped_lines(desc, widths[5])
        if r in group_starts:
            meta_plain = ws.cell(row=r, column=6).value or ""
            memo_plain = ws.cell(row=r, column=7).value or ""
            need = max(
                desc_lines,
                estimate_wrapped_lines(meta_plain, widths[6]),
                estimate_wrapped_lines(memo_plain, widths[7]),
            )
        else:
            need = desc_lines
        ws.row_dimensions[r].height = max(1, need) * LINE_HEIGHT_PT

    # 틀 고정
    ws.freeze_panes = "A2"

    # 메모리로 저장
    bio = BytesIO()
    wb.save(bio)
    bio.seek(0)
    return bio.getvalue()


def photo_json_to_xlsx_bytes(data: Dict[str, Any]) -> bytes:
    """
    datalyManager에서 호출하는 공개 API
    """
    rows = to_rows(data)
    if not rows:
        # 빈 통합문서라도 반환(다운로드 버튼 활성 목적)
        return _write_excel_to_bytes([])
    return _write_excel_to_bytes(rows)
