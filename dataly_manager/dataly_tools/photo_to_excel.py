# -*- coding: utf-8 -*-
"""
사진(이미지) JSON -> Excel 변환기 (+ 엑셀 수정본을 JSON에 역반영)
- 입력: dict(JSON 파싱 결과)
- 출력: bytes(XLSX), datalyManager에서 st.download_button으로 바로 다운로드
- 같은 document.id 묶음에서 [id, worker_id_cnst, Medium_category, metadata, mdfcn_memo] 세로 병합
- metadata: 멀티라인 텍스트 + 같은 id 첫 행에만 URL 하이퍼링크(파란색, 밑줄 없음)

추가:
- apply_excel_desc_to_json_from_zip(zip_bytes, sheet_name=None, skip_blank=True)
  : ZIP(엑셀+단일 JSON)을 받아 엑셀의 '설명 문장'을 JSON에 반영해 반환
"""

import json
import math
import re
import zipfile
from io import BytesIO
from pathlib import Path
from typing import Any, Dict, Iterable, List, Tuple, Optional
from collections import defaultdict

import pandas as pd
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
TYPE_BRACKET_RE = re.compile(r"^\s*\[(.+?)\]\s*(.*)$")


# =========================
# JSON -> Excel (정방향)
# =========================
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


# ==========================================
# Excel('설명 문장') → JSON (역방향, ZIP 지원)
# ==========================================

def _delete_slot(slot_descriptor):
    """
    ('list', list_obj, idx)  -> list_obj.pop(idx)
    ('dict_scalar', dict_obj, key) -> dict_obj.pop(key, None)
    """
    mode = slot_descriptor[0]
    if mode == "list":
        lst, idx = slot_descriptor[1], slot_descriptor[2]
        if 0 <= idx < len(lst):
            lst.pop(idx)
    elif mode == "dict_scalar":
        obj, key = slot_descriptor[1], slot_descriptor[2]
        if isinstance(obj, dict):
            obj.pop(key, None)


def _cleanup_exp_sentences(doc: Dict[str, Any]) -> None:
    """
    빈 문자열/빈 리스트/빈 딕셔너리를 걷어내서 exp_sentence 구조를 가볍게 정리.
    (키 자체도 제거)
    """
    ex_list = doc.get("EX", [])
    if not isinstance(ex_list, list):
        return

    for ex in ex_list:
        exp = ex.get("exp_sentence")
        if exp is None:
            continue

        # list 형태
        if isinstance(exp, list):
            new_exp = []
            for item in exp:
                if isinstance(item, dict):
                    new_item = {}
                    for k, v in item.items():
                        if isinstance(v, list):
                            vv = [str(s).strip() for s in v if str(s or "").strip()]
                            if vv:
                                new_item[k] = vv
                        else:
                            s = str(v or "").strip()
                            if s:
                                new_item[k] = s
                    if new_item:
                        new_exp.append(new_item)
                # dict가 아니면 버림
            ex["exp_sentence"] = new_exp
            if not new_exp:
                # 완전히 비었으면 키 제거
                ex.pop("exp_sentence", None)

        # dict 형태
        elif isinstance(exp, dict):
            new_exp = {}
            for k, v in exp.items():
                if isinstance(v, list):
                    vv = [str(s).strip() for s in v if str(s or "").strip()]
                    if vv:
                        new_exp[k] = vv
                else:
                    s = str(v or "").strip()
                    if s:
                        new_exp[k] = s
            if new_exp:
                ex["exp_sentence"] = new_exp
            else:
                ex.pop("exp_sentence", None)


def _compose_text_with_type(old_text: str, new_sentence: str, excel_type: str) -> str:
    """
    엑셀 '유형'만 바꿔도 문장은 유지하면서 타입을 교체.
    - new_sentence가 비어 있어도 excel_type이 있으면 타입만 바꾼다.
    - excel_type에 대괄호가 이미 있으면 이중 대괄호를 방지한다.
    """
    s_new = "" if new_sentence is None else str(new_sentence).strip()
    t_new = "" if excel_type is None else str(excel_type).strip()

    # [타입] 파싱
    old_text_str = str(old_text or "")
    m = TYPE_BRACKET_RE.match(old_text_str)
    old_type = (m.group(1).strip() if m else "")
    old_body = (m.group(2).strip() if m else old_text_str.strip())

    # 엑셀 유형에 대괄호가 들어온 경우 이중 괄호 방지
    if t_new.startswith("[") and t_new.endswith("]"):
        t_new = t_new[1:-1].strip()

    # 본문은 새 문장 있으면 교체, 없으면 기존 유지
    body = s_new if s_new else old_body

    # 타입 우선순위: 엑셀 유형 > 기존 유형 > 없음
    final_type = t_new if t_new else old_type

    return f"[{final_type}] {body}".strip() if final_type else body


def _iter_sentence_slots_with_old(doc: Dict[str, Any]):
    """
    사진 JSON의 EX[*].exp_sentence에서 실제 '문장 슬롯'을 순서대로 찾아
    (slot_descriptor, old_text) 를 yield.
    slot_descriptor는 ('list', list_obj, idx) 또는 ('dict_scalar', dict_obj, key)
    같은 형태로, _assign_text_to_slot에서 사용.
    """
    ex_list = doc.get("EX", [])
    if not isinstance(ex_list, list):
        return

    for ex in ex_list:
        exp = ex.get("exp_sentence")
        if exp is None:
            continue

        # 최빈 구조: list[ dict{key: list[str] or str}, ... ]
        if isinstance(exp, list):
            for item in exp:
                if isinstance(item, dict):
                    for k, v in item.items():
                        if isinstance(v, list):
                            for i, s in enumerate(v):
                                yield (("list", v, i), ("" if s is None else str(s)))
                        else:
                            yield (("dict_scalar", item, k), ("" if v is None else str(v)))
                else:
                    # list 안에 문자열이 직접 들어오는 경우도 방어적 무시(드묾)
                    continue

        elif isinstance(exp, dict):
            for k, v in exp.items():
                if isinstance(v, list):
                    for i, s in enumerate(v):
                        yield (("list", v, i), ("" if s is None else str(s)))
                else:
                    yield (("dict_scalar", exp, k), ("" if v is None else str(v)))

        elif isinstance(exp, str):
            # 문자열 하나 전체가 슬롯인 경우
            yield (("dict_scalar", ex, "exp_sentence"), exp)


def _assign_text_to_slot(slot_descriptor, new_text: str):
    mode = slot_descriptor[0]
    if mode == "list":
        lst, idx = slot_descriptor[1], slot_descriptor[2]
        lst[idx] = new_text
    elif mode == "dict_scalar":
        obj, key = slot_descriptor[1], slot_descriptor[2]
        obj[key] = new_text


def _collect_excel_pairs_by_id(df: pd.DataFrame, skip_blank: bool = True) -> Dict[str, List[Tuple[str, str]]]:
    """
    엑셀에서 id별 (유형, 설명 문장) 시퀀스를 원본 행 순서대로 수집.
    - 병합 셀로 인해 비는 id/유형은 ffill로 채움
    - skip_blank=True면 빈 '설명 문장'은 건너뜀
    반환: { id: [(type, sentence), ...], ... }
    """
    required = {"id", "설명 문장"}
    if not required.issubset(set(df.columns)):
        raise ValueError("엑셀에 'id', '설명 문장' 컬럼이 필요합니다.")

    tmp = df.copy()
    # 병합 셀 보정
    if "id" in tmp.columns:
        tmp["id"] = tmp["id"].ffill()
    if "유형" in tmp.columns:
        tmp["유형"] = tmp["유형"].ffill()

    tmp["id"] = tmp["id"].astype(str)
    if "유형" not in tmp.columns:
        tmp["유형"] = ""  # 유형 컬럼이 없어도 동작하게
    tmp["유형"] = tmp["유형"].fillna("").astype(str)
    tmp["설명 문장"] = tmp["설명 문장"].fillna("").astype(str)

    bucket: Dict[str, List[Tuple[str, str]]] = defaultdict(list)
    for _, row in tmp.iterrows():
        _id = row["id"].strip()
        typ = row["유형"].strip()
        sent = row["설명 문장"].strip()
        if skip_blank and not sent:
            continue
        bucket[_id].append((typ, sent))
    return bucket


def apply_excel_desc_to_photo_json(
    json_obj: Dict[str, Any],
    excel_df: pd.DataFrame,
    skip_blank: bool = False
) -> Dict[str, Any]:
    """
    사진 JSON에 엑셀의 '설명 문장'을 반영.
    - 같은 id 내에서 '엑셀 행 순서'와 '기존 JSON 슬롯 순서'를 1:1로 맞춰 반영
    - 규칙:
      1) 유형·문장 모두 빈값("")이면 해당 슬롯을 '삭제'
      2) 유형만 있고 문장 빈값이면 '유형만 교체, 본문 유지'
      3) 문장만 있고 유형 빈값이면 '본문만 교체, 기존 유형 유지'
      4) 엑셀 행 수 < JSON 슬롯 수면, 남은 슬롯은 뒤에서부터 삭제하여
         최종 슬롯 개수를 엑셀과 '정확히 동일'하게 맞춤
    """
    mapping = _collect_excel_pairs_by_id(excel_df, skip_blank=skip_blank)

    docs = json_obj.get("document", [])
    if not isinstance(docs, list):
        return json_obj

    for doc in docs:
        doc_id = str(doc.get("id", ""))
        seq = mapping.get(doc_id, [])
        # 현재 문서의 슬롯 스냅샷(삭제 안전 처리를 위해 list로 고정)
        slots = list(_iter_sentence_slots_with_old(doc))

        # 엑셀-JSON 매칭 길이
        n = min(len(seq), len(slots))
        delete_slot_indices = []

        # 1) 앞에서부터 n개 매칭
        for i in range(n):
            (slot_desc, old_text) = slots[i]
            typ, new_sent = seq[i]
            typ_clean = (typ or "").strip()
            sent_clean = (new_sent or "").strip()

            # 둘 다 비면 '삭제' 지시로 간주
            if typ_clean == "" and sent_clean == "":
                delete_slot_indices.append(i)
                continue

            # 변경(문장/유형 교체 규칙)
            composed = _compose_text_with_type(old_text, sent_clean, typ_clean)
            _assign_text_to_slot(slot_desc, composed)

        # 2) 엑셀 길이보다 JSON 슬롯이 더 길면, 남은 슬롯은 모두 삭제
        if len(slots) > len(seq):
            delete_slot_indices.extend(range(n, len(slots)))

        # 3) 삭제 실제 수행 (뒤에서부터 지워서 인덱스 안정성 확보)
        for idx in sorted(delete_slot_indices, reverse=True):
            _delete_slot(slots[idx][0])

        # 4) 비어버린 컨테이너/키 정리
        _cleanup_exp_sentences(doc)

    return json_obj



def apply_excel_desc_to_json_from_zip(
    zip_bytes: bytes,
    sheet_name: Optional[str] = None,
    skip_blank: bool = True,
) -> Tuple[bytes, str]:
    """
    (사진 전용) ZIP(엑셀+단일 JSON)을 받아 엑셀의 '설명 문장'을 JSON에 반영.
    반환: (updated_json_bytes, suggested_filename)
    """
    if not isinstance(zip_bytes, (bytes, bytearray)):
        raise TypeError("zip_bytes는 bytes/bytearray여야 합니다.")

    with zipfile.ZipFile(BytesIO(zip_bytes), "r") as zf:
        # 구성 파일 선택
        json_members = [m for m in zf.namelist() if m.lower().endswith(".json")]
        xlsx_members = [m for m in zf.namelist() if m.lower().endswith(".xlsx")]
        xls_members  = [m for m in zf.namelist() if m.lower().endswith(".xls")]

        json_member = None
        for m in json_members:
            if Path(m).name.startswith("project_"):
                json_member = m
                break
        if json_member is None and json_members:
            json_member = json_members[0]

        excel_member = xlsx_members[0] if xlsx_members else (xls_members[0] if xls_members else None)

        if not json_member:
            raise FileNotFoundError("ZIP 안에 JSON 파일이 없습니다.")
        if not excel_member:
            raise FileNotFoundError("ZIP 안에 Excel 파일(.xlsx/.xls)이 없습니다.")

        # JSON 로드
        with zf.open(json_member) as jf:
            json_obj = json.loads(jf.read().decode("utf-8"))

        # Excel 로드
        with zf.open(excel_member) as ef:
            df = pd.read_excel(ef, sheet_name=sheet_name) if sheet_name else pd.read_excel(ef)

        # 반영
        updated = apply_excel_desc_to_photo_json(json_obj, df, skip_blank=skip_blank)

        # 파일명 제안
        base = Path(json_member).name
        out_name = (base[:-5] if base.lower().endswith(".json") else base) + "_updated.json"

        text = json.dumps(updated, ensure_ascii=False, indent=2)
        return text.encode("utf-8"), out_name
