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
from typing import Dict, Any, List, Optional, Tuple

import pandas as pd
from openpyxl.styles import Alignment, PatternFill, Font, Border, Side

import zipfile
from pathlib import Path
# ===== 유형 매핑 =====
REF_MAP = {
    "table_ref": "표 설명 문장",
    "row_ref": "행 설명 문장",
    "col_ref": "열 설명 문장",
    "cell_ref": "불연속 영역 설명 문장",
}
TYPE_TAGS = {"table_ref", "row_ref", "col_ref", "cell_ref"}

# 역매핑: 엑셀 '유형' → JSON reference_type
REF_MAP_INV = {v: k for k, v in REF_MAP.items()}


def _set_exp_sentence_on_dict(d: Dict[str, Any], new_sentence: str, prefer_existing: bool = True) -> None:
    """
    d(dict) 내부의 '설명문장' 계열 키(공백/변형 포함)를 모두 정리하고 하나의 키로만 기록.
    - prefer_existing=True: 기존에 쓰던 키명이 있으면 그 키를 유지하여 overwrite
    - 기존 키가 없다면 기본 키 '설명문장'으로 기록
    - 기록 형태는 호환성을 위해 list[str] 유지
    """
    if not isinstance(d, dict):
        return

    # 후보 키 수집(공백 제거 후 비교)
    candidates = []
    for k in list(d.keys()):
        kn = str(k).replace(" ", "")
        if kn in ("설명문장", "설명문장들", "설명문", "설명"):
            candidates.append(k)

    # 사용할 타깃 키 결정
    target_key = candidates[0] if (prefer_existing and candidates) else "설명문장"

    # 중복 방지: 기존 후보 키 제거
    for k in candidates:
        try:
            del d[k]
        except Exception:
            pass

    d[target_key] = ["" if new_sentence is None else str(new_sentence)]

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



def _collect_excel_sentences_by_id(df: pd.DataFrame) -> Dict[str, List[str]]:
    if "id" not in df.columns or "설명 문장" not in df.columns:
        raise ValueError("엑셀에 'id'와 '설명 문장' 컬럼이 필요합니다.")

    tmp = df.copy()

    # ★ 병합 셀 복원
    if "id" in tmp.columns:
        tmp["id"] = tmp["id"].ffill()

    # 문자열화 + NaN 처리
    tmp["id"] = tmp["id"].astype(str)
    tmp["설명 문장"] = tmp["설명 문장"].fillna("").astype(str)

    bucket: Dict[str, List[str]] = defaultdict(list)
    for _, row in tmp.iterrows():
        _id = row["id"]
        sent = row["설명 문장"].strip()
        bucket[_id].append(sent)
    return bucket



def _iter_exp_slots(ex_obj):
    """
    ex_obj 내부의 exp_sentence '슬롯'을 순서대로 반환.
    각 슬롯은 ('list'|'dict'|'str', container, index_or_key) 형태로 기술.
    - list: (mode='list', list_obj, idx)
    - dict: (mode='dict', dict_obj, None) → dict['설명문장']에 기록
    - str : (mode='str',  ex_obj(부모), 'exp_sentence') → ex_obj['exp_sentence'] 문자열에 기록
    기타 타입/누락 시, 기본 슬롯을 생성해서 반환.
    """
    if not isinstance(ex_obj, dict):
        return  # 방어적 처리

    if "exp_sentence" not in ex_obj or ex_obj["exp_sentence"] is None:
        # 기본 한 슬롯을 만들어 준다(문자열)
        ex_obj["exp_sentence"] = ""
        yield ("str", ex_obj, "exp_sentence")
        return

    raw = ex_obj["exp_sentence"]

    if isinstance(raw, list):
        # 각 원소가 dict/str 등 무엇이든 슬롯으로 간주
        for i in range(len(raw)):
            yield ("list", raw, i)
        return

    if isinstance(raw, dict):
        # dict 한 덩어리를 하나의 슬롯으로 간주: dict['설명문장']에 기록
        yield ("dict", raw, None)
        return

    if isinstance(raw, str):
        # 문자열 하나면 그 자체가 슬롯
        yield ("str", ex_obj, "exp_sentence")
        return

    # 알 수 없는 타입이면 문자열 슬롯으로 강제
    ex_obj["exp_sentence"] = ""
    yield ("str", ex_obj, "exp_sentence")


def _assign_sentence_to_slot(slot, new_sentence: str):
    """
    슬롯 정의에 따라 new_sentence를 해당 위치에 기록.
    dict/list-dict 슬롯은 _set_exp_sentence_on_dict()로 정규화하여
    '설명문장' 계열 키가 둘 이상 생기지 않도록 하나의 키로만 overwrite.
    """
    mode, container, pos = slot
    s = "" if new_sentence is None else str(new_sentence)

    if mode == "list":
        item = container[pos]
        if isinstance(item, dict):
            # 기존/변형 키 정리 후 하나의 키로 overwrite
            _set_exp_sentence_on_dict(item, s, prefer_existing=True)
        else:
            container[pos] = s
        return

    if mode == "dict":
        # 기존/변형 키 정리 후 하나의 키로 overwrite
        _set_exp_sentence_on_dict(container, s, prefer_existing=True)
        return

    if mode == "str":
        container[pos] = s
        return


def apply_excel_desc_to_json(json_obj: Dict[str, Any], excel_df: pd.DataFrame) -> Dict[str, Any]:
    """
    엑셀의 '설명 문장' 값을 JSON의 exp_sentence에 반영하여 JSON 객체를 반환.
    매핑 규칙:
      - 같은 id 블록 내에서 행 순서대로 EX의 exp_sentence '슬롯'에 순차 매핑
      - 슬롯 수 > 엑셀 문장 수 → 남는 슬롯은 원본 유지
      - 슬롯 수 < 엑셀 문장 수 → 초과 문장은 무시
    """
    mapping = _collect_excel_sentences_by_id(excel_df)

    docs = json_obj.get("document", [])
    if not isinstance(docs, list):
        return json_obj  # 형식 방어

    for doc in docs:
        doc_id = str(doc.get("id", ""))
        seq = mapping.get(doc_id, [])
        if not seq:
            continue

        used = 0
        ex_list = doc.get("EX", [])
        if not isinstance(ex_list, list):
            continue

        for ex in ex_list:
            for slot in _iter_exp_slots(ex):
                if used >= len(seq):
                    break
                _assign_sentence_to_slot(slot, seq[used])
                used += 1

            if used >= len(seq):
                break

    return json_obj


def _pick_zip_members(zf: zipfile.ZipFile):
    """
    ZIP에서 JSON 1개, XLSX 1개를 추출 대상으로 선택.
    - JSON은 project_*.json 우선, 그 외 첫 번째 .json
    - Excel은 .xlsx 우선(.xls는 의존성에 따라 미지원일 수 있음)
    """
    json_members = [m for m in zf.namelist() if m.lower().endswith(".json")]
    xlsx_members = [m for m in zf.namelist() if m.lower().endswith(".xlsx")]
    xls_members  = [m for m in zf.namelist() if m.lower().endswith(".xls")]

    # JSON 우선순위: project_*
    json_member = None
    for m in json_members:
        name = Path(m).name
        if name.startswith("project_"):
            json_member = m
            break
    if json_member is None and json_members:
        json_member = json_members[0]

    # Excel 우선순위: .xlsx → (가능하면 .xls 대체)
    excel_member = xlsx_members[0] if xlsx_members else (xls_members[0] if xls_members else None)

    return json_member, excel_member

def _collect_excel_sentences_by_id_type(df: pd.DataFrame, skip_blank: bool = False) -> Dict[str, Dict[str, List[str]]]:
    required = {"id", "유형", "설명 문장"}
    if not required.issubset(set(df.columns)):
        raise ValueError("엑셀에 'id', '유형', '설명 문장' 컬럼이 모두 필요합니다.")

    tmp = df.copy()

    # ★ 병합 셀 복원
    if "id" in tmp.columns:
        tmp["id"] = tmp["id"].ffill()
    if "유형" in tmp.columns:
        tmp["유형"] = tmp["유형"].ffill()

    tmp["id"] = tmp["id"].astype(str)
    tmp["유형"] = tmp["유형"].astype(str)
    tmp["설명 문장"] = tmp["설명 문장"].fillna("").astype(str)

    bucket: Dict[str, Dict[str, List[str]]] = defaultdict(lambda: defaultdict(list))
    for _, row in tmp.iterrows():
        _id = row["id"].strip()
        label = row["유형"].strip()
        sent = row["설명 문장"].strip()
        if skip_blank and not sent:
            continue
        ref_type = _label_to_ref_type(label)   # ▼ 패치2에서 추가되는 정규화 함수 사용
        bucket[_id][ref_type].append(sent)
    return bucket

def _label_to_ref_type(label: Any) -> str:
    s = "" if label is None else str(label).strip()
    # 1차: 정확 매칭(기존 역매핑)
    if s in REF_MAP_INV:
        return REF_MAP_INV[s]
    # 2차: 공백/밑줄 제거 후 느슨 매칭
    s2 = s.replace(" ", "").replace("_", "")
    norm = {
        "표설명문장": "table_ref", "표설명": "table_ref", "표": "table_ref",
        "행설명문장": "row_ref",   "행설명": "row_ref",   "행": "row_ref",
        "열설명문장": "col_ref",   "열설명": "col_ref",   "열": "col_ref",
        "불연속영역설명문장": "cell_ref", "불연속영역설명": "cell_ref",
        "불연속영역": "cell_ref", "불연속": "cell_ref"
    }
    if s2 in norm:
        return norm[s2]
    # 3차: 시작어로 판별(예: '불연속영역 설명...'과 같은 변형)
    if s2.startswith("불연속"):
        return "cell_ref"
    return s  # 마지막 수단: 원문 라벨 그대로(매칭 실패 시 건너뜀)

def apply_excel_desc_to_json(json_obj: Dict[str, Any], excel_df: pd.DataFrame, skip_blank: bool = True) -> Dict[str, Any]:
    """
    엑셀의 '설명 문장'을 JSON의 exp_sentence에 반영.
    - '유형' 컬럼이 있으면 id+유형별 정밀 매핑(권장)
    - 없으면(구버전 엑셀) id 순서대로 일괄 배분(이전 호환)
    - skip_blank=True면 엑셀의 빈 문자열은 원본을 덮어쓰지 않음
    """
    docs = json_obj.get("document", [])
    if not isinstance(docs, list):
        return json_obj

    has_type_col = "유형" in excel_df.columns

    if has_type_col:
        # 유형 정밀 매핑
        mapping_by_type = _collect_excel_sentences_by_id_type(excel_df, skip_blank=skip_blank)

        for doc in docs:
            doc_id = str(doc.get("id", ""))
            type_map = mapping_by_type.get(doc_id, {})
            if not type_map:
                continue

            ex_list = doc.get("EX", [])
            if not isinstance(ex_list, list):
                continue

            # 유형별 소비 인덱스
            used_by_type: Dict[str, int] = defaultdict(int)

            for ex in ex_list:
                ref_type = ""
                ref = ex.get("reference", {})
                if isinstance(ref, dict):
                    ref_type = str(ref.get("reference_type", "")).strip()

                seq = type_map.get(ref_type, [])
                if not seq:
                    continue

                used = used_by_type[ref_type]

                for slot in _iter_exp_slots(ex):
                    if used >= len(seq):
                        break
                    s = seq[used]
                    # skip_blank=True인 경우, mapping 단계에서 이미 공백은 제거됨
                    _assign_sentence_to_slot(slot, s)
                    used += 1

                used_by_type[ref_type] = used

        return json_obj

    # === '유형' 컬럼이 없는 구버전 엑셀 호환 경로 ===
    mapping = _collect_excel_sentences_by_id(excel_df)

    for doc in docs:
        doc_id = str(doc.get("id", ""))
        seq = mapping.get(doc_id, [])
        if not seq:
            continue

        used = 0
        ex_list = doc.get("EX", [])
        if not isinstance(ex_list, list):
            continue

        for ex in ex_list:
            for slot in _iter_exp_slots(ex):
                if used >= len(seq):
                    break
                s = seq[used]
                if skip_blank and (s is None or str(s).strip() == ""):
                    used += 1
                    continue
                _assign_sentence_to_slot(slot, s)
                used += 1

            if used >= len(seq):
                break

    return json_obj


def apply_excel_desc_to_json_from_zip(
    zip_bytes: bytes,
    sheet_name: Optional[str] = None,
    skip_blank: bool = True,
) -> Tuple[bytes, str]:
    """
    ZIP 바이트(엑셀 + 단일 JSON)를 받아,
    엑셀의 '설명 문장'을 JSON에 반영한 뒤 (updated_json_bytes, suggested_filename)을 반환.
    - 엑셀에 '유형' 컬럼이 있으면 id+유형 정밀 매핑 사용
    - skip_blank=True면 엑셀의 빈 문자열은 원본을 덮어쓰지 않음
    """
    if not isinstance(zip_bytes, (bytes, bytearray)):
        raise TypeError("zip_bytes는 bytes이어야 합니다.")

    with zipfile.ZipFile(BytesIO(zip_bytes), "r") as zf:
        json_member, excel_member = _pick_zip_members(zf)
        if not json_member:
            raise FileNotFoundError("ZIP 안에 JSON 파일이 없습니다.")
        if not excel_member:
            raise FileNotFoundError("ZIP 안에 Excel(.xlsx) 파일이 없습니다.")

        # JSON 로드
        with zf.open(json_member) as jf:
            json_obj = json.loads(jf.read().decode("utf-8"))

        # Excel 로드
        with zf.open(excel_member) as ef:
            df = pd.read_excel(ef, sheet_name=sheet_name) if sheet_name else pd.read_excel(ef)

        # 반영 (유형 컬럼이 있으면 정밀 매핑 분기 사용)
        updated = apply_excel_desc_to_json(json_obj, df, skip_blank=skip_blank)

        # 출력 파일명 제안
        base = Path(json_member).name
        out_name = (base[:-5] if base.lower().endswith(".json") else base) + "_updated.json"

        return json.dumps(updated, ensure_ascii=False, indent=2).encode("utf-8"), out_name
