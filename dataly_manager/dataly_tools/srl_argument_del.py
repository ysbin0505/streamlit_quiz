# -*- coding: utf-8 -*-
from __future__ import annotations

"""
SRL 정리 엔진 (엑셀/CSV 없음)

규칙
- 라벨 보정: argument.label 의 'PTR' -> 'PRT'
- 프레디케이트 삭제(유일 조건): SRL의 predicate가 가리키는 word_id들의 형태소 라벨 중
  'V'로 시작하는 라벨들의 집합이 정확히 {'VX'}(= V계열이 오직 VX 뿐)이라면
  → 해당 SRL 프레임 전체( predicate + argument ) 삭제
  예) VV+EC+VX → 유지,  VX+EC → 삭제

※ 인자(argument) 관련 삭제/정리는 더 이상 수행하지 않음.
   (argument가 비어 있어도 프레임 유지, 빈 라벨/범위/VX 여부 등 무시)

엑셀 생성:
- make_vx_removed_only_excel(result)  → 'predicate_removed_vx_only' 행만 필터링
"""

import io
import json
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Union, Callable


# ---------------- 내부 유틸 ----------------
def _predicate_surface(srl_item: Dict[str, Any]) -> str:
    pred = srl_item.get("predicate")
    if isinstance(pred, list) and pred:
        return str(pred[0].get("form") or "")
    if isinstance(pred, dict):
        return str(pred.get("form") or "")
    return ""


def _to_int_safe(x: Any) -> Optional[int]:
    try:
        if isinstance(x, bool):
            return None
        return int(x)
    except Exception:
        return None


def _collect_words(sent: Dict[str, Any]) -> List[Dict[str, Any]]:
    w = sent.get("word")
    return w if isinstance(w, list) else []


def _collect_morph_labels_by_word(sent: Dict[str, Any]) -> Dict[int, List[str]]:
    """
    word_id(int) -> [morph.label, ...]
    morph.word_id 가 문자열일 수 있어 안전 변환.
    """
    out: Dict[int, List[str]] = {}
    morph_list = sent.get("morph")
    if not isinstance(morph_list, list):
        return out
    for m in morph_list:
        if not isinstance(m, dict):
            continue
        wid = _to_int_safe(m.get("word_id"))
        lab = m.get("label")
        if wid is None or lab is None:
            continue
        out.setdefault(wid, []).append(str(lab))
    return out


def _iter_json_files(path: Path):
    if path.is_file() and path.suffix.lower() == ".json":
        return [path], None
    if path.is_dir():
        return list(path.rglob("*.json")), path
    return [], None


def _save_json(obj: Dict[str, Any], in_file: Path) -> None:
    in_file.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


# --------- 라벨 보정 유틸 ---------
def _normalize_label(v: Any) -> str:
    if v is None:
        return ""
    return str(v).strip().upper()


def _patch_srl_labels(doc: Dict[str, Any]) -> int:
    """
    document[].sentence[].SRL[].argument[].label 에서 'PTR' -> 'PRT'
    반환: 치환 건수
    """
    replaced = 0
    documents = doc.get("document", [])
    if not isinstance(documents, list):
        return 0

    for d in documents:
        sentences = d.get("sentence", [])
        if not isinstance(sentences, list):
            continue

        for sent in sentences:
            srl_list = sent.get("SRL", [])
            if not isinstance(srl_list, list):
                continue

            for frame in srl_list:
                args = frame.get("argument", [])
                if isinstance(args, dict):
                    args = [args]
                if not isinstance(args, list):
                    continue

                for arg in args:
                    if not isinstance(arg, dict):
                        continue
                    raw = arg.get("label", None)
                    if _normalize_label(raw) == "PTR":
                        arg["label"] = "PRT"
                        replaced += 1
    return replaced


# --------- 프레디케이트 VX-only 판단 ---------
def _collect_predicate_word_ids(srl_item: Dict[str, Any]) -> Set[int]:
    res: Set[int] = set()
    pred = srl_item.get("predicate")
    if isinstance(pred, dict):
        wid = _to_int_safe(pred.get("word_id"))
        if wid is not None:
            res.add(wid)
    elif isinstance(pred, list):
        for p in pred:
            if not isinstance(p, dict):
                continue
            wid = _to_int_safe(p.get("word_id"))
            if wid is not None:
                res.add(wid)
    return res


def _predicate_is_vx_only(
    srl_item: Dict[str, Any],
    morph_by_wid: Dict[int, List[str]],
) -> bool:
    """
    프레디케이트 word_id들의 morph.label 중 'V'로 시작하는 라벨의 집합이 정확히 {'VX'}면 True.
    (V계열 라벨이 비어있으면 False)
    """
    wids = _collect_predicate_word_ids(srl_item)
    if not wids:
        return False

    v_like: Set[str] = set()
    for wid in wids:
        for lab in morph_by_wid.get(wid, []):
            nl = _normalize_label(lab)
            if nl.startswith("V"):   # V*, 예: VV, VA, VX, VCP, VCN 등
                v_like.add(nl)

    return (len(v_like) > 0) and (v_like == {"VX"})


# --------- JSON 처리 ---------
def _process_json_obj(
    obj: Dict[str, Any],
    file_path: Path,
    log_rows: List[List[str]],
) -> bool:
    changed = False

    # 0) 파일 단위 라벨 보정(PTR -> PRT)
    patched = _patch_srl_labels(obj)
    if patched > 0:
        changed = True
        log_rows.append([str(file_path), "", "", "", f"label_PTR->PRT:{patched}"])

    # 1) 프레디케이트 VX-only 규칙만 적용
    documents = obj.get("document") or []
    for doc in documents:
        sents = doc.get("sentence") or []
        for sent in sents:
            srl_list = sent.get("SRL")

            if not isinstance(srl_list, list):
                if "SRL" in sent and srl_list is not None:
                    sent["SRL"] = []
                    changed = True
                continue

            if not srl_list:
                continue

            morph_by_wid = _collect_morph_labels_by_word(sent)

            new_srl: List[Dict[str, Any]] = []
            sentence_changed = False

            for srl in srl_list:
                if not isinstance(srl, dict):
                    sentence_changed = True
                    continue

                if _predicate_is_vx_only(srl, morph_by_wid):
                    sentence_changed = True
                    changed = True
                    log_rows.append([
                        str(file_path),
                        str(sent.get("id") or ""),
                        _predicate_surface(srl),
                        "",
                        "predicate_removed_vx_only",
                    ])
                    continue  # 프레임 전체 제거

                # ✅ 인자에 대한 삭제/보정 없음
                new_srl.append(srl)

            if sentence_changed or len(new_srl) != len(srl_list):
                changed = True

            sent["SRL"] = new_srl

    return changed


# ---------------- 공개 API ----------------
def srl_argument_cleanup(
    in_path: Union[str, Path],
    write_back: bool = False,
    progress_cb: Optional[Callable[[int, int, Path], None]] = None,
) -> Dict[str, Any]:
    """
    in_path(파일/폴더) 내 JSON을 정리.
    write_back=True 이면 실제 파일을 덮어씁니다(임시폴더에서 사용할 것).
    """
    p_in = Path(in_path)
    if not p_in.exists():
        raise FileNotFoundError(f"경로가 존재하지 않습니다: {p_in}")

    files, _root = _iter_json_files(p_in)

    log_rows: List[List[str]] = [["file", "sentence_id", "predicate_form", "argument_form", "action"]]
    changed_cnt, skipped_cnt = 0, 0
    changed_files: List[str] = []
    total = len(files)

    for idx, f in enumerate(files, start=1):
        if progress_cb:
            progress_cb(idx, total, f)

        try:
            obj = json.loads(f.read_text(encoding="utf-8"))
        except Exception as e:
            skipped_cnt += 1
            log_rows.append([str(f), "", "", "", f"load_failed: {e}"])
            continue

        changed = _process_json_obj(obj, f, log_rows)
        if changed:
            changed_cnt += 1
            changed_files.append(str(f))
            if write_back:
                try:
                    _save_json(obj, f)
                except Exception as e:
                    log_rows.append([str(f), "", "", "", f"save_failed: {e}"])
        else:
            skipped_cnt += 1

    return {
        "total_files": total,
        "changed_files": changed_cnt,
        "skipped_files": skipped_cnt,
        "changed_files_list": changed_files,
        "log_rows": log_rows,
    }


# ---------------- 엑셀: VX-only 삭제 항목만 ----------------
def make_vx_removed_only_excel(result: Dict[str, Any]) -> bytes:
    """
    action == 'predicate_removed_vx_only' 인 행만 필터링하여 xlsx 바이트로 반환
    - 컬럼: file, sentence_id, predicate_form
    """
    import pandas as pd
    buf = io.BytesIO()

    rows = result.get("log_rows") or []
    header = rows[0] if rows else ["file", "sentence_id", "predicate_form", "argument_form", "action"]
    body = rows[1:] if len(rows) > 1 else []

    # 인덱스 찾기(유연하게)
    try:
        idx_file = header.index("file")
        idx_sid  = header.index("sentence_id")
        idx_pred = header.index("predicate_form")
        idx_act  = header.index("action")
    except ValueError:
        idx_file, idx_sid, idx_pred, idx_act = 0, 1, 2, 4

    filtered = []
    for r in body:
        if len(r) > idx_act and str(r[idx_act]) == "predicate_removed_vx_only":
            filtered.append([
                r[idx_file] if len(r) > idx_file else "",
                r[idx_sid]  if len(r) > idx_sid  else "",
                r[idx_pred] if len(r) > idx_pred else "",
            ])

    df = pd.DataFrame(filtered, columns=["file", "sentence_id", "predicate_form"])

    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="VX_Removed", index=False)

    buf.seek(0)
    return buf.getvalue()
