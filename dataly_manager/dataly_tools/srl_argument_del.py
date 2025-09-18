# -*- coding: utf-8 -*-
from __future__ import annotations

"""
SRL argument 정리 엔진 (엑셀/CSV 없이 JSON만 처리)

규칙
- 라벨 보정: argument.label 의 'PTR' -> 'PRT'
- 인자 삭제: argument.label 이 비어 있고(없음/None/공백) AND
  argument가 커버하는 단어들 중 morph.label == "VX" 가 하나라도 있으면 → 그 argument 삭제
  (⚠︎ argument가 모두 사라져도 SRL 프레임은 유지)
- 프레디케이트 삭제(신규): SRL의 predicate가 가리키는 word_id들의 형태소 라벨 중
  'V'로 시작하는 라벨을 모았을 때 집합이 {'VX'}(= V계열이 오직 VX 뿐)이라면
  → 해당 SRL 프레임 전체( predicate + argument ) 삭제
  예) VV+EC+VX → 유지,  VX+EC → 삭제

호출
- srl_argument_cleanup(in_path, write_back=True/False, progress_cb=None)
  - write_back=True 이면 실제 JSON 파일을 덮어씁니다(임시폴더에서 사용할 것).
  - 반환: 요약/로그
"""

import json
from pathlib import Path
from typing import Any, Dict, List, Optional, Set, Union, Callable, Tuple


# ---------------- 내부 유틸 ----------------
def _is_empty_label(arg: Dict[str, Any]) -> bool:
    if "label" not in arg:
        return True
    v = arg.get("label")
    if v is None:
        return True
    return str(v).strip() == ""


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
    morph.word_id 가 문자열로 들어오는 데이터 특성에 맞춰 정규화.
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


def _arg_word_ids_from_word_id_field(arg: Dict[str, Any]) -> Set[int]:
    """
    argument.word_id 가 int 또는 list 로 올 수 있음 -> set[int] 로 정규화.
    """
    res: Set[int] = set()
    if "word_id" not in arg:
        return res
    wid_val = arg.get("word_id")
    if isinstance(wid_val, list):
        for v in wid_val:
            iv = _to_int_safe(v)
            if iv is not None:
                res.add(iv)
    else:
        iv = _to_int_safe(wid_val)
        if iv is not None:
            res.add(iv)
    return res


def _arg_word_ids_from_span(arg: Dict[str, Any], sent: Dict[str, Any]) -> Set[int]:
    """
    argument.begin~end 문자 범위로 포함되는 word.id 를 수집.
    word.begin >= arg.begin AND word.end <= arg.end 로 판정.
    """
    res: Set[int] = set()
    ab = _to_int_safe(arg.get("begin"))
    ae = _to_int_safe(arg.get("end"))
    if ab is None or ae is None:
        return res

    for w in _collect_words(sent):
        if not isinstance(w, dict):
            continue
        wid = _to_int_safe(w.get("id"))
        wb = _to_int_safe(w.get("begin"))
        we = _to_int_safe(w.get("end"))
        if wid is None or wb is None or we is None:
            continue
        if wb >= ab and we <= ae:
            res.add(wid)
    return res


def _extract_arg_word_ids(arg: Dict[str, Any], sent: Dict[str, Any]) -> Set[int]:
    """
    argument 가 커버하는 word_id 집합을 추출:
    1) word_id 필드 우선 사용
    2) 없거나 비어 있으면 begin~end 범위로 추출
    """
    wids = _arg_word_ids_from_word_id_field(arg)
    if not wids:
        wids = _arg_word_ids_from_span(arg, sent)
    return wids


def _argument_has_VX(arg: Dict[str, Any], sent: Dict[str, Any], morph_by_wid: Dict[int, List[str]]) -> bool:
    """
    argument 가 커버하는 단어들 중 morph.label == 'VX' 가 하나라도 있으면 True.
    """
    wid_set = _extract_arg_word_ids(arg, sent)
    if not wid_set:
        return False
    for wid in wid_set:
        labels = morph_by_wid.get(wid, [])
        if any(lab == "VX" for lab in labels):
            return True
    return False


def _iter_json_files(path: Path) -> Tuple[List[Path], Optional[Path]]:
    """입력 경로에서 JSON 파일 목록과 루트 디렉터리 반환."""
    if path.is_file() and path.suffix.lower() == ".json":
        return [path], None
    if path.is_dir():
        return list(path.rglob("*.json")), path
    return [], None


def _save_json(obj: Dict[str, Any], in_file: Path) -> None:
    """원본 파일 덮어쓰기 저장."""
    in_file.write_text(json.dumps(obj, ensure_ascii=False, indent=2), encoding="utf-8")


# --------- 라벨 보정 유틸 ---------
def _normalize_label(v: Any) -> str:
    """라벨 비교를 위한 정규화: 문자열화 + strip + 대문자."""
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
                # argument 가 dict 단일 객체로 오는 데이터 방어
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
    """
    srl_item.predicate 에서 word_id들을 set[int]로 수집.
    """
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
    프레디케이트의 word_id들에 매칭되는 morph.label 중
    'V'로 시작하는 라벨들의 집합이 {'VX'}이면 True.
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

    # 1) 인자 삭제 규칙 & 2) 프레디케이트 VX-only 규칙
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

                # 2) 프레디케이트 VX-only → 프레임 삭제
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
                    continue  # 이 SRL 프레임 자체 제거

                # 1) 인자 삭제 규칙 (argument가 0개여도 프레임 유지)
                args = srl.get("argument")
                if not isinstance(args, list):
                    args = []

                kept_args: List[Dict[str, Any]] = []
                removed_count = 0

                for a in args:
                    if not isinstance(a, dict):
                        removed_count += 1
                        continue

                    if _is_empty_label(a) and _argument_has_VX(a, sent, morph_by_wid):
                        removed_count += 1
                        log_rows.append([
                            str(file_path),
                            str(sent.get("id") or ""),
                            _predicate_surface(srl),
                            str(a.get("form") or ""),
                            "argument_removed_empty_label_with_VX",
                        ])
                    else:
                        kept_args.append(a)

                if removed_count > 0:
                    sentence_changed = True

                # ✅ argument가 0개여도 프레임은 유지
                srl["argument"] = kept_args
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
