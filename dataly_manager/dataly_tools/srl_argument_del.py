# -*- coding: utf-8 -*-
from __future__ import annotations

import json
import csv
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple, Set, Union, Callable


# ---------- 내부 유틸 ----------
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
    res: Set[int] = set()
    if "word_id" not in arg:
        return res
    wid_val: Union[int, str, List[Any]] = arg.get("word_id")

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
    wids = _arg_word_ids_from_word_id_field(arg)
    if not wids:
        wids = _arg_word_ids_from_span(arg, sent)
    return wids


def _argument_has_VX(arg: Dict[str, Any], sent: Dict[str, Any], morph_by_wid: Dict[int, List[str]]) -> bool:
    wid_set = _extract_arg_word_ids(arg, sent)
    if not wid_set:
        return False
    for wid in wid_set:
        labels = morph_by_wid.get(wid, [])
        if any(lab == "VX" for lab in labels):
            return True
    return False


def _iter_json_files(path: Path) -> Tuple[List[Path], Optional[Path]]:
    if path.is_file() and path.suffix.lower() == ".json":
        return [path], None
    if path.is_dir():
        return list(path.rglob("*.json")), path
    return [], None


def _save_json(obj: Dict[str, Any], in_file: Path, out_dir: Optional[Path], root_dir: Optional[Path]) -> Path:
    text = json.dumps(obj, ensure_ascii=False, indent=2)
    if out_dir is None:
        in_file.write_text(text, encoding="utf-8")
        return in_file
    if root_dir and in_file.is_relative_to(root_dir):
        rel = in_file.relative_to(root_dir)
    else:
        rel = in_file.name
    out_path = out_dir / rel
    out_path.parent.mkdir(parents=True, exist_ok=True)
    out_path.write_text(text, encoding="utf-8")
    return out_path


def _process_json_obj(
    obj: Dict[str, Any],
    file_path: Path,
    log_rows: List[List[str]]
) -> bool:
    changed = False
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
                            "argument_removed_empty_label_with_VX"
                        ])
                    else:
                        kept_args.append(a)

                if removed_count > 0:
                    sentence_changed = True

                if len(kept_args) == 0:
                    log_rows.append([
                        str(file_path),
                        str(sent.get("id") or ""),
                        _predicate_surface(srl),
                        "",
                        "srl_removed_no_arguments"
                    ])
                    continue

                srl["argument"] = kept_args
                new_srl.append(srl)

            if sentence_changed or len(new_srl) != len(srl_list):
                changed = True

            sent["SRL"] = new_srl

    return changed


# ---------- 공개 엔진 함수 ----------
def srl_argument_cleanup(
    in_path: Union[str, Path],
    out_dir: Optional[Union[str, Path]] = None,
    report_csv: Optional[Union[str, Path]] = None,
    progress_cb: Optional[Callable[[int, int, Path], None]] = None,
) -> Dict[str, Any]:
    """
    SRL argument 중 (label 비어 있고) AND (VX 포함) 조건을 삭제하고,
    argument가 비면 해당 SRL 항목 자체를 제거.

    Returns:
        {
          "total_files": int,
          "changed_files": int,
          "skipped_files": int,
          "log_rows": List[List[str]],
          "outputs": List[{"src": Path, "dst": Path}],
        }
    """
    p_in = Path(in_path)
    p_out = Path(out_dir) if out_dir else None
    p_csv = Path(report_csv) if report_csv else None

    if not p_in.exists():
        raise FileNotFoundError(f"경로가 존재하지 않습니다: {p_in}")

    if p_out:
        p_out.mkdir(parents=True, exist_ok=True)

    files, root = _iter_json_files(p_in)
    log_rows: List[List[str]] = [["file", "sentence_id", "predicate_form", "argument_form", "action"]]
    outputs: List[Dict[str, Path]] = []

    changed_cnt = 0
    skipped_cnt = 0

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

        if _process_json_obj(obj, f, log_rows):
            dst = _save_json(obj, f, p_out, root)
            outputs.append({"src": f, "dst": dst})
            changed_cnt += 1
        else:
            skipped_cnt += 1

    if p_csv and len(log_rows) > 1:
        with p_csv.open("w", newline="", encoding="utf-8") as fp:
            writer = csv.writer(fp)
            writer.writerows(log_rows)

    return {
        "total_files": total,
        "changed_files": changed_cnt,
        "skipped_files": skipped_cnt,
        "log_rows": log_rows,
        "outputs": outputs,
        "report_csv": str(p_csv) if p_csv else None,
    }
