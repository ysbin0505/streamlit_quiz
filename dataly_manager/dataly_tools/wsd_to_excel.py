# dataly_manager/dataly_tools/wsd_to_excel.py
import os
import re
import json
from typing import List, Dict, Any
import pandas as pd

__all__ = ["jsons_to_wsd_excel"]


def jsons_to_wsd_excel(
    base_dir: str,
    excel_name: str = "WSD_sense_tagging_simple.xlsx",
    include_memo_sheet: bool = True,
    memo_placement: str = "by_row",  # "by_row" | "first" | "repeat"
    memo_sep: str = " | ",
) -> str:
    """
    폴더(하위폴더 포함)를 재귀 순회하며 *.json을 스캔해 엑셀로 변환합니다.

    생성 컬럼 (WSD 시트):
      - file_name, doc_id, sent_id, sentence
      - word_id, word, morph, WSD Form
      - head, DP Label
      - SRL Span                : SRL.argument.word_id (여러 개면 ' + ' 결합)
      - SRL Label               : SRL.argument.label (여러 개면 ' + ' 결합)
      - SRL Predicate Lamma     : SRL.predicate 의 "word_id/lemma" (여러 개면 ' + ' 결합)
      - ant_sen_id              : ZA_argument.sentence_id 의 말미 숫자들 (예: '3.1')
      - ant_word_id             : ZA_argument.word_id
      - ant_form                : ZA_argument.form
      - restored_form           : antecedent[].form  결합
      - restored_type           : antecedent[].type  결합
      - prev_word, prev_morph, prev_WSD Form
      - memo_count, memos       : 메모 배치 옵션에 따름

    반환값: 생성한 엑셀의 절대경로
    """
    excel_rows: List[Dict[str, Any]] = []
    memo_rows: List[Dict[str, Any]] = []

    # ---------- helpers ----------
    def _normalize_memos(m) -> List[Dict[str, str]]:
        norm = []
        if isinstance(m, list):
            for item in m:
                if isinstance(item, dict):
                    norm.append(
                        {
                            "row": str(item.get("row", "")).strip(),
                            "text": str(item.get("text", "")).strip(),
                        }
                    )
                else:
                    norm.append({"row": "", "text": str(item).strip()})
        elif isinstance(m, dict):
            norm.append(
                {"row": str(m.get("row", "")).strip(), "text": str(m.get("text", "")).strip()}
            )
        elif isinstance(m, (str, int, float)):
            norm.append({"row": "", "text": str(m).strip()})
        return [x for x in norm if (x.get("row") or x.get("text"))]

    def _join(lst: List[str], sep: str = memo_sep) -> str:
        return sep.join([s for s in lst if s])

    def _short_sid(sid: str) -> str:
        """문장ID 꼬리의 '숫자.숫자' 또는 '숫자'만 추출 (미존재 시 원문)."""
        if not sid:
            return ""
        m = re.search(r"(\d+\.\d+)$", str(sid))
        if m:
            return m.group(1)
        m = re.search(r"(\d+)$", str(sid))
        return m.group(1) if m else str(sid)

    def _uniq_join(vals: List[str]) -> str:
        if not vals:
            return ""
        seen, out = set(), []
        for v in vals:
            if v not in seen:
                seen.add(v)
                out.append(v)
        return " + ".join(out)

    # ---------- gather json files (recursive) ----------
    json_paths: List[str] = []
    for root, _, files in os.walk(base_dir):
        for fn in files:
            if fn.lower().endswith(".json"):
                json_paths.append(os.path.join(root, fn))

    # ---------- parse ----------
    for path in json_paths:
        fname = os.path.basename(path)
        try:
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
        except Exception:
            continue

        for doc in data.get("document", []):
            doc_id = doc.get("id", "")
            doc_level_memos = _normalize_memos(doc.get("memos", []))

            for sentence in doc.get("sentence", []):
                sent_id = sentence.get("id", "")
                sent_form = sentence.get("form", "")

                word_list = sentence.get("word", []) or []
                morph_list = sentence.get("morph", []) or []
                wsd_list = sentence.get("WSD", []) or []
                dp_list = sentence.get("DP", []) or []
                srl_list = sentence.get("SRL", []) or []
                za_list = sentence.get("ZA", []) or []

                # ----- memos -----
                raw_memos = sentence.get("memos", []) or doc_level_memos
                memos_norm = _normalize_memos(raw_memos)

                if include_memo_sheet and memos_norm:
                    for order, mm in enumerate(memos_norm, 1):
                        memo_rows.append(
                            {
                                "file_name": fname,
                                "doc_id": doc_id,
                                "sent_id": sent_id,
                                "sentence": sent_form,
                                "memo_order": order,
                                "memo_row": mm.get("row", ""),
                                "memo_text": mm.get("text", ""),
                            }
                        )

                memos_by_row: Dict[str, List[str]] = {}
                unmapped_buffer: List[str] = []
                for mm in memos_norm:
                    r = (mm.get("row") or "").strip()
                    t = (mm.get("text") or "").strip()
                    if r.isdigit():
                        memos_by_row.setdefault(r, []).append(t)
                    elif t:
                        unmapped_buffer.append(t)

                # ----- morph map (by word_id) -----
                morphs_by_wordid: Dict[str, List[str]] = {}
                for morph in morph_list:
                    wid = str(morph.get("word_id"))
                    morphs_by_wordid.setdefault(wid, []).append(
                        f"{morph.get('form','')}/{morph.get('label','')}"
                    )

                # ----- WSD map (by word_id_display, fallback: word_id) -----
                wsds_by_wordid: Dict[str, List[str]] = {}
                for wsd in wsd_list:
                    base_wid = wsd.get("word_id_display", wsd.get("word_id"))
                    if base_wid is None:
                        continue
                    wid = str(base_wid)
                    wsds_by_wordid.setdefault(wid, []).append(
                        f"{wsd.get('form','')}/{wsd.get('sense_id','')}"
                    )

                # ----- DP map (by word_id) -----
                dp_by_wordid = {str(dp.get("word_id")): dp for dp in dp_list}

                # ----- SRL maps -----
                srl_span_by_wid: Dict[str, List[str]] = {}
                srl_label_by_wid: Dict[str, List[str]] = {}
                srl_predlemma_by_wid: Dict[str, List[str]] = {}

                for frame in srl_list:
                    preds = frame.get("predicate", []) or []
                    args = frame.get("argument", []) or []

                    # predicate 디스크립터들 (word_id/lemma)
                    pred_descs: List[str] = []
                    for p in preds:
                        p_wid = p.get("word_id")
                        lemma = p.get("lemma") or p.get("form") or ""
                        if p_wid is not None:
                            pred_descs.append(f"{p_wid}/{str(lemma).strip()}")
                        elif lemma:
                            pred_descs.append(str(lemma).strip())

                    # argument 의 word_id를 각 단어행에 연결
                    for arg in args:
                        label = str(arg.get("label", "")).strip()
                        wids = arg.get("word_id")

                        if isinstance(wids, list):
                            wid_list = [str(w) for w in wids if w not in (None, "")]
                        elif wids in (None, ""):
                            wid_list = []
                        else:
                            wid_list = [str(wids)]

                        for wid in wid_list:
                            srl_span_by_wid.setdefault(wid, []).append(wid)  # 숫자 id 기록
                            if label:
                                srl_label_by_wid.setdefault(wid, []).append(label)
                            if pred_descs:
                                srl_predlemma_by_wid.setdefault(wid, []).extend(pred_descs)

                # uniq join
                for wid in list(srl_span_by_wid.keys()):
                    srl_span_by_wid[wid] = [_uniq_join(srl_span_by_wid[wid])]
                for wid in list(srl_label_by_wid.keys()):
                    srl_label_by_wid[wid] = [_uniq_join(srl_label_by_wid[wid])]
                for wid in list(srl_predlemma_by_wid.keys()):
                    srl_predlemma_by_wid[wid] = [_uniq_join(srl_predlemma_by_wid[wid])]

                # ----- ZA map -----
                # ZA_argument 기준: 해당 word_id 를 가진 단어행에 아래 필드 노출
                # ant_* : ZA_argument 에서, restored_* : antecedent[] 집합
                za_by_wid: Dict[str, List[tuple]] = {}
                for item in za_list:
                    za_arg = item.get("ZA_argument") or {}
                    z_form = str(za_arg.get("form", "")).strip()
                    z_sid = _short_sid(str(za_arg.get("sentence_id", "")).strip())
                    z_wid = za_arg.get("word_id")

                    if isinstance(z_wid, list):
                        z_wid_str = (
                            str(z_wid[0]) if z_wid and z_wid[0] not in (None, "", "#") else None
                        )
                    else:
                        z_wid_str = str(z_wid) if z_wid not in (None, "", "#") else None

                    ants = item.get("antecedent", []) or []
                    ant_forms = [str(a.get("form", "")).strip() for a in ants if isinstance(a, dict)]
                    ant_types = [str(a.get("type", "")).strip() for a in ants if isinstance(a, dict)]
                    restored_form = " + ".join([x for x in ant_forms if x])
                    restored_type = " + ".join([x for x in ant_types if x])

                    if z_wid_str:
                        za_by_wid.setdefault(z_wid_str, []).append(
                            (z_sid, z_wid_str, z_form, restored_form, restored_type)
                        )

                # ----- sentence-level memo string -----
                sentence_memo_all = _join([txt for arr in memos_by_row.values() for txt in arr])
                if not sentence_memo_all and unmapped_buffer:
                    sentence_memo_all = _join(unmapped_buffer)

                # ----- row emit -----
                prev_word = prev_morph = prev_wsd = ""

                for i, w in enumerate(word_list):
                    wid = str(w.get("id"))
                    word_form = w.get("form", "")

                    morph_str = " + ".join(morphs_by_wordid.get(wid, []))
                    wsd_str = " + ".join(wsds_by_wordid.get(wid, []))

                    head = label = ""
                    if wid in dp_by_wordid:
                        head = str(dp_by_wordid[wid].get("head", ""))
                        label = str(dp_by_wordid[wid].get("label", ""))

                    # 메모 배치
                    if memo_placement == "by_row":
                        row_memos = memos_by_row.get(wid, [])
                        memos_for_row = _join(row_memos)
                        memo_count_for_row = len(row_memos)
                    elif memo_placement == "first":
                        memos_for_row = sentence_memo_all if i == 0 else ""
                        memo_count_for_row = (
                            len(sentence_memo_all.split(memo_sep))
                            if i == 0 and sentence_memo_all
                            else ""
                        )
                    else:  # repeat
                        memos_for_row = sentence_memo_all
                        memo_count_for_row = (
                            len(sentence_memo_all.split(memo_sep)) if sentence_memo_all else 0
                        )

                    # SRL/ZA
                    srl_span = (srl_span_by_wid.get(wid) or [""])[0]
                    srl_label = (srl_label_by_wid.get(wid) or [""])[0]
                    srl_plemma = (srl_predlemma_by_wid.get(wid) or [""])[0]

                    if wid in za_by_wid:
                        tuples = za_by_wid[wid]
                        ant_sen_id = " + ".join([t[0] for t in tuples if t[0]])
                        ant_word_id = " + ".join([t[1] for t in tuples if t[1]])
                        ant_form = " + ".join([t[2] for t in tuples if t[2]])
                        restored_form = " + ".join([t[3] for t in tuples if t[3]])
                        restored_type = " + ".join([t[4] for t in tuples if t[4]])
                    else:
                        ant_sen_id = ant_word_id = ant_form = restored_form = restored_type = ""

                    excel_rows.append(
                        {
                            "file_name": fname,
                            "doc_id": doc_id,
                            "sent_id": sent_id,
                            "sentence": sent_form,
                            "word_id": wid,
                            "word": word_form,
                            "morph": morph_str,
                            "WSD Form": wsd_str,
                            "head": head,
                            "DP Label": label,
                            "SRL Span": srl_span,
                            "SRL Label": srl_label,
                            "SRL Predicate Lamma": srl_plemma,
                            "ant_sen_id": ant_sen_id,
                            "ant_word_id": ant_word_id,
                            "ant_form": ant_form,
                            "restored_form": restored_form,
                            "restored_type": restored_type,
                            "prev_word": prev_word,
                            "prev_morph": prev_morph,
                            "prev_WSD Form": prev_wsd,
                            "memo_count": memo_count_for_row,
                            "memos": memos_for_row,
                        }
                    )

                    prev_word = word_form
                    prev_morph = morph_str
                    prev_wsd = wsd_str

    # ---------- save ----------
    df = pd.DataFrame(excel_rows)
    excel_save_path = os.path.join(base_dir, excel_name)
    if include_memo_sheet and memo_rows:
        df_memos = pd.DataFrame(memo_rows)
        with pd.ExcelWriter(excel_save_path, engine="openpyxl") as writer:
            df.to_excel(writer, index=False, sheet_name="WSD")
            df_memos.to_excel(writer, index=False, sheet_name="Memos")
    else:
        df.to_excel(excel_save_path, index=False)

    return excel_save_path
