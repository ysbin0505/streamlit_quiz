# dataly_manager/dataly_tools/wsd_to_excel.py
import os
import json
import re
import pandas as pd
from typing import List

__all__ = ["jsons_to_wsd_excel"]

def jsons_to_wsd_excel(
    base_dir: str,
    excel_name: str = "WSD_sense_tagging_simple.xlsx",
    include_memo_sheet: bool = True,
    memo_placement: str = "by_row",  # "by_row" | "first" | "repeat"
    memo_sep: str = " | "
) -> str:
    """
    폴더 내 *.json을 스캔해 엑셀로 변환.
    추가 컬럼:
      - SRL Span: argument의 word_id (여러 개면 + 결합)
      - SRL Label: argument.label (여러 개면 + 결합)
      - SRL Predicate Lamma: 'predicate word_id/lemma' (여러 개면 + 결합)
      - ant_sen_id: ZA_argument.sentence_id의 꼬리(예: '3.1')
      - ant_word_id: ZA_argument.word_id
      - ant_form: ZA_argument.form
      - restored_form: antecedent[].form (+ 결합)
      - restored_type: antecedent[].type (+ 결합)
    """
    excel_rows: List[dict] = []
    memo_rows:  List[dict] = []

    def _normalize_memos(m):
        norm = []
        if isinstance(m, list):
            for item in m:
                if isinstance(item, dict):
                    norm.append({"row": str(item.get("row", "")).strip(),
                                 "text": str(item.get("text", "")).strip()})
                else:
                    norm.append({"row": "", "text": str(item).strip()})
        elif isinstance(m, dict):
            norm.append({"row": str(m.get("row", "")).strip(), "text": str(m.get("text", "")).strip()})
        elif isinstance(m, (str, int, float)):
            norm.append({"row": "", "text": str(m).strip()})
        return [x for x in norm if (x.get("row") or x.get("text"))]

    def _join(lst, sep=memo_sep):
        return sep.join([s for s in lst if s])

    def _short_sid(sid: str) -> str:
        """NZRW...3.1 처럼 끝의 '숫자.숫자'만 추출; 없으면 원문."""
        if not sid: return ""
        m = re.search(r'(\d+\.\d+)$', str(sid))
        if m: return m.group(1)
        m = re.search(r'(\d+)$', str(sid))
        return m.group(1) if m else str(sid)

    for fname in os.listdir(base_dir):
        if not fname.endswith(".json"):
            continue
        path = os.path.join(base_dir, fname)
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
                word_list = sentence.get("word", [])
                morph_list = sentence.get("morph", [])
                wsd_list = sentence.get("WSD", [])
                dp_list = sentence.get("DP", [])
                srl_list = sentence.get("SRL", []) or []
                za_list  = sentence.get("ZA", []) or []

                # sentence-level 우선, 없으면 doc-level 사용
                raw_memos = sentence.get("memos", []) or doc_level_memos
                memos_norm = _normalize_memos(raw_memos)

                if include_memo_sheet and memos_norm:
                    for order, mm in enumerate(memos_norm, 1):
                        memo_rows.append({
                            "file_name": fname, "doc_id": doc_id, "sent_id": sent_id,
                            "sentence": sent_form, "memo_order": order,
                            "memo_row": mm.get("row", ""), "memo_text": mm.get("text", "")
                        })

                # by_row 배치용
                memos_by_row, unmapped_buffer = {}, []
                for mm in memos_norm:
                    r = (mm.get("row") or "").strip()
                    t = (mm.get("text") or "").strip()
                    if r.isdigit():
                        memos_by_row.setdefault(r, []).append(t)
                    elif t:
                        unmapped_buffer.append(t)

                # morphs: word_id별 합치기
                morphs_by_wordid = {}
                for morph in morph_list:
                    wid = str(morph.get("word_id"))
                    morphs_by_wordid.setdefault(wid, []).append(f"{morph.get('form','')}/{morph.get('label','')}")

                # WSD: word_id_display별 합치기
                wsds_by_wordid = {}
                for wsd in wsd_list:
                    wid = str(wsd.get("word_id_display"))
                    wsds_by_wordid.setdefault(wid, []).append(f"{wsd.get('form','')}/{wsd.get('sense_id','')}")

                # DP: word_id 기준
                dp_by_wordid = {str(dp.get("word_id")): dp for dp in dp_list}

                # ===== SRL 맵 =====
                srl_span_by_wid, srl_label_by_wid, srl_predlemma_by_wid = {}, {}, {}
                for frame in srl_list:
                    preds = frame.get("predicate", []) or []
                    args  = frame.get("argument", []) or []

                    # 'word_id/lemma' 표기(여러 predicate 가능)
                    pred_descs = []
                    for p in preds:
                        p_wid = p.get("word_id")
                        lemma = p.get("lemma") or p.get("form") or ""
                        if p_wid is not None:
                            pred_descs.append(f"{p_wid}/{str(lemma).strip()}")
                        elif lemma:
                            pred_descs.append(str(lemma).strip())

                    # argument 단어에 매핑
                    for arg in args:
                        label = str(arg.get("label", "")).strip()
                        wids  = arg.get("word_id")
                        if isinstance(wids, list):
                            wid_list = [str(w) for w in wids if w not in (None, "")]
                        elif wids in (None, ""):
                            wid_list = []
                        else:
                            wid_list = [str(wids)]

                        for wid in wid_list:
                            srl_span_by_wid.setdefault(wid, []).append(wid)  # 숫자 id
                            if label:
                                srl_label_by_wid.setdefault(wid, []).append(label)
                            if pred_descs:
                                srl_predlemma_by_wid.setdefault(wid, []).extend(pred_descs)

                def _uniq_join(vals):
                    if not vals: return ""
                    seen, out = set(), []
                    for v in vals:
                        if v not in seen:
                            seen.add(v); out.append(v)
                    return " + ".join(out)

                for wid in list(srl_span_by_wid.keys()):
                    srl_span_by_wid[wid] = _uniq_join(srl_span_by_wid[wid])
                for wid in list(srl_label_by_wid.keys()):
                    srl_label_by_wid[wid] = _uniq_join(srl_label_by_wid[wid])
                for wid in list(srl_predlemma_by_wid.keys()):
                    srl_predlemma_by_wid[wid] = _uniq_join(srl_predlemma_by_wid[wid])

                # ===== ZA 맵 =====
                za_by_wid = {}
                for item in za_list:
                    za_arg = item.get("ZA_argument") or {}
                    z_form = str(za_arg.get("form", "")).strip()
                    z_sid  = _short_sid(str(za_arg.get("sentence_id", "")).strip())
                    z_wid  = za_arg.get("word_id")

                    if isinstance(z_wid, list):
                        z_wid_str = str(z_wid[0]) if z_wid and z_wid[0] not in (None, "", "#") else None
                    else:
                        z_wid_str = str(z_wid) if z_wid not in (None, "", "#") else None

                    ants = item.get("antecedent", []) or []
                    ant_forms = [str(a.get("form","")).strip() for a in ants if isinstance(a, dict)]
                    ant_types = [str(a.get("type","")).strip() for a in ants if isinstance(a, dict)]
                    restored_form = " + ".join([x for x in ant_forms if x])
                    restored_type = " + ".join([x for x in ant_types if x])

                    if z_wid_str:
                        za_by_wid.setdefault(z_wid_str, []).append(
                            (z_sid, z_wid_str, z_form, restored_form, restored_type)
                        )

                # 메모 문자열
                sentence_memo_all = _join([txt for arr in memos_by_row.values() for txt in arr])
                if not sentence_memo_all and unmapped_buffer:
                    sentence_memo_all = _join(unmapped_buffer)

                prev_word = prev_morph = prev_wsd = ""

                for i, word in enumerate(word_list):
                    wid = str(word.get("id"))
                    word_form = word.get("form", "")
                    morph_str = " + ".join(morphs_by_wordid.get(wid, []))
                    wsd_str = " + ".join(wsds_by_wordid.get(wid, []))

                    head = label = ""
                    if wid in dp_by_wordid:
                        head = dp_by_wordid[wid].get("head", "")
                        label = dp_by_wordid[wid].get("label", "")

                    # 메모 배치
                    if memo_placement == "by_row":
                        row_memos = memos_by_row.get(wid, [])
                        memos_for_row = _join(row_memos)
                        memo_count_for_row = len(row_memos)
                    elif memo_placement == "first":
                        memos_for_row = sentence_memo_all if i == 0 else ""
                        memo_count_for_row = (len(sentence_memo_all.split(memo_sep)) if i == 0 and sentence_memo_all else "")
                    else:
                        memos_for_row = sentence_memo_all
                        memo_count_for_row = (len(sentence_memo_all.split(memo_sep)) if sentence_memo_all else 0)

                    # SRL/ZA 값
                    srl_span   = srl_span_by_wid.get(wid, "")
                    srl_label  = srl_label_by_wid.get(wid, "")
                    srl_plemma = srl_predlemma_by_wid.get(wid, "")

                    if wid in za_by_wid:
                        tuples = za_by_wid[wid]
                        ant_sen_id     = " + ".join([t[0] for t in tuples if t[0]])
                        ant_word_id    = " + ".join([t[1] for t in tuples if t[1]])
                        ant_form       = " + ".join([t[2] for t in tuples if t[2]])
                        restored_form  = " + ".join([t[3] for t in tuples if t[3]])
                        restored_type  = " + ".join([t[4] for t in tuples if t[4]])
                    else:
                        ant_sen_id = ant_word_id = ant_form = restored_form = restored_type = ""

                    excel_rows.append({
                        "file_name": fname, "doc_id": doc_id, "sent_id": sent_id, "sentence": sent_form,
                        "word_id": wid, "word": word_form, "morph": morph_str, "WSD Form": wsd_str,
                        "head": head, "DP Label": label,

                        "SRL Span": srl_span,
                        "SRL Label": srl_label,
                        "SRL Predicate Lamma": srl_plemma,
                        "ant_sen_id": ant_sen_id,
                        "ant_word_id": ant_word_id,
                        "ant_form": ant_form,
                        "restored_form": restored_form,
                        "restored_type": restored_type,

                        "prev_word": prev_word, "prev_morph": prev_morph, "prev_WSD Form": prev_wsd,
                        "memo_count": memo_count_for_row, "memos": memos_for_row
                    })

                    prev_word = word_form
                    prev_morph = morph_str
                    prev_wsd = wsd_str

    # 저장
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
