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
                # ----- SRL (GUI와 동일: 인자 span은 '마지막 word_id' 행에만 표기, span은 wid 나열) -----
                srl_by_wid: Dict[str, Dict[str, str]] = {}
                seen_keys: set = set()

                for frame in srl_list:
                    preds = frame.get("predicate", []) or []
                    pred_lemma = ""
                    pred_cell = ""  # 표시용 예: "10/체결하다"
                    if preds:
                        p0 = preds[0]
                        pred_lemma = str(p0.get("lemma", "") or "")
                        pw = str(p0.get("word_id", "")).strip()
                        # GUI 로직과 맞춤: word_id와 lemma가 있으면 "id/lemma", 아니면 있는 값만
                        pred_cell = f"{pw}/{pred_lemma}" if pw.isdigit() and pred_lemma else (pw or pred_lemma)

                    for arg in (frame.get("argument", []) or []):
                        wids = arg.get("word_id", [])
                        if isinstance(wids, int):
                            wids = [wids]
                        elif not isinstance(wids, list):
                            wids = [wids] if wids else []

                        # 정수 wid만 모아 정렬/중복제거
                        span_sorted = sorted({int(x) for x in wids if str(x).isdigit()})
                        if not span_sorted:
                            continue

                        # GUI와 동일한 표시: ", "로 조인 (예: "3, 4, 5")
                        span_str = ", ".join(str(x) for x in span_sorted)
                        label = str(arg.get("label", "") or "")

                        # 중복 방지
                        key = (tuple(span_sorted), label, pred_lemma)
                        if key in seen_keys:
                            continue
                        seen_keys.add(key)

                        # 마지막 토큰 행에만 기록
                        target_wid = str(max(span_sorted))
                        cell = srl_by_wid.setdefault(target_wid, {"span": "", "label": "", "pred": ""})

                        # 여러 인자/프레임이 같은 행에 겹치면 " / "로 구분
                        cell["span"] = (cell["span"] + " / " if cell["span"] else "") + span_str
                        cell["label"] = (cell["label"] + " / " if cell["label"] else "") + label
                        if pred_cell:
                            if cell["pred"]:
                                # 같은 pred_cell 중복 연결 방지
                                if pred_cell not in cell["pred"].split(" / "):
                                    cell["pred"] += " / " + pred_cell
                            else:
                                cell["pred"] = pred_cell

                # ----- ZA map (antecedent 기준으로 각 wid 행에 기록) -----
                za_by_wid: Dict[str, List[tuple]] = {}

                for item in za_list:
                    za_arg = item.get("ZA_argument") or {}
                    rest_form = str(za_arg.get("form", "")).strip()
                    rest_type = str(za_arg.get("type", "")).strip()

                    ants = item.get("antecedent", []) or []
                    ant_sids, ant_wids, ant_forms = [], [], []

                    for a in ants:
                        # sentence_id 꼬리만 추출
                        sid_tail = _short_sid(str(a.get("sentence_id", "")).strip())
                        if sid_tail:
                            ant_sids.append(sid_tail)

                        # antecedent.word_id (단일/리스트 모두 처리)
                        w = a.get("word_id")
                        if isinstance(w, list):
                            for x in w:
                                s = str(x).strip()
                                if s.isdigit():
                                    ant_wids.append(s)
                        else:
                            s = str(w).strip()
                            if s.isdigit():
                                ant_wids.append(s)

                        # antecedent.form
                        f = str(a.get("form", "")).strip()
                        if f:
                            ant_forms.append(f)

                    # antecedent 정보가 하나도 없으면 스킵
                    if not (ant_sids or ant_wids or ant_forms):
                        continue

                    # 항목 내부 중복 제거(순서 보존)
                    def _uniq_keep_order(lst: List[str]) -> List[str]:
                        seen, out = set(), []
                        for v in lst:
                            if v not in seen:
                                seen.add(v);
                                out.append(v)
                        return out

                    ant_sen_id_join = " + ".join(_uniq_keep_order(ant_sids))
                    ant_word_id_join = " + ".join(_uniq_keep_order(ant_wids))
                    ant_form_join = " + ".join(_uniq_keep_order(ant_forms))

                    # 이 ZA 항목을 "모든 antecedent wid" 행에 붙인다
                    for tw in _uniq_keep_order(ant_wids):
                        za_by_wid.setdefault(tw, []).append(
                            (ant_sen_id_join, ant_word_id_join, ant_form_join, rest_form, rest_type)
                        )

                # ----- sentence-level memo string -----
                sentence_memo_all = _join([txt for arr in memos_by_row.values() for txt in arr])
                if not sentence_memo_all and unmapped_buffer:
                    sentence_memo_all = _join(unmapped_buffer)

                # ----- row emit -----
                # prev_word = prev_morph = prev_wsd = ""

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
                    cell = srl_by_wid.get(wid, {})
                    srl_span = cell.get("span", "")
                    srl_label = cell.get("label", "")
                    srl_plemma = cell.get("pred", "")

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
                            # "prev_word": prev_word,
                            # "prev_morph": prev_morph,
                            # "prev_WSD Form": prev_wsd,
                            "memo_count": memo_count_for_row,
                            "memos": memos_for_row,
                        }
                    )

                    # prev_word = word_form
                    # prev_morph = morph_str
                    # prev_wsd = wsd_str

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
