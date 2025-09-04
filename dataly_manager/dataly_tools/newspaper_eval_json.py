import os
import json
import copy
import re
import random
import shutil  # zip 압축용

def ensure_folder(path):
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)

def strip_prefix(fname):
    return re.sub(r'^\d+_', '', fname)

def find_subfolder(parent, candidates):
    # A/A팀, B/B팀 등 다양한 폴더명 지원
    for c in candidates:
        p = os.path.join(parent, c)
        if os.path.isdir(p):
            return p
    return None

def merge_newspaper_eval(week_num=1, files_per_week=102, base_dir=None):
    if base_dir is None:
        raise ValueError("base_dir는 반드시 지정되어야 합니다. (ZIP 해제 경로)")
    dir_a = find_subfolder(base_dir, ["A", "A팀"])
    dir_b = find_subfolder(base_dir, ["B", "B팀"])
    merge_base = os.path.join(base_dir, "merged")
    output_dir = os.path.join(merge_base, f"{week_num}주차")
    ensure_folder(output_dir)

    if not dir_a or not dir_b:
        raise FileNotFoundError(f"'A' 또는 'B' 폴더를 찾을 수 없습니다. base_dir: {base_dir}")

    default_eval = {
        "id": "evaluatorAJ",
        "content": {"description": None, "claims": None, "arguments": None, "comment": ""},
        "organization": {"completion": None, "comment": ""},
        "expression": {"accuracy": None, "comment": ""}
    }

    a_files = {strip_prefix(fn): fn for fn in os.listdir(dir_a) if fn.endswith(".json")}
    b_files = {strip_prefix(fn): fn for fn in os.listdir(dir_b) if fn.endswith(".json")}
    candidate_keys = sorted(set(a_files.keys()) & set(b_files.keys()))

    used_keys = set()
    for prev_week in range(1, week_num):
        prev_dir = os.path.join(merge_base, f"{prev_week}주차")
        if not os.path.isdir(prev_dir):
            continue
        for f in os.listdir(prev_dir):
            if f.endswith(".json"):
                key = strip_prefix(f)
                used_keys.add(key)
    remain_keys = [k for k in candidate_keys if k not in used_keys]
    remain_keys = remain_keys[:files_per_week]

    out_files = []
    count = 0
    for key in remain_keys:
        a_path = os.path.join(dir_a, a_files[key])
        b_path = os.path.join(dir_b, b_files[key])
        with open(a_path, 'r', encoding='utf-8') as fa:
            data_a = json.load(fa)
        with open(b_path, 'r', encoding='utf-8') as fb:
            data_b = json.load(fb)

        sc1_b = data_b.get("SC1")
        if not isinstance(sc1_b, dict):
            print(f"'{key}'에 B팀 SC1 없음, 건너뜀")
            continue
        sc2 = copy.deepcopy(sc1_b)
        sc2["ai_flag"] = False
        sc2["evaluation"] = copy.deepcopy(default_eval)

        if isinstance(data_a.get("document"), list):
            articles = data_a["document"]
            corpus_id = data_a.get("id")
            corpus_meta = data_a.get("metadata")
        else:
            articles = [data_a]
            corpus_id = data_a.get("id")
            corpus_meta = data_a.get("metadata")

        for art in articles:
            if isinstance(art.get("SC1"), dict):
                art["SC1"]["ai_flag"] = False
                art["SC1"]["evaluation"] = copy.deepcopy(default_eval)
            else:
                art["SC1"] = {"ai_flag": False, "evaluation": copy.deepcopy(default_eval)}
            art["SC2"] = copy.deepcopy(sc2)

        merged = {
            "id": corpus_id,
            "metadata": corpus_meta,
            "document": articles
        }

        out_name = f"{week_num}_{key}.json"
        out_path = os.path.join(output_dir, out_name)
        with open(out_path, 'w', encoding='utf-8') as fo:
            json.dump(merged, fo, ensure_ascii=False, indent=2)
        out_files.append(out_path)
        print(f"완료: {out_name}")
        count += 1

    msg = f"{week_num}주차: 병합 완료 (총 {count}건, 이전 주차 사용 {len(used_keys)}건, 남은 후보 {len(remain_keys)})"
    print(msg)

    # zip 파일 생성
    zip_path = os.path.join(output_dir, f"merged_{week_num}주차.zip")
    shutil.make_archive(zip_path.replace('.zip', ''), 'zip', output_dir)

    return msg, output_dir, zip_path
