#newspaper_eval_json.py

import os
import json
import copy
import re
import random
import platform

def get_base_dir():
    # 맥
    if platform.system() == "Darwin":
        home = os.path.expanduser("~")
        return os.path.join(home, "Desktop", "말뭉치배포", "신문")
    # 윈도우
    elif platform.system() == "Windows":
        home = os.path.expanduser("~")
        return os.path.join(home, "Desktop", "말뭉치배포", "신문")
    # 기타 OS
    else:
        return os.path.join(os.getcwd(), "말뭉치배포", "신문")

def ensure_folder(path):
    if not os.path.exists(path):
        os.makedirs(path, exist_ok=True)

def strip_prefix(fname):
    # 접두어 숫자_ 제거 (1_, 2_ 등)
    return re.sub(r'^\d+_', '', fname)

def merge_newspaper_eval(week_num=1, files_per_week=102):
    # 기본 경로
    base_dir = get_base_dir()
    dir_a = os.path.join(base_dir, "A")
    dir_b = os.path.join(base_dir, "B")
    merge_base = os.path.join(base_dir, "merged")
    output_dir = os.path.join(merge_base, f"{week_num}주차")
    ensure_folder(output_dir)

    default_eval = {
        "id": "evaluatorAJ",
        "content": {"description": None, "claims": None, "arguments": None, "comment": ""},
        "organization": {"completion": None, "comment": ""},
        "expression": {"accuracy": None, "comment": ""}
    }

    # 1. 파일 후보(교집합) 확보
    a_files = {strip_prefix(fn): fn for fn in os.listdir(dir_a) if fn.endswith(".json")}
    b_files = {strip_prefix(fn): fn for fn in os.listdir(dir_b) if fn.endswith(".json")}
    candidate_keys = sorted(set(a_files.keys()) & set(b_files.keys()))

    # 2. 이전 주차의 파일명 모으기 (접두어 제거해서 비교)
    used_keys = set()
    for prev_week in range(1, week_num):
        prev_dir = os.path.join(merge_base, f"{prev_week}주차")
        if not os.path.isdir(prev_dir):
            continue
        for f in os.listdir(prev_dir):
            if f.endswith(".json"):
                key = strip_prefix(f)
                used_keys.add(key)
    # 3. 이번 주차에서 사용할 102개 선정 (이전 주차 사용분 제외)
    remain_keys = [k for k in candidate_keys if k not in used_keys]
    # remain_keys = random.sample(remain_keys, min(files_per_week, len(remain_keys)))  # 랜덤으로 하고 싶으면 주석 해제
    remain_keys = remain_keys[:files_per_week]  # 이름순(혹은 순서)으로 102개

    count = 0
    for key in remain_keys:
        a_path = os.path.join(dir_a, a_files[key])
        b_path = os.path.join(dir_b, b_files[key])
        with open(a_path, 'r', encoding='utf-8') as fa:
            data_a = json.load(fa)
        with open(b_path, 'r', encoding='utf-8') as fb:
            data_b = json.load(fb)

        # B팀 SC1 → SC2 생성
        sc1_b = data_b.get("SC1")
        if not isinstance(sc1_b, dict):
            print(f"'{key}'에 B팀 SC1 없음, 건너뜀")
            continue
        sc2 = copy.deepcopy(sc1_b)
        sc2["ai_flag"] = False
        sc2["evaluation"] = copy.deepcopy(default_eval)

        # A팀 데이터 리스트화
        if isinstance(data_a.get("document"), list):
            articles = data_a["document"]
            corpus_id = data_a.get("id")
            corpus_meta = data_a.get("metadata")
        else:
            articles = [data_a]
            corpus_id = data_a.get("id")
            corpus_meta = data_a.get("metadata")

        # 각 article에 SC1·SC2 추가
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

        # 결과 저장 (주차 접두어 붙이기)
        out_name = f"{week_num}_{key}"
        out_path = os.path.join(output_dir, out_name)
        with open(out_path, 'w', encoding='utf-8') as fo:
            json.dump(merged, fo, ensure_ascii=False, indent=2)
        print(f"완료: {out_name}")
        count += 1

    msg = f"{week_num}주차: 병합 완료 (총 {count}건, 이전 주차 사용 {len(used_keys)}건, 남은 후보 {len(remain_keys)})"
    print(msg)
    return msg
