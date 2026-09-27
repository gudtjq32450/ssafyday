# -*- coding: utf-8 -*-
"""
SSAFY 인물퀴즈 동일 단계 내 문제 무작위 셔플(Shuffle) 스크립트
- 각 카테고리/난이도(15개 라운드, 각 20문제) 내부에서만 문제 순서를 랜덤으로 섞습니다.
- 문제와 사진(imageUrl), 힌트(hint), 정답(answer)은 1:1로 정확하게 유지됩니다.
- index.html의 createInitialQuizzes()를 교체하고, 브라우저 캐시 무효화를 위해
  LocalStorage 버전을 v9 -> v10으로 승격합니다.
"""

import os
import sys
import json
import random
import re

sys.stdout.reconfigure(encoding='utf-8')

BASE_DIR = r"C:\Users\wognl\OneDrive\바탕 화면\윤형섭\SSAFY관련\싸피데이용 게임\인물퀴즈"
HTML_PATH = os.path.join(BASE_DIR, "index.html")

def shuffle_quizzes():
    print(f">>> index.html 로드 중: {HTML_PATH}")
    with open(HTML_PATH, "r", encoding="utf-8") as f:
        html = f.read()

    # createInitialQuizzes() 배열 추출
    start_idx = html.find("createInitialQuizzes()")
    if start_idx == -1:
        raise ValueError("createInitialQuizzes() not found")

    array_start = html.find("[", start_idx)
    bracket_depth = 0
    array_end = -1
    in_string = False
    escape = False
    quote_char = None

    for i in range(array_start, len(html)):
        c = html[i]
        if escape:
            escape = False
            continue
        if c == '\\':
            escape = True
            continue
        if in_string:
            if c == quote_char:
                in_string = False
            continue
        else:
            if c in ('"', "'"):
                in_string = True
                quote_char = c
                continue
            if c == '[':
                bracket_depth += 1
            elif c == ']':
                bracket_depth -= 1
                if bracket_depth == 0:
                    array_end = i + 1
                    break

    quizzes = json.loads(html[array_start:array_end])
    print(f">>> 총 {len(quizzes)}개 퀴즈 파싱 완료!")

    # 15개 라운드별로 문제 분류
    round_order = []
    rounds = {}
    for q in quizzes:
        rk = q['roundKey']
        if rk not in rounds:
            round_order.append(rk)
            rounds[rk] = []
        rounds[rk].append(q)

    print(f">>> 총 {len(round_order)}개 라운드 확인됨:")

    # 난수 생성기 시드 (필요시 랜덤)
    rng = random.Random(42) # 고정 시드를 주어 일관성 보장하거나, time 기반
    # 사용자가 원한 것은 가나다순 탈피 및 자연스러운 랜덤 셔플!
    import time
    rng = random.Random(int(time.time()))

    all_shuffled_quizzes = []

    for rk in round_order:
        round_items = rounds[rk]
        old_answers = [item['answer'] for item in round_items]
        
        # 셔플 수행 (같은 단계/라운드 내에서만!)
        rng.shuffle(round_items)
        new_answers = [item['answer'] for item in round_items]

        # 새 qNum 및 id 재부여
        for new_idx, item in enumerate(round_items, start=1):
            item['qNum'] = new_idx
            item['id'] = f"{rk}_q{new_idx}"
            all_shuffled_quizzes.append(item)

        print(f"\n[{rk}] (20문제 셔플 완료)")
        print(f"  기존 앞 3개: {old_answers[:3]}")
        print(f"  새로운 앞 3개: {new_answers[:3]}")

    print(f"\n>>> 전체 {len(all_shuffled_quizzes)}개 퀴즈 셔플 및 번호 재할당 완료!")

    # 1. createInitialQuizzes() 대체
    formatted_json = json.dumps(all_shuffled_quizzes, ensure_ascii=False, indent=8)
    new_func = f"""createInitialQuizzes() {{
            return {formatted_json};
        }}"""

    # 정규식으로 createInitialQuizzes() { ... return [ ... ]; } 교체
    pattern = r"createInitialQuizzes\(\)\s*\{[\s\S]*?return\s*\[[\s\S]*?\];\s*\}"
    html, count = re.subn(pattern, new_func, html, count=1)
    if count == 0:
        raise ValueError("createInitialQuizzes 정규식 치환 실패")
    print(">>> index.html createInitialQuizzes() 코드 치환 성공!")

    # 2. LocalStorage 버전 v14 -> v15로 승격 (브라우저가 무조건 새 데이터를 읽도록)
    v14_keys = [
        "ssafy_quiz_list_v14",
        "ssafy_solved_status_v14",
        "ssafy_round_status_v14",
        "ssafy_team_scores_v14",
        "ssafy_game_set_v14",
        "ssafy_last_solver_v14"
    ]
    for k in v14_keys:
        k_new = k.replace("_v14", "_v15")
        html = html.replace(k, k_new)
    print(">>> LocalStorage 버전을 v14 -> v15로 승격 완료 (캐시 무효화)")

    # 3. index.html 저장
    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(html)
    print(">>> index.html 파일 저장 완료!")

if __name__ == "__main__":
    shuffle_quizzes()
