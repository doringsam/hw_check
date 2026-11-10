"""
테스트용 데이터 생성 스크립트:
- 원본 ipynb 파일 생성
- 학생 제출 ipynb 파일 4개 생성 (그대로 제출, 코드 수정, 주석/공백만 수정, 변수값 수정)
- 수강생 엑셀 파일(students.xlsx) 생성 (미제출자 1명 포함 총 5명)
"""
import os
import json
import openpyxl

# 1. 테스트 폴더 생성
os.makedirs("test_env/submissions", exist_ok=True)

def create_notebook(code_cells):
    """코드 셀 리스트를 받아 ipynb JSON 딕셔너리를 생성하는 헬퍼 함수"""
    cells = []
    for code in code_cells:
        cells.append({
            "cell_type": "code",
            "execution_count": 1,
            "metadata": {},
            "outputs": [],
            "source": code.splitlines(keepends=True)
        })
    return {
        "cells": cells,
        "metadata": {
            "language_info": {"name": "python"}
        },
        "nbformat": 4,
        "nbformat_minor": 2
    }

# 원본 코드
orig_code = [
    "# 과제: 두 수의 합을 구하는 함수를 작성하세요.\ndef add(a, b):\n    return a + b\n\nresult = add(10, 20)\nprint('결과:', result)"
]

# (1) 원본 파일 저장
with open("test_env/original.ipynb", "w", encoding="utf-8") as f:
    json.dump(create_notebook(orig_code), f, ensure_ascii=False, indent=2)

# (2) 학생 1: 원본 그대로 제출 (김철수 -> 0점 예상)
with open("test_env/submissions/20240001_김철수_과제1.ipynb", "w", encoding="utf-8") as f:
    json.dump(create_notebook(orig_code), f, ensure_ascii=False, indent=2)

# (3) 학생 2: 코드 수정 제출 (이영희 -> 1점 예상)
modified_code_1 = [
    "# 과제: 두 수의 합을 구하는 함수를 작성하세요.\ndef add(a, b):\n    print('계산 시작')\n    return a + b\n\nresult = add(10, 20)\nprint('결과:', result)"
]
with open("test_env/submissions/20240002_이영희_과제1.ipynb", "w", encoding="utf-8") as f:
    json.dump(create_notebook(modified_code_1), f, ensure_ascii=False, indent=2)

# (4) 학생 3: 주석과 띄어쓰기/빈 줄만 추가/수정 (박민수 -> 스마트 비교로 0점 예상)
comment_only_code = [
    "# [학생 작성 주석] 과제입니다.\n# 주석만 잔뜩 추가함\n\ndef add(a, b):\n\n    return a + b\n\n\nresult = add(10, 20)\nprint('결과:', result) # 출력문 주석"
]
with open("test_env/submissions/20240003_박민수_과제1.ipynb", "w", encoding="utf-8") as f:
    json.dump(create_notebook(comment_only_code), f, ensure_ascii=False, indent=2)

# (5) 학생 4: 숫자 값 수정 (최수진 -> 1점 예상)
modified_code_2 = [
    "# 과제: 두 수의 합을 구하는 함수를 작성하세요.\ndef add(a, b):\n    return a + b\n\nresult = add(30, 50)\nprint('결과:', result)"
]
with open("test_env/submissions/20240004_최수진_과제1.ipynb", "w", encoding="utf-8") as f:
    json.dump(create_notebook(modified_code_2), f, ensure_ascii=False, indent=2)

# (6) 수강생 엑셀 파일 생성
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "수강생목록"

# 헤더 작성: 학번, 이름, 점수
ws.append(["학번", "이름", "점수"])

# 학생 데이터
students = [
    ("20240001", "김철수", None),
    ("20240002", "이영희", None),
    ("20240003", "박민수", None),
    ("20240004", "최수진", None),
    ("20240005", "정미제출", None), # 미제출 학생 (0점 예상)
]

for s in students:
    ws.append(list(s))

wb.save("test_env/students.xlsx")
print("테스트 데이터 생성 완료!")
