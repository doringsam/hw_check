# -*- coding: utf-8 -*-
"""
====================================================================
           [숙제 자동 채점 프로그램 - Homework Grader (다회차 누적형)]
====================================================================
* 주요 특징:
  1. [단일 엑셀 열(Column) 누적 관리]:
     - 매 회차 과제를 채점할 때마다 같은 엑셀 파일에 새로운 과제명 열을
       오른쪽에 계속 추가하여 한 학기 성적을 한눈에 누적 관리합니다.
     - 기존에 채점했던 과제를 다시 채점(재제출 반영 등)할 경우, 중복 열을
       만들지 않고 해당 열의 점수만 최신으로 깔끔하게 갱신(Update)합니다.
  2. [유연한 채점 모드 선택]:
     - 모드 1 (스마트 비교): 주석/공백만 바꾼 꼼수는 0점, 실제 코드 수정자만 1점 부여
     - 모드 2 (제출자 전원 인정): 출제 오류나 단순 제출 확인 과제 시,
       제출하기만 했으면 무조건 1점 부여 (미제출자만 0점)
  3. [주피터 노트북(.ipynb & .ipynb.txt) 완벽 지원]:
     - 주피터 노트북 파일 내부의 순수 'code 셀'만 추출하여 비교합니다.
  4. [데이터 안전 백업]:
     - 엑셀 파일을 갱신하기 전 자동으로 백업 파일(.bak)을 생성하여 실수를 방지합니다.
====================================================================
"""

import os
import sys
import json
import ast
import re
import shutil
from pathlib import Path
import openpyxl

# 윈도우 콘솔 환경에서 한글 및 기호 출력 인코딩 안전 처리
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
        sys.stderr.reconfigure(encoding="utf-8")
    except Exception:
        pass


def extract_code_from_file(file_path: str) -> str:
    """
    파일에서 소스코드를 추출하는 함수입니다.
    - 파일이 주피터 노트북(.ipynb 또는 .ipynb.txt)인 경우: 'code' 셀에 들어있는 코드만 모아서 반환합니다.
    - 파일이 일반 파이썬(.py) 또는 텍스트인 경우: 전체 텍스트를 그대로 읽어옵니다.
    """
    path = Path(file_path)
    filename_lower = path.name.lower()

    # 주피터 노트북(.ipynb 또는 다운로드 시 확장자가 붙은 .ipynb.txt) 파일 처리
    if filename_lower.endswith(".ipynb") or filename_lower.endswith(".ipynb.txt"):
        try:
            with open(path, "r", encoding="utf-8") as f:
                notebook_data = json.load(f)

            code_lines = []
            # 노트북 안의 모든 셀(cell)을 하나씩 확인합니다.
            for cell in notebook_data.get("cells", []):
                # 셀 종류가 'code'(코드 셀)인 경우에만 코드를 수집합니다.
                if cell.get("cell_type") == "code":
                    source = cell.get("source", [])
                    # source가 리스트 형태로 들어있는 경우 문자열로 합칩니다.
                    if isinstance(source, list):
                        code_lines.append("".join(source))
                    elif isinstance(source, str):
                        code_lines.append(source)

            # 셀과 셀 사이에 줄바꿈을 두어 하나의 완성된 파이썬 코드로 합칩니다.
            return "\n".join(code_lines)

        except Exception as e:
            print(f"  [경고] 노트북 파일 읽기 실패 ({path.name}): {e}")
            return ""

    # 일반 텍스트 또는 .py 파일인 경우
    else:
        try:
            with open(path, "r", encoding="utf-8") as f:
                return f.read()
        except UnicodeDecodeError:
            # utf-8로 읽기 실패 시 cp949(한국어 윈도우 기본)로 재시도
            with open(path, "r", encoding="cp949") as f:
                return f.read()


def clean_for_ast(code_str: str) -> str:
    """
    주피터 노트북 특유의 매직 명령어(!pip, %matplotlib 등)를 파이썬 AST 파서가
    에러로 인식하지 않도록 주석(#)으로 안전하게 변환해주는 함수입니다.
    """
    cleaned_lines = []
    for line in code_str.splitlines():
        stripped = line.strip()
        # 느낌표(!)나 퍼센트(%)로 시작하는 주피터 명령어는 주석 처리
        if stripped.startswith("!") or stripped.startswith("%"):
            cleaned_lines.append("# " + line)
        else:
            cleaned_lines.append(line)
    return "\n".join(cleaned_lines)


def normalize_code_text(code_str: str) -> str:
    """
    주석과 모든 공백을 제거하여 순수한 핵심 코드 문자열만 추출하는 보조 함수입니다.
    (AST 분석이 불가능한 문법 에러 코드 비교 시 백업용으로 활용됩니다.)
    """
    result = []
    for line in code_str.splitlines():
        line_without_comment = re.sub(r"#.*$", "", line)
        line_no_spaces = re.sub(r"\s+", "", line_without_comment)
        if line_no_spaces:
            result.append(line_no_spaces)
    return "".join(result)


def is_code_modified_smart(original_code: str, student_code: str) -> bool:
    """
    [스마트 비교 알고리즘]
    원본 코드와 학생 코드를 비교하여 실질적인 수정 여부를 판별합니다.
    - True  -> 코드를 수정함 (1점)
    - False -> 코드를 수정하지 않고 그대로 냄 (0점)
    """
    orig_cleaned = clean_for_ast(original_code)
    stud_cleaned = clean_for_ast(student_code)

    orig_ast_dump = None
    stud_ast_dump = None

    try:
        orig_tree = ast.parse(orig_cleaned)
        orig_ast_dump = ast.dump(orig_tree)
    except SyntaxError:
        pass

    try:
        stud_tree = ast.parse(stud_cleaned)
        stud_ast_dump = ast.dump(stud_tree)
    except SyntaxError:
        # 학생이 코드 수정 중 문법 에러를 낸 경우 -> 코드를 수정한 것으로 인정
        return True

    # 양쪽 모두 정상적인 파이썬 코드인 경우 AST 구조 비교
    if orig_ast_dump is not None and stud_ast_dump is not None:
        return orig_ast_dump != stud_ast_dump

    # 백업 로직: 주석 및 공백 제거 텍스트 직접 비교
    orig_norm = normalize_code_text(original_code)
    stud_norm = normalize_code_text(student_code)
    return orig_norm != stud_norm


def scan_submission_files(folder_path: str) -> dict:
    """
    숙제 폴더를 검사하여 학생들의 과제 파일 목록을 수집합니다.
    - 파일명 형식: 학번_이름_*.ipynb (또는 .ipynb.txt)
    - 반환값: { '학번': { 'name': 이름, 'path': 경로, 'filename': 파일명 } }
    """
    submissions = {}
    folder = Path(folder_path)

    if not folder.exists() or not folder.is_dir():
        print(f"[오류] 폴더를 찾을 수 없습니다: {folder_path}")
        return submissions

    for file in folder.iterdir():
        if not file.is_file():
            continue

        name_lower = file.name.lower()
        if not (name_lower.endswith(".ipynb") or name_lower.endswith(".ipynb.txt")):
            continue

        # 확장자를 제거한 파일명 추출
        filename = file.name
        if name_lower.endswith(".ipynb.txt"):
            filename_no_ext = filename[:-10]
        elif name_lower.endswith(".ipynb"):
            filename_no_ext = filename[:-6]
        else:
            filename_no_ext = file.stem

        # 언더스코어(_)로 파일명 분리
        parts = filename_no_ext.split("_")
        if len(parts) >= 2:
            student_id = parts[0].strip()   # 학번
            student_name = parts[1].strip() # 이름
            submissions[student_id] = {
                "name": student_name,
                "path": str(file),
                "filename": file.name
            }
        else:
            print(f"  [알림] '학번_이름_*.ipynb' 규칙과 맞지 않는 파일 건너뜀: {file.name}")

    return submissions


def find_or_create_assignment_column(ws, assignment_name: str) -> tuple:
    """
    엑셀 시트에서 이번 과제 점수를 기록할 열(Column) 번호를 결정합니다.
    - 반환값: (target_col_idx, mode_description)
    """
    header_row = 1
    headers = {}
    last_valid_col = 0

    # 1행의 모든 셀을 스캔하여 실제 텍스트가 있는 헤더 확인
    for col_idx in range(1, ws.max_column + 1):
        cell_val = ws.cell(row=header_row, column=col_idx).value
        if cell_val is not None and str(cell_val).strip() != "":
            clean_name = str(cell_val).strip()
            headers[clean_name] = col_idx
            last_valid_col = max(last_valid_col, col_idx)

    # 1) 이미 동일한 과제명의 열이 존재하는 경우 -> 갱신 모드
    if assignment_name in headers:
        target_col = headers[assignment_name]
        return target_col, f"기존 '{assignment_name}' 열 갱신 (열 {target_col})"

    # 2) 기본 '점수' 열이 있는 경우 -> 이번 과제명으로 치환하여 사용
    if "점수" in headers:
        score_col = headers["점수"]
        ws.cell(row=header_row, column=score_col, value=assignment_name)
        return score_col, f"기본 '점수' 열을 '{assignment_name}' 열로 지정 (열 {score_col})"

    # 3) 새로운 과제인 경우 -> 마지막 유효 헤더 바로 다음 열에 새 열 추가
    target_col = last_valid_col + 1
    ws.cell(row=header_row, column=target_col, value=assignment_name)
    return target_col, f"오른쪽에 새로운 '{assignment_name}' 열 추가 (열 {target_col})"


def grade_homework(folder_path: str, original_code_path: str, excel_path: str, assignment_name: str = None, give_all_to_submitted: bool = False):
    """
    숙제 자동 채점을 수행하고 통합 엑셀 파일에 결과를 누적 기록하는 메인 함수입니다.
    
    * 매개변수:
      - folder_path: 학생 과제들이 들어있는 폴더 경로
      - original_code_path: 교사용 원본 소스코드 파일 경로
      - excel_path: 수강생 성적표 엑셀 파일 경로
      - assignment_name: 엑셀에 들어갈 열 이름 (기본값: 폴더명)
      - give_all_to_submitted: True일 경우 제출자 전원 1점 부여 (미제출만 0점)
    """
    if not assignment_name:
        assignment_name = Path(folder_path).name

    mode_title = "제출자 전원 점수 부여 모드 (제출=1점, 미제출=0점)" if give_all_to_submitted else "스마트 비교 모드 (수정=1점, 미수정/미제출=0점)"

    print("\n" + "=" * 60)
    print("           [ 숙제 자동 채점을 시작합니다 ]")
    print("=" * 60)
    print(f"* 이번 과제명 (열 이름): {assignment_name}")
    print(f"* 채점 방식            : {mode_title}")
    print(f"* 숙제 폴더 경로       : {folder_path}")
    if not give_all_to_submitted:
        print(f"* 원본 소스코드 경로   : {original_code_path}")
    print(f"* 수강생 성적표 엑셀   : {excel_path}")
    print("-" * 60)

    # 1. 스마트 비교 모드인 경우 원본 소스코드 읽기
    original_code = ""
    if not give_all_to_submitted:
        if not os.path.exists(original_code_path):
            print(f"[오류] 원본 소스코드 파일을 찾을 수 없습니다: {original_code_path}")
            return
        original_code = extract_code_from_file(original_code_path)
        print("[V] 원본 소스코드 로드 완료!")

    # 2. 숙제 폴더 내 학생 제출 파일 수집
    submissions = scan_submission_files(folder_path)
    print(f"[V] 제출된 과제 파일 수: {len(submissions)}개 발견")

    # 3. 수강생 엑셀 파일 열기
    if not os.path.exists(excel_path):
        print(f"[오류] 엑셀 파일을 찾을 수 없습니다: {excel_path}")
        return

    # 안전을 위한 자동 백업 생성 (.bak)
    try:
        backup_path = excel_path + ".bak"
        shutil.copy2(excel_path, backup_path)
    except Exception:
        pass

    try:
        wb = openpyxl.load_workbook(excel_path)
        ws = wb.active
    except Exception as e:
        print(f"[오류] 엑셀 파일을 여는 중 문제가 발생했습니다: {e}")
        return

    # 4. 엑셀의 학번/이름 열 및 이번 과제 열 위치 결정
    id_col = 1
    name_col = 2
    start_row = 2

    for col_idx in range(1, ws.max_column + 1):
        cell_val = str(ws.cell(row=1, column=col_idx).value or "").strip()
        if "학번" in cell_val:
            id_col = col_idx
        elif "이름" in cell_val:
            name_col = col_idx

    target_score_col, col_mode_desc = find_or_create_assignment_column(ws, assignment_name)

    print(f"[V] 엑셀 매핑: 학번={id_col}열, 이름={name_col}열 | {col_mode_desc}")
    print("-" * 60)

    # 5. 각 학생별 채점 진행
    total_students = 0
    submitted_count = 0
    modified_count = 0    # 1점
    unmodified_count = 0  # 0점 (그대로 제출)
    missing_count = 0     # 0점 (미제출)

    print(f"{'순번':<4} | {'학번':<12} | {'이름':<8} | {'채점결과':<8} | {'상세내용'}")
    print("-" * 60)

    for row_idx in range(start_row, ws.max_row + 1):
        id_cell = ws.cell(row=row_idx, column=id_col)
        name_cell = ws.cell(row=row_idx, column=name_col)
        score_cell = ws.cell(row=row_idx, column=target_score_col)

        if id_cell.value is None or str(id_cell.value).strip() == "":
            continue

        total_students += 1
        student_id = str(id_cell.value).strip()
        student_name = str(name_cell.value).strip() if name_cell.value is not None else ""

        # 해당 학생의 과제 파일 확인
        if student_id in submissions:
            submitted_count += 1
            file_info = submissions[student_id]

            # [모드 2]: 제출자 전원 점수 부여 모드
            if give_all_to_submitted:
                score = 1
                modified_count += 1
                detail = "과제 제출 확인 (전원 인정 통과)"

            # [모드 1]: 스마트 비교 모드
            else:
                student_code = extract_code_from_file(file_info["path"])
                is_modified = is_code_modified_smart(original_code, student_code)

                if is_modified:
                    score = 1
                    modified_count += 1
                    detail = "코드 수정 확인 (통과)"
                else:
                    score = 0
                    unmodified_count += 1
                    detail = "원본 코드 그대로 제출 (미수정)"
        else:
            score = 0
            missing_count += 1
            detail = "과제 미제출"

        score_cell.value = score
        print(f"{total_students:<4} | {student_id:<12} | {student_name:<8} | {score}점      | {detail}")

    # 6. 통합 엑셀 파일 저장
    wb.save(excel_path)

    # 7. 요약 통계 출력
    print("=" * 60)
    print("                 [ 채점 결과 종합 요약 ]")
    print("=" * 60)
    print(f" * 과제명 (열 이름)   : {assignment_name}")
    print(f" * 적용된 채점 방식   : {mode_title}")
    print(f" * 총 수강생 수       : {total_students}명")
    print(f" * 과제 제출자 수     : {submitted_count}명")
    print(f" * 1점 부여 학생 수   : {modified_count}명")
    if not give_all_to_submitted:
        print(f" * 0점 (그대로 제출)  : {unmodified_count}명")
    print(f" * 0점 (과제 미제출)  : {missing_count}명")
    print("-" * 60)
    print(f" [V] 성적표 엑셀 파일이 성공적으로 갱신되었습니다!")
    print(f" [V] 저장 파일: {os.path.abspath(excel_path)}")
    print("=" * 60 + "\n")


def main():
    """
    프로그램의 진입점:
    1) 명령행 인자가 전달된 경우:
       python homework_grader.py [폴더경로] [원본코드경로] [엑셀경로] [과제명(선택)] [채점모드(all 또는 smart)]
    2) 인자가 없는 경우 콘솔 대화형으로 경로와 옵션을 입력받습니다.
    """
    if len(sys.argv) >= 4:
        folder_path = sys.argv[1]
        orig_path = sys.argv[2]
        excel_path = sys.argv[3]
        assignment_name = sys.argv[4] if len(sys.argv) >= 5 else None
        
        # 5번째 인자가 'all' 또는 '2'이면 전원 인정 모드
        give_all = False
        if len(sys.argv) >= 6:
            arg_mode = sys.argv[5].strip().lower()
            if arg_mode in ("all", "2", "give_all", "true"):
                give_all = True

        grade_homework(folder_path, orig_path, excel_path, assignment_name, give_all_to_submitted=give_all)
        return

    # 대화형 프롬프트
    print("\n" + "=" * 60)
    print("      [ 파이썬 과제 자동 채점 프로그램에 오신 것을 환영합니다 ]")
    print("=" * 60)
    print("* 각 항목에서 엔터(Enter)를 누르시면 괄호 [기본값]이 그대로 적용됩니다.\n")

    # 1. 채점 방식 선택
    print("1. 채점 방식을 선택하세요:")
    print("   [1] 스마트 비교 모드 (코드 수정자만 1점, 그대로 제출 0점) - 기본값")
    print("   [2] 제출자 전원 점수 부여 (제출자 모두 1점, 미제출만 0점)")
    mode_input = input("선택 [1]: ").strip()
    give_all = (mode_input == "2")

    # 2. 숙제 폴더 입력
    default_folder = "./submissions/10_19과제" if os.path.exists("./submissions/10_19과제") else "./submissions"
    folder_input = input(f"\n2. 숙제 파일들이 들어있는 폴더 경로 [{default_folder}]: ").strip()
    folder_path = folder_input if folder_input else default_folder

    # 3. 원본 소스코드 파일 입력 (전원 인정 모드가 아닐 때만 중요)
    default_orig = "./sources/10.19_source.ipynb" if os.path.exists("./sources/10.19_source.ipynb") else "./sources"
    if not give_all:
        orig_input = input(f"3. 교사용 원본 소스코드 파일 경로 [{default_orig}]: ").strip()
        orig_path = orig_input if orig_input else default_orig
    else:
        orig_path = default_orig

    # 4. 통합 수강생 엑셀 파일 입력
    default_excel = "./수강생.xlsx"
    excel_input = input(f"4. 수강생 성적표 엑셀 파일(.xlsx) 경로 [{default_excel}]: ").strip()
    excel_path = excel_input if excel_input else default_excel

    # 5. 이번 과제명 (엑셀 열 이름)
    default_assignment = Path(folder_path).name
    assign_input = input(f"5. 엑셀에 추가할 과제 열 이름 [{default_assignment}]: ").strip()
    assignment_name = assign_input if assign_input else default_assignment

    # 채점 시작
    grade_homework(folder_path, orig_path, excel_path, assignment_name, give_all_to_submitted=give_all)


if __name__ == "__main__":
    main()
