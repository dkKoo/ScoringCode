# 🎓 채점 프로그램

이 프로그램은 학생들의 답안을 자동으로 채점하고 결과를 출력하는 Python 스크립트입니다.

## 📁 파일 구성

- `채점프로그램.py`: 채점 프로그램의 주요 Python 스크립트입니다.
- `정답파일업로드서식.xlsx`: 정답 파일의 서식을 보여주는 예시 파일입니다.
- `학생답안파일업로드서식.xlsx`: 학생 답안 파일의 서식을 보여주는 예시 파일입니다.

## 🚀 사용 방법

1. `채점프로그램.py` 파일을 실행합니다.
2. 프로그램이 실행되면 "정답 파일 업로드" 버튼을 클릭하여 정답 파일을 선택합니다.
   - 정답 파일은 `정답파일업로드서식.xlsx`와 같은 형식으로 작성되어야 합니다.
   - 첫 번째 행은 헤더로 사용되며, 두 번째 행부터 정답이 입력되어야 합니다.
   - 정답은 1번 문제부터 순서대로 입력되어야 합니다.
3. "학생 답안 파일 업로드" 버튼을 클릭하여 학생들의 답안 파일을 선택합니다.
   - 학생 답안 파일은 `학생답안파일업로드서식.xlsx`와 같은 형식으로 작성되어야 합니다.
   - 첫 번째 행은 헤더로 사용되며, 두 번째 행부터 학생들의 답안이 입력되어야 합니다.
   - 각 학생의 답안은 학번, 이름, 1번 문제부터 순서대로 입력되어야 합니다.
4. "채점" 버튼을 클릭하여 채점을 시작합니다.
5. 채점이 완료되면 결과가 화면에 출력됩니다.
   - 각 학생의 이름, 학번, 점수, 순위, 틀린 문제 번호와 정답이 표시됩니다.
6. 채점 결과는 자동으로 Excel 파일로 저장됩니다.
   - 결과 파일에는 학번, 이름, 점수, 순위, 틀린 문제 정보가 포함됩니다.

## ⚠️ 주의 사항

- 정답 파일과 학생 답안 파일은 반드시 지정된 서식에 맞게 작성되어야 합니다.
- 파일 업로드 시 정답 파일과 학생 답안 파일을 정확히 선택해야 합니다.
- 프로그램 실행 중에는 파일을 수정하거나 삭제하지 마세요.
- 결과 파일은 프로그램이 실행된 디렉토리에 자동으로 저장됩니다.
