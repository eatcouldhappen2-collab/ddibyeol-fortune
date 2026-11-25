# AGENTS.md

이 프로젝트의 정본 문서는 **[PROJECT_GUIDE.md](PROJECT_GUIDE.md)** 다. 작업 전에 반드시 읽는다.

- 목표·산출물, 디렉토리 구조, 환경 요구사항, 워크플로우 7단계, 완료 체크리스트, SUMMARIES 작성
  규칙 전문이 모두 PROJECT_GUIDE.md에 있다.
- **파이썬은 `py -X utf8`만** 사용한다(`python`/`python3` 금지).
- 규칙 전문을 이 문서나 다른 곳에 복제하지 않는다 — 항상 PROJECT_GUIDE.md를 정본으로 참조한다.
  규칙을 바꿀 때는 PROJECT_GUIDE.md와 `2_스크립트/qa.py`를 동시에 수정한다.
- 작업 흐름: `2_스크립트/extract.py`로 원문 추출 → `fortune_memory.py harvest` → `fortune_memory.py draft MMDD`로
  `build_MMDD.py` 조립(빈칸은 `todo_MMDD.txt`만 작성) →
  `py -X utf8 qa.py MMDD`가 `RESULT: PASS`일 때까지 반복 → `build_MMDD.py` 실행 →
  `embed_export.ps1`로 글꼴 임베드 + PNG 내보내기. 순서를 지키지 않으면 글꼴 임베드가 빠질 수 있다
  (PROJECT_GUIDE.md §5 "함정" 참조).
