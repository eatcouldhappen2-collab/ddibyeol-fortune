---
name: ddibyeol-fortune
description: 신문 「오늘의 운세」 띠별 슬라이드(표지+12간지 13장) 파이프라인 작업 스킬 — 새 날짜(MMDD) 추가, 원문 추출, 운세 문장(SUMMARIES) 작성과 qa.py 검증·FAIL 수정, build 스크립트 작성·pptx 생성, 글꼴 임베드·PNG 내보내기, 확정 문장 메모리 갱신에 사용한다. 날짜 코드와 단계 이름만 온 요청("0805 운세 변환", "0805 이미지 추출")도 포함.
---

# 띠별운세 작업 스킬

이 스킬은 워크플로우 요약이다. **규칙 전문·환경 요구사항·완료 체크리스트는 항상
[PROJECT_GUIDE.md](../../../PROJECT_GUIDE.md)를 정본으로 참조한다** — 이 파일에는 규칙을 복제하지 않는다.

## 3단계로 나눠 실행하기

전체 7단계를 한 번에 하지 않고 **"날짜 코드 + 단계 이름"만 말해서 한 단계씩** 시킬 수 있다. 각 단계는
전용 슬래시 명령으로도 존재한다(예: `/운세변환 0805`) — 어느 쪽으로 요청하든 그 단계만 하고 멈춘다.

| 사용자가 말하는 단계 | 해당 명령 | 아래 7단계 중 | 하는 일 | 안 하는 일 |
| --- | --- | --- | --- | --- |
| 운세 변환 | `/운세변환` | 2~4단계 | 원문 추출 → SUMMARIES 작성 → qa PASS | pptx 생성, 이미지 추출 |
| 운세 삽입 | `/운세삽입` | 5단계 | `build_MMDD.py` 실행 → pptx 저장 | 텍스트 작성/qa, 이미지 추출 |
| 이미지 추출 | `/이미지추출` | 6단계 | 글꼴 임베드 + PNG 내보내기 | 텍스트 작성/qa, pptx 생성 |

세 단계를 순서대로 다 거쳐야 배포 가능한 최종 산출물이 된다(중간에 멈춰도 됨 — 다음에 이어서 해당
단계 명령만 다시 부르면 된다). 각 명령의 사전 조건·범위·완료 확인 방법은 명령 파일 자체에 있다.

## 워크플로우 7단계 (요약)

`2_스크립트/` 폴더에서 `py -X utf8`로 실행한다.

1. **원본 엑셀 넣기** — `N월 운세.xlsx`를 `1_원본엑셀/`에 복사.
2. **원문 추출** — `py -X utf8 extract.py "../1_원본엑셀/7월 운세.xlsx" 0709` → `src_MMDD.txt` 생성.
3. **빌드 스크립트 조립** — `py -X utf8 fortune_memory.py harvest` → `py -X utf8 fortune_memory.py draft 0709`.
   사용자가 확정한 문장 메모리로 72칸 대부분이 채워진다. `todo_MMDD.txt`의 빈칸과 qa 지적 줄만
   `fortune-writer` 서브에이전트가 채운다(PROJECT_GUIDE.md §4-3).
4. **검증** — `py -X utf8 qa.py 0709` → `RESULT: PASS`까지 SUMMARIES 반복 수정.
5. **PPTX 생성** — `py -X utf8 build_0709.py` → `4_PPTX/0709.pptx` (아직 글꼴 미임베드).
6. **글꼴 임베드 + PNG 내보내기** — `powershell -ExecutionPolicy Bypass -File embed_export.ps1 0709`.
   **build를 다시 돌렸다면 반드시 이 단계도 다시 실행한다** — 안 하면 글꼴 임베드가 조용히 빠진다.
7. **(선택) 배포 zip** — `4_PPTX`·`5_PNG`를 묶어 `6_ZIP/`에 저장.

완료 체크리스트(qa PASS → build 실행 → embed_export 재실행 → `presentation.xml`의
`embedTrueTypeFonts="1"` 확인)는 PROJECT_GUIDE.md §5를 그대로 따른다.

## 자주 나오는 문장 문제 패턴

`qa.py`가 잡지 못하는 반복 문체 문제(화자 시점·의문형·구어체·소재 중복)는 PROJECT_GUIDE.md §6-13에
있다 — SUMMARIES 작성·리뷰 시 직접 검토한다.

## 관련 파일

- 정본: `PROJECT_GUIDE.md`
- 검증기: `2_스크립트/qa.py`
- 추출기: `2_스크립트/extract.py`
- 임베드/내보내기: `2_스크립트/embed_export.ps1`
- SUMMARIES 작성 서브에이전트: `.claude/agents/fortune-writer.md`
- 단계별 명령: `.claude/commands/운세변환.md`, `.claude/commands/운세삽입.md`, `.claude/commands/이미지추출.md`
