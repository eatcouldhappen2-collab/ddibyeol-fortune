---
name: ddibyeol-fortune
description: 신문 「오늘의 운세」 코너 띠별 슬라이드(13장) 작업 스킬. "운세 써줘", "운세 슬라이드 만들어줘", "build 스크립트 작성해줘", "qa 통과 안 됨", "qa.py FAIL", "MMDD 만들어줘", "새 날짜 추가", "글꼴 임베드", "PNG 내보내기" 같은 요청에 사용한다.
---

# 띠별운세 작업 스킬

이 스킬은 워크플로우 요약이다. **규칙 전문·환경 요구사항·완료 체크리스트는 항상
[PROJECT_GUIDE.md](../../../PROJECT_GUIDE.md)를 정본으로 참조한다** — 이 파일에는 규칙을 복제하지 않는다.

## 워크플로우 7단계 (요약)

`2_스크립트/` 폴더에서 `py -X utf8`로 실행한다.

1. **원본 엑셀 넣기** — `N월 운세.xlsx`를 `1_원본엑셀/`에 복사.
2. **원문 추출** — `py -X utf8 extract.py "../1_원본엑셀/7월 운세.xlsx" 0709` → `src_MMDD.txt` 생성.
3. **빌드 스크립트 작성** — 최근 `build_MMDD.py`를 복제해 표지 날짜·`SCORES`·`SUMMARIES`·저장 파일명만 변경.
   SUMMARIES 작성은 `fortune-writer` 서브에이전트에 위임 가능.
4. **검증** — `py -X utf8 qa.py 0709` → `RESULT: PASS`까지 SUMMARIES 반복 수정.
5. **PPTX 생성** — `py -X utf8 build_0709.py` → `4_PPTX/0709.pptx` (아직 글꼴 미임베드).
6. **글꼴 임베드 + PNG 내보내기** — `powershell -ExecutionPolicy Bypass -File embed_export.ps1 0709`.
   **build를 다시 돌렸다면 반드시 이 단계도 다시 실행한다** — 안 하면 글꼴 임베드가 조용히 빠진다.
7. **(선택) 배포 zip** — `4_PPTX`·`5_PNG`를 묶어 `6_ZIP/`에 저장.

완료 체크리스트(qa PASS → build 실행 → embed_export 재실행 → `presentation.xml`의
`embedTrueTypeFonts="1"` 확인)는 PROJECT_GUIDE.md §5를 그대로 따른다.

## 자주 나오는 문장 문제 패턴

SUMMARIES(2줄 요약) 작성·리뷰 시 아래 네 가지가 반복적으로 나온다. `qa.py`가 잡지 못하는 문체 문제이므로
직접 검토한다.

| 패턴 | 문제 | 나쁜 예 | 좋은 예 |
| --- | --- | --- | --- |
| 화자 시점 어긋남 | 운세가 예언인데 남 얘기하듯 추측형으로 끝맺어 화자 위치가 흔들림 | "노력의 대가가 따르니 최선을 다하겠죠." | "최선을 다해야겠어요." |
| 의문형 종결 | `~까요?`로 끝내 예언의 단정성이 흔들림(qa.py도 물음표 종결은 강제 금지) | "합격과 승진으로 보람을 느끼게 될까요?" | "합격과 승진 소식이 반갑겠어요." |
| 구어체 지나침 | `~거든요` 등 구어투가 지면 문체와 어울리지 않음 | "오늘은 상대에게 져 주는 게 답이거든요." | "상대에게 양보하는 게 현명해요." |
| 소재 중복 | 같은 간지 안에서 같은 소재(예: 여행)가 여러 슬롯에 반복돼 단조로움 | slot1·slot3 모두 '여행' 언급 | slot3 소재를 가족·방향 등 다른 화제로 교체 |

## 관련 파일

- 정본: `PROJECT_GUIDE.md`
- 검증기: `2_스크립트/qa.py`
- 추출기: `2_스크립트/extract.py`
- 임베드/내보내기: `2_스크립트/embed_export.ps1`
- SUMMARIES 작성 서브에이전트: `.claude/agents/fortune-writer.md`
