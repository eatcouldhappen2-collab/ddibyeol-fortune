# -*- coding: utf-8 -*-
"""각 build_MMDD.py에서 PPT 텍스트 값(표지 날짜 + 띠별 점수 + 6칸 2줄)을 뽑아 요일별 txt로 묶는다.
   - 문장 앞 번호·들여쓰기 없음(왼쪽 정렬)
   - 점수는 총운·재물·애정 모두 표기 (운세 문구는 연령대 6칸만 — 총운 문구 없음)
   - 2줄(슬롯)마다 빈 줄 한 번
   출력: ../7월_PPT텍스트/{요일}.txt"""
import io, os, re, ast

WEEKDAYS = {
 "월요일": ["0706","0713","0720","0727"],
 "화요일": ["0707","0714","0721","0728"],
 "수요일": ["0701","0708","0715","0722","0729"],
 "목요일": ["0702","0709","0716","0723","0730"],
 "금요일": ["0703","0710","0717","0724","0731"],
}
ZOD = ["쥐","소","호랑이","토끼","용","뱀","말","양","원숭이","닭","개","돼지"]
OUT = "../7월_PPT텍스트"
os.makedirs(OUT, exist_ok=True)

def grab(code):
    txt = io.open(f"build_{code}.py", encoding="utf-8").read()
    date = re.search(r'"(2026년 7월\d+일 [가-힣]+요일)"', txt).group(1)
    scores = ast.literal_eval(re.search(r'^SCORES = (\{.*\})\s*$', txt, re.M).group(1))
    summ = ast.literal_eval(re.search(r'^SUMMARIES = (\{.*?\n\})', txt, re.S | re.M).group(1))
    return date, scores, summ

for wd, codes in WEEKDAYS.items():
    out = io.open(f"{OUT}/{wd}.txt", "w", encoding="utf-8")
    out.write(f"[2026년 7월 · {wd}] 띠별 오늘의 운세 — PPT 텍스트 값\n")
    for code in codes:
        date, sc, su = grab(code)
        out.write("\n" + "=" * 46 + f"\n  {date}\n" + "=" * 46 + "\n")
        for z in ZOD:
            c, j, a = sc[z]
            out.write(f"\n[{z}]  총운 {c} · 재물 {j} · 애정 {a}\n")
            for l1, l2 in su[z]:
                out.write(f"{l1}\n{l2}\n\n")
    out.close()
    print(f"wrote {wd}.txt  ({len(codes)}일)")
