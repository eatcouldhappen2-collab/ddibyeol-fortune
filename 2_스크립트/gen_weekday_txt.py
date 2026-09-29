# -*- coding: utf-8 -*-
"""build_MMDD.py의 PPT 텍스트 값(표지 날짜 + 띠별 점수 + 6칸 2줄)을 요일별 txt로 묶는다.

사용법:
  py -X utf8 gen_weekday_txt.py 10                     -> ../10월_PPT텍스트/{요일}.txt
  py -X utf8 gen_weekday_txt.py 10 "C:/Users/.../Desktop/10월_띠별운세_텍스트"

- 해당 월의 build_MM*.py 가 있는 날짜를 요일(월~금)별로 묶는다. 출력 폴더가 이미 있으면 덮어쓰지 않고 멈춘다.
- 문장 앞 번호·들여쓰기 없음, 점수는 총운·재물·애정 모두 표기, 2줄(슬롯)마다 빈 줄 한 번.
"""
import datetime
import glob
import importlib
import io
import os
import sys

BASE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, BASE)
ZOD = ["쥐", "소", "호랑이", "토끼", "용", "뱀", "말", "양", "원숭이", "닭", "개", "돼지"]
NAMES = "월화수목금토일"


def main():
    if len(sys.argv) < 2 or not sys.argv[1].isdigit():
        print(__doc__)
        return
    month = int(sys.argv[1])
    out_dir = sys.argv[2] if len(sys.argv) > 2 else os.path.join(BASE, "..", f"{month}월_PPT텍스트")
    codes = sorted(os.path.basename(p)[6:10] for p in glob.glob(os.path.join(BASE, f"build_{month:02d}[0-3][0-9].py")))
    if not codes:
        sys.exit(f"build_{month:02d}MM.py 가 없습니다")
    year = datetime.date.today().year
    groups = {}
    for code in codes:
        weekday = NAMES[datetime.date(year, month, int(code[2:])).weekday()] + "요일"
        groups.setdefault(weekday, []).append(code)
    os.makedirs(out_dir, exist_ok=False)
    for weekday, day_codes in groups.items():
        with io.open(os.path.join(out_dir, f"{weekday}.txt"), "w", encoding="utf-8") as out:
            out.write(f"[{year}년 {month}월 · {weekday}] 띠별 오늘의 운세 — PPT 텍스트 값\n")
            for code in day_codes:
                mod = importlib.import_module("build_" + code)
                out.write("\n" + "=" * 46 + f"\n  {mod.COVER}\n" + "=" * 46 + "\n")
                for zodiac in ZOD:
                    total, money, love = mod.SCORES[zodiac]
                    out.write(f"\n[{zodiac}]  총운 {total} · 재물 {money} · 애정 {love}\n")
                    for line1, line2 in mod.SUMMARIES[zodiac]:
                        if not (line1 and line2):
                            sys.exit(f"빈칸 남음: {code} {zodiac}")
                        out.write(f"{line1}\n{line2}\n\n")
        print(f"wrote {weekday}.txt  ({len(day_codes)}일)")


if __name__ == "__main__":
    main()
