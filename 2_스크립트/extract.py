# -*- coding: utf-8 -*-
"""월별 원본 엑셀에서 특정 날짜의 운세 원문을 추출한다 (build 스크립트 작성용).

사용법:
  py extract.py "../1_원본엑셀/7월 운세.xlsx" 0709 0710 0711 ...
또는 한 달 전체:
  py extract.py "../1_원본엑셀/7월 운세.xlsx" all

각 날짜마다 src_MMDD.txt 를 만든다. 파일 1행은 그대로 build 스크립트에 붙여넣을 수 있는
SCORES = {...} 이고, 그 아래에 12간지별 [총운설명] + slot1~slot6 원문이 있다.
(이 파일들은 build 스크립트를 작성한 뒤 지워도 된다.)
"""
import sys, io, datetime
import openpyxl

ZOD = ["쥐","소","호랑이","토끼","용","뱀","말","양","원숭이","닭","개","돼지"]

def main():
    if len(sys.argv) < 3:
        print(__doc__); return
    xlsx = sys.argv[1]
    wb = openpyxl.load_workbook(xlsx, data_only=True)
    ws = wb.active
    # map every date-block start row
    starts = {}
    for r in range(1, ws.max_row + 1):
        v = ws.cell(r, 1).value
        if isinstance(v, (int, float)):
            d = datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))
            starts[d] = r
    # which dates?
    args = sys.argv[2:]
    if args == ["all"]:
        dates = sorted(starts)
    else:
        dates = []
        for code in args:
            mm, dd = int(code[:2]), int(code[2:])
            yr = max(d.year for d in starts) if starts else 2026
            # pick the year present in the workbook for that month/day
            cand = [d for d in starts if d.month == mm and d.day == dd]
            if not cand:
                print(f"!! {code}: 해당 날짜를 엑셀에서 찾지 못함"); continue
            dates.append(cand[0])
    for d in dates:
        st = starts[d]
        code = f"{d.month:02d}{d.day:02d}"
        out = io.open(f"src_{code}.txt", "w", encoding="utf-8")
        sc = {}
        for ci, z in enumerate(ZOD):
            col = 2 + ci
            s = [ws.cell(st + 6, col).value, ws.cell(st + 7, col).value, ws.cell(st + 8, col).value]
            sc[z] = [int(str(x).split(":")[-1].strip()) for x in s]
        out.write("SCORES = {" + ", ".join('"%s": (%d,%d,%d)' % (z, *sc[z]) for z in ZOD) + "}\n")
        out.write(f"# {d.isoformat()}  (cover date string 예: \"{d.year}년 {d.month}월{d.day}일 ?요일\")\n")
        for ci, z in enumerate(ZOD):
            col = 2 + ci
            out.write("\n--- %s ---  (총운%d 재물%d 애정%d)\n" % (z, *sc[z]))
            out.write("[총운설명] %s\n" % (ws.cell(st + 4, col).value))
            for k in range(6):
                yr = ws.cell(st + 10 + 4 * k, col).value
                tx = ws.cell(st + 12 + 4 * k, col).value
                out.write("  slot%d %s: %s\n" % (k + 1, yr, tx))
        out.close()
        print(f"wrote src_{code}.txt  ({d.isoformat()}, startrow {st})")

if __name__ == "__main__":
    main()
