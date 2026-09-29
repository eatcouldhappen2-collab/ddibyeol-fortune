# -*- coding: utf-8 -*-
"""QA for daily-fortune summaries. Enforces, per date module build_MMDD:
  (HARD) every line fits ONE physical line in the slide textbox (no wrap),
         measured in pixels with the rendering metric (KBIZ Myeongjo M @ export scale);
  (VAR)  no readable vertical ending pattern across the 12 zodiacs per slot/line;
  (VAR)  no monotone ending within a zodiac;
  (MIN)  no over-compressed line (>= MIN_CHARS).
Usage: py qa.py 0709    (run from this folder; arg = MMDD code, e.g. 0709 = July 9)"""
import importlib, os, sys, re, io
from collections import Counter
from PIL import ImageFont

ZOD = ["쥐","소","호랑이","토끼","용","뱀","말","양","원숭이","닭","개","돼지"]

# ---- physical-line width model (mirrors PNG export) ----
# slide 7.5in -> 1500px (200 px/in); fortune font 20pt; textbox usable width 5.075in.
FONT_PX  = 20/72*200          # ~55.56 px
# Calibrated to ground truth: PowerPoint COM (real font a아시아명조M @20pt, 5.075in box).
# Re-measured 2026-09-11 by writing test strings into the live shape via COM and reading
# back the autofit-resized shape height: no-wrap confirmed up to 951px, wrap confirmed at
# 1001px (KBIZ-proxy width). Old 915px boundary was stale/over-conservative (documented
# false-positive overflow flags on 0907/0914/0831 that never actually wrapped in PowerPoint).
USABLE   = 950              # px wrap boundary (KBIZ proxy, ground-truth calibrated 2026-09-11)
WIDTH_LIMIT = 935          # px hard cap
_FONT_PATH = os.environ.get("FORTUNE_QA_FONT", "C:/Windows/Fonts/KBIZ한마음명조 M.ttf")
if not os.path.isfile(_FONT_PATH):
    sys.exit(f"폰트를 찾을 수 없습니다: {_FONT_PATH}\n"
             f"'KBIZ한마음명조 M' 폰트를 설치하거나, 환경변수 FORTUNE_QA_FONT로 폰트 경로를 지정하세요.")
_FONT = ImageFont.truetype(_FONT_PATH, int(round(FONT_PX)))
def width_px(line):
    return _FONT.getbbox(line)[2]

MIN_CHARS = 12

# ---- house style: quantity + unit written WITHOUT a space ----
# '한잔'(O) / '한 잔'(X), '두번'(O) / '두 번'(X)
_QTY  = "한두|두세|서너|너덧|다섯|여섯|일곱|여덟|아홉|한|두|세|네|열|몇"
_UNIT = ("잔|번|개|명|마리|병|그릇|끼|살|달|해|시간|가지|걸음|송이|모금|술|입|"
         "방울|자루|켤레|권|장|통|줄|판|알|척|대|칸|평|채|점|조각|봉지")
_QU = re.compile(rf"(?<![가-힣])({_QTY})\s+({_UNIT})(?=[을를이가은는에도와과로으의만\s\.\!\?\,]|$)")
def qty_unit_violations(line):
    return [m.group(0) for m in _QU.finditer(line)]

# ---- house style: auxiliary '보세요' spacing by preceding-word length ----
# 앞말 2글자 이하 -> 붙임('먹어보세요'); 3글자 이상 -> 띔('건드려 보세요').
# Only applies to auxiliary 보다 (앞말 = ~아/어/여 connective form, open syllable).
_CONN_MED = {0,1,3,4,5,6,9,10,14,15}   # ㅏㅐㅒㅓㅔㅕㅘㅙㅝㅞ
def _is_conn(ch):
    if not ("가" <= ch <= "힣"): return False
    b = ord(ch) - 0xAC00
    return b % 28 == 0 and (b // 28) % 21 in _CONN_MED   # open syllable + 아/어/여 vowel
# ---- forbidden awkward collocations (from editorial corrections) ----
BAD_PHRASES = {
    "결과를 맺": "결실을 맺",
    "지혜를 펴": "지혜를 발휘",
    "불안한 수": "불안한 운/기운",
    "걱정 만들": "걱정 키우",
    "걱정을 만들": "걱정을 키우",
    "정을 만들": "정을 쌓",
    "불안을 던지": "불안을 떨치",
    "불안은 던지": "불안은 떨치",
    "서둘 일": "서두를 일",
}
def phrase_violations(line):
    return [(b, g) for b, g in BAD_PHRASES.items() if b in line]

def boseyo_violations(line):
    out = []
    for m in re.finditer("보세요", line):
        i = m.start()
        if i == 0: continue
        before = line[:i]
        if before.endswith(" "):
            toks = before.split()
            if toks and _is_conn(toks[-1][-1]) and len(toks[-1]) <= 2:
                out.append(f"'{toks[-1]} 보세요' → 붙여쓰기 '{toks[-1]}보세요'")
        else:
            k, run = i - 1, ""
            while k >= 0 and "가" <= line[k] <= "힣":
                run = line[k] + run; k -= 1
            if run and _is_conn(run[-1]) and len(run) >= 3:
                out.append(f"'{run}보세요' → 띄어쓰기 '{run} 보세요'")
    return out

PATTERNS = [
    ("겠어요", r"겠어요$"), ("겠네요", r"겠네요$"),
    ("거예요", r"(거예요|거에요|꺼예요)$"),
    ("ㄹ까요", r"까요\?$"),
    ("세요",  r"세요$"),
    ("이에요", r"(이에요|에요|예요)$"),
    ("는데요", r"데요$"),
    ("군요",  r"군요$"), ("네요", r"네요$"), ("봐요", r"봐요$"),
    ("죠",    r"(죠|지요)$"),
    ("어요",  r"(아요|어요|여요|해요|와요|워요|러요|가요)$"),
    ("요",    r"요$"),
]
def ecls(line):
    s = re.sub(r"[\s\.\!]+$", "", line)
    for name, pat in PATTERNS:
        if re.search(pat, s):
            return name
    return "(?)"+s[-2:]
def clen(line):
    return len(re.sub(r"[\.\?\!]+$", "", line.strip()))

COL_MAX, ZOD_MAX = 4, 5   # 열(세로) ≤4, 간지 내부 ≤5 (불가피한 1회 반복 허용)

def line_problems(z, line):
    """Return the existing per-line QA messages without printing them."""
    problems = []
    if clen(line) < MIN_CHARS:
        problems.append(f"{z}: too short ({clen(line)}<{MIN_CHARS}): {line}")
    for v in qty_unit_violations(line):
        problems.append(f"{z}: 수량+단위 띄어쓰기 '{v}' → 붙여쓰기 '{v.replace(' ','')}': {line}")
    for v in boseyo_violations(line):
        problems.append(f"{z}: 보세요 띄어쓰기 {v}: {line}")
    if ecls(line) == "는데요":
        problems.append(f"{z}: ~데요 어미 금지(군요·네요·입니다 등으로 대체): {line}")
    for b, g in phrase_violations(line):
        problems.append(f"{z}: 어색한 표현 '{b}…' → '{g}…': {line}")
    if re.sub(r"[\s]+$", "", line).endswith("?"):
        problems.append(f"{z}: 물음표 종결 지양(평서문으로): {line}")
    return problems

def check(S):
    """Run the original full QA sequence and return (problems, overflow, worst)."""
    problems, overflow = [], []
    worst = 0
    for z in ZOD:
        for i,(a,b) in enumerate(S[z],1):
            for li,l in ((1,a),(2,b)):
                w = width_px(l)
                worst = max(worst, w)
                if w > WIDTH_LIMIT:
                    overflow.append((z,i,li,w,l))
    if overflow:
        problems.append(f"{len(overflow)} line(s) exceed {WIDTH_LIMIT}px (would wrap)")

    for slot in range(6):
        for li in range(2):
            cnt = Counter(ecls(S[z][slot][li]) for z in ZOD)
            top,n = cnt.most_common(1)[0]
            if n > COL_MAX:
                problems.append(f"slot{slot+1} L{li+1}: ending '{top}' x{n} (>{COL_MAX}) across zodiacs")
    for z in ZOD:
        lines=[l for p in S[z] for l in p]
        cnt=Counter(ecls(l) for l in lines)
        top,n=cnt.most_common(1)[0]
        if n>ZOD_MAX:
            problems.append(f"{z}: ending '{top}' x{n} within zodiac (>{ZOD_MAX})")
        for l in lines:
            problems.extend(line_problems(z, l))
    return problems, overflow, worst

def main():
    """Preserve the legacy command-line report while keeping imports side-effect free."""
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    date = sys.argv[1]
    # importing a build module triggers prs.save(); QA only needs SUMMARIES -> no-op save.
    import pptx.presentation
    pptx.presentation.Presentation.save = lambda *a, **k: None
    mod = importlib.import_module(f"build_{date}")
    S = mod.SUMMARIES
    problems, overflow, worst = check(S)

    print(f"=== width check (limit {WIDTH_LIMIT}px, boundary {USABLE}px) ===")
    print(f" widest line = {worst}px ; overflow count = {len(overflow)}")
    for z,i,li,w,l in overflow:
        print(f"   OVERFLOW {z} slot{i} L{li} {w}px: {l}")
    print("\nRESULT:", "PASS" if not problems else f"FAIL - {len(problems)} issue(s)")
    for p in problems:
        print("  -", p)

if __name__ == "__main__":
    main()
