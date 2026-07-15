# -*- coding: utf-8 -*-
"""Verify 0622 summaries: (1) 12 distinct endings per zodiac,
   (2) within each pair the two lines differ by <= 3 chars."""
import io, re, sys
from build_0622 import SUMMARIES

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

PATTERNS = [
    ("겠네요", r"겠네요$"), ("겠어요", r"겠어요$"), ("거예요", r"거예요$|꺼예요$"),
    ("ㄹ까요", r"까요\?$"), ("이에요", r"이에요$|예요$"), ("는데요", r"데요$"),
    ("군요", r"군요$"), ("네요", r"네요$"), ("세요", r"세요$"),
    ("봐요", r"봐요$"), ("죠", r"죠$|지요$"), ("요", r"요$"),
]

def ending_class(line):
    s = re.sub(r"[\s\.\!]+$", "", line)   # keep '?' so ~까요? is detected
    for name, pat in PATTERNS:
        if re.search(pat, s):
            return name
    return "(?)" + s[-2:]

# char length: full string minus the single trailing . or ? (spaces kept)
def length(line):
    return len(re.sub(r"[\.\?]+$", "", line))

ok = True
for zod, pairs in SUMMARIES.items():
    lines = [l for pair in pairs for l in pair]
    classes = [ending_class(l) for l in lines]
    dups = [c for c in set(classes) if classes.count(c) > 1]
    bad_len = []
    for i, (l1, l2) in enumerate(pairs):
        d = abs(length(l1) - length(l2))
        if d > 3:
            bad_len.append((i + 1, length(l1), length(l2), d))
    status = "OK " if not dups and not bad_len else "BAD"
    if dups or bad_len:
        ok = False
    print(f"[{status}] {zod}")
    if dups:
        print(f"      repeated endings: {dups}")
    if bad_len:
        for slot, a, b, d in bad_len:
            print(f"      slot{slot} len diff {d} ({a} vs {b})")

# also show max diff overall
maxd = max(abs(length(a) - length(b)) for p in SUMMARIES.values() for a, b in p)
print(f"\nmax pair length diff across all = {maxd}")
print("RESULT:", "all constraints satisfied" if ok else "CONSTRAINTS VIOLATED")
