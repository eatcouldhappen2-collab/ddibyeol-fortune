# -*- coding: utf-8 -*-
"""Verify that, within each zodiac, no two summary lines share the same ending."""
import re
from build_0615 import SUMMARIES

# ordered longest-first so specific endings win over generic ones
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

ok = True
for zod, pairs in SUMMARIES.items():
    lines = [l for pair in pairs for l in pair]
    classes = [ending_class(l) for l in lines]
    dups = [c for c in set(classes) if classes.count(c) > 1]
    l1 = [ending_class(p[0]) for p in pairs]
    l2 = [ending_class(p[1]) for p in pairs]
    status = "OK " if not dups else "DUP"
    if dups:
        ok = False
    print(f"[{status}] {zod}: L1={l1} L2={l2}" + (f"  <-- repeated:{dups}" if dups else ""))

print("\nRESULT:", "all zodiac signs have 12 distinct line endings" if ok
      else "SOME zodiac signs still repeat an ending")
