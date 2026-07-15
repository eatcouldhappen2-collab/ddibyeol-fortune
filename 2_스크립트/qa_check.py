# -*- coding: utf-8 -*-
import importlib, sys, re
from collections import Counter
ZOD=["쥐","소","호랑이","토끼","용","뱀","말","양","원숭이","닭","개","돼지"]
MIN_CHARS=12
_QTY="한두|두세|서너|너덧|다섯|여섯|일곱|여덟|아홉|한|두|세|네|열|몇"
_UNIT=("잔|번|개|명|마리|병|그릇|끼|살|달|해|시간|가지|걸음|송이|모금|술|입|"
       "방울|자루|켤레|권|장|통|줄|판|알|척|대|칸|평|채|점|조각|봉지")
_QU=re.compile(rf"(?<![가-힣])({_QTY})\s+({_UNIT})(?=[을를이가은는에도와과로으의만\s\.\!\?\,]|$)")
def qty_unit_violations(line): return [m.group(0) for m in _QU.finditer(line)]
_CONN_MED={0,1,3,4,5,6,9,10,14,15}
def _is_conn(ch):
    if not ("가"<=ch<="힣"): return False
    b=ord(ch)-0xAC00
    return b%28==0 and (b//28)%21 in _CONN_MED
def boseyo_violations(line):
    out=[]
    for m in re.finditer("보세요",line):
        i=m.start()
        if i==0: continue
        before=line[:i]
        if before.endswith(" "):
            toks=before.split()
            if toks and _is_conn(toks[-1][-1]) and len(toks[-1])<=2:
                out.append(f"'{toks[-1]} 보세요'→'{toks[-1]}보세요'")
        else:
            k,run=i-1,""
            while k>=0 and "가"<=line[k]<="힣":
                run=line[k]+run; k-=1
            if run and _is_conn(run[-1]) and len(run)>=3:
                out.append(f"'{run}보세요'→'{run} 보세요'")
    return out
PATTERNS=[("겠어요",r"겠어요$"),("겠네요",r"겠네요$"),("거예요",r"(거예요|거에요|꺼예요)$"),("ㄹ까요",r"까요\?$"),("세요",r"세요$"),("이에요",r"(이에요|에요|예요)$"),("는데요",r"데요$"),("군요",r"군요$"),("네요",r"네요$"),("봐요",r"봐요$"),("죠",r"(죠|지요)$"),("어요",r"(아요|어요|여요|해요|와요|워요|러요|가요)$"),("요",r"요$")]
def ecls(line):
    s=re.sub(r"[\s\.\!]+$","",line)
    for n,p in PATTERNS:
        if re.search(p,s): return n
    return "(?)"+s[-2:]
def clen(line): return len(re.sub(r"[\.\?\!]+$","",line.strip()))
import pptx.presentation
pptx.presentation.Presentation.save=lambda *a,**k:None
S=importlib.import_module(f"build_{sys.argv[1]}").SUMMARIES
problems=[]
for slot in range(6):
    for li in range(2):
        c=Counter(ecls(S[z][slot][li]) for z in ZOD); t,n=c.most_common(1)[0]
        if n>4: problems.append(f"slot{slot+1} L{li+1}: '{t}' x{n}")
for z in ZOD:
    lines=[l for p in S[z] for l in p]
    c=Counter(ecls(l) for l in lines); t,n=c.most_common(1)[0]
    if n>4: problems.append(f"{z}: ending '{t}' x{n}")
    for l in lines:
        if clen(l)<MIN_CHARS: problems.append(f"{z}: short: {l}")
        for v in qty_unit_violations(l): problems.append(f"{z}: 수량단위 {v}: {l}")
        for v in boseyo_violations(l): problems.append(f"{z}: 보세요 {v}: {l}")
print("RESULT:","PASS" if not problems else f"FAIL-{len(problems)}")
for p in problems: print("  -",p)
# false-positive / true-positive tests for 보세요 rule
print("--- boseyo tests ---")
tests={"먹어 보세요":["'먹어 보세요'→'먹어보세요'"],"건드려보세요":["'건드려보세요'→'건드려 보세요'"],
 "잡아 보세요":["'잡아 보세요'→'잡아보세요'"],"의논해 보세요":[],"즐겨보세요":[],
 "여기 보세요":[], "천천히 보세요":[], "잘 보세요":[], "하늘을 보세요":[], "다시 보세요":[]}
fail=0
for s,exp in tests.items():
    got=boseyo_violations(s); ok=got==exp; fail+=0 if ok else 1
    print(("OK " if ok else "XX "),repr(s),"->",got)
print("test failures:",fail)
