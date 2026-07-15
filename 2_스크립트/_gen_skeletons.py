# -*- coding: utf-8 -*-
import io, re
DATES = [
 ("0701",1,"수요일"),("0702",2,"목요일"),("0703",3,"금요일"),
 ("0706",6,"월요일"),("0707",7,"화요일"),("0708",8,"수요일"),("0709",9,"목요일"),("0710",10,"금요일"),
 ("0713",13,"월요일"),("0714",14,"화요일"),("0715",15,"수요일"),("0716",16,"목요일"),("0717",17,"금요일"),
 ("0720",20,"월요일"),("0721",21,"화요일"),("0722",22,"수요일"),("0723",23,"목요일"),("0724",24,"금요일"),
 ("0727",27,"월요일"),("0728",28,"화요일"),("0729",29,"수요일"),("0730",30,"목요일"),("0731",31,"금요일"),
]
base = io.open("build_0629.py", encoding="utf-8").read()
ZOD = ["쥐","소","호랑이","토끼","용","뱀","말","양","원숭이","닭","개","돼지"]
blank = "SUMMARIES = {\n" + "".join(' "%s": [("",""),("",""),("",""),("",""),("",""),("","")],\n' % z for z in ZOD) + "}"
# locate existing SUMMARIES block (from 'SUMMARIES = {' to the line '}' just before '# Slide order')
pat = re.compile(r"SUMMARIES = \{.*?\n\}", re.S)
for code,d,wd in DATES:
    src1 = io.open(f"src_{code}.txt", encoding="utf-8").readline().strip()  # SCORES = {...}
    t = base
    t = pat.sub(blank, t, count=1)
    t = re.sub(r"^SCORES = \{.*\}$", src1.replace("\\","\\\\"), t, count=1, flags=re.M)
    t = t.replace("2026년 6월29일 월요일", f"2026년 7월{d}일 {wd}")
    t = t.replace("0629.pptx", f"{code}.pptx")
    t = t.replace("June-29", f"July-{d:02d}")
    io.open(f"build_{code}.py", "w", encoding="utf-8").write(t)
    print(f"build_{code}.py  ({wd}, SCORES set, SUMMARIES blanked)")
