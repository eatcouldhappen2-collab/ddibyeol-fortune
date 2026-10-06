# -*- coding: utf-8 -*-
"""0511.pptx -> 0511_flat.pptx. 고정 요소는 슬라이드별 배경 PNG 1장으로 합치고
가변 텍스트(날짜·점수 3·운세 본문)만 텍스트상자로 남긴다.
usage: py -X utf8 make_flat_template.py strip   # 가변 텍스트 제거본 생성 (렌더 입력)
       (PowerPoint로 PNG 내보내기 후)
       py -X utf8 make_flat_template.py flat    # 배경 PNG + 가변 텍스트상자만 남긴 템플릿 생성
"""
import os, sys
from pptx import Presentation

BASE = os.path.dirname(os.path.abspath(__file__))
TPL = os.path.join(BASE, "..", "3_템플릿", "0511.pptx")
FLAT = os.path.join(BASE, "..", "3_템플릿", "0511_flat.pptx")
WORK = os.path.join(BASE, "..", "3_템플릿", "_flat_work")
STRIPPED = os.path.join(WORK, "stripped.pptx")

def dynamic(slide, idx):
    out = []
    for sh in slide.shapes:
        if not sh.has_text_frame:
            continue
        t = sh.text_frame.text.strip()
        if idx == 0:
            if "5월11일" in t:
                out.append(sh)
        elif (t.endswith("점") and len(t) <= 5) or (sh.height and sh.height > 4000000):
            out.append(sh)
    return out

def drop(sh):
    sh._element.getparent().remove(sh._element)

os.makedirs(WORK, exist_ok=True)
prs = Presentation(TPL)
mode = sys.argv[1]
if mode == "strip":
    for i, s in enumerate(prs.slides):
        for sh in dynamic(s, i):
            drop(sh)
    for sh in prs.slides[0].shapes:  # 배경에 굽는 자료 출처: 이 PC 설치명으로 교체 후 렌더
        if sh.has_text_frame and "자료=" in sh.text_frame.text:
            for t in sh._element.iter("{http://schemas.openxmlformats.org/drawingml/2006/main}latin", "{http://schemas.openxmlformats.org/drawingml/2006/main}ea"):
                if t.get("typeface") == "a옛날이발관L-Compix":
                    t.set("typeface", "a옛날이발관L")
                    t.attrib.pop("panose", None)
    prs.save(STRIPPED)
    print("saved", STRIPPED)
elif mode == "flat":
    from PIL import Image
    for i in range(len(prs.slides)):
        p = os.path.join(WORK, f"bg_{i+1:02d}")
        Image.open(p + ".png").convert("RGB").save(p + ".jpg", quality=95, subsampling=0)
    W, H = prs.slide_width, prs.slide_height
    for i, s in enumerate(prs.slides):
        keep = [sh._element for sh in dynamic(s, i)]
        for sh in list(s.shapes):
            if not any(sh._element is e for e in keep):
                drop(sh)
        pic = s.shapes.add_picture(os.path.join(WORK, f"bg_{i+1:02d}.jpg"), 0, 0, W, H)
        pic.name = "배경(고정요소 합본)"
        s.shapes._spTree.remove(pic._element)
        s.shapes._spTree.insert(2, pic._element)
    prs.save(FLAT)
    print("saved", FLAT)
