# -*- coding: utf-8 -*-
"""Fill the shared 0511 PowerPoint template.

Use ``build(code, cover, scores, summaries)`` from a date module's main guard.
"""
import os
from pptx import Presentation

_BASE = os.path.dirname(os.path.abspath(__file__))
_TPL = os.path.join(_BASE, "..", "3_템플릿", "0511_flat.pptx")
_DEFAULT_OUT = os.path.join(_BASE, "..", "4_PPTX")

SLIDE_ZODIAC = {
    2:"쥐", 3:"소", 4:"호랑이", 5:"토끼", 6:"용", 7:"뱀",
    8:"말", 9:"양", 10:"원숭이", 11:"닭", 12:"개", 13:"돼지",
}

def set_run_only(para, text):
    """Set paragraph text into its first run, drop the rest (keeps font)."""
    runs = para.runs
    if not runs:
        para.text = text
        return
    runs[0].text = text
    for r in runs[1:]:
        r._r.getparent().remove(r._r)

def set_score(box, val):
    """Set a 'NN점' box to f'{val}점' robustly, regardless of run count."""
    para = box.text_frame.paragraphs[0]
    runs = para.runs
    para.runs[0].text = str(val)
    jeom_idx = next((i for i, r in enumerate(runs) if "점" in r.text), None)
    if jeom_idx is None:
        runs[0].text = f"{val}점"
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        runs[jeom_idx].text = "점"
        for i, r in enumerate(runs):
            if i not in (0, jeom_idx):
                r._r.getparent().remove(r._r)

def fortune_slots(tf):
    """Group non-empty paragraphs separated by blank ones -> list of [paras]."""
    groups, cur = [], []
    for para in tf.paragraphs:
        if para.text.strip() == "":
            if cur:
                groups.append(cur); cur = []
        else:
            cur.append(para)
    if cur:
        groups.append(cur)
    return groups

def build(code, cover, scores, summaries, out_dir=None):
    """Apply one day's values to the template and save ``code.pptx``."""
    out_dir = _DEFAULT_OUT if out_dir is None else os.fspath(out_dir)
    prs = Presentation(_TPL)
    warnings = []

    for sh in prs.slides[0].shapes:
        if sh.has_text_frame and "5월11일" in sh.text_frame.text:
            set_run_only(sh.text_frame.paragraphs[0], cover)
    for sidx, zod in SLIDE_ZODIAC.items():
        slide = prs.slides[sidx-1]
        chong, jaemul, aejeong = scores[zod]

        score_boxes = [sh for sh in slide.shapes
                       if sh.has_text_frame and sh.text_frame.text.strip().endswith("점")
                       and len(sh.text_frame.text.strip()) <= 5]
        score_boxes.sort(key=lambda s: s.left)
        if len(score_boxes) != 3:
            warnings.append(f"slide {sidx} {zod}: found {len(score_boxes)} score boxes")
        for box, val in zip(score_boxes, (chong, jaemul, aejeong)):
            set_score(box, val)

        big = [sh for sh in slide.shapes
               if sh.has_text_frame and sh.height and sh.height > 4000000]
        big = max(big, key=lambda s: s.height)
        slots = fortune_slots(big.text_frame)
        if len(slots) != 6:
            warnings.append(f"slide {sidx} {zod}: {len(slots)} fortune slots (expected 6)")
        for slot, (l1, l2) in zip(slots, summaries[zod]):
            if len(slot) >= 2:
                set_run_only(slot[0], l1)
                set_run_only(slot[1], l2)
                for extra in slot[2:]:
                    set_run_only(extra, "")
            elif len(slot) == 1:
                set_run_only(slot[0], l1 + " " + l2)

    prs.save(os.path.join(out_dir, f"{code}.pptx"))
    print(f"Saved {code}.pptx")
    if warnings:
        print("WARNINGS:")
        for w in warnings: print(" -", w)
    else:
        print("No warnings; all 12 slides updated.")
