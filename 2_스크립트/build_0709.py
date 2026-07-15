# -*- coding: utf-8 -*-
"""Fill the 0511.pptx template with summarized July-09 fortunes -> 0709.pptx."""
from pptx import Presentation
import os
_BASE = os.path.dirname(os.path.abspath(__file__))
_TPL = os.path.join(_BASE, "..", "3_템플릿")
_OUT = os.path.join(_BASE, "..", "4_PPTX")

# ---- Scores per zodiac (총운, 재물, 애정) read from 6월 운세.xlsx, June 29 ----
SCORES = {"쥐": (73,72,81), "소": (97,97,87), "호랑이": (88,96,85), "토끼": (70,79,82), "용": (69,79,91), "뱀": (85,77,92), "말": (90,89,93), "양": (91,94,92), "원숭이": (88,82,91), "닭": (83,82,86), "개": (81,84,83), "돼지": (73,91,87)}

# ---- 2-line summaries, age order oldest->youngest (matches badges) ----
SUMMARIES = {
 "쥐": [
  ("오늘은 외출을 접고 집으로 가야겠군요","따뜻한 물로 씻으며 피로를 풀어보세요"),
  ("금전운이 활짝 열려 기분 좋은 날이에요","계획 있게 움직여 좋은 기회를 잡아봐요"),
  ("복잡한 마음을 정리하고 긴장을 푸세요","다툼이 이별로 번지지 않게 조심하세요"),
  ("흐름이 하락세를 타고 있는 형국이군요","새로운 만남과 시작에 복이 깃들었죠"),
  ("가족의 건강과 마음을 살뜰히 돌보세요","어려움을 나누면 격려와 위안을 받아요"),
  ("뜻밖의 돈이나 선물이 생길지도 모르죠","인기가 높으니 반가운 고백도 기대해요"),
 ],
 "소": [
  ("뜻하지 않은 지출이 생길 수 있겠네요","외출은 줄이고 집에서 마음을 정리해요"),
  ("이별을 고민하는 마음이 참 무겁겠군요","성급한 결정은 서로를 더 힘들게 만들죠"),
  ("걱정과 달리 아픈 곳 없이 건강하겠네요","무엇보다 건강이 우선임을 기억해야죠"),
  ("모든 상황이 불리하게 돌아가고 있네요","조언해 줄 사람을 서둘러 찾아야 해요"),
  ("갑작스러운 지출로 살림이 빠듯해져요","돈을 효율적으로 굴리도록 애써야겠어요"),
  ("아무리 애써도 이기기 힘든 게임이군요","한눈팔지 말고 나의 본업에 충실하세요"),
 ],
 "호랑이": [
  ("원하는 대로 일이 술술 풀리는 하루예요","가벼운 마음으로 나들이를 즐겨보세요"),
  ("가벼운 여행을 떠나기에 좋은 운이군요","무거운 짐은 줄이고 가볍게 다녀와요"),
  ("믿었던 사람이 등을 돌리는 날이죠","정신을 바짝 차리고 신중해야 하겠어요"),
  ("부부 사이에 언짢은 다툼이 생기겠네요","상대가 원하는 것을 차분히 들어봐요"),
  ("급하게 서두르면 손해가 크게 불어나요","혼자보다 지혜로운 조언자가 필요해요"),
  ("오늘은 애정운이 살짝 저조한 하루군요","만남은 미루고 나를 위한 시간을 보내요"),
 ],
 "토끼": [
  ("자녀의 일로 근심이 생길 수 있겠군요","부드러운 말로 상대의 마음을 다독여요"),
  ("걱정은 뒤로 미루고 밀고 나가도 좋아요","망설이지 말고 적극적으로 나아가봐요"),
  ("금전운이 하락세로 기우는 형국이군요","포기하지 말고 동쪽에서 풀어나가봐요"),
  ("서두르지만 않으면 목표를 이루겠네요","조급한 마음을 버리고 차근차근 가봐요"),
  ("연인과 이별할 징조가 보이는 날이네요","인연이 아니라 싶으면 빨리 정리하세요"),
  ("칭찬과 좋은 소식에 어깨가 으쓱해져요","매력이 넘치는 날이니 바삐 움직여봐요"),
 ],
 "용": [
  ("오늘은 차량 사고가 염려되는 날이에요","되도록 차보다 대중교통을 이용하세요"),
  ("오래 묵은 병이 좋아질 기미가 보여요","마음을 다독일 가벼운 여행도 좋겠어요"),
  ("서둘러 손쓰지 않으면 곤란해지겠어요","막연한 기대를 버리고 두루 점검해봐요"),
  ("힘들던 일에 해결의 실마리가 보이네요","반가운 손님과 계약이 이익을 안겨줘요"),
  ("현재에 충실하면 원하는 것을 얻겠네요","일을 미루지 말고 오늘 안에 끝내야죠"),
  ("전반적으로 기운이 가라앉는 날이군요","가족과 친구들 속에서 기쁨을 찾으세요"),
 ],
 "뱀": [
  ("순풍에 돛 단 듯이 순탄한 하루이군요","생각만 해오던 일을 이뤄도 좋은 날이죠"),
  ("진행하던 일이 드디어 인정을 받겠어요","주저하지 말고 적극적으로 나서보세요"),
  ("여행이나 외출이 술술 풀리는 날이에요","동료들과 단합 여행을 떠나도 즐거워요"),
  ("가장 왕성하게 움직여 볼 만한 시기예요","머뭇대지 말고 추진력 있게 밀고 가세요"),
  ("애정운이 최상인 하루가 되겠어요","모임에 나가서 새로운 인연을 만나봐요"),
  ("어떤 위기도 거뜬히 헤쳐가는 능력자죠","나만의 시간에 멋진 계획을 세워보세요"),
 ],
 "말": [
  ("계획하던 일에 문제가 생길 수 있어요","느긋하게 동쪽에서 올 귀인을 기다려요"),
  ("언 땅에 눈까지 내리는 고단한 하루군요","마음을 비우고 가볍게 하루를 보내봐요"),
  ("모처럼 찾아온 기회를 놓치면 아쉬워요","늘 하던 대로 소신 있게 밀어붙여 보세요"),
  ("오늘은 화해와 평화가 깃드는 하루예요","중요한 결정도 스스로 밀고 나가봐요"),
  ("만족할 만한 성과가 기다리고 있겠네요","바삐 움직여 투자와 매매를 시도해봐요"),
  ("특별한 인연이 찾아오는 날이 되겠네요","자신감을 갖고 대인관계를 이끌어가세요"),
 ],
 "양": [
  ("가족을 향한 배려가 어긋날 수도 있어요","각자의 시간을 갖고 부드럽게 대화해요"),
  ("외출은 가벼운 산책 정도가 알맞겠네요","차보다는 도보로 다니면서 조심하세요"),
  ("아무리 애써도 성과가 더딘 하루이군요","충분한 휴식으로 에너지를 채워봐요"),
  ("무리한 결정은 후회를 부를 수 있겠어요","가족과 상의하고 조언을 귀담아들어요"),
  ("애정 전선에 작은 문제가 생기겠군요","급한 결정과 약속은 잠시 뒤로 미루세요"),
  ("멀리서 반가운 친구와 연락이 찾아와요","스트레스를 풀 수 있는 여유를 가져봐요"),
 ],
 "원숭이": [
  ("사랑의 기운이 포근히 싹트는 날이에요","소중한 가족의 건강을 함께 살펴보세요"),
  ("마음의 안정을 찾는 일이 가장 중요해요","독서나 목욕으로 나른하게 쉬어보세요"),
  ("불안한 마음을 지우고 즐기면 복이 와요","계산을 접고 사람들과 즐겁게 어울려요"),
  ("일을 방해하는 사람이 나타날 수 있어요","당황 말고 스스로 일을 매듭지어 보세요"),
  ("컨디션이 최상으로 오르는 하루이군요","자신 있게 밀고 나가면 승산이 있죠"),
  ("노력한 만큼 대가가 따르는 날이네요","욕심과 말은 줄이고 집에 일찍 들어가요"),
 ],
 "닭": [
  ("감정이 상해서 언성이 높아질 수 있어요","한번 더 참고 상대를 이해하면 좋아요"),
  ("뜻밖의 사람에게 금전 도움을 받겠네요","겸손히 받아들이되 신중하게 결정하세요"),
  ("모든 면에서 유리한 자리에 서게 되네요","말과 행동을 마지막까지 조심하세요"),
  ("힘들던 일이 좋은 결실이 되어 돌아와요","귀인의 도움으로 일이 순조롭게 풀려요"),
  ("사랑이 활짝 만발하는 운명의 날이에요","가까운 곳의 인연을 적극 찾아 나서봐요"),
  ("의리 있는 친구가 큰 힘이 되어주겠군요","마음을 열면 기쁜 일이 절로 찾아와요"),
 ],
 "개": [
  ("일은 고되고 소득은 작게 느껴지는군요","먼 미래를 내다보며 다시 기운을 내봐요"),
  ("몸과 마음이 모두 지치기 쉬운 날이에요","남을 탓하지 말고 건강부터 챙겨야죠"),
  ("사방에서 물량이 쏟아져 나오고 있군요","발 빠르게 움직여 기회를 잡아야 해요"),
  ("몸은 바쁜데도 이익은 더디게 따라와요","가족과 의논하며 해결점을 찾아보세요"),
  ("혈기 탓에 직장에서 다툼이 생기겠네요","자제력을 기르고 성실히 움직여 보세요"),
  ("공부가 지겹고 집중이 잘 안되는 날이죠","과감하게 쉬며 마음의 평화를 찾아봐요"),
 ],
 "돼지": [
  ("남쪽에서 기쁘고 좋은 소식이 들려와요","마음을 가라앉히고 차분하게 지켜봐요"),
  ("혈색이 좋아 온화한 기운이 가득하군요","사랑하는 사람들과 추억을 만들어봐요"),
  ("일이 생각보다 훨씬 빠르게 풀리겠어요","의사 결정을 신속히 내려 결과를 얻어요"),
  ("다툼이 있던 상대에게 중재가 들어와요","적당한 선에서 슬기로운 타협을 해봐요"),
  ("금전운이 최상으로 차오르는 하루예요","조력자와 함께한 투자로 이익을 거둬요"),
  ("가까운 친구와 가족에게 경사가 생겨요","두려워 말고 씩씩하게 앞으로 나아가요"),
 ],
}

# Slide order (1-based) -> zodiac, matching Excel column order
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
    para.runs[0].text = str(val)        # number -> first run (keeps its font)
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

prs = Presentation(os.path.join(_TPL, "0511.pptx"))
warnings = []

# --- cover slide: update the date line ---
for sh in prs.slides[0].shapes:
    if sh.has_text_frame and "5월11일" in sh.text_frame.text:
        set_run_only(sh.text_frame.paragraphs[0], "2026년 7월9일 목요일")
for sidx, zod in SLIDE_ZODIAC.items():
    slide = prs.slides[sidx-1]
    chong, jaemul, aejeong = SCORES[zod]

    # --- score boxes: the 3 textboxes whose text is "NN점" ---
    score_boxes = [sh for sh in slide.shapes
                   if sh.has_text_frame and sh.text_frame.text.strip().endswith("점")
                   and len(sh.text_frame.text.strip()) <= 5]
    score_boxes.sort(key=lambda s: s.left)  # left->right : 총운, 재물, 애정
    if len(score_boxes) != 3:
        warnings.append(f"slide {sidx} {zod}: found {len(score_boxes)} score boxes")
    for box, val in zip(score_boxes, (chong, jaemul, aejeong)):
        set_score(box, val)

    # --- big fortune text box (height > 4,000,000) ---
    big = [sh for sh in slide.shapes
           if sh.has_text_frame and sh.height and sh.height > 4000000]
    big = max(big, key=lambda s: s.height)
    slots = fortune_slots(big.text_frame)
    if len(slots) != 6:
        warnings.append(f"slide {sidx} {zod}: {len(slots)} fortune slots (expected 6)")
    for slot, (l1, l2) in zip(slots, SUMMARIES[zod]):
        if len(slot) >= 2:
            set_run_only(slot[0], l1)
            set_run_only(slot[1], l2)
            for extra in slot[2:]:
                set_run_only(extra, "")
        elif len(slot) == 1:
            set_run_only(slot[0], l1 + " " + l2)

prs.save(os.path.join(_OUT, "0709.pptx"))
print("Saved 0709.pptx")
if warnings:
    print("WARNINGS:")
    for w in warnings: print(" -", w)
else:
    print("No warnings; all 12 slides updated.")
