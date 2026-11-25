# -*- coding: utf-8 -*-
"""Harvest approved fortune text and draft a new build module from that memory.

Usage: ``py -X utf8 fortune_memory.py harvest`` or
``py -X utf8 fortune_memory.py draft MMDD [--force] [--out PATH] [--include-self]``.
"""
import argparse
import ast
import datetime as dt
import json
import os
import re
from pathlib import Path

from pptx import Presentation

import qa

BASE = Path(__file__).resolve().parent
PPTX_DIR = BASE.parent / "4_PPTX"
MEMORY_PATH = BASE / "fortune_memory.json"
CORRECTIONS_PATH = BASE / "fortune_corrections.json"
ZOD = ["쥐", "소", "호랑이", "토끼", "용", "뱀", "말", "양", "원숭이", "닭", "개", "돼지"]
WEEKDAYS = "월화수목금토일"
SLOT_RE = re.compile(r"^\s*slot([1-6])\s+[^:]*:\s*(.*)$")
HEADER_RE = re.compile(r"^---\s*(쥐|소|호랑이|토끼|용|뱀|말|양|원숭이|닭|개|돼지)\s*---")


def normalized(text):
    """Return the punctuation-free key used for original text and variants."""
    return re.sub(r"[^0-9A-Za-z가-힣]", "", text)


def fortune_groups(text_frame):
    """Return groups of non-empty paragraphs separated by blank paragraphs."""
    groups, current = [], []
    for paragraph in text_frame.paragraphs:
        if paragraph.text.strip():
            current.append(paragraph)
        elif current:
            groups.append(current)
            current = []
    if current:
        groups.append(current)
    return groups


def group_lines(group):
    """Read a fortune group, including old one-paragraph merged two-line text."""
    clean = lambda text: text.replace("\x0b", "").strip()
    if len(group) >= 2:
        return clean(group[0].text), clean(group[1].text)
    if len(group) != 1:
        return None
    text = group[0].text
    if "\x0b" in text or "\n" in text:
        left, right = re.split(r"[\x0b\n]", text, maxsplit=1)
    else:
        sentence_end = re.search(r"[.!?]\s+", text)
        if not sentence_end:
            return None
        left, right = text[:sentence_end.start() + 1], text[sentence_end.end():]
    left, right = clean(left), clean(right)
    return (left, right) if left and right else None


def parse_source(path):
    """Parse one src_MMDD.txt, returning its scores line, date and 12x6 sources."""
    lines = path.read_text(encoding="utf-8").splitlines()
    if len(lines) < 2 or not lines[0].startswith("SCORES ="):
        raise ValueError(f"{path.name}: 첫 줄 SCORES를 읽을 수 없습니다")
    match = re.search(r"#\s*(\d{4})-(\d{2})-(\d{2})", lines[1])
    if not match:
        raise ValueError(f"{path.name}: 둘째 줄 날짜를 읽을 수 없습니다")
    date = match.group(2) + match.group(3)
    slots = {zodiac: {} for zodiac in ZOD}
    zodiac = None
    for line in lines[2:]:
        header = HEADER_RE.match(line)
        if header:
            zodiac = header.group(1)
            continue
        slot = SLOT_RE.match(line)
        if zodiac and slot:
            slots[zodiac][int(slot.group(1))] = slot.group(2).strip()
    missing = [f"{zodiac} slot{slot}" for zodiac in ZOD for slot in range(1, 7)
               if not slots[zodiac].get(slot)]
    if missing:
        raise ValueError(f"{path.name}: 72칸이 완전하지 않습니다 ({', '.join(missing)})")
    return {
        "scores_line": lines[0],
        "year": int(match.group(1)),
        "month": int(match.group(2)),
        "day": int(match.group(3)),
        "date": date,
        "sources": {zodiac: [slots[zodiac][slot] for slot in range(1, 7)] for zodiac in ZOD},
    }


def parse_draft(path):
    """Read SUMMARIES with AST so an old build module is never imported."""
    if not path.is_file():
        return None
    try:
        tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    except (OSError, SyntaxError, UnicodeError):
        return None
    for node in tree.body:
        if not isinstance(node, ast.Assign):
            continue
        if any(isinstance(target, ast.Name) and target.id == "SUMMARIES" for target in node.targets):
            try:
                value = ast.literal_eval(node.value)
            except (ValueError, TypeError):
                return None
            return value if isinstance(value, dict) else None
    return None


def extract_pptx(path):
    """Extract the six two-line approved summaries from slides 2 through 13."""
    approved, warnings = {}, []
    try:
        presentation = Presentation(path)
    except Exception as exc:
        return approved, [f"{path.name}: pptx를 열 수 없습니다 ({exc})"], False
    if len(presentation.slides) < 13:
        return approved, [f"{path.name}: 슬라이드가 13장보다 적습니다"], True
    for index, zodiac in enumerate(ZOD, start=2):
        slide = presentation.slides[index - 1]
        boxes = [shape for shape in slide.shapes
                 if shape.has_text_frame and shape.height and shape.height > 4_000_000]
        if not boxes:
            warnings.append(f"{path.name} {zodiac}: 운세 상자를 찾지 못했습니다")
            continue
        groups = fortune_groups(max(boxes, key=lambda shape: shape.height).text_frame)
        lines = [group_lines(group) for group in groups]
        if len(groups) != 6 or any(pair is None for pair in lines):
            warnings.append(f"{path.name} {zodiac}: 운세 그룹이 6개가 아니거나 두 줄이 아닙니다")
            continue
        approved[zodiac] = lines
    return approved, warnings, True


def source_files():
    """Yield real src files in code order, excluding any non-date lookalikes."""
    for path in sorted(BASE.glob("src_????.txt")):
        if re.fullmatch(r"src_\d{4}\.txt", path.name):
            yield path


def save_memory(memory):
    tmp = MEMORY_PATH.with_name(MEMORY_PATH.name + ".tmp")
    tmp.write_text(json.dumps(memory, ensure_ascii=False, indent=1) + "\n", encoding="utf-8")
    os.replace(tmp, MEMORY_PATH)


def load_memory():
    """Load the memory file, returning an empty valid memory when it is absent."""
    if not MEMORY_PATH.is_file():
        return {"version": 1, "entries": {}}
    try:
        value = json.loads(MEMORY_PATH.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"메모리 파일을 읽을 수 없습니다: {exc}") from exc
    if not isinstance(value, dict) or value.get("version") != 1 or not isinstance(value.get("entries"), dict):
        raise ValueError("메모리 파일 형식이 올바르지 않습니다")
    return value


def load_corrections():
    """Load the corrections ledger, returning an empty list when it is absent."""
    if not CORRECTIONS_PATH.is_file():
        return []
    try:
        value = json.loads(CORRECTIONS_PATH.read_text(encoding="utf-8"))
    except (OSError, UnicodeError, json.JSONDecodeError) as exc:
        raise ValueError(f"교정 원장을 읽을 수 없습니다: {exc}") from exc
    if not isinstance(value, dict) or value.get("version") != 1 or not isinstance(value.get("corrections"), list):
        raise ValueError("교정 원장 형식이 올바르지 않습니다")
    for item in value["corrections"]:
        if (not isinstance(item, dict) or any(name not in item for name in ("key", "old", "new"))
                or not isinstance(item["key"], str)):
            raise ValueError("교정 원장 형식이 올바르지 않습니다")
        for pair in (item["old"], item["new"]):
            if not (isinstance(pair, list) and len(pair) == 2 and all(isinstance(line, str) for line in pair)):
                raise ValueError("교정 원장 형식이 올바르지 않습니다")
    return value["corrections"]


def apply_corrections(final_entries, corrections):
    """Apply each ledger correction once to original lines, without cascading."""
    applied = 0
    original, touched, seen_keys = {}, [], set()
    for correction in corrections:
        item = final_entries.get(correction["key"])
        if not item:
            continue
        key = correction["key"]
        if key not in seen_keys:
            seen_keys.add(key)
            touched.append(key)
            for variant in item["variants"]:
                original[id(variant)] = (variant["l1"], variant["l2"])
        old_pair = tuple(correction["old"])
        new_l1, new_l2 = correction["new"]
        for variant in item["variants"]:
            # Match the pre-correction pair so a rewritten line cannot feed the next row.
            if original.get(id(variant)) != old_pair:
                continue
            original[id(variant)] = None
            variant["l1"] = new_l1
            variant["l2"] = new_l2
            variant["edited"] = True  # A ledger correction is a reviewed, approved sentence.
            applied += 1
    for key in touched:
        item = final_entries[key]
        seen, kept = set(), []
        for variant in item["variants"]:
            pair = (variant["l1"], variant["l2"])
            if pair in seen:
                continue
            seen.add(pair)
            kept.append(variant)
        # Keep the earlier variant so edited-first, newest-date order survives.
        item["variants"] = kept
    return applied


def approved_variant(draft, approved, zodiac, slot):
    """Return whether an approved pair differs from the matching source build draft."""
    if not draft:
        return False
    try:
        candidate = draft[zodiac][slot - 1]
        return normalized(candidate[0] + candidate[1]) != normalized(approved[0] + approved[1])
    except (KeyError, IndexError, TypeError):
        return False


def merge_existing(entries, existing_entries, replaced):
    """Keep variants whose (date, zodiac) was not re-harvested, merging sentences already collected."""
    preserved = 0
    for key, item in existing_entries.items():
        if (not isinstance(item, dict) or not isinstance(item.get("source"), str)
                or not isinstance(item.get("variants"), list)):
            raise ValueError(f"메모리 파일 형식이 올바르지 않습니다: {key[:20]}")
        for variant in item["variants"]:
            if (not isinstance(variant, dict) or not isinstance(variant.get("l1"), str)
                    or not isinstance(variant.get("l2"), str) or not isinstance(variant.get("date"), str)
                    or not isinstance(variant.get("zodiac"), str) or not isinstance(variant.get("edited"), bool)):
                raise ValueError(f"메모리 파일 형식이 올바르지 않습니다: {key[:20]}")
            if (variant["date"], variant["zodiac"]) in replaced:
                # This date and zodiac are harvested again, so only carry edited onto the same new sentence.
                target = entries.get(key)
                if target is not None:
                    text_key = normalized(variant["l1"] + variant["l2"])
                    current = target["variants"].get(text_key)
                    if current is not None:
                        current["edited"] = current["edited"] or variant["edited"]
                continue
            target = entries.setdefault(key, {"source": item["source"], "variants": {}})
            text_key = normalized(variant["l1"] + variant["l2"])
            current = target["variants"].get(text_key)
            if current is None:
                target["variants"][text_key] = dict(variant)
                preserved += 1
                continue
            # A newer date wins. The same date keeps the fresh harvest and only ORs edited.
            if variant["date"] > current["date"]:
                chosen = dict(variant)
            else:
                chosen = current
            chosen["edited"] = current["edited"] or variant["edited"]
            target["variants"][text_key] = chosen
    return preserved


def harvest():
    """Merge today's source/PPTX harvest into fortune_memory.json, preserving variants whose (date, zodiac) was not read."""
    existing = load_memory()
    entries, warnings, dates = {}, [], set()
    replaced = set()
    for src_path in source_files():
        try:
            source = parse_source(src_path)
        except ValueError as exc:
            warnings.append(str(exc))
            continue
        code = source["date"]
        pptx_path = PPTX_DIR / f"{code}.pptx"
        if not pptx_path.is_file() or pptx_path.name.startswith("~$"):
            warnings.append(f"{code}: 확정 pptx가 없어 건너뜁니다")
            continue
        approved, pptx_warnings, opened = extract_pptx(pptx_path)
        if opened:
            dates.add(code)
        warnings.extend(pptx_warnings)
        draft = parse_draft(BASE / f"build_{code}.py")
        for zodiac in ZOD:
            if zodiac not in approved:
                continue
            replaced.add((code, zodiac))
            for slot, pair in enumerate(approved[zodiac], start=1):
                raw = source["sources"][zodiac][slot - 1]
                key = normalized(raw)
                text_key = normalized(pair[0] + pair[1])
                item = entries.setdefault(key, {"source": raw, "variants": {}})
                edited = approved_variant(draft, pair, zodiac, slot)
                old = item["variants"].get(text_key)
                current = {"l1": pair[0], "l2": pair[1], "date": code,
                           "zodiac": zodiac, "slot": slot, "edited": edited}
                if old is None or code >= old["date"]:
                    if old is not None:
                        current["edited"] = old["edited"] or edited
                    item["variants"][text_key] = current
                elif edited:
                    old["edited"] = True
    preserved = merge_existing(entries, existing["entries"], replaced)
    final_entries = {}
    for key, item in entries.items():
        # Stable sorts express the required priority: edited first, then newest date.
        variants = sorted(item["variants"].values(), key=lambda variant: variant["date"], reverse=True)
        variants.sort(key=lambda variant: not variant["edited"])
        final_entries[key] = {"source": item["source"], "variants": variants}
    corrections = load_corrections()
    applied = apply_corrections(final_entries, corrections)
    memory = {"version": 1, "entries": final_entries}
    save_memory(memory)
    count = sum(len(item["variants"]) for item in final_entries.values())
    edited = sum(variant["edited"] for item in final_entries.values() for variant in item["variants"])
    print(f"수집 날짜: {len(dates)}")
    print(f"엔트리: {len(final_entries)}")
    print(f"변형: {count} (edited {edited})")
    print(f"보존(pptx 없는 날짜): {preserved}")
    print(f"교정 원장 적용: {applied}/{len(corrections)}")
    if warnings:
        print("경고:")
        for warning in warnings:
            print(" -", warning)


def valid_candidates(memory, raw, code, include_self):
    """Return QA-safe variants for raw text, in stored priority order."""
    item = memory["entries"].get(normalized(raw))
    if not item:
        return []
    candidates = []
    for variant in item.get("variants", []):
        if not include_self and variant.get("date") == code:
            continue
        l1 = terminalized(variant.get("l1", ""))
        l2 = terminalized(variant.get("l2", ""))
        if qa.line_problems("", l1) or qa.line_problems("", l2):
            continue
        if qa.width_px(l1) > qa.WIDTH_LIMIT or qa.width_px(l2) > qa.WIDTH_LIMIT:
            continue
        candidate = dict(variant)
        candidate["l1"], candidate["l2"] = l1, l2
        candidates.append(candidate)
    return candidates


def terminalized(line):
    """Ensure a drafted memory line ends in sentence punctuation before QA checks."""
    line = line.rstrip()
    return line if line.endswith((".", "!", "?")) else line + "."


def placeholder(zodiac_index, slot, line):
    """QA-neutral stand-in used only while TODO cells are excluded from QA scoring."""
    # qa.ecls uses only the final two characters for unknown endings.  A two-digit
    # hexadecimal id keeps every temporary line distinct, including within one zodiac.
    unique_id = zodiac_index * 12 + (slot - 1) * 2 + (line - 1)
    return f"대체문장이없어검토가필요합니다{unique_id:02x}"


def qa_problems(summaries, available):
    """Run qa.check after replacing unavailable TODO cells with unique neutral text."""
    checked = {}
    for zi, zodiac in enumerate(ZOD):
        checked[zodiac] = []
        for slot, pair in enumerate(summaries[zodiac], start=1):
            if (zodiac, slot) in available:
                checked[zodiac].append(pair)
            else:
                checked[zodiac].append((placeholder(zi, slot, 1), placeholder(zi, slot, 2)))
    return qa.check(checked)[0]


def improve_summaries(summaries, candidates, available):
    """Deterministically accept candidate replacements only when QA issues decrease."""
    problems = qa_problems(summaries, available)
    attempts = 0
    changed = True
    while problems and changed and attempts < 5000:
        changed = False
        for zodiac in ZOD:
            for slot in range(1, 7):
                options = candidates[(zodiac, slot)]
                if len(options) < 2:
                    continue
                current = summaries[zodiac][slot - 1]
                for option in options:
                    pair = (option["l1"], option["l2"])
                    if pair == current:
                        continue
                    attempts += 1
                    summaries[zodiac][slot - 1] = pair
                    candidate_problems = qa_problems(summaries, available)
                    if len(candidate_problems) < len(problems):
                        problems = candidate_problems
                        changed = True
                        break
                    summaries[zodiac][slot - 1] = current
                    if attempts >= 5000:
                        break
                if changed or attempts >= 5000:
                    break
            if changed or attempts >= 5000:
                break
    return problems


def cover_text(source):
    """Make the required cover string from the source file's actual calendar date."""
    weekday = WEEKDAYS[dt.date(source["year"], source["month"], source["day"]).weekday()]
    return f'{source["year"]}년 {source["month"]}월{source["day"]}일 {weekday}요일'


def render_build(code, source, summaries, todos):
    """Render the intentionally small, import-safe build module used for new dates."""
    lines = [
        "# -*- coding: utf-8 -*-",
        '"""Build this date\'s fortunes from the shared template. Run as a script."""',
        "from build_common import build",
        "",
        f"COVER = {cover_text(source)!r}",
        source["scores_line"],
        "SUMMARIES = {",
    ]
    todo_set = {(zodiac, slot) for zodiac, slot, _ in todos}
    raw_lookup = source["sources"]
    for zodiac in ZOD:
        lines.append(f" {zodiac!r}: [")
        for slot, pair in enumerate(summaries[zodiac], start=1):
            if (zodiac, slot) in todo_set:
                lines.append(f"  # TODO 원문: {raw_lookup[zodiac][slot - 1]}")
            lines.append(f"  ({pair[0]!r}, {pair[1]!r}),")
        lines.append(" ],")
    lines.extend([
        "}",
        "",
        'if __name__ == "__main__":',
        f"    build({code!r}, COVER, SCORES, SUMMARIES)",
        "",
    ])
    return "\n".join(lines)


def draft(code, force=False, out=None, include_self=False):
    """Draft a build_MMDD.py from memory, preserving missing cells as explicit TODOs."""
    if not re.fullmatch(r"\d{4}", code):
        raise ValueError("MMDD는 숫자 네 자리여야 합니다")
    src_path = BASE / f"src_{code}.txt"
    source = parse_source(src_path)
    memory = load_memory()
    output = Path(out) if out else BASE / f"build_{code}.py"
    if not output.is_absolute():
        output = BASE / output
    if output.exists() and not force:
        raise FileExistsError(f"이미 파일이 있습니다: {output} (--force로 덮어쓰기)")

    summaries = {zodiac: [] for zodiac in ZOD}
    candidates, available, todos = {}, set(), []
    memory_edited = memory_plain = 0
    for zodiac in ZOD:
        for slot, raw in enumerate(source["sources"][zodiac], start=1):
            options = valid_candidates(memory, raw, code, include_self)
            candidates[(zodiac, slot)] = options
            if options:
                chosen = options[0]
                summaries[zodiac].append((chosen["l1"], chosen["l2"]))
                available.add((zodiac, slot))
                if chosen["edited"]:
                    memory_edited += 1
                else:
                    memory_plain += 1
            else:
                summaries[zodiac].append(("", ""))
                todos.append((zodiac, slot, raw))

    problems = improve_summaries(summaries, candidates, available)
    # Count the variants actually chosen after local search rather than only initial picks.
    memory_edited = memory_plain = 0
    for zodiac, slot in available:
        pair = summaries[zodiac][slot - 1]
        selected = next((variant for variant in candidates[(zodiac, slot)]
                         if (variant["l1"], variant["l2"]) == pair), None)
        if selected and selected["edited"]:
            memory_edited += 1
        elif selected:
            memory_plain += 1

    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(render_build(code, source, summaries, todos), encoding="utf-8")
    todo_path = output.parent / f"todo_{code}.txt"
    if todos:
        todo_path.write_text("\n".join(f"{zodiac} slot{slot}: {raw}" for zodiac, slot, raw in todos) + "\n",
                             encoding="utf-8")
    elif todo_path.exists():
        todo_path.unlink()
    print(f"메모리 채움: {memory_edited + memory_plain}칸 (edited {memory_edited}, 비edited {memory_plain})")
    print(f"TODO 칸: {len(todos)}")
    print(f"남은 qa problems: {len(problems)}")
    print(f"작성: {output}")


def main():
    """Parse the harvest/draft CLI without doing work when this module is imported."""
    parser = argparse.ArgumentParser(description="확정 운세 문장 메모리 도구")
    subparsers = parser.add_subparsers(dest="command", required=True)
    subparsers.add_parser("harvest", help="src와 확정 PPTX에서 메모리를 다시 수집")
    draft_parser = subparsers.add_parser("draft", help="메모리로 build_MMDD.py 초안을 작성")
    draft_parser.add_argument("code", metavar="MMDD")
    draft_parser.add_argument("--force", action="store_true")
    draft_parser.add_argument("--out")
    draft_parser.add_argument("--include-self", action="store_true")
    args = parser.parse_args()
    try:
        if args.command == "harvest":
            harvest()
        else:
            draft(args.code, args.force, args.out, args.include_self)
    except (ValueError, FileNotFoundError, FileExistsError) as exc:
        parser.error(str(exc))


if __name__ == "__main__":
    main()
