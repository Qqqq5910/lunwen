from docx.oxml.ns import qn
from analyzer.docx_edit import run_has_cross_reference, iter_text_runs

SPACE_CHARS = {
    " ", "\u00a0", "\u2000", "\u2001", "\u2002", "\u2003", "\u2004", "\u2005",
    "\u2006", "\u2007", "\u2008", "\u2009", "\u200a", "\u202f", "\u205f", "\u3000",
    "\t", "\r", "\n"
}


def is_space(ch):
    return ch in SPACE_CHARS


def is_cjk(ch):
    return bool(ch) and "\u4e00" <= ch <= "\u9fff"


def is_latin_digit(ch):
    return bool(ch) and ch.isascii() and (ch.isalpha() or ch.isdigit())


def is_open_punct(ch):
    return ch in "（([【〔《“‘"


def is_close_punct(ch):
    return ch in "）)]】〕》”’"


def is_cjk_punct(ch):
    return ch in "，。；：、！？（）。；：、！？【】〔〕《》“”‘’"


def run_is_field_or_hyperlink(run):
    if run_has_cross_reference(run):
        return True
    node = run._element
    while node is not None:
        if node.tag in {qn("w:hyperlink"), qn("w:fldSimple")}:
            return True
        node = node.getparent()
    return run._element.find(qn("w:fldChar")) is not None or run._element.find(qn("w:instrText")) is not None


def should_remove_boundary_space(prev_ch, next_ch):
    if not prev_ch or not next_ch:
        return False

    # Citation markers should be adjacent to surrounding Chinese prose.
    if prev_ch == "]" or next_ch == "[":
        return True
    if prev_ch == "[" or next_ch == "]":
        return True

    # Chinese punctuation/brackets do not require inner spaces. Avoid removing
    # normal English spaces before "(" such as "Generation (RAG)".
    if is_open_punct(prev_ch) or is_close_punct(next_ch):
        return True
    if is_open_punct(next_ch) and (is_cjk(prev_ch) or is_cjk_punct(prev_ch)):
        return True
    if is_close_punct(prev_ch) and (is_cjk(next_ch) or is_cjk_punct(next_ch)):
        return True
    if is_cjk_punct(prev_ch) or is_cjk_punct(next_ch):
        return True

    # In Chinese thesis body text, AI/Word copy often leaves visible blank runs
    # between Chinese and English terms, names, algorithms, or "等人".
    if is_cjk(prev_ch) and is_cjk(next_ch):
        return True
    if (is_cjk(prev_ch) and is_latin_digit(next_ch)) or (is_latin_digit(prev_ch) and is_cjk(next_ch)):
        return True

    return False


def clean_paragraph_whitespace(paragraph):
    chars = []
    for run_index, run in enumerate(iter_text_runs(paragraph)):
        text = run.text or ""
        protected = run_is_field_or_hyperlink(run)
        for char_index, ch in enumerate(text):
            chars.append({"ch": ch, "run": run_index, "index": char_index, "protected": protected, "run_obj": run})

    if not chars:
        return 0

    remove_positions = set()
    total = len(chars)

    for pos, item in enumerate(chars):
        if not is_space(item["ch"]) or item["protected"]:
            continue

        left = pos - 1
        while left >= 0 and is_space(chars[left]["ch"]):
            left -= 1
        right = pos + 1
        while right < total and is_space(chars[right]["ch"]):
            right += 1

        prev_ch = chars[left]["ch"] if left >= 0 else ""
        next_ch = chars[right]["ch"] if right < total else ""
        if should_remove_boundary_space(prev_ch, next_ch):
            remove_positions.add((item["run"], item["index"]))

    # Collapse repeated editable spaces that are not removed by boundary rules.
    pos = 0
    while pos < total:
        if not is_space(chars[pos]["ch"]):
            pos += 1
            continue
        start = pos
        while pos < total and is_space(chars[pos]["ch"]):
            pos += 1
        editable = [
            item for item in chars[start:pos]
            if not item["protected"] and (item["run"], item["index"]) not in remove_positions
        ]
        if len(editable) > 1:
            for item in editable[1:]:
                remove_positions.add((item["run"], item["index"]))

    removed = 0
    for run_index, run in enumerate(iter_text_runs(paragraph)):
        if run_is_field_or_hyperlink(run):
            continue
        old = run.text or ""
        if not old:
            continue
        new_chars = []
        for char_index, ch in enumerate(old):
            if (run_index, char_index) in remove_positions:
                removed += 1
                continue
            new_chars.append(" " if is_space(ch) else ch)
        new = "".join(new_chars)
        if new != old:
            run.text = new
    return removed


def clean_body_whitespace(document, body_paragraphs):
    fixed = 0
    for item in body_paragraphs:
        paragraph = item.get("paragraph")
        if paragraph is None:
            continue
        fixed += clean_paragraph_whitespace(paragraph)
    return fixed
