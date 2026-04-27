from copy import deepcopy
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm


def set_r_text(r_element, text):
    r_element.clear_content()
    if text:
        t = OxmlElement("w:t")
        if text.startswith(" ") or text.endswith(" "):
            t.set(qn("xml:space"), "preserve")
        t.text = text
        r_element.append(t)


def set_r_superscript(r_element):
    rpr = r_element.get_or_add_rPr()
    existing = rpr.find(qn("w:vertAlign"))
    if existing is None:
        existing = OxmlElement("w:vertAlign")
        rpr.append(existing)
    existing.set(qn("w:val"), "superscript")


def set_run_font(run, font_name):
    run.font.name = font_name
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
    rfonts.set(qn("w:eastAsia"), font_name)
    rfonts.set(qn("w:ascii"), font_name)
    rfonts.set(qn("w:hAnsi"), font_name)


def get_run_spans(paragraph):
    spans = []
    cursor = 0
    for run in paragraph.runs:
        text = run.text
        spans.append((cursor, cursor + len(text), run))
        cursor += len(text)
    return spans


def replace_range_with_run(paragraph, start, end, replacement, superscript=True):
    spans = get_run_spans(paragraph)
    touched = []
    for left, right, run in spans:
        if right <= start:
            continue
        if left >= end:
            break
        if right > start and left < end:
            touched.append((left, right, run))
    if not touched:
        return 0
    first_left, first_right, first_run = touched[0]
    last_left, last_right, last_run = touched[-1]
    prefix = first_run.text[:max(0, start - first_left)]
    suffix = last_run.text[max(0, end - last_left):]
    parent = first_run._r.getparent()
    insert_index = parent.index(first_run._r)
    new_elements = []
    if prefix:
        prefix_r = deepcopy(first_run._r)
        set_r_text(prefix_r, prefix)
        new_elements.append(prefix_r)
    replacement_r = deepcopy(first_run._r)
    set_r_text(replacement_r, replacement)
    if superscript:
        set_r_superscript(replacement_r)
    new_elements.append(replacement_r)
    if suffix:
        suffix_r = deepcopy(last_run._r)
        set_r_text(suffix_r, suffix)
        new_elements.append(suffix_r)
    for _, _, run in touched:
        parent.remove(run._r)
    for offset, element in enumerate(new_elements):
        parent.insert(insert_index + offset, element)
    return 1


def set_paragraph_basic_format(paragraph, rule):
    fmt = paragraph.paragraph_format
    if rule.get("alignment_value") is not None:
        paragraph.alignment = rule["alignment_value"]
    if rule.get("line_spacing_value") is not None:
        fmt.line_spacing = rule["line_spacing_value"]
    if rule.get("first_line_indent_cm") is not None:
        fmt.first_line_indent = Cm(rule["first_line_indent_cm"])


def set_paragraph_runs_format(paragraph, rule):
    font_name = rule.get("font")
    font_size = rule.get("size_pt")
    for run in paragraph.runs:
        if not run.text:
            continue
        if font_name:
            set_run_font(run, font_name)
        if font_size:
            run.font.size = Pt(font_size)
