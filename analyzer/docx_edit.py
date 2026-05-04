from copy import deepcopy
import re
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Cm
from docx.text.run import Run


def set_r_text(r_element, text):
    r_element.clear_content()
    if text:
        t = OxmlElement("w:t")
        if text.startswith(" ") or text.endswith(" "):
            t.set(qn("xml:space"), "preserve")
        t.text = text
        r_element.append(t)


def set_r_superscript(r_element, enabled=True):
    rpr = r_element.get_or_add_rPr()
    existing = rpr.find(qn("w:vertAlign"))
    if existing is None:
        existing = OxmlElement("w:vertAlign")
        rpr.append(existing)
    existing.set(qn("w:val"), "superscript" if enabled else "baseline")


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


def iter_text_runs(paragraph):
    """Yield all visible text runs in paragraph text order.

    python-docx Paragraph.runs only returns direct w:r children and skips runs
    stored inside w:hyperlink/w:fldSimple. After this project creates clickable
    citation links, those hyperlink runs still contribute to paragraph.text. If
    they are skipped, later range-based fixes calculate wrong offsets and may
    overwrite the character immediately after a citation (e.g. PC -> C,
    NOTEARS -> OTEARS) or duplicate the citation marker.
    """
    for child in paragraph._p.iterchildren():
        if child.tag == qn("w:r"):
            yield Run(child, paragraph)
        elif child.tag in {qn("w:hyperlink"), qn("w:fldSimple")}:
            for sub in child.iterchildren():
                if sub.tag == qn("w:r"):
                    yield Run(sub, paragraph)


def get_run_spans(paragraph):
    spans = []
    cursor = 0
    for run in iter_text_runs(paragraph):
        text = run.text or ""
        spans.append((cursor, cursor + len(text), run))
        cursor += len(text)
    return spans


def touched_runs(paragraph, start, end):
    result = []
    for left, right, run in get_run_spans(paragraph):
        if right <= start:
            continue
        if left >= end:
            break
        if right > start and left < end:
            result.append((left, right, run))
    return result


def run_has_cross_reference(run):
    node = run._element
    while node is not None:
        if node.tag == qn("w:hyperlink"):
            return True
        if node.tag == qn("w:fldSimple"):
            return True
        node = node.getparent()
    if run._element.find(qn("w:fldChar")) is not None:
        return True
    if run._element.find(qn("w:instrText")) is not None:
        return True
    return False


def range_has_cross_reference(paragraph, start, end):
    return any(run_has_cross_reference(run) for _, _, run in touched_runs(paragraph, start, end))


def paragraph_has_bookmark(paragraph, name):
    for item in paragraph._p.iter(qn("w:bookmarkStart")):
        if item.get(qn("w:name")) == name:
            return True
    return False


def add_bookmark_to_paragraph(paragraph, name, bookmark_id):
    if paragraph_has_bookmark(paragraph, name):
        return
    start = OxmlElement("w:bookmarkStart")
    start.set(qn("w:id"), str(bookmark_id))
    start.set(qn("w:name"), name)
    end = OxmlElement("w:bookmarkEnd")
    end.set(qn("w:id"), str(bookmark_id))
    paragraph._p.insert(0, start)
    paragraph._p.append(end)


def make_bookmark_hyperlink(run_element, display_text, bookmark_name, superscript=True):
    hyperlink = OxmlElement("w:hyperlink")
    hyperlink.set(qn("w:anchor"), bookmark_name)
    hyperlink.set(qn("w:history"), "1")
    link_run = deepcopy(run_element)
    set_r_text(link_run, display_text)
    set_r_superscript(link_run, superscript)
    hyperlink.append(link_run)
    return hyperlink


def replace_range_with_run(paragraph, start, end, replacement, superscript=True):
    if range_has_cross_reference(paragraph, start, end):
        return 0
    touched = touched_runs(paragraph, start, end)
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
        set_r_superscript(replacement_r, True)
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


def replace_range_with_ref_field(paragraph, start, end, display_text, bookmark_name, superscript=True):
    if range_has_cross_reference(paragraph, start, end):
        return 0
    touched = touched_runs(paragraph, start, end)
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
    new_elements.append(make_bookmark_hyperlink(first_run._r, display_text, bookmark_name, superscript))
    if suffix:
        suffix_r = deepcopy(last_run._r)
        set_r_text(suffix_r, suffix)
        new_elements.append(suffix_r)
    for _, _, run in touched:
        parent.remove(run._r)
    for offset, element in enumerate(new_elements):
        parent.insert(insert_index + offset, element)
    return 1


def convert_generated_ref_fields_to_hyperlinks(document):
    """Convert old generated REF ThesisRef_* fldSimple fields to direct links.

    Some PDF exporters keep TOC links but do not export REF fields as clickable
    annotations. Direct w:hyperlink anchors are more reliable for both Word and
    PDF export, while existing user-created references with other bookmark names
    are left untouched.
    """
    converted = 0
    pattern = re.compile(r"REF\s+(ThesisRef_\d+)\b")
    for paragraph in document.paragraphs:
        for field in list(paragraph._p.iter(qn("w:fldSimple"))):
            instr = field.get(qn("w:instr")) or ""
            match = pattern.search(instr)
            if not match:
                continue
            bookmark_name = match.group(1)
            parent = field.getparent()
            if parent is None:
                continue
            index = parent.index(field)
            text = "".join(t.text or "" for t in field.iter(qn("w:t")))
            source_run = None
            for child in field:
                if child.tag == qn("w:r"):
                    source_run = child
                    break
            if source_run is None:
                source_run = OxmlElement("w:r")
            hyperlink = make_bookmark_hyperlink(source_run, text, bookmark_name, True)
            parent.remove(field)
            parent.insert(index, hyperlink)
            converted += 1
    return converted


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
    for run in iter_text_runs(paragraph):
        if not run.text:
            continue
        if font_name:
            set_run_font(run, font_name)
        if font_size:
            run.font.size = Pt(font_size)
