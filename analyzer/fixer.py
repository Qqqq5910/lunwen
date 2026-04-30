from pathlib import Path
from docx.shared import Pt, Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from analyzer.citation_utils import CITATION_PATTERN, CITATION_SEQUENCE_PATTERN, compact_sequence_text, normalize_single_marker_text, format_numbers, expand_citation_numbers
from analyzer.docx_edit import replace_range_with_run, replace_range_with_ref_field, range_has_cross_reference, add_bookmark_to_paragraph
from analyzer.marker_fix import plain_spans

BOOKMARK_PREFIX = "ThesisRef"


def _set_run_fonts(run, east_asia=None, latin=None, size_pt=None, bold=None):
    if range_has_cross_reference(run._parent, 0, 0):
        pass
    rpr = run._element.get_or_add_rPr()
    for old in list(rpr.findall(qn("w:rFonts"))):
        rpr.remove(old)
    rfonts = OxmlElement("w:rFonts")
    rpr.insert(0, rfonts)
    if east_asia:
        rfonts.set(qn("w:eastAsia"), east_asia)
    if latin:
        rfonts.set(qn("w:ascii"), latin)
        rfonts.set(qn("w:hAnsi"), latin)
        rfonts.set(qn("w:cs"), latin)
    if size_pt:
        run.font.size = Pt(size_pt)
    if bold is not None:
        run.font.bold = bold


def _set_paragraph_format(paragraph, rule):
    fmt = paragraph.paragraph_format
    if rule.get("alignment_value") is not None:
        paragraph.alignment = rule["alignment_value"]
    if rule.get("line_spacing_pt") is not None:
        fmt.line_spacing = Pt(rule["line_spacing_pt"])
    elif rule.get("line_spacing_value") is not None:
        fmt.line_spacing = rule["line_spacing_value"]
    if rule.get("space_before_pt") is not None:
        fmt.space_before = Pt(rule["space_before_pt"])
    if rule.get("space_after_pt") is not None:
        fmt.space_after = Pt(rule["space_after_pt"])
    if rule.get("first_line_indent_cm") is not None:
        fmt.first_line_indent = Cm(rule["first_line_indent_cm"])


def build_reference_bookmarks(references, reference_paragraphs):
    paragraphs = {item["index"]: item["paragraph"] for item in reference_paragraphs}
    result = {}
    for ref in references:
        number = ref.get("number")
        paragraph = paragraphs.get(ref.get("paragraph_index"))
        if not number or paragraph is None:
            continue
        name = f"{BOOKMARK_PREFIX}_{number}"
        add_bookmark_to_paragraph(paragraph, name, 70000 + int(number))
        result[number] = name
    return result


def fix_plain_citations_in_body(document, body_paragraphs, reference_numbers=None, citation_rule=None, bookmark_map=None):
    fixed = 0
    rule = citation_rule or {"bracket_style": "[]", "range_separator": "~", "list_separator": ",", "superscript": True}
    bookmark_map = bookmark_map or {}
    for item in body_paragraphs:
        paragraph = item["paragraph"]
        text = paragraph.text
        for start, end, number in reversed(plain_spans(text, reference_numbers)):
            if range_has_cross_reference(paragraph, start, end):
                continue
            replacement = format_numbers([number], rule)
            bookmark = bookmark_map.get(number)
            if bookmark:
                fixed += replace_range_with_ref_field(paragraph, start, end, replacement, bookmark, rule.get("superscript", True))
            else:
                fixed += replace_range_with_run(paragraph, start, end, replacement, rule.get("superscript", True))
    return fixed


def fix_superscript_in_body(document, body_paragraphs, citation_rule=None, bookmark_map=None):
    fixed = 0
    rule = citation_rule or {"bracket_style": "[]", "range_separator": "~", "list_separator": ",", "superscript": True}
    superscript = rule.get("superscript", True)
    bookmark_map = bookmark_map or {}
    for item in body_paragraphs:
        paragraph = item["paragraph"]
        text = paragraph.text
        matches = list(CITATION_PATTERN.finditer(text))
        for match in reversed(matches):
            replacement = normalize_single_marker_text(match.group(0), rule)
            if range_has_cross_reference(paragraph, match.start(), match.end()):
                continue
            numbers = expand_citation_numbers(match.group(0))
            bookmark = bookmark_map.get(numbers[0]) if numbers else None
            if bookmark:
                fixed += replace_range_with_ref_field(paragraph, match.start(), match.end(), replacement, bookmark, superscript)
            else:
                fixed += replace_range_with_run(paragraph, match.start(), match.end(), replacement, superscript)
    return fixed


def fix_citation_ranges_in_body(document, body_paragraphs, citation_rule=None, bookmark_map=None):
    fixed = 0
    rule = citation_rule or {"bracket_style": "[]", "range_separator": "~", "list_separator": ",", "superscript": True}
    superscript = rule.get("superscript", True)
    bookmark_map = bookmark_map or {}
    for item in body_paragraphs:
        paragraph = item["paragraph"]
        text = paragraph.text
        matches = list(CITATION_SEQUENCE_PATTERN.finditer(text))
        for match in reversed(matches):
            replacement = compact_sequence_text(match.group(0), rule)
            if range_has_cross_reference(paragraph, match.start(), match.end()):
                continue
            numbers = expand_citation_numbers(match.group(0))
            bookmark = bookmark_map.get(numbers[0]) if numbers else None
            if bookmark:
                fixed += replace_range_with_ref_field(paragraph, match.start(), match.end(), replacement, bookmark, superscript)
            elif replacement != match.group(0):
                fixed += replace_range_with_run(paragraph, match.start(), match.end(), replacement, superscript)
    return fixed


def fix_school_format(document, categorized_paragraphs, rules):
    fixed = 0
    for item in categorized_paragraphs:
        rule = rules.get(item["category"])
        if not rule:
            continue
        paragraph = item["paragraph"]
        _set_paragraph_format(paragraph, rule)
        for run in paragraph.runs:
            if run.text:
                _set_run_fonts(run, rule.get("font_east_asia"), rule.get("font_latin"), rule.get("size_pt"), rule.get("bold") if "bold" in rule else None)
        fixed += 1
    return fixed


def save_fixed_document(document, output_path):
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(output))
