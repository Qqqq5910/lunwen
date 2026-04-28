from pathlib import Path
from docx.shared import Pt, Cm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from analyzer.citation_utils import CITATION_PATTERN, CITATION_SEQUENCE_PATTERN, compact_sequence_text, normalize_single_marker_text
from analyzer.docx_edit import replace_range_with_run


def _set_run_fonts(run, east_asia=None, latin=None, size_pt=None, bold=None):
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.append(rfonts)
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


def fix_superscript_in_body(document, body_paragraphs, citation_rule=None):
    fixed = 0
    rule = citation_rule or {"bracket_style": "[]", "range_separator": "~", "list_separator": ",", "superscript": True}
    superscript = rule.get("superscript", True)
    for item in body_paragraphs:
        paragraph = item["paragraph"]
        text = paragraph.text
        matches = list(CITATION_PATTERN.finditer(text))
        for match in reversed(matches):
            replacement = normalize_single_marker_text(match.group(0), rule)
            fixed += replace_range_with_run(paragraph, match.start(), match.end(), replacement, superscript)
    return fixed


def fix_citation_ranges_in_body(document, body_paragraphs, citation_rule=None):
    fixed = 0
    rule = citation_rule or {"bracket_style": "[]", "range_separator": "~", "list_separator": ",", "superscript": True}
    superscript = rule.get("superscript", True)
    for item in body_paragraphs:
        paragraph = item["paragraph"]
        text = paragraph.text
        matches = list(CITATION_SEQUENCE_PATTERN.finditer(text))
        for match in reversed(matches):
            replacement = compact_sequence_text(match.group(0), rule)
            if replacement != match.group(0):
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
