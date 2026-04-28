from pathlib import Path
from analyzer.citation_utils import CITATION_PATTERN, CITATION_SEQUENCE_PATTERN, compact_sequence_text, normalize_single_marker_text
from analyzer.docx_edit import replace_range_with_run, set_paragraph_basic_format, set_paragraph_runs_format


def fix_superscript_in_body(document, body_paragraphs, citation_rule=None):
    fixed = 0
    rule = citation_rule or {}
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
    rule = citation_rule or {}
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
        category = item["category"]
        paragraph = item["paragraph"]
        rule = rules.get(category)
        if not rule:
            continue
        if rule.get("font") or rule.get("size_pt"):
            set_paragraph_runs_format(paragraph, rule)
            fixed += 1
        if rule.get("alignment_value") is not None or rule.get("line_spacing_value") is not None or rule.get("first_line_indent_cm") is not None:
            set_paragraph_basic_format(paragraph, rule)
            fixed += 1
    return fixed


def save_fixed_document(document, output_path):
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(output))
