from pathlib import Path
from analyzer.citation_utils import CITATION_PATTERN, CITATION_SEQUENCE_PATTERN, compact_sequence_text
from analyzer.docx_edit import replace_range_with_run


def fix_superscript_in_body(document, body_paragraphs):
    fixed = 0
    for item in body_paragraphs:
        paragraph = item["paragraph"]
        text = paragraph.text
        matches = list(CITATION_PATTERN.finditer(text))
        for match in reversed(matches):
            spans = []
            cursor = 0
            for run in paragraph.runs:
                spans.append((cursor, cursor + len(run.text), run))
                cursor += len(run.text)
            touched = []
            for left, right, run in spans:
                if right <= match.start():
                    continue
                if left >= match.end():
                    break
                if right > match.start() and left < match.end():
                    touched.append(run)
            if touched and not all(run.font.superscript is True for run in touched if run.text):
                fixed += replace_range_with_run(paragraph, match.start(), match.end(), match.group(0), True)
    return fixed


def fix_citation_ranges_in_body(document, body_paragraphs):
    fixed = 0
    for item in body_paragraphs:
        paragraph = item["paragraph"]
        text = paragraph.text
        matches = list(CITATION_SEQUENCE_PATTERN.finditer(text))
        for match in reversed(matches):
            replacement = compact_sequence_text(match.group(0))
            if replacement != match.group(0):
                fixed += replace_range_with_run(paragraph, match.start(), match.end(), replacement, True)
    return fixed


def fix_school_format(document, categorized_paragraphs, rules):
    return 0


def save_fixed_document(document, output_path):
    output = Path(output_path)
    output.parent.mkdir(parents=True, exist_ok=True)
    document.save(str(output))
