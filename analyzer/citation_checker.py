from models.issue import Issue
from analyzer.labels import issue_label
from analyzer.citation_utils import CITATION_PATTERN, CITATION_SEQUENCE_PATTERN, normalize_marker, expand_citation_numbers, compact_sequence_text, normalize_single_marker_text


def get_run_spans(paragraph):
    spans=[]
    cursor=0
    for run in paragraph.runs:
        text=run.text
        spans.append((cursor,cursor+len(text),run))
        cursor+=len(text)
    return spans


def runs_for_range(spans,start,end):
    result=[]
    for left,right,run in spans:
        if right<=start:
            continue
        if left>=end:
            break
        if right>start and left<end:
            result.append(run)
    return result


def _run_is_superscript(run):
    if run.font.superscript is True:
        return True
    style = getattr(run, "style", None)
    if style is not None and getattr(style, "font", None) is not None:
        return style.font.superscript is True
    return False


def is_marker_superscript(paragraph,start,end):
    runs=runs_for_range(get_run_spans(paragraph),start,end)
    if not runs:
        return False
    text_runs=[run for run in runs if run.text]
    return bool(text_runs) and all(_run_is_superscript(run) for run in text_runs)


def find_citations(paragraphs):
    citations=[]
    for item in paragraphs:
        paragraph=item["paragraph"]
        text=paragraph.text
        for match in CITATION_PATTERN.finditer(text):
            raw=match.group(0)
            numbers=expand_citation_numbers(raw)
            if not numbers:
                continue
            citations.append({"paragraph_index":item["index"],"raw":raw,"normalized":normalize_marker(raw),"numbers":numbers,"text":text.strip(),"start":match.start(),"end":match.end(),"is_superscript":is_marker_superscript(paragraph,match.start(),match.end())})
    return citations


def find_citation_sequences(paragraphs, citation_rule=None):
    sequences=[]
    for item in paragraphs:
        paragraph=item["paragraph"]
        text=paragraph.text
        for match in CITATION_SEQUENCE_PATTERN.finditer(text):
            raw=match.group(0)
            compacted=compact_sequence_text(raw, citation_rule)
            if compacted!=raw:
                sequences.append({"paragraph_index":item["index"],"raw":raw,"compacted":compacted,"text":text.strip(),"start":match.start(),"end":match.end()})
    return sequences


def get_cited_number_set(citations):
    result=set()
    for item in citations:
        result.update(item["numbers"])
    return result


def check_citation_format(citations, citation_rule=None):
    issues=[]
    rule=citation_rule or {}
    require_superscript=rule.get("superscript", True)
    for item in citations:
        raw=item["raw"]
        expected=normalize_single_marker_text(raw, rule)
        if require_superscript and not item["is_superscript"]:
            t="citation_not_superscript"
            issues.append(Issue(type=t,label=issue_label(t),group="auto",paragraph_index=item["paragraph_index"],text=item["text"],problem=f"正文引用 {raw} 未设置为上标。",suggestion=f"学校要求引用为上标，建议将 {raw} 设置为上标格式。",auto_fixable=True,meta={"expected":expected,"citation_rule":rule}))
        if not require_superscript and item["is_superscript"]:
            t="citation_should_not_superscript"
            issues.append(Issue(type=t,label=issue_label(t),group="auto",paragraph_index=item["paragraph_index"],text=item["text"],problem=f"正文引用 {raw} 不应设置为上标。",suggestion=f"学校要求引用不上标，建议改为 {expected}。",auto_fixable=True,meta={"expected":expected,"citation_rule":rule}))
        if expected != raw:
            t="citation_style_mismatch"
            issues.append(Issue(type=t,label=issue_label(t),group="auto",paragraph_index=item["paragraph_index"],text=item["text"],problem=f"正文引用 {raw} 与学校要求格式不一致。",suggestion=f"建议按学校要求改为 {expected}。",auto_fixable=True,meta={"expected":expected,"citation_rule":rule}))
    return issues


def check_citation_sequences(sequences):
    issues=[]
    for item in sequences:
        t="citation_sequence_not_compacted"
        issues.append(Issue(type=t,label=issue_label(t),group="auto",paragraph_index=item["paragraph_index"],text=item["text"],problem=f"连续引用 {item['raw']} 可以合并为 {item['compacted']}。",suggestion=f"建议将连续引用统一为 {item['compacted']}。",auto_fixable=True,meta={"expected":item["compacted"]}))
    return issues
