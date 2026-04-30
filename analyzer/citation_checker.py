from models.issue import Issue
from analyzer.labels import issue_label
from analyzer.citation_utils import CITATION_PATTERN, CITATION_SEQUENCE_PATTERN, normalize_marker, expand_citation_numbers, compact_sequence_text
from docx.oxml.ns import qn
import re


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


def run_has_superscript_xml(run):
    rpr=run._element.rPr
    if rpr is None:
        return False
    vert=rpr.find(qn("w:vertAlign"))
    return vert is not None and vert.get(qn("w:val"))=="superscript"


def run_is_field_or_hyperlink(run):
    node=run._element
    while node is not None:
        if node.tag in {qn("w:hyperlink"), qn("w:fldSimple")}:
            return True
        node=node.getparent()
    return run._element.find(qn("w:fldChar")) is not None or run._element.find(qn("w:instrText")) is not None


def is_marker_superscript(paragraph,start,end):
    runs=[run for run in runs_for_range(get_run_spans(paragraph),start,end) if run.text]
    if not runs:
        return False
    if any(run_is_field_or_hyperlink(run) for run in runs):
        return True
    return all((run.font.superscript is True) or run_has_superscript_xml(run) for run in runs)


def paragraph_field_text(paragraph):
    parts=[]
    for node in paragraph._p.iter():
        if node.tag==qn("w:t") and node.text:
            parts.append(node.text)
    return "".join(parts)


def unique_citations(citations):
    seen=set()
    out=[]
    for item in citations:
        key=(item["paragraph_index"], tuple(item["numbers"]), item["raw"])
        if key in seen:
            continue
        seen.add(key)
        out.append(item)
    return out


def find_citations(paragraphs):
    citations=[]
    for item in paragraphs:
        paragraph=item["paragraph"]
        texts=[paragraph.text]
        rich_text=paragraph_field_text(paragraph)
        if rich_text and rich_text!=paragraph.text:
            texts.append(rich_text)
        for text in texts:
            for match in CITATION_PATTERN.finditer(text):
                raw=match.group(0)
                numbers=expand_citation_numbers(raw)
                if not numbers:
                    continue
                is_super=True if text is rich_text and text!=paragraph.text else is_marker_superscript(paragraph,match.start(),match.end())
                citations.append({"paragraph_index":item["index"],"raw":raw,"normalized":normalize_marker(raw),"numbers":numbers,"text":text.strip(),"start":match.start(),"end":match.end(),"is_superscript":is_super})
    return unique_citations(citations)


def find_citation_sequences(paragraphs,citation_rule=None):
    sequences=[]
    for item in paragraphs:
        paragraph=item["paragraph"]
        text=paragraph.text
        for match in CITATION_SEQUENCE_PATTERN.finditer(text):
            raw=match.group(0)
            compacted=compact_sequence_text(raw,citation_rule)
            if compacted!=raw:
                sequences.append({"paragraph_index":item["index"],"raw":raw,"compacted":compacted,"text":text.strip(),"start":match.start(),"end":match.end()})
    return sequences


def get_cited_number_set(citations):
    result=set()
    for item in citations:
        result.update(item["numbers"])
    return result


def check_citation_format(citations):
    issues=[]
    for item in citations:
        raw=item["raw"]
        if not item["is_superscript"]:
            t="citation_not_superscript"
            issues.append(Issue(type=t,label=issue_label(t),group="auto",paragraph_index=item["paragraph_index"],text=item["text"],problem=f"正文引用 {raw} 未设置为上标。",suggestion=f"建议将正文引用 {raw} 设置为上标格式。",auto_fixable=True))
        if "［" in raw or "］" in raw:
            t="citation_fullwidth_bracket"
            issues.append(Issue(type=t,label=issue_label(t),group="auto",paragraph_index=item["paragraph_index"],text=item["text"],problem=f"正文引用 {raw} 使用了全角方括号。",suggestion=f"建议改为半角方括号格式，例如 {normalize_marker(raw)}。",auto_fixable=True))
        if ", " in raw or " ," in raw or "，" in raw:
            t="citation_spacing_or_comma"
            issues.append(Issue(type=t,label=issue_label(t),group="auto",paragraph_index=item["paragraph_index"],text=item["text"],problem=f"正文引用 {raw} 中存在空格或中文逗号。",suggestion="建议统一为英文逗号且无空格格式。",auto_fixable=True))
    return issues


def check_citation_sequences(sequences):
    issues=[]
    for item in sequences:
        t="citation_sequence_not_compacted"
        issues.append(Issue(type=t,label=issue_label(t),group="auto",paragraph_index=item["paragraph_index"],text=item["text"],problem=f"连续引用 {item['raw']} 可以合并为 {item['compacted']}。",suggestion=f"建议将连续引用统一为 {item['compacted']}。",auto_fixable=True))
    return issues
