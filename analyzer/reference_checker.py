import re
from models.issue import Issue
from analyzer.labels import issue_label

REFERENCE_TITLE_PATTERN=re.compile(r"^\s*(参考文献|References|REFERENCES|Bibliography|BIBLIOGRAPHY)\s*[:：]?\s*$|^\s*参考文献\s+(References|REFERENCES)?\s*$|^\s*参\s*考\s*文\s*献\s*$")
REFERENCE_NUMBER_PATTERN=re.compile(r"^\s*(?:\[|［)?\s*(\d{1,4})\s*(?:\]|］|\.|．|、|\))\s*(.+)?$")
STOP_TITLE_PATTERN=re.compile(r"^\s*(致谢|致\s*谢|附录|附\s*录|声明|作者简介)\s*$")
STOP_KEYWORDS=["在读期间","公开发表","承担科研项目","取得成果","攻读学位期间","学术成果","科研项目"]

def is_reference_title(text):
    value=text.strip()
    if value in {"参考文献","References","REFERENCES","Bibliography","BIBLIOGRAPHY"}:
        return True
    return bool(REFERENCE_TITLE_PATTERN.match(value))

def is_stop_title(text):
    value=text.strip()
    if STOP_TITLE_PATTERN.match(value):
        return True
    return any(keyword in value for keyword in STOP_KEYWORDS)

def find_reference_start_pos(paragraphs):
    for pos,paragraph in enumerate(paragraphs):
        if is_reference_title(paragraph["text"]):
            return pos
    return None

def split_body_and_reference(paragraphs):
    start_pos=find_reference_start_pos(paragraphs)
    if start_pos is None:
        return paragraphs,[],None
    body=paragraphs[:start_pos]
    tail=paragraphs[start_pos+1:]
    refs=[]
    for item in tail:
        if is_stop_title(item["text"]):
            break
        refs.append(item)
    return body,refs,paragraphs[start_pos]

def looks_like_reference_text(text):
    value=text.strip()
    if len(value)<8:
        return False
    if re.search(r"\[[A-Z]{1,3}(?:/[A-Z]{1,3})?\]",value):
        return True
    if re.search(r"\b(19|20)\d{2}\b",value):
        return True
    if "DOI" in value.upper():
        return True
    if re.search(r"[\u4e00-\u9fa5]{2,}.*[.．]",value):
        return True
    return False

def extract_references_from_reference_paragraphs(reference_paragraphs):
    references=[]
    inferred_number=1
    explicit_numbers=[]
    for paragraph in reference_paragraphs:
        text=paragraph["text"].strip()
        match=REFERENCE_NUMBER_PATTERN.match(text)
        if match:
            number=int(match.group(1))
            explicit_numbers.append(number)
            references.append({"number":number,"paragraph_index":paragraph["index"],"text":match.group(2).strip() if match.group(2) else text,"raw_text":text,"number_source":"explicit"})
        elif looks_like_reference_text(text):
            references.append({"number":inferred_number,"paragraph_index":paragraph["index"],"text":text,"raw_text":text,"number_source":"inferred"})
            inferred_number+=1
    if explicit_numbers:
        for idx,ref in enumerate(references,start=1):
            if ref["number_source"]=="inferred":
                ref["number"]=idx
    return references

def get_reference_number_set(references):
    return {item["number"] for item in references}

def check_reference_numbers(references,reference_title_found):
    issues=[]
    if not reference_title_found:
        t="reference_section_missing"
        issues.append(Issue(type=t,label=issue_label(t),group="manual",problem="未识别到参考文献区域。",suggestion="请确认文末是否存在单独一行“参考文献”。",auto_fixable=False))
        return issues
    if not references:
        t="reference_items_missing"
        issues.append(Issue(type=t,label=issue_label(t),group="manual",problem="已识别到参考文献标题，但未识别到具体参考文献条目。",suggestion="请确认参考文献是否采用编号列表，或每条参考文献是否独立成段。",auto_fixable=False))
        return issues
    numbers=[item["number"] for item in references]
    if min(numbers)!=1:
        t="reference_not_start_from_1"
        first_ref=references[0]
        issues.append(Issue(type=t,label=issue_label(t),group="manual",paragraph_index=first_ref["paragraph_index"],text=first_ref["raw_text"],problem=f"参考文献编号最小值为 [{min(numbers)}]，不是从 [1] 开始。",suggestion="请检查参考文献编号是否缺失或排序异常。",auto_fixable=False))
    seen=set()
    for ref in references:
        number=ref["number"]
        if number in seen:
            t="reference_duplicate_number"
            issues.append(Issue(type=t,label=issue_label(t),group="manual",paragraph_index=ref["paragraph_index"],text=ref["raw_text"],problem=f"参考文献编号 [{number}] 重复出现。",suggestion="请检查是否存在重复编号或编号错误。",auto_fixable=False))
        seen.add(number)
    ordered=sorted(set(numbers))
    expected=set(range(1,ordered[-1]+1))
    for number in sorted(expected-set(ordered)):
        t="reference_number_missing"
        issues.append(Issue(type=t,label=issue_label(t),group="manual",problem=f"参考文献编号缺少 [{number}]。",suggestion="请检查参考文献列表是否存在跳号。",auto_fixable=False))
    for ref in references:
        text=ref["raw_text"]
        if not re.search(r"\[[A-Z]{1,3}(?:/[A-Z]{1,3})?\]",text):
            t="reference_type_missing"
            issues.append(Issue(type=t,label=issue_label(t),group="reminder",paragraph_index=ref["paragraph_index"],text=text,problem=f"参考文献 [{ref['number']}] 可能缺少文献类型标识。",suggestion="请核对是否需要补充 [J]、[C]、[M]、[D] 或 [EB/OL]。",auto_fixable=False))
    return issues

def check_citation_reference_mapping(cited_numbers,reference_numbers,references):
    issues=[]
    ref_by_number={item["number"]:item for item in references}
    for number in sorted(cited_numbers-reference_numbers):
        t="citation_missing_reference"
        issues.append(Issue(type=t,label=issue_label(t),group="manual",problem=f"正文引用了 [{number}]，但参考文献列表中未找到第 [{number}] 条。",suggestion=f"请补充第 [{number}] 条参考文献，或修改正文引用编号。",auto_fixable=False))
    for number in sorted(reference_numbers-cited_numbers):
        t="reference_not_cited"
        ref=ref_by_number.get(number,{})
        issues.append(Issue(type=t,label=issue_label(t),group="reminder",paragraph_index=ref.get("paragraph_index"),text=ref.get("raw_text"),problem=f"参考文献 [{number}] 未在正文中被引用。",suggestion=f"请确认参考文献 [{number}] 是否需要保留，或在正文中补充引用。",auto_fixable=False))
    return issues
