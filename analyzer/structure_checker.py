import re
from collections import defaultdict
from models.issue import Issue
from analyzer.labels import issue_label

FIGURE_PATTERN = re.compile(r"图\s*(\d+)[-.－.](\d+)")
TABLE_PATTERN = re.compile(r"表\s*(\d+)[-.－.](\d+)")
EQUATION_PATTERN = re.compile(r"(?:公式|式)\s*[（(]\s*(\d+)[-.－.](\d+)\s*[）)]")
HEADING_PATTERN = re.compile(r"^\s*(\d+(?:\.\d+){1,3})\s+(.+)$")

def extract_numbered_items(paragraphs):
    items=[]
    patterns=[("figure",FIGURE_PATTERN),("table",TABLE_PATTERN),("equation",EQUATION_PATTERN)]
    for para in paragraphs:
        text=para["text"]
        for kind,pattern in patterns:
            for m in pattern.finditer(text):
                items.append({"kind":kind,"chapter":int(m.group(1)),"index":int(m.group(2)),"raw":m.group(0),"paragraph_index":para["index"],"text":text})
    return items

def check_numbered_items(paragraphs):
    issues=[]
    items=extract_numbered_items(paragraphs)
    grouped=defaultdict(list)
    name_map={"figure":"图","table":"表","equation":"公式"}
    for item in items:
        grouped[(item["kind"],item["chapter"])].append(item)
    for (kind,chapter),group in grouped.items():
        indexes=[item["index"] for item in group]
        seen=set()
        for item in group:
            if item["index"] in seen:
                t=f"{kind}_number_duplicate"
                issues.append(Issue(type=t,label=issue_label(t),group="manual",paragraph_index=item["paragraph_index"],text=item["text"],problem=f"{name_map[kind]} {chapter}-{item['index']} 重复出现。",suggestion="请检查编号是否重复。",auto_fixable=False))
            seen.add(item["index"])
        unique=sorted(set(indexes))
        if unique:
            expected=set(range(unique[0],unique[-1]+1))
            for miss in sorted(expected-set(unique)):
                t=f"{kind}_number_jump"
                issues.append(Issue(type=t,label=issue_label(t),group="manual",problem=f"{name_map[kind]} {chapter}-{miss} 疑似缺失，编号存在跳号。",suggestion="请检查图、表或公式编号是否连续。",auto_fixable=False))
    for para in paragraphs:
        text=para["text"]
        if "如下图" in text or "见下图" in text:
            t="ambiguous_figure_reference"
            issues.append(Issue(type=t,label=issue_label(t),group="reminder",paragraph_index=para["index"],text=text,problem="正文中出现“如下图”或“见下图”，但可能缺少具体图编号。",suggestion="建议改为“如图 x-x 所示”。",auto_fixable=False))
        if "如下表" in text or "见下表" in text:
            t="ambiguous_table_reference"
            issues.append(Issue(type=t,label=issue_label(t),group="reminder",paragraph_index=para["index"],text=text,problem="正文中出现“如下表”或“见下表”，但可能缺少具体表编号。",suggestion="建议改为“如表 x-x 所示”。",auto_fixable=False))
    return issues

def find_headings(paragraphs):
    headings=[]
    for para in paragraphs:
        text=para["text"].strip()
        match=HEADING_PATTERN.match(text)
        if match:
            headings.append({"number":match.group(1),"title":match.group(2),"level":match.group(1).count(".")+1,"paragraph_index":para["index"],"text":text})
    return headings

def check_headings(paragraphs):
    issues=[]
    headings=find_headings(paragraphs)
    seen={}
    for h in headings:
        number=h["number"]
        if number in seen:
            t="heading_number_duplicate"
            issues.append(Issue(type=t,label=issue_label(t),group="manual",paragraph_index=h["paragraph_index"],text=h["text"],problem=f"标题编号 {number} 重复出现。",suggestion="请检查标题编号是否重复。",auto_fixable=False))
        seen[number]=h
    grouped=defaultdict(list)
    for h in headings:
        parts=h["number"].split(".")
        if len(parts)>=2:
            grouped[".".join(parts[:-1])].append(int(parts[-1]))
    for parent,nums in grouped.items():
        unique=sorted(set(nums))
        if unique:
            expected=set(range(unique[0],unique[-1]+1))
            for miss in sorted(expected-set(unique)):
                t="heading_number_jump"
                issues.append(Issue(type=t,label=issue_label(t),group="manual",problem=f"标题编号 {parent}.{miss} 疑似缺失。",suggestion="请检查标题层级编号是否连续。",auto_fixable=False))
    return issues
