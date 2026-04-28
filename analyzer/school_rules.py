import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.oxml.ns import qn
from models.issue import Issue
from analyzer.labels import issue_label
from analyzer.docx_reader import load_document, read_docx_paragraphs

SIZE_MAP={"初号":42,"小初":36,"一号":26,"小一":24,"二号":22,"小二":18,"三号":16,"小三":15,"四号":14,"小四":12,"五号":10.5,"小五":9,"六号":7.5,"小六":6.5}
FONT_NAMES=["Times New Roman","Arial","Cambria","宋体","黑体","楷体","仿宋","微软雅黑"]
CATEGORY_KEYWORDS={"body":["正文","论文正文"],"heading1":["一级标题","章标题","章名"],"heading2":["二级标题","节标题"],"heading3":["三级标题"],"abstract":["中文摘要","摘要"],"english_abstract":["英文摘要","ABSTRACT"],"keywords":["关键词","Key Words","Keywords"],"reference":["参考文献"],"figure_caption":["图题","图名","图注","图标题"],"table_caption":["表题","表名","表标题"]}
ALIGNMENT_MAP={"居中":WD_ALIGN_PARAGRAPH.CENTER,"左对齐":WD_ALIGN_PARAGRAPH.LEFT,"右对齐":WD_ALIGN_PARAGRAPH.RIGHT,"两端对齐":WD_ALIGN_PARAGRAPH.JUSTIFY}

def clean_text(text):
    return re.sub(r"\s+"," ",text.strip())

def infer_category(text):
    for category,words in CATEGORY_KEYWORDS.items():
        if any(word in text for word in words):
            return category
    return None

def extract_font(text):
    for font in FONT_NAMES:
        if font in text:
            return font
    return None

def extract_size(text):
    for name,pt in SIZE_MAP.items():
        if name in text:
            return name,pt
    m=re.search(r"(\d+(?:\.\d+)?)\s*磅",text)
    if m:
        return f"{m.group(1)}磅",float(m.group(1))
    return None,None

def extract_alignment(text):
    for key,value in ALIGNMENT_MAP.items():
        if key in text:
            return key,value
    return None,None

def extract_line_spacing(text):
    m=re.search(r"(\d+(?:\.\d+)?)\s*倍行距",text)
    if m:
        return f"{m.group(1)}倍行距",float(m.group(1))
    m=re.search(r"固定值\s*(\d+(?:\.\d+)?)\s*磅",text)
    if m:
        return f"固定值{m.group(1)}磅",Pt(float(m.group(1)))
    return None,None

def extract_first_line_indent(text):
    if "首行缩进" in text:
        m=re.search(r"首行缩进\s*(\d+(?:\.\d+)?)\s*(字符|字)",text)
        if m:
            return f"首行缩进{m.group(1)}字符",float(m.group(1))*0.37
        if "2字符" in text or "2 字符" in text or "两字符" in text:
            return "首行缩进2字符",0.74
    return None,None

def extract_citation_rule_from_text(text):
    rule={}
    if not any(k in text for k in ["引用","引文","文献标注","参考文献标注","正文标注","顺序编码"]):
        return rule
    if "【" in text or "】" in text or "【1】" in text:
        rule["bracket_style"]="【】"
    elif "〔" in text or "〕" in text or "〔1〕" in text:
        rule["bracket_style"]="〔〕"
    elif "［" in text or "］" in text or "［1］" in text:
        rule["bracket_style"]="［］"
    elif "[" in text or "]" in text or "[1]" in text:
        rule["bracket_style"]="[]"
    if "1~3" in text or "1～3" in text or "~" in text or "～" in text:
        rule["range_separator"]="~"
    elif "1-3" in text or "1–3" in text or "1—3" in text:
        rule["range_separator"]="-"
    if "1，2" in text or "中文逗号" in text:
        rule["list_separator"]="，"
    elif "1、2" in text or "顿号" in text:
        rule["list_separator"]="、"
    elif "1;2" in text or "1；2" in text or "分号" in text:
        rule["list_separator"]="；"
    elif "1,2" in text or "英文逗号" in text or ",[" in text:
        rule["list_separator']=','"
    if "不上标" in text or "非上标" in text or "不以上标" in text or "不采用上标" in text:
        rule["superscript"]=False
    elif "上标" in text or "角标" in text:
        rule["superscript"]=True
    return rule

def normalize_citation_rule(rule):
    result={"bracket_style":"[]","range_separator":"~","list_separator":",","superscript":True}
    result.update(rule or {})
    if result.get("list_separator")=="；":
        result["list_separator"]=";"
    if result.get("list_separator")=="，":
        result["list_separator"]="," if result.get("prefer_english_comma") else "，"
    return result

def parse_school_requirement_docx(file_path):
    document=load_document(file_path)
    paragraphs=read_docx_paragraphs(document, include_tables=True)
    rules={}
    raw_lines=[]
    citation_rule={}
    for para in paragraphs:
        text=clean_text(para["text"])
        if not text:
            continue
        raw_lines.append(text)
        extracted_citation_rule=extract_citation_rule_from_text(text)
        if extracted_citation_rule:
            citation_rule.update(extracted_citation_rule)
        category=infer_category(text)
        if not category:
            continue
        rule=rules.get(category,{})
        font=extract_font(text)
        size_name,size_pt=extract_size(text)
        align_name,align_value=extract_alignment(text)
        line_name,line_value=extract_line_spacing(text)
        indent_name,indent_value=extract_first_line_indent(text)
        if font:
            rule["font"]=font
        if size_pt:
            rule["size_name"]=size_name
            rule["size_pt"]=size_pt
        if align_name:
            rule["alignment_name"]=align_name
            rule["alignment_value"]=align_value
        if line_name:
            rule["line_spacing_name"]=line_name
            rule["line_spacing_value"]=line_value
        if indent_name:
            rule["first_line_indent_name"]=indent_name
            rule["first_line_indent_cm"]=indent_value
        rule["source_text"]=text
        rules[category]=rule
    return {"rules":rules,"citation_rule":normalize_citation_rule(citation_rule),"raw_text_preview":raw_lines[:80],"rule_count":len(rules)}

def paragraph_category(item,in_reference=False):
    text=item["text"].strip()
    style_name=item["paragraph"].style.name if item["paragraph"].style else ""
    if in_reference:
        return "reference"
    if text in {"摘  要","摘要"} or "摘要" in style_name:
        return "abstract"
    if text.upper()=="ABSTRACT" or text.startswith("ABSTRACT"):
        return "english_abstract"
    if text.startswith("关键词") or text.startswith("Key Words") or text.startswith("Keywords"):
        return "keywords"
    if re.match(r"^第[一二三四五六七八九十]+章",text) or "Heading 1" in style_name or "标题 1" in style_name:
        return "heading1"
    if re.match(r"^\d+\.\d+\s+",text) or "Heading 2" in style_name or "标题 2" in style_name:
        return "heading2"
    if re.match(r"^\d+\.\d+\.\d+\s+",text) or "Heading 3" in style_name or "标题 3" in style_name:
        return "heading3"
    if re.match(r"^图\s*\d+",text):
        return "figure_caption"
    if re.match(r"^表\s*\d+",text):
        return "table_caption"
    if len(text)>20:
        return "body"
    return None

def categorize_paragraphs(body_paragraphs,reference_paragraphs):
    result=[]
    for item in body_paragraphs:
        category=paragraph_category(item,False)
        if category:
            result.append({**item,"category":category})
    for item in reference_paragraphs:
        category=paragraph_category(item,True)
        if category:
            result.append({**item,"category":category})
    return result

def get_run_font_name(run):
    rpr=run._element.rPr
    if rpr is not None and rpr.rFonts is not None:
        for attr in [qn("w:eastAsia"),qn("w:ascii"),qn("w:hAnsi")]:
            value=rpr.rFonts.get(attr)
            if value:
                return value
    if run.font.name:
        return run.font.name
    return None

def get_paragraph_main_font(paragraph):
    fonts=[]
    for run in paragraph.runs:
        if run.text.strip():
            name=get_run_font_name(run)
            if name:
                fonts.append(name)
    if fonts:
        return max(set(fonts),key=fonts.count)
    style=paragraph.style
    if style and style.font and style.font.name:
        return style.font.name
    return None

def get_paragraph_main_size(paragraph):
    sizes=[]
    for run in paragraph.runs:
        if run.text.strip() and run.font.size:
            sizes.append(round(run.font.size.pt,1))
    if sizes:
        return max(set(sizes),key=sizes.count)
    style=paragraph.style
    if style and style.font and style.font.size:
        return round(style.font.size.pt,1)
    return None

def compare_float(a,b,tolerance=0.2):
    if a is None or b is None:
        return True
    return abs(float(a)-float(b))<=tolerance

def get_paragraph_line_spacing(paragraph):
    spacing = paragraph.paragraph_format.line_spacing
    if spacing is not None:
        return spacing
    style = paragraph.style
    if style and style.paragraph_format and style.paragraph_format.line_spacing is not None:
        return style.paragraph_format.line_spacing
    return None

def get_paragraph_first_line_indent_cm(paragraph):
    indent = paragraph.paragraph_format.first_line_indent
    if indent is None:
        style = paragraph.style
        if style and style.paragraph_format:
            indent = style.paragraph_format.first_line_indent
    if indent is None:
        return None
    return indent.cm

def compare_line_spacing(actual, expected):
    if actual is None or expected is None:
        return True
    if hasattr(actual, "pt") and hasattr(expected, "pt"):
        return compare_float(actual.pt, expected.pt, 0.5)
    if not hasattr(actual, "pt") and not hasattr(expected, "pt"):
        return compare_float(actual, expected, 0.05)
    return True

def check_school_format(categorized_paragraphs,school_rules):
    issues=[]
    rules=school_rules.get("rules",{}) if school_rules else {}
    if not rules:
        t="school_rule_not_extracted"
        issues.append(Issue(type=t,label=issue_label(t),group="school",problem="未能从学校格式要求 Word 中抽取到足够规则。",suggestion="建议确认学校要求中是否包含“正文、标题、参考文献、字体、字号、行距”等明确文字。",auto_fixable=False))
        return issues
    for item in categorized_paragraphs:
        category=item["category"]
        rule=rules.get(category)
        if not rule:
            continue
        paragraph=item["paragraph"]
        text=item["text"]
        if rule.get("font"):
            actual=get_paragraph_main_font(paragraph)
            expected=rule["font"]
            if actual and expected not in actual:
                t="school_font_mismatch"
                issues.append(Issue(type=t,label=issue_label(t),group="school",paragraph_index=item["index"],text=text,problem=f"{category} 字体疑似为 {actual}，学校要求为 {expected}。",suggestion=f"建议将该段字体调整为 {expected}。",auto_fixable=True,meta={"category":category,"expected":expected,"actual":actual}))
        if rule.get("size_pt"):
            actual_size=get_paragraph_main_size(paragraph)
            expected_size=rule["size_pt"]
            if actual_size and not compare_float(actual_size,expected_size):
                t="school_size_mismatch"
                issues.append(Issue(type=t,label=issue_label(t),group="school",paragraph_index=item["index"],text=text,problem=f"{category} 字号疑似为 {actual_size} 磅，学校要求为 {rule.get('size_name',expected_size)}。",suggestion=f"建议将该段字号调整为 {rule.get('size_name',expected_size)}。",auto_fixable=True,meta={"category":category,"expected":expected_size,"actual":actual_size}))
        if rule.get("alignment_value") is not None and paragraph.alignment is not None and paragraph.alignment!=rule["alignment_value"]:
            t="school_alignment_mismatch"
            issues.append(Issue(type=t,label=issue_label(t),group="school",paragraph_index=item["index"],text=text,problem=f"{category} 对齐方式与学校要求不一致。",suggestion=f"建议调整为 {rule.get('alignment_name')}。",auto_fixable=True,meta={"category":category}))
        if rule.get("line_spacing_value") is not None:
            actual_spacing=get_paragraph_line_spacing(paragraph)
            expected_spacing=rule["line_spacing_value"]
            if actual_spacing is not None and not compare_line_spacing(actual_spacing, expected_spacing):
                t="school_line_spacing_mismatch"
                issues.append(Issue(type=t,label=issue_label(t),group="school",paragraph_index=item["index"],text=text,problem=f"{category} 行距与学校要求不一致。",suggestion=f"建议调整为 {rule.get('line_spacing_name')}。",auto_fixable=True,meta={"category":category,"expected":rule.get("line_spacing_name"),"actual":str(actual_spacing)}))
        if rule.get("first_line_indent_cm") is not None:
            actual_indent=get_paragraph_first_line_indent_cm(paragraph)
            expected_indent=rule["first_line_indent_cm"]
            if actual_indent is not None and not compare_float(actual_indent, expected_indent, 0.15):
                t="school_first_line_indent_mismatch"
                issues.append(Issue(type=t,label=issue_label(t),group="school",paragraph_index=item["index"],text=text,problem=f"{category} 首行缩进与学校要求不一致。",suggestion=f"建议调整为 {rule.get('first_line_indent_name')}。",auto_fixable=True,meta={"category":category,"expected":expected_indent,"actual":actual_indent}))
    return issues
