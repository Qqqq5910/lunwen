import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt
from docx.oxml.ns import qn
from models.issue import Issue
from analyzer.labels import issue_label
from analyzer.docx_reader import load_document, read_docx_paragraphs

SIZE_MAP = {"初号": 42, "小初": 36, "一号": 26, "小一": 24, "二号": 22, "小二": 18, "三号": 16, "小三": 15, "四号": 14, "小四": 12, "五号": 10.5, "小五": 9, "六号": 7.5, "小六": 6.5}
SIZE_KEYS = ["小初", "初号", "小一", "一号", "小二", "二号", "小三", "三号", "小四", "四号", "小五", "五号", "小六", "六号"]
FONT_NAMES = ["Times New Roman", "Arial", "Cambria", "宋体", "黑体", "楷体", "仿宋", "微软雅黑"]
ALIGNMENT_MAP = {"居中": WD_ALIGN_PARAGRAPH.CENTER, "左对齐": WD_ALIGN_PARAGRAPH.LEFT, "右对齐": WD_ALIGN_PARAGRAPH.RIGHT, "两端对齐": WD_ALIGN_PARAGRAPH.JUSTIFY}
DEFAULT_CITATION_RULE = {"bracket_style": "[]", "range_separator": "~", "list_separator": ",", "superscript": True}
CATEGORY_KEYWORDS = {
    "body": ["正文字体", "论文所有的文字", "正文"],
    "abstract": ["摘要内容", "中文摘要内容"],
    "english_abstract": ["英文摘要内容"],
    "heading1": ["章标题", "各章标题", "章名"],
    "heading2": ["一级节标题"],
    "heading3": ["二级节标题"],
    "toc": ["目录内容"],
    "reference": ["参考文献的正文", "成果内容格式"],
    "figure_caption": ["图序与图名", "图题", "图名"],
    "table_caption": ["表序与表名", "表题", "表名"],
    "keywords": ["中文论文关键词"],
    "english_keywords": ["英文论文关键词"]
}

def clean_text(text):
    return re.sub(r"\s+", " ", text.strip())

def extract_size(text):
    for name in SIZE_KEYS:
        if name in text:
            return name, SIZE_MAP[name]
    m = re.search(r"(\d+(?:\.\d+)?)\s*磅", text)
    if m:
        return f"{m.group(1)}磅", float(m.group(1))
    return None, None

def extract_alignment(text):
    for key, value in ALIGNMENT_MAP.items():
        if key in text:
            return key, value
    return None, None

def extract_line_spacing(text):
    m = re.search(r"固定值\s*(\d+(?:\.\d+)?)\s*磅", text)
    if m:
        return f"固定值{m.group(1)}磅", float(m.group(1))
    m = re.search(r"(\d+(?:\.\d+)?)\s*倍行距", text)
    if m:
        return f"{m.group(1)}倍行距", float(m.group(1))
    if "单倍行距" in text:
        return "单倍行距", 1.0
    return None, None

def extract_spacing(text):
    before = None
    after = None
    m = re.search(r"段前空\s*(\d+(?:\.\d+)?)\s*磅", text)
    if m:
        before = float(m.group(1))
    m = re.search(r"段后空\s*(\d+(?:\.\d+)?)\s*磅", text)
    if m:
        after = float(m.group(1))
    return before, after

def extract_first_line_indent(text):
    if "首行缩进" not in text:
        return None
    m = re.search(r"首行缩进\s*(\d+(?:\.\d+)?)\s*(?:个)?汉?字符", text)
    if m:
        return float(m.group(1)) * 0.37
    if "两个汉字符" in text or "2字符" in text or "2 字符" in text or "两字符" in text:
        return 0.74
    return None

def infer_category(text):
    if "英文摘要内容" in text:
        return "english_abstract"
    if "摘要内容" in text and "英文" not in text:
        return "abstract"
    if "中文论文关键词" in text:
        return "keywords"
    if "英文论文关键词" in text:
        return "english_keywords"
    if "参考文献的正文" in text:
        return "reference"
    if "图序与图名" in text:
        return "figure_caption"
    if "表序与表名" in text:
        return "table_caption"
    if "各章标题" in text or "章标题" in text:
        return "heading1"
    if "一级节标题" in text:
        return "heading2"
    if "二级节标题" in text:
        return "heading3"
    if "目录内容" in text:
        return "toc"
    if "正文字体" in text or "论文所有的文字" in text:
        return "body"
    for category, words in CATEGORY_KEYWORDS.items():
        if any(word in text for word in words):
            return category
    return None

def extract_font_rule(text, category):
    rule = {}
    if "中文用宋体" in text:
        rule["font_east_asia"] = "宋体"
    elif "中文用" in text:
        for font in FONT_NAMES:
            if f"中文用{font}" in text:
                rule["font_east_asia"] = font
                break
    else:
        for font in FONT_NAMES:
            if font in text and font != "Times New Roman":
                rule["font_east_asia"] = font
                break
    if "数字和英文用 Times New Roman" in text or "英文用 Times New Roman" in text or "Times New Roman" in text:
        rule["font_latin"] = "Times New Roman"
    if category and category.startswith("english"):
        rule["font_latin"] = "Times New Roman"
        rule["font_east_asia"] = rule.get("font_east_asia", "Times New Roman")
    if "加粗" in text:
        rule["bold"] = True
    return rule

def extract_citation_rule(text):
    rule = {}
    if not any(k in text for k in ["引用", "引文", "文献标注", "参考文献标注", "正文标注", "顺序编码"]):
        return rule
    if "【" in text or "】" in text:
        rule["bracket_style"] = "【】"
    elif "〔" in text or "〕" in text:
        rule["bracket_style"] = "〔〕"
    elif "［" in text or "］" in text:
        rule["bracket_style"] = "［］"
    elif "[" in text or "]" in text or "方括号" in text:
        rule["bracket_style"] = "[]"
    if "1~3" in text or "10~12" in text or "～" in text or "~" in text:
        rule["range_separator"] = "~"
    elif "1-3" in text or "1–3" in text or "1—3" in text:
        rule["range_separator"] = "-"
    if "1，2" in text or "中文逗号" in text:
        rule["list_separator"] = "，"
    elif "1、2" in text or "顿号" in text:
        rule["list_separator"] = "、"
    elif "1;2" in text or "1；2" in text or "分号" in text:
        rule["list_separator"] = ";"
    elif "1,2" in text or "英文逗号" in text or "半角符号" in text:
        rule["list_separator"] = ","
    if "不上标" in text or "非上标" in text or "不采用上标" in text:
        rule["superscript"] = False
    elif "上标" in text or "右上角" in text or "角标" in text:
        rule["superscript"] = True
    return rule

def normalize_rule(category, rule):
    if category == "body":
        rule.setdefault("font_east_asia", "宋体")
        rule.setdefault("font_latin", "Times New Roman")
        rule.setdefault("size_pt", 12)
        rule.setdefault("line_spacing_pt", 20)
    if category in {"abstract"}:
        rule.setdefault("font_east_asia", "宋体")
        rule.setdefault("font_latin", "Times New Roman")
        rule.setdefault("size_pt", 12)
        rule.setdefault("line_spacing_pt", 20)
        rule.setdefault("first_line_indent_cm", 0.74)
    if category in {"english_abstract"}:
        rule.setdefault("font_east_asia", "Times New Roman")
        rule.setdefault("font_latin", "Times New Roman")
        rule.setdefault("size_pt", 12)
        rule.setdefault("line_spacing_pt", 20)
        rule.setdefault("first_line_indent_cm", 0.74)
    if category == "heading1":
        rule.setdefault("font_east_asia", "黑体")
        rule.setdefault("font_latin", "Times New Roman")
        rule.setdefault("size_pt", 18)
        rule.setdefault("alignment_value", WD_ALIGN_PARAGRAPH.CENTER)
        rule.setdefault("space_before_pt", 24)
        rule.setdefault("space_after_pt", 18)
        rule.setdefault("line_spacing_value", 1.0)
    if category == "heading2":
        rule.setdefault("font_east_asia", "宋体")
        rule.setdefault("font_latin", "Times New Roman")
        rule.setdefault("size_pt", 14)
        rule.setdefault("bold", True)
        rule.setdefault("alignment_value", WD_ALIGN_PARAGRAPH.LEFT)
        rule.setdefault("line_spacing_pt", 20)
        rule.setdefault("space_before_pt", 24)
        rule.setdefault("space_after_pt", 6)
    if category == "heading3":
        rule.setdefault("font_east_asia", "宋体")
        rule.setdefault("font_latin", "Times New Roman")
        rule.setdefault("size_pt", 12)
        rule.setdefault("bold", True)
        rule.setdefault("alignment_value", WD_ALIGN_PARAGRAPH.LEFT)
        rule.setdefault("line_spacing_pt", 20)
        rule.setdefault("space_before_pt", 12)
        rule.setdefault("space_after_pt", 6)
    if category in {"figure_caption", "table_caption"}:
        rule.setdefault("font_east_asia", "宋体")
        rule.setdefault("font_latin", "Times New Roman")
        rule.setdefault("size_pt", 12)
        rule.setdefault("alignment_value", WD_ALIGN_PARAGRAPH.CENTER)
        rule.setdefault("line_spacing_value", 1.0)
    if category == "figure_caption":
        rule.setdefault("space_before_pt", 6)
        rule.setdefault("space_after_pt", 12)
    if category == "table_caption":
        rule.setdefault("space_before_pt", 12)
        rule.setdefault("space_after_pt", 6)
    if category == "reference":
        rule.setdefault("font_east_asia", "宋体")
        rule.setdefault("font_latin", "Times New Roman")
        rule.setdefault("size_pt", 12)
        rule.setdefault("line_spacing_pt", 16)
        rule.setdefault("space_before_pt", 3)
        rule.setdefault("space_after_pt", 0)
    if category == "keywords":
        rule.setdefault("font_east_asia", "宋体")
        rule.setdefault("font_latin", "Times New Roman")
        rule.setdefault("size_pt", 14)
        rule.setdefault("bold", True)
    if category == "english_keywords":
        rule.setdefault("font_east_asia", "Times New Roman")
        rule.setdefault("font_latin", "Times New Roman")
        rule.setdefault("size_pt", 14)
        rule.setdefault("bold", True)
    return rule

def parse_school_requirement_docx(file_path):
    document = load_document(file_path)
    paragraphs = read_docx_paragraphs(document)
    rules = {}
    citation_rule = {}
    raw_lines = []
    for para in paragraphs:
        text = clean_text(para["text"])
        if not text:
            continue
        raw_lines.append(text)
        found = extract_citation_rule(text)
        if found:
            citation_rule.update(found)
        category = infer_category(text)
        if not category:
            continue
        rule = rules.get(category, {})
        size_name, size_pt = extract_size(text)
        align_name, align_value = extract_alignment(text)
        line_name, line_value = extract_line_spacing(text)
        space_before, space_after = extract_spacing(text)
        indent = extract_first_line_indent(text)
        rule.update(extract_font_rule(text, category))
        if size_pt:
            rule["size_name"] = size_name
            rule["size_pt"] = size_pt
        if align_name:
            rule["alignment_name"] = align_name
            rule["alignment_value"] = align_value
        if line_name:
            rule["line_spacing_name"] = line_name
            if isinstance(line_value, float) and line_name.startswith("固定值"):
                rule["line_spacing_pt"] = line_value
            else:
                rule["line_spacing_value"] = line_value
        if space_before is not None:
            rule["space_before_pt"] = space_before
        if space_after is not None:
            rule["space_after_pt"] = space_after
        if indent is not None:
            rule["first_line_indent_cm"] = indent
        rule["source_text"] = text
        rules[category] = normalize_rule(category, rule)
    for category in ["body", "abstract", "english_abstract", "heading1", "heading2", "heading3", "figure_caption", "table_caption", "reference"]:
        rules[category] = normalize_rule(category, rules.get(category, {}))
    merged = DEFAULT_CITATION_RULE.copy()
    merged.update(citation_rule)
    return {"rules": rules, "citation_rule": merged, "raw_text_preview": raw_lines[:80], "rule_count": len(rules)}

def paragraph_category(item, in_reference=False, state=None):
    if state is None:
        state = {"main_started": False, "abstract_mode": None}
    text = item["text"].strip()
    paragraph = item["paragraph"]
    style_name = paragraph.style.name if paragraph.style else ""
    if in_reference:
        return "reference"
    if text in {"摘  要", "摘要"} or text == "中文摘要":
        state["abstract_mode"] = "abstract"
        return "heading1"
    if text == "ABSTRACT":
        state["abstract_mode"] = "english_abstract"
        return "heading1"
    if text in {"目  录", "目录"}:
        state["abstract_mode"] = None
        return "heading1"
    if text.startswith("关键词"):
        return "keywords"
    if text.startswith("Key Words") or text.startswith("Keywords") or text.startswith("Key Word"):
        return "english_keywords"
    if re.match(r"^图\s*\d+[.．-]\d+", text):
        return "figure_caption"
    if re.match(r"^表\s*\d+[.．-]\d+", text):
        return "table_caption"
    if re.match(r"^第[一二三四五六七八九十百]+章", text) or "Heading 1" in style_name or "标题 1" in style_name:
        state["main_started"] = True
        state["abstract_mode"] = None
        return "heading1"
    if re.match(r"^\d+\.\d+\.\d+\s+", text) or "Heading 3" in style_name or "标题 3" in style_name:
        return "heading3"
    if re.match(r"^\d+\.\d+\s+", text) or "Heading 2" in style_name or "标题 2" in style_name:
        return "heading2"
    if state.get("abstract_mode") and len(text) > 20 and not state.get("main_started"):
        return state["abstract_mode"]
    if state.get("main_started") and len(text) > 20:
        return "body"
    return None

def categorize_paragraphs(body_paragraphs, reference_paragraphs):
    result = []
    state = {"main_started": False, "abstract_mode": None}
    for item in body_paragraphs:
        category = paragraph_category(item, False, state)
        if category:
            result.append({**item, "category": category})
    for item in reference_paragraphs:
        category = paragraph_category(item, True, state)
        if category:
            result.append({**item, "category": category})
    return result

def get_run_font_name(run):
    rpr = run._element.rPr
    if rpr is not None and rpr.rFonts is not None:
        for attr in [qn("w:eastAsia"), qn("w:ascii"), qn("w:hAnsi")]:
            value = rpr.rFonts.get(attr)
            if value:
                return value
    return run.font.name

def get_paragraph_main_font(paragraph):
    fonts = []
    for run in paragraph.runs:
        if run.text.strip():
            name = get_run_font_name(run)
            if name:
                fonts.append(name)
    if fonts:
        return max(set(fonts), key=fonts.count)
    if paragraph.style and paragraph.style.font and paragraph.style.font.name:
        return paragraph.style.font.name
    return None

def get_paragraph_main_size(paragraph):
    sizes = []
    for run in paragraph.runs:
        if run.text.strip() and run.font.size:
            sizes.append(round(run.font.size.pt, 1))
    if sizes:
        return max(set(sizes), key=sizes.count)
    if paragraph.style and paragraph.style.font and paragraph.style.font.size:
        return round(paragraph.style.font.size.pt, 1)
    return None

def compare_float(a, b, tolerance=0.2):
    if a is None or b is None:
        return True
    return abs(float(a) - float(b)) <= tolerance

def check_school_format(categorized_paragraphs, school_rules):
    issues = []
    rules = school_rules.get("rules", {}) if school_rules else {}
    if not rules:
        t = "school_rule_not_extracted"
        issues.append(Issue(type=t, label=issue_label(t), group="school", problem="未能从学校格式要求 Word 中抽取到足够规则。", suggestion="建议确认学校要求中是否包含正文、标题、参考文献、字体、字号等明确文字。", auto_fixable=False))
        return issues
    for item in categorized_paragraphs:
        category = item["category"]
        rule = rules.get(category)
        if not rule:
            continue
        paragraph = item["paragraph"]
        text = item["text"]
        if rule.get("font_east_asia"):
            actual = get_paragraph_main_font(paragraph)
            expected = rule["font_east_asia"]
            if actual and expected not in actual and category not in {"english_abstract", "english_keywords"}:
                t = "school_font_mismatch"
                issues.append(Issue(type=t, label=issue_label(t), group="school", paragraph_index=item["index"], text=text, problem=f"{category} 字体疑似为 {actual}，学校要求中文为 {expected}。", suggestion="建议按学校要求调整该段字体。", auto_fixable=True, meta={"category": category, "expected": expected, "actual": actual}))
        if rule.get("size_pt"):
            actual_size = get_paragraph_main_size(paragraph)
            expected_size = rule["size_pt"]
            if actual_size and not compare_float(actual_size, expected_size):
                t = "school_size_mismatch"
                issues.append(Issue(type=t, label=issue_label(t), group="school", paragraph_index=item["index"], text=text, problem=f"{category} 字号疑似为 {actual_size} 磅，学校要求为 {rule.get('size_name', expected_size)}。", suggestion=f"建议将该段字号调整为 {rule.get('size_name', expected_size)}。", auto_fixable=True, meta={"category": category, "expected": expected_size, "actual": actual_size}))
        if rule.get("alignment_value") is not None and paragraph.alignment is not None and paragraph.alignment != rule["alignment_value"]:
            t = "school_alignment_mismatch"
            issues.append(Issue(type=t, label=issue_label(t), group="school", paragraph_index=item["index"], text=text, problem=f"{category} 对齐方式与学校要求不一致。", suggestion=f"建议调整为 {rule.get('alignment_name', '学校要求的对齐方式')}。", auto_fixable=True, meta={"category": category}))
    return issues
