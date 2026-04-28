import re
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from models.issue import Issue
from analyzer.labels import issue_label
from analyzer.docx_reader import load_document, read_docx_paragraphs

SIZE_MAP = {"初号": 42, "小初": 36, "一号": 26, "小一": 24, "二号": 22, "小二": 18, "三号": 16, "小三": 15, "四号": 14, "小四": 12, "五号": 10.5, "小五": 9, "六号": 7.5, "小六": 6.5}
SIZE_KEYS = ["小初", "初号", "小一", "一号", "小二", "二号", "小三", "三号", "小四", "四号", "小五", "五号", "小六", "六号"]
FONT_NAMES = ["Times New Roman", "Arial", "Cambria", "宋体", "黑体", "楷体", "仿宋", "微软雅黑"]
ALIGNMENT_MAP = {"居中": WD_ALIGN_PARAGRAPH.CENTER, "左对齐": WD_ALIGN_PARAGRAPH.LEFT, "右对齐": WD_ALIGN_PARAGRAPH.RIGHT, "两端对齐": WD_ALIGN_PARAGRAPH.JUSTIFY}
DEFAULT_CITATION_RULE = {"bracket_style": "[]", "range_separator": "~", "list_separator": ",", "superscript": True}


def clean_text(text):
    return re.sub(r"\s+", " ", text.strip())


def extract_size(text):
    for name in SIZE_KEYS:
        if name in text:
            return name, SIZE_MAP[name]
    match = re.search(r"(\d+(?:\.\d+)?)\s*磅", text)
    if match:
        return f"{match.group(1)}磅", float(match.group(1))
    return None, None


def extract_alignment(text):
    for key, value in ALIGNMENT_MAP.items():
        if key in text:
            return key, value
    return None, None


def extract_line_spacing(text):
    match = re.search(r"固定值\s*(\d+(?:\.\d+)?)\s*磅", text)
    if match:
        return f"固定值{match.group(1)}磅", float(match.group(1))
    match = re.search(r"(\d+(?:\.\d+)?)\s*倍行距", text)
    if match:
        return f"{match.group(1)}倍行距", float(match.group(1))
    if "单倍行距" in text:
        return "单倍行距", 1.0
    return None, None


def extract_spacing(text):
    before = None
    after = None
    match = re.search(r"段前空\s*(\d+(?:\.\d+)?)\s*磅", text)
    if match:
        before = float(match.group(1))
    match = re.search(r"段后空\s*(\d+(?:\.\d+)?)\s*磅", text)
    if match:
        after = float(match.group(1))
    return before, after


def extract_first_line_indent(text):
    if "首行缩进" not in text:
        return None
    match = re.search(r"首行缩进\s*(\d+(?:\.\d+)?)\s*(?:个)?汉?字符", text)
    if match:
        return float(match.group(1)) * 0.37
    if "两个汉字符" in text or "2字符" in text or "2 字符" in text or "两字符" in text:
        return 0.74
    return None


def infer_category(text):
    if "中文摘要部分的标题" in text:
        return "abstract_title"
    if "英文摘要部分的标题" in text:
        return "english_abstract_title"
    if "目录部分的标题" in text:
        return "toc_title"
    if "摘要内容" in text and "英文" not in text:
        return "abstract"
    if "英文摘要内容" in text:
        return "english_abstract"
    if "中文论文关键词" in text:
        return "keywords"
    if "英文论文关键词" in text:
        return "english_keywords"
    if "参考文献的正文" in text or "成果内容格式与参考文献格式要求相同" in text:
        return "reference"
    if "参考文献”四个字" in text or "参考文献\"四个字" in text:
        return "reference_title"
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
    return None


def extract_font_rule(text, category):
    rule = {}
    if "中文用宋体" in text:
        rule["font_east_asia"] = "宋体"
    else:
        for font in FONT_NAMES:
            if font in text and font != "Times New Roman":
                rule["font_east_asia"] = font
                break
    if "Times New Roman" in text:
        rule["font_latin"] = "Times New Roman"
    if category in {"english_abstract", "english_abstract_title", "english_keywords"}:
        rule["font_east_asia"] = "Times New Roman"
        rule["font_latin"] = "Times New Roman"
    if "加粗" in text:
        rule["bold"] = True
    return rule


def extract_citation_rule(text):
    rule = {}
    if not any(key in text for key in ["引用", "引文", "文献标注", "参考文献标注", "正文标注", "顺序编码"]):
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


def apply_default(rule, **kwargs):
    for key, value in kwargs.items():
        rule[key] = value
    return rule


def normalize_rule(category, rule):
    if category == "body":
        return apply_default(rule, font_east_asia="宋体", font_latin="Times New Roman", size_pt=12, alignment_value=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing_pt=20, first_line_indent_cm=0.74, space_before_pt=0, space_after_pt=0, bold=False)
    if category == "abstract":
        return apply_default(rule, font_east_asia="宋体", font_latin="Times New Roman", size_pt=12, alignment_value=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing_pt=20, first_line_indent_cm=0.74, space_before_pt=0, space_after_pt=0, bold=False)
    if category == "english_abstract":
        return apply_default(rule, font_east_asia="Times New Roman", font_latin="Times New Roman", size_pt=12, alignment_value=WD_ALIGN_PARAGRAPH.JUSTIFY, line_spacing_pt=20, first_line_indent_cm=0.74, space_before_pt=0, space_after_pt=0, bold=False)
    if category in {"heading1", "abstract_title", "toc_title", "reference_title"}:
        return apply_default(rule, font_east_asia="黑体", font_latin="Times New Roman", size_pt=18, alignment_value=WD_ALIGN_PARAGRAPH.CENTER, space_before_pt=24, space_after_pt=18, line_spacing_value=1.0)
    if category == "english_abstract_title":
        return apply_default(rule, font_east_asia="Times New Roman", font_latin="Times New Roman", size_pt=18, alignment_value=WD_ALIGN_PARAGRAPH.CENTER, space_before_pt=24, space_after_pt=18, line_spacing_value=1.0)
    if category == "heading2":
        return apply_default(rule, font_east_asia="宋体", font_latin="Times New Roman", size_pt=14, bold=True, alignment_value=WD_ALIGN_PARAGRAPH.LEFT, line_spacing_pt=20, space_before_pt=24, space_after_pt=6)
    if category == "heading3":
        return apply_default(rule, font_east_asia="宋体", font_latin="Times New Roman", size_pt=12, bold=True, alignment_value=WD_ALIGN_PARAGRAPH.LEFT, line_spacing_pt=20, space_before_pt=12, space_after_pt=6)
    if category == "figure_caption":
        return apply_default(rule, font_east_asia="宋体", font_latin="Times New Roman", size_pt=12, alignment_value=WD_ALIGN_PARAGRAPH.CENTER, line_spacing_value=1.0, space_before_pt=6, space_after_pt=12)
    if category == "table_caption":
        return apply_default(rule, font_east_asia="宋体", font_latin="Times New Roman", size_pt=12, alignment_value=WD_ALIGN_PARAGRAPH.CENTER, line_spacing_value=1.0, space_before_pt=12, space_after_pt=6)
    if category == "reference":
        return apply_default(rule, font_east_asia="宋体", font_latin="Times New Roman", size_pt=12, alignment_value=WD_ALIGN_PARAGRAPH.LEFT, line_spacing_pt=16, space_before_pt=3, space_after_pt=0, bold=False)
    if category == "keywords":
        return apply_default(rule, font_east_asia="宋体", font_latin="Times New Roman", size_pt=14, alignment_value=WD_ALIGN_PARAGRAPH.LEFT, line_spacing_pt=20, space_before_pt=0, space_after_pt=0, bold=True)
    if category == "english_keywords":
        return apply_default(rule, font_east_asia="Times New Roman", font_latin="Times New Roman", size_pt=14, alignment_value=WD_ALIGN_PARAGRAPH.LEFT, line_spacing_pt=20, space_before_pt=0, space_after_pt=0, bold=True)
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
    for category in ["body", "abstract_title", "abstract", "english_abstract_title", "english_abstract", "heading1", "heading2", "heading3", "figure_caption", "table_caption", "reference_title", "reference", "keywords", "english_keywords"]:
        rules[category] = normalize_rule(category, rules.get(category, {}))
    merged = DEFAULT_CITATION_RULE.copy()
    merged.update(citation_rule)
    return {"rules": rules, "citation_rule": merged, "raw_text_preview": raw_lines[:80], "rule_count": len(rules)}


def _is_heading_style(style_name, level):
    if level == 1:
        return "Heading 1" in style_name or "标题 1" in style_name or style_name in {"1", "2"}
    if level == 2:
        return "Heading 2" in style_name or "标题 2" in style_name or style_name in {"3"}
    if level == 3:
        return "Heading 3" in style_name or "标题 3" in style_name or style_name in {"4"}
    return False


def _is_chapter_heading(text, style_name):
    if _is_heading_style(style_name, 1):
        return len(text) <= 80
    if re.match(r"^第[一二三四五六七八九十百]+章", text) and len(text) <= 80:
        return True
    if text in {"绪论", "总结与展望", "致谢", "致  谢"}:
        return True
    return False


def paragraph_category(item, in_reference=False, state=None):
    if state is None:
        state = {"main_started": False, "abstract_mode": None, "toc_started": False}
    text = item["text"].strip()
    paragraph = item["paragraph"]
    style_name = paragraph.style.name if paragraph.style else ""
    if in_reference:
        return "reference"
    if text in {"参考文献", "参 考 文 献"}:
        return "reference_title"
    if text in {"摘  要", "摘要", "中文摘要"}:
        state["abstract_mode"] = "abstract"
        state["toc_started"] = False
        return "abstract_title"
    if text == "ABSTRACT":
        state["abstract_mode"] = "english_abstract"
        state["toc_started"] = False
        return "english_abstract_title"
    if text in {"目  录", "目录"}:
        state["abstract_mode"] = None
        state["toc_started"] = True
        return "toc_title"
    if text.startswith("关键词"):
        return "keywords"
    if text.startswith("Key Words") or text.startswith("Keywords") or text.startswith("Key Word"):
        return "english_keywords"
    if state.get("abstract_mode") and not state.get("main_started") and len(text) > 20:
        return state["abstract_mode"]
    if re.match(r"^图\s*\d+[.．-]\d+", text):
        return "figure_caption"
    if re.match(r"^表\s*\d+[.．-]\d+", text):
        return "table_caption"
    if _is_chapter_heading(text, style_name):
        state["main_started"] = True
        state["abstract_mode"] = None
        state["toc_started"] = False
        return "heading1"
    if _is_heading_style(style_name, 3) or re.match(r"^\d+\.\d+\.\d+\s+", text):
        return "heading3" if len(text) <= 100 else "body"
    if _is_heading_style(style_name, 2) or re.match(r"^\d+\.\d+\s+", text):
        return "heading2" if len(text) <= 100 else "body"
    if state.get("main_started") and len(text) > 20:
        return "body"
    return None


def categorize_paragraphs(body_paragraphs, reference_paragraphs):
    result = []
    state = {"main_started": False, "abstract_mode": None, "toc_started": False}
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
        issue_type = "school_rule_not_extracted"
        issues.append(Issue(type=issue_type, label=issue_label(issue_type), group="school", problem="未能从学校格式要求 Word 中抽取到足够规则。", suggestion="建议确认学校要求中是否包含正文、标题、参考文献、字体、字号等明确文字。", auto_fixable=False))
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
            if actual and expected not in actual and category not in {"english_abstract", "english_abstract_title", "english_keywords"}:
                issue_type = "school_font_mismatch"
                issues.append(Issue(type=issue_type, label=issue_label(issue_type), group="school", paragraph_index=item["index"], text=text, problem=f"{category} 字体疑似为 {actual}，学校要求中文为 {expected}。", suggestion="建议按学校要求调整该段字体。", auto_fixable=True, meta={"category": category, "expected": expected, "actual": actual}))
        if rule.get("size_pt"):
            actual_size = get_paragraph_main_size(paragraph)
            expected_size = rule["size_pt"]
            if actual_size and not compare_float(actual_size, expected_size):
                issue_type = "school_size_mismatch"
                issues.append(Issue(type=issue_type, label=issue_label(issue_type), group="school", paragraph_index=item["index"], text=text, problem=f"{category} 字号疑似为 {actual_size} 磅，学校要求为 {rule.get('size_name', expected_size)}。", suggestion=f"建议将该段字号调整为 {rule.get('size_name', expected_size)}。", auto_fixable=True, meta={"category": category, "expected": expected_size, "actual": actual_size}))
        if rule.get("alignment_value") is not None and paragraph.alignment is not None and paragraph.alignment != rule["alignment_value"]:
            issue_type = "school_alignment_mismatch"
            issues.append(Issue(type=issue_type, label=issue_label(issue_type), group="school", paragraph_index=item["index"], text=text, problem=f"{category} 对齐方式与学校要求不一致。", suggestion=f"建议调整为 {rule.get('alignment_name', '学校要求的对齐方式')}。", auto_fixable=True, meta={"category": category}))
    return issues
