ISSUE_LABELS = {
    "citation_not_superscript": "正文引用未设置为上标",
    "citation_should_not_superscript": "正文引用不应设置为上标",
    "citation_fullwidth_bracket": "正文引用使用全角方括号",
    "citation_spacing_or_comma": "正文引用逗号或空格不规范",
    "citation_sequence_not_compacted": "连续引用未合并",
    "citation_style_mismatch": "正文引用样式与学校要求不一致",
    "citation_missing_reference": "正文引用缺少对应参考文献",
    "reference_not_cited": "参考文献未在正文中引用",
    "reference_section_missing": "未识别到参考文献区域",
    "reference_items_missing": "未识别到参考文献条目",
    "reference_not_start_from_1": "参考文献编号未从 1 开始",
    "reference_duplicate_number": "参考文献编号重复",
    "reference_number_missing": "参考文献编号跳号",
    "reference_type_missing": "参考文献疑似缺少类型标识",
    "figure_number_jump": "图编号疑似跳号",
    "table_number_jump": "表编号疑似跳号",
    "equation_number_jump": "公式编号疑似跳号",
    "figure_number_duplicate": "图编号重复",
    "table_number_duplicate": "表编号重复",
    "equation_number_duplicate": "公式编号重复",
    "ambiguous_figure_reference": "图引用表述不明确",
    "ambiguous_table_reference": "表引用表述不明确",
    "heading_number_duplicate": "标题编号重复",
    "heading_number_jump": "标题编号疑似跳号",
    "school_rule_not_extracted": "学校规则抽取不足",
    "school_font_mismatch": "字体与学校要求不一致",
    "school_size_mismatch": "字号与学校要求不一致",
    "school_alignment_mismatch": "对齐方式与学校要求不一致",
    "school_line_spacing_mismatch": "行距与学校要求不一致",
    "school_first_line_indent_mismatch": "首行缩进与学校要求不一致"
}
GROUP_LABELS = {
    "fixed": "已自动修复",
    "auto": "可自动修复",
    "manual": "需人工确认",
    "reminder": "格式提醒",
    "school": "学校格式检查"
}
def issue_label(issue_type):
    return ISSUE_LABELS.get(issue_type, issue_type)
