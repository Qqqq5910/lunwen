import json
from pathlib import Path
from datetime import datetime
from analyzer.labels import GROUP_LABELS


HIDDEN_DETAIL_GROUPS_WHEN_FIXED = {"fixed", "auto"}


def issues_requiring_attention(issues, fixed_info=None):
    """Return only issues that should be shown in the user-facing detail list.

    Once automatic repair is enabled, fixed/auto-fixable items should not keep
    appearing as reminders. The detail list is reserved for issues that still need
    review or cannot be safely repaired automatically.
    """
    fixed_enabled = bool((fixed_info or {}).get("enabled"))
    if not fixed_enabled:
        return issues
    return [
        issue for issue in issues
        if (issue.group not in HIDDEN_DETAIL_GROUPS_WHEN_FIXED) and not getattr(issue, "fixed", False)
    ]


def issue_type_counts_zh(issues):
    result = {}
    for issue in issues:
        key = issue.label or issue.type
        result[key] = result.get(key, 0) + 1
    return result


def group_counts(issues):
    result = {"fixed": 0, "auto": 0, "manual": 0, "reminder": 0, "school": 0}
    for issue in issues:
        key = issue.group if issue.group in result else "manual"
        result[key] += 1
    return result


def split_issues_by_group(issues):
    result = {"fixed": [], "auto": [], "manual": [], "reminder": [], "school": []}
    for issue in issues:
        key = issue.group if issue.group in result else "manual"
        result[key].append(issue.model_dump())
    return result


def compact_summary(body_paragraphs, reference_paragraphs, citations, citation_sequences, references, issues):
    cited_numbers = set()
    for citation in citations:
        cited_numbers.update(citation["numbers"])
    groups = group_counts(issues)
    return {
        "body_paragraph_count": len(body_paragraphs),
        "reference_paragraph_count": len(reference_paragraphs),
        "citation_marker_count": len(citations),
        "citation_sequence_count": len(citation_sequences),
        "cited_reference_number_count": len(cited_numbers),
        "reference_count": len(references),
        "total_issues": len(issues),
        "auto_fixable_count": groups["auto"],
        "manual_confirm_count": groups["manual"],
        "reminder_count": groups["reminder"],
        "school_format_count": groups["school"],
        "fixed_count": groups["fixed"]
    }


def build_report(filename, body_paragraphs, reference_paragraphs, citations, citation_sequences, references, issues, fixed_info=None, before_summary=None, school_rules=None, token=None):
    detail_issues = issues_requiring_attention(issues, fixed_info)
    return {
        "filename": filename,
        "summary": compact_summary(body_paragraphs, reference_paragraphs, citations, citation_sequences, references, issues),
        "before_summary": before_summary,
        "issue_type_counts": issue_type_counts_zh(detail_issues),
        "group_counts": group_counts(detail_issues),
        "group_labels": GROUP_LABELS,
        "issues_by_group": split_issues_by_group(detail_issues),
        "fixed": fixed_info,
        "school_rules": school_rules,
        "download_token": token,
        "issues_preview": [issue.model_dump() for issue in detail_issues[:10]]
    }


def write_reports(report, issues, output_dir, keep_history=True):
    base = Path(output_dir)
    base.mkdir(parents=True, exist_ok=True)
    json_path = base / "report.json"
    txt_path = base / "report.txt"
    json_path.write_text(json.dumps(report, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
    txt_path.write_text(render_txt(report), encoding="utf-8")
    history = {}
    if keep_history:
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        history_dir = base / "history"
        history_dir.mkdir(parents=True, exist_ok=True)
        history_json = history_dir / f"report_{stamp}.json"
        history_txt = history_dir / f"report_{stamp}.txt"
        history_json.write_text(json.dumps(report, ensure_ascii=False, indent=2, default=str), encoding="utf-8")
        history_txt.write_text(render_txt(report), encoding="utf-8")
        history = {"json": str(history_json), "txt": str(history_txt)}
    return {"json": str(json_path), "txt": str(txt_path), "history": history}


def render_txt(report):
    lines = []
    summary = report["summary"]
    lines.append("论文格式检查报告")
    lines.append("")
    lines.append("一、检查概况")
    lines.append(f"文件名：{report.get('filename', '')}")
    lines.append(f"正文引用标记数：{summary['citation_marker_count']}")
    lines.append(f"连续引用组数：{summary['citation_sequence_count']}")
    lines.append(f"正文实际引用编号数：{summary['cited_reference_number_count']}")
    lines.append(f"参考文献条目数：{summary['reference_count']}")
    lines.append(f"发现问题：{summary['total_issues']}")
    lines.append(f"需人工确认：{summary['manual_confirm_count']}")
    lines.append(f"格式提醒：{summary['reminder_count']}")
    lines.append(f"学校格式检查：{summary['school_format_count']}")
    fixed = report.get("fixed") or {}
    lines.append("")
    lines.append("二、自动修复")
    lines.append(f"是否启用：{fixed.get('enabled')}")
    lines.append(f"上标修复片段数：{fixed.get('superscript_fixed_count')}")
    lines.append(f"连续引用合并数：{fixed.get('citation_range_fixed_count')}")
    lines.append(f"学校格式修复段落数：{fixed.get('school_format_fixed_count')}")
    lines.append(f"异常空格清理数：{fixed.get('whitespace_fixed_count')}")
    lines.append(f"引用链接转换数：{fixed.get('hyperlink_converted_count')}")
    if fixed.get("output_docx"):
        lines.append(f"修复文件：{fixed.get('output_docx')}")
    before = report.get("before_summary")
    if before:
        lines.append("")
        lines.append("三、修复前后对比")
        lines.append(f"修复前问题数：{before.get('total_issues')}")
        lines.append(f"修复后问题数：{summary.get('total_issues')}")
        lines.append(f"修复前可自动修复项：{before.get('auto_fixable_count')}")
        lines.append(f"修复后可自动修复项：{summary.get('auto_fixable_count')}")
    school = report.get("school_rules") or {}
    if school:
        lines.append("")
        lines.append("四、学校格式规则")
        lines.append(f"抽取规则数量：{school.get('rule_count', 0)}")
        for key, rule in (school.get("rules") or {}).items():
            parts = []
            for field in ["font", "size_name", "alignment_name", "line_spacing_name", "first_line_indent_name"]:
                if rule.get(field):
                    parts.append(str(rule.get(field)))
            if parts:
                lines.append(f"{key}: {'，'.join(parts)}")
    lines.append("")
    lines.append("五、仍需关注的问题类型统计")
    counts = report.get("issue_type_counts", {})
    if counts:
        for key, value in counts.items():
            lines.append(f"{key}: {value}")
    else:
        lines.append("暂无需要人工关注的问题。")
    lines.append("")
    lines.append("六、仍需关注的问题明细")
    has_detail = False
    for group_key, group_label in report.get("group_labels", {}).items():
        items = (report.get("issues_by_group") or {}).get(group_key) or []
        if not items:
            continue
        has_detail = True
        lines.append("")
        lines.append(group_label)
        for idx, issue in enumerate(items, start=1):
            lines.append(f"{idx}. {issue.get('problem')}")
            if issue.get("paragraph_index") is not None:
                lines.append(f"段落：{issue.get('paragraph_index')}")
            if issue.get("text"):
                lines.append(f"原文：{issue.get('text')}")
            lines.append(f"建议：{issue.get('suggestion')}")
    if not has_detail:
        lines.append("暂无需要人工关注的问题。")
    return "\n".join(lines)
