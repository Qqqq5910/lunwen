import argparse
import json
from analyzer.pipeline import analyze_docx


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("docx_path")
    parser.add_argument("--school-requirement")
    parser.add_argument("--fix-superscript", action="store_true")
    parser.add_argument("--fix-citation-ranges", action="store_true")
    parser.add_argument("--fix-school-format", action="store_true")
    parser.add_argument("--fix", action="store_true")
    parser.add_argument("--no-history", action="store_true")
    args = parser.parse_args()
    report = analyze_docx(args.docx_path, school_requirement_path=args.school_requirement, fix_superscript=args.fix_superscript or args.fix, fix_citation_ranges=args.fix_citation_ranges or args.fix, fix_school=args.fix_school_format, keep_history=not args.no_history, job_id="local")
    preview = {"summary": report["summary"], "before_summary": report["before_summary"], "issue_type_counts": report["issue_type_counts"], "group_counts": report["group_counts"], "fixed": report["fixed"], "school_rules": {"rule_count": (report.get("school_rules") or {}).get("rule_count", 0), "rules": (report.get("school_rules") or {}).get("rules", {})}, "report_files": report["report_files"], "issues_preview": report["issues_preview"]}
    print(json.dumps(preview, ensure_ascii=False, indent=2, default=str))


if __name__ == "__main__":
    main()
