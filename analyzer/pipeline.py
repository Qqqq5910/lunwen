from pathlib import Path
from datetime import datetime
import shutil
import uuid
from analyzer.docx_reader import load_document, read_docx_paragraphs
from analyzer.reference_checker import split_body_and_reference, extract_references_from_reference_paragraphs, get_reference_number_set, check_reference_numbers, check_citation_reference_mapping
from analyzer.citation_checker import find_citations, find_citation_sequences, get_cited_number_set, check_citation_format, check_citation_sequences
from analyzer.structure_checker import check_numbered_items, check_headings
from analyzer.school_rules import parse_school_requirement_docx, categorize_paragraphs, check_school_format
from analyzer.fixer import fix_superscript_in_body, fix_citation_ranges_in_body, fix_school_format, save_fixed_document
from analyzer.reporter import build_report, write_reports, compact_summary
from analyzer.security import generate_token, write_token


def get_citation_rule(school_rules):
    if school_rules and school_rules.get("citation_rule"):
        return school_rules.get("citation_rule")
    return {"bracket_style":"[]","range_separator":"~","list_separator":",","superscript":True}


def analyze_document_state(document, school_rules=None):
    citation_rule = get_citation_rule(school_rules)
    paragraphs = read_docx_paragraphs(document)
    body_paragraphs, reference_paragraphs, reference_title = split_body_and_reference(paragraphs)
    references = extract_references_from_reference_paragraphs(reference_paragraphs)
    citations = find_citations(body_paragraphs)
    citation_sequences = find_citation_sequences(body_paragraphs, citation_rule)
    cited_numbers = get_cited_number_set(citations)
    reference_numbers = get_reference_number_set(references)
    issues = []
    issues.extend(check_citation_format(citations, citation_rule))
    issues.extend(check_citation_sequences(citation_sequences))
    issues.extend(check_reference_numbers(references, reference_title is not None))
    issues.extend(check_citation_reference_mapping(cited_numbers, reference_numbers, references))
    issues.extend(check_numbered_items(body_paragraphs))
    issues.extend(check_headings(body_paragraphs))
    categorized = categorize_paragraphs(body_paragraphs, reference_paragraphs)
    if school_rules:
        issues.extend(check_school_format(categorized, school_rules))
    return {"paragraphs": paragraphs, "body_paragraphs": body_paragraphs, "reference_paragraphs": reference_paragraphs, "references": references, "citations": citations, "citation_sequences": citation_sequences, "issues": issues, "categorized": categorized}


def analyze_docx(file_path, school_requirement_path=None, fix_superscript=False, fix_citation_ranges=False, fix_school=False, keep_history=True, job_id=None, report_base_dir="reports"):
    source = Path(file_path)
    current_job_id = job_id or uuid.uuid4().hex
    token = generate_token()
    report_dir = Path(report_base_dir) / current_job_id
    write_token(report_dir, token)
    school_rules = parse_school_requirement_docx(school_requirement_path) if school_requirement_path else None
    citation_rule = get_citation_rule(school_rules)
    document = load_document(source)
    before_state = analyze_document_state(document, school_rules)
    before_summary = compact_summary(before_state["body_paragraphs"], before_state["reference_paragraphs"], before_state["citations"], before_state["citation_sequences"], before_state["references"], before_state["issues"])
    fixed_info = {"enabled": False, "superscript_fixed_count": 0, "citation_range_fixed_count": 0, "school_format_fixed_count": 0, "output_docx": None, "latest_docx": None, "download_url": None}
    final_document = document
    if fix_superscript or fix_citation_ranges or fix_school:
        superscript_fixed_count = 0
        citation_range_fixed_count = 0
        school_format_fixed_count = 0
        if fix_citation_ranges:
            citation_range_fixed_count = fix_citation_ranges_in_body(final_document, before_state["body_paragraphs"], citation_rule)
            before_state = analyze_document_state(final_document, school_rules)
        if fix_superscript:
            superscript_fixed_count = fix_superscript_in_body(final_document, before_state["body_paragraphs"], citation_rule)
            before_state = analyze_document_state(final_document, school_rules)
        if fix_school and school_rules:
            school_format_fixed_count = fix_school_format(final_document, before_state["categorized"], school_rules.get("rules", {}))
            before_state = analyze_document_state(final_document, school_rules)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        fixed_dir = report_dir / "fixed"
        fixed_dir.mkdir(parents=True, exist_ok=True)
        fixed_output = fixed_dir / f"fixed_{stamp}_{source.name}"
        latest_docx = fixed_dir / "fixed_latest.docx"
        save_fixed_document(final_document, fixed_output)
        shutil.copyfile(fixed_output, latest_docx)
        fixed_info = {"enabled": True, "superscript_fixed_count": superscript_fixed_count, "citation_range_fixed_count": citation_range_fixed_count, "school_format_fixed_count": school_format_fixed_count, "output_docx": str(fixed_output), "latest_docx": str(latest_docx), "download_url": f"/download/{current_job_id}/fixed/fixed_latest.docx?token={token}"}
    final_state = analyze_document_state(final_document, school_rules)
    report = build_report(source.name, final_state["body_paragraphs"], final_state["reference_paragraphs"], final_state["citations"], final_state["citation_sequences"], final_state["references"], final_state["issues"], fixed_info, before_summary, school_rules, token)
    report["job_id"] = current_job_id
    saved = write_reports(report, final_state["issues"], report_dir, keep_history)
    report["report_files"] = {"json": saved["json"], "txt": saved["txt"], "history": saved["history"], "txt_download_url": f"/download/{current_job_id}/report/report.txt?token={token}", "json_download_url": f"/download/{current_job_id}/report/report.json?token={token}"}
    return report
