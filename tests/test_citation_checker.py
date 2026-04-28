import unittest

from analyzer.citation_checker import check_citation_format, find_citations
from analyzer.citation_utils import expand_citation_numbers, normalize_marker


class FakeFont:
    def __init__(self, superscript=None):
        self.superscript = superscript


class FakeRun:
    def __init__(self, text, superscript=None):
        self.text = text
        self.font = FakeFont(superscript)
        self.style = None


class FakeParagraph:
    def __init__(self, runs):
        self.runs = runs
        self.text = "".join(run.text for run in runs)


class CitationCheckerTest(unittest.TestCase):
    def test_common_chinese_brackets_are_detected(self):
        self.assertEqual("[1,3]", normalize_marker("【1，3】"))
        self.assertEqual([1, 2, 3], expand_citation_numbers("〔1-3〕"))

    def test_non_superscript_chinese_bracket_citation_is_reported(self):
        paragraph = FakeParagraph([FakeRun("已有研究【12】说明该问题。", superscript=False)])
        citations = find_citations([{"index": 3, "text": paragraph.text, "paragraph": paragraph}])
        issues = check_citation_format(citations)

        self.assertEqual(["【12】"], [item["raw"] for item in citations])
        self.assertEqual(1, len([issue for issue in issues if issue.type == "citation_not_superscript"]))

    def test_superscript_marker_is_not_reported(self):
        paragraph = FakeParagraph([FakeRun("已有研究", None), FakeRun("[12]", True), FakeRun("说明该问题。", None)])
        citations = find_citations([{"index": 3, "text": paragraph.text, "paragraph": paragraph}])
        issues = check_citation_format(citations)

        self.assertEqual(1, len(citations))
        self.assertFalse([issue for issue in issues if issue.type == "citation_not_superscript"])


if __name__ == "__main__":
    unittest.main()
