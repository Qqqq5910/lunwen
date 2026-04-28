import unittest

from analyzer.structure_checker import check_numbered_items, extract_numbered_items


class StructureCheckerTest(unittest.TestCase):
    def test_body_references_do_not_count_as_duplicate_captions(self):
        paragraphs = [
            {"index": 1, "text": "如图 1.1 所示，系统包含多个模块。"},
            {"index": 2, "text": "图1.1 双路径因果问题解决框架图[35]"},
            {"index": 3, "text": "再次见图 1.1 可以发现流程差异。"},
            {"index": 4, "text": "图1.2 靶向推断攻击示意图[9]"},
        ]

        items = extract_numbered_items(paragraphs)
        issues = check_numbered_items(paragraphs)

        self.assertEqual([(item["kind"], item["chapter"], item["index"]) for item in items], [("figure", 1, 1), ("figure", 1, 2)])
        self.assertFalse([issue for issue in issues if issue.type == "figure_number_duplicate"])

    def test_repeated_caption_is_reported(self):
        paragraphs = [
            {"index": 1, "text": "表2-1 实验参数"},
            {"index": 2, "text": "表 2.1 消融实验结果"},
        ]

        issues = check_numbered_items(paragraphs)

        self.assertEqual(1, len([issue for issue in issues if issue.type == "table_number_duplicate"]))


if __name__ == "__main__":
    unittest.main()
