import unittest

from pipeline.core import (
    BlockSpec,
    anchor_metric_paragraphs,
    build_block_render_elements,
    clean_docx_body_text,
    internal_artifact_mentions,
    normalize_reference_line,
    split_paragraphs,
)


class TextOutputCleanupTests(unittest.TestCase):
    def test_single_newline_blocks_are_rendered_as_separate_paragraphs(self):
        text = "1.1.1 临床边界\n第一段仍然是正文[1]\n第二段也应成为独立自然段[2]"

        self.assertEqual(
            split_paragraphs(text),
            ["1.1.1 临床边界", "第一段仍然是正文[1]", "第二段也应成为独立自然段[2]"],
        )

    def test_internal_newlines_do_not_survive_inside_docx_body_text(self):
        text = "诊疗路径[1][2]\n仍需在72小时内复评 [ 3 ]。"

        self.assertEqual(clean_docx_body_text(text), "诊疗路径仍需在72小时内复评。")

    def test_reference_cleanup_removes_xlex_token(self):
        line = "[1] 机构. 标题[EB/OL]. 2024. xlex https://example.com"

        self.assertEqual(normalize_reference_line(line), "[1] 机构. 标题[EB/OL]. 2024. https://example.com")

    def test_render_elements_keep_headings_and_clean_body_paragraphs(self):
        spec = BlockSpec("1.1", 1, "1.1 测试", 100, [], "", "")
        elements = build_block_render_elements(spec, "1.1.1 标题\n正文第一段[1]\n正文第二段[2]")

        cleaned_body = [clean_docx_body_text(text) for kind, text in elements if kind == "body"]

        self.assertEqual(elements[0], ("heading3", "1.1.1 标题"))
        self.assertEqual(cleaned_body, ["正文第一段", "正文第二段"])
        self.assertTrue(all("\n" not in text for text in cleaned_body))

    def test_anchor_metrics_ignore_explicit_third_level_headings(self):
        text = "1.1.1 标题\n正文第一段包含2025Q3锚点\n正文第二段包含72小时锚点"

        self.assertEqual(
            anchor_metric_paragraphs(text),
            ["正文第一段包含2025Q3锚点", "正文第二段包含72小时锚点"],
        )

    def test_internal_artifact_mentions_detect_report_internal_filenames(self):
        text = "根据 market_data_codex_extract.json 与 骨关节炎.xlsx 可见结论。"

        hits = internal_artifact_mentions(text)

        self.assertIn("market_data_codex_extract.json", hits)
        self.assertIn("骨关节炎.xlsx", hits)


if __name__ == "__main__":
    unittest.main()
