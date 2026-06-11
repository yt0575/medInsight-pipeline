import unittest
import json
from pathlib import Path
from tempfile import TemporaryDirectory

from pipeline import core


class TemplateProfileTests(unittest.TestCase):
    def tearDown(self):
        core._TEMPLATE_PROFILE_CACHE = None
        core.configure_runtime(
            disease_name="示例医学主题",
            excel_path=core.default_excel_path("示例医学主题"),
            template_path=Path("template.docx"),
            out_base=Path("autofile"),
            template_id="legacy_default",
        )

    def test_builtin_registry_contains_legacy_and_template_ids(self):
        registry = core.load_template_registry()

        self.assertIn("legacy_default", registry)
        self.assertIn("disease_template_01", registry)
        self.assertIn("drug_template_03", registry)
        self.assertTrue(core.TEMPLATE_PROFILE_CONFIG_PATH.exists())
        self.assertEqual(registry["disease_template_01"].path, Path("template") / "疾病对应药品市场报告模板" / "template1.docx")
        self.assertEqual(registry["drug_template_03"].market_data_chapter, 4)

    def test_template_profile_config_contains_explicit_chapters_and_blocks(self):
        raw = json.loads(core.TEMPLATE_PROFILE_CONFIG_PATH.read_text(encoding="utf-8"))
        entry = next(item for item in raw["templates"] if item["template_id"] == "disease_template_09")

        self.assertIn("chapters", entry)
        self.assertIn("blocks", entry)
        self.assertEqual(entry["chapters"]["9"], "第九章 市场机会与风险分析")
        self.assertTrue(any(block["block_id"] == "9.4" for block in entry["blocks"]))

    def test_template_id_drives_runtime_chapters_and_block_files(self):
        core.configure_runtime(
            disease_name="测试主题",
            template_id="disease_template_09",
            out_base=Path("autofile"),
        )

        specs = core.runtime_block_specs()

        self.assertEqual(core.ACTIVE_TEMPLATE_ID, "disease_template_09")
        self.assertEqual(core.TEMPLATE_PATH, Path("template") / "疾病对应药品市场报告模板" / "template9.docx")
        self.assertEqual(core.template_chapters(), list(range(1, 10)))
        self.assertEqual(max(spec.chapter for spec in specs), 9)
        self.assertTrue(any(spec.block_id == "9.4" and "本章小结" in spec.subtitle for spec in specs))
        self.assertEqual(core.chapter_title(9), "第九章 市场机会与风险分析")
        self.assertEqual(core.chapter_text_filename(9), "ch09.txt")

    def test_toc_driven_profile_recovers_missing_numbered_chapters(self):
        core.configure_runtime(
            disease_name="测试主题",
            template_id="disease_template_01",
            out_base=Path("autofile"),
        )

        profile = core.active_template_profile()
        specs = core.runtime_block_specs()

        self.assertEqual(core.template_chapters(), [1, 2, 3, 4, 5, 6, 7])
        self.assertEqual(profile.chapters[3], "第三章 临床诊断与治疗现状")
        self.assertTrue(any(spec.block_id == "3.1" and spec.chapter == 3 for spec in specs))

    def test_drug_template_can_recover_eight_chapter_structure_from_toc(self):
        core.configure_runtime(
            disease_name="测试主题",
            template_id="drug_template_01",
            out_base=Path("autofile"),
        )

        profile = core.active_template_profile()

        self.assertEqual(core.template_chapters(), [1, 2, 3, 4, 5, 6, 7, 8])
        self.assertEqual(profile.market_data_chapter, 3)
        self.assertEqual(profile.chapters[8], "第八章 风险分析与战略建议")

    def test_excel_block_scope_is_supported(self):
        core.configure_runtime(
            disease_name="测试主题",
            template_id="disease_template_05",
            out_base=Path("autofile"),
        )

        profile = core.active_template_profile()

        self.assertEqual(profile.market_data_chapter, 2)
        self.assertEqual(profile.market_data_block_id, "2.1")
        self.assertEqual(core.market_data_base_fig_id(), "fig_2_4")
        self.assertIn("market_quarterly_trend", core.market_figure_id_map())
        self.assertTrue(core.is_market_data_block(core.BlockSpec("2.1", 2, "2.1 标准治疗药物", 100, [], "", "")))
        self.assertFalse(core.is_market_data_block(core.BlockSpec("2.2", 2, "2.2 已上市创新药", 100, [], "", "")))

    def test_market_figure_ids_shift_when_market_chapter_has_semantic_figures(self):
        core.configure_runtime(
            disease_name="测试主题",
            template_id="disease_template_09",
            out_base=Path("autofile"),
        )

        fig_map = core.market_figure_id_map()

        self.assertEqual(fig_map["market_quarterly_trend"], "fig_5_5")
        self.assertEqual(fig_map["market_cr5"], "fig_5_12")

    def test_neutral_market_data_file_precedes_legacy_ch04(self):
        with TemporaryDirectory() as tmp:
            out_base = Path(tmp)
            core.configure_runtime(
                disease_name="测试主题",
                template_id="drug_template_03",
                out_base=out_base,
            )
            core.ensure_runtime_dirs()
            neutral = core.OUT_ROOT / core.MARKET_DATA_CODEX_EXTRACT_NAME
            legacy = core.OUT_ROOT / "ch04_codex_extract.json"
            neutral.write_text('{"schema_version":"market_data_codex_extract_v1"}', encoding="utf-8")
            legacy.write_text('{"schema_version":"ch4_codex_extract_v1"}', encoding="utf-8")

            self.assertEqual(core.market_data_extract_path(), neutral)

            neutral.unlink()
            self.assertEqual(core.market_data_extract_path(), legacy)

    def test_unknown_template_id_raises_helpful_error(self):
        with self.assertRaisesRegex(ValueError, "Unknown template_id"):
            core.resolve_template_profile("missing_template", None)

    def test_registry_can_be_driven_by_explicit_config_without_docx(self):
        with TemporaryDirectory() as tmp:
            config_path = Path(tmp) / "template_profiles.json"
            config_path.write_text(
                json.dumps(
                    {
                        "templates": [
                            {
                                "template_id": "tmp_profile",
                                "family": "custom",
                                "path": "missing-template.docx",
                                "chapters": {"1": "第一章 自定义章节"},
                                "blocks": [
                                    {
                                        "block_id": "1.1",
                                        "chapter": 1,
                                        "subtitle": "1.1 自定义小节",
                                        "target_chars": 1200,
                                        "topics": ["自定义主题"],
                                        "evidence_ids": "E01",
                                        "fig_ids": "fig_1_1",
                                    }
                                ],
                                "market_data_scope": {"chapter": 1},
                                "chapter_min_chars": {"1": 3000},
                            }
                        ]
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )

            original_config_path = core.TEMPLATE_PROFILE_CONFIG_PATH
            original_cache = core._TEMPLATE_PROFILE_CACHE
            try:
                core.TEMPLATE_PROFILE_CONFIG_PATH = config_path
                core._TEMPLATE_PROFILE_CACHE = None
                registry = core.load_template_registry()
            finally:
                core.TEMPLATE_PROFILE_CONFIG_PATH = original_config_path
                core._TEMPLATE_PROFILE_CACHE = original_cache

            profile = registry["tmp_profile"]
            self.assertEqual(profile.chapters[1], "第一章 自定义章节")
            self.assertEqual(profile.blocks[0].subtitle, "1.1 自定义小节")
            self.assertEqual(profile.market_data_chapter, 1)

    def test_codex_blueprint_references_template_chapter_files(self):
        core.configure_runtime(
            disease_name="测试主题",
            template_id="disease_template_09",
            out_base=Path("autofile"),
        )

        blueprint = core.build_codex_content_blueprint(core.runtime_block_specs())

        self.assertIn("ch09.txt", blueprint)
        self.assertIn("ch01.txt", blueprint)
        self.assertIn("summary.txt", blueprint)

    def test_assist_mode_error_mentions_profile_chapter_files(self):
        with TemporaryDirectory() as tmp:
            out_base = Path(tmp)
            core.configure_runtime(
                disease_name="测试主题",
                template_id="disease_template_10",
                out_base=out_base,
            )

            with self.assertRaisesRegex(RuntimeError, "ch08.txt"):
                core.run_assist_pipeline()

    def test_resolved_fig_ids_for_block_can_read_manifest_assignments(self):
        spec = core.BlockSpec("5.1", 5, "5.1 示例", 100, [], "", "")
        fig_rows = [
            {"fig_id": "fig_5_5", "插入到哪个block之后": "5.1"},
            {"fig_id": "fig_5_6", "插入到哪个block之后": "5.1"},
        ]

        self.assertEqual(core.resolved_fig_ids_for_block(spec, fig_rows), ["fig_5_5", "fig_5_6"])

    def test_assemble_docx_can_render_eight_chapter_template(self):
        with TemporaryDirectory() as tmp:
            out_base = Path(tmp)
            core.configure_runtime(
                disease_name="测试主题",
                template_id="drug_template_01",
                out_base=out_base,
            )
            core.ensure_runtime_dirs()

            specs = core.runtime_block_specs()
            block_text = {}
            for spec in specs:
                body = f"{spec.block_id}.1 小标题\n这是 {spec.block_id} 的测试正文，包含2025Q1、72小时和CR5等锚点。"
                block_text[spec.block_id] = body
            summary_text = "这是总结，包含2025Q1和关键发现。"
            refs_text = "[1] 机构. 标题[EB/OL]. 2025. https://example.com"

            core.assemble_docx(specs, block_text, summary_text, refs_text, [])

            self.assertTrue(core.FINAL_DOCX.exists())

    def test_refresh_progress_bootstraps_codex_prompt_assets_for_new_topic(self):
        with TemporaryDirectory() as tmp:
            out_base = Path(tmp)
            topic = "骨关节炎"
            source_xlsx = Path("data") / f"{topic}.xlsx"
            temp_data_dir = out_base / "data"
            temp_data_dir.mkdir(parents=True, exist_ok=True)
            copied_xlsx = temp_data_dir / source_xlsx.name
            copied_xlsx.write_bytes(source_xlsx.read_bytes())

            core.run_refresh_progress(
                disease=topic,
                template_id="legacy_default",
                out_base=str(out_base),
            )

            topic_dir = out_base / topic
            self.assertTrue((topic_dir / core.CODEX_CONTENT_BLUEPRINT_NAME).exists())
            self.assertTrue((topic_dir / core.MARKET_DATA_CODEX_PROMPT_NAME).exists())
            self.assertTrue((topic_dir / "00_evidence.txt").exists())

    def test_legacy_parser_supports_combined_channel_sheets(self):
        xlsx = Path("data") / "痤疮抗炎治疗药物.xlsx"

        ch4 = core.build_ch4_data_from_legacy_parser(xlsx)

        self.assertEqual(ch4.latest_quarter, "2025Q4")
        self.assertFalse(ch4.quarterly.empty)
        self.assertGreater(float(ch4.quarterly.iloc[-1]["hospital"]), 0.0)
        self.assertGreater(float(ch4.quarterly.iloc[-1]["drugstore"]), 0.0)
        self.assertGreater(float(ch4.quarterly.iloc[-1]["online"]), 0.0)
        self.assertFalse(ch4.top_hospital.empty)
        self.assertFalse(ch4.top_drugstore.empty)
        self.assertFalse(ch4.top_online.empty)

    def test_txt_stage_check_rejects_internal_artifact_mentions(self):
        core.configure_runtime(
            disease_name="测试主题",
            template_id="legacy_default",
            out_base=Path("autofile"),
        )
        specs = core.runtime_block_specs()
        block_text = {}
        for spec in specs:
            body = f"{spec.block_id}.1 小标题\n这是 {spec.block_id} 的测试正文，包含2025Q3、72小时和CR5锚点。"
            block_text[spec.block_id] = body
        block_text["4.1"] += "\n根据 market_data_codex_extract.json 可见趋势。"
        summary_text = "这是总结，包含2025Q3和72小时锚点。"

        report, passed = core.run_txt_stage_checks(specs, block_text, summary_text)

        self.assertFalse(passed)
        self.assertIn("内部文件名", report)


if __name__ == "__main__":
    unittest.main()
