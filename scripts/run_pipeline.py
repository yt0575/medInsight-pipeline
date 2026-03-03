#!/usr/bin/env python
# -*- coding: utf-8 -*-

from __future__ import annotations

import argparse
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from pipeline import core as orch


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Run disease report pipeline by stage or role")
    parser.add_argument("--disease", default=None, help="疾病名，例如：糖尿病")
    parser.add_argument("--from-readme", action="store_true", help="从README读取“疾病名：”配置（未传--disease时生效）")
    parser.add_argument("--readme", default="README.md", help="README文件路径（默认README.md）")
    parser.add_argument("--xlsx", default=None, help="第四章Excel路径，默认：<疾病名>第四章数据.xlsx")
    parser.add_argument("--template", default="template.docx", help="Word模板路径")
    parser.add_argument("--out-base", default="autofile", help="输出根目录")
    parser.add_argument("--reuse-text", dest="reuse_text", action="store_true", default=True, help="阶段3优先复用已有ch01~ch07与summary（默认开启）")
    parser.add_argument("--no-reuse-text", dest="reuse_text", action="store_false", help="阶段3忽略已有文本并重新生成")
    parser.add_argument("--stage", default="all", help="all/1/2/3/4/5 或 stage1...stage5")
    parser.add_argument("--role", default=None, help="all/evidence/content/docx/qa（设置后优先于--stage）")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    disease_name = orch.resolve_disease_name(
        disease=args.disease,
        from_readme=args.from_readme,
        readme_path=args.readme,
    )
    orch.run(
        disease=disease_name,
        stage=args.stage,
        xlsx=args.xlsx,
        template=args.template,
        out_base=args.out_base,
        role=args.role,
        reuse_text=args.reuse_text,
    )


if __name__ == "__main__":
    main()
