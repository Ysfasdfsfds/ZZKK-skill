from __future__ import annotations

import argparse
import json
import sys
from pathlib import Path

CURRENT_DIR = Path(__file__).resolve().parent
if str(CURRENT_DIR) not in sys.path:
    sys.path.insert(0, str(CURRENT_DIR))

try:
    from lib.pipeline import (  # noqa: E402
        PipelineConfig,
        plan_outputs,
        run_pipeline,
        verify_outputs,
    )
except ModuleNotFoundError as exc:  # pragma: no cover
    if exc.name == "openpyxl":
        print(
            'Missing dependency: openpyxl. Install it with "python3 -m pip install -r requirements.txt".',
            file=sys.stderr,
        )
        raise SystemExit(1) from exc
    raise


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(description="ZZKK PRD workflow helper")
    sub = parser.add_subparsers(dest="command", required=True)

    def add_common(p: argparse.ArgumentParser) -> None:
        p.add_argument("--workbook", required=True)
        p.add_argument("--customer-template", required=True)
        p.add_argument("--product-template", required=True)
        p.add_argument("--review-template", required=True)
        p.add_argument("--module-code", required=True)
        p.add_argument("--module-name", required=True)
        p.add_argument("--draft-version", required=True)
        p.add_argument("--final-version", required=True)
        p.add_argument("--output-dir", required=True)
        p.add_argument("--run-date", default=None, help="YYYYMMDD, defaults to today")

    plan = sub.add_parser("plan", help="Show computed filenames and identifiers")
    add_common(plan)

    run = sub.add_parser("run", help="Generate the six output files")
    add_common(run)
    run.add_argument("--overwrite", action="store_true")

    verify = sub.add_parser("verify", help="Verify generated files and identifiers")
    add_common(verify)

    return parser


def config_from_args(args: argparse.Namespace) -> PipelineConfig:
    return PipelineConfig(
        workbook=Path(args.workbook).expanduser().resolve(),
        customer_template=Path(args.customer_template).expanduser().resolve(),
        product_template=Path(args.product_template).expanduser().resolve(),
        review_template=Path(args.review_template).expanduser().resolve(),
        module_code=args.module_code.strip(),
        module_name=args.module_name.strip(),
        draft_version=args.draft_version.strip(),
        final_version=args.final_version.strip(),
        output_dir=Path(args.output_dir).expanduser().resolve(),
        run_date=args.run_date.strip() if args.run_date else None,
        overwrite=getattr(args, "overwrite", False),
    )


def main() -> int:
    parser = build_parser()
    args = parser.parse_args()
    config = config_from_args(args)

    if args.command == "plan":
        print(json.dumps(plan_outputs(config), ensure_ascii=False, indent=2))
        return 0
    if args.command == "run":
        print(json.dumps(run_pipeline(config), ensure_ascii=False, indent=2))
        return 0
    if args.command == "verify":
        print(json.dumps(verify_outputs(config), ensure_ascii=False, indent=2))
        return 0
    parser.error(f"unsupported command: {args.command}")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
