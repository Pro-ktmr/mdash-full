#!/usr/bin/env python3
"""Run all year-specific extractors and build combined QA reports."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any

from curriculum_extractor import SUPPORTED_YEARS, extract_year, write_json


def failure_record(year: str, university: dict[str, Any]) -> dict[str, Any]:
    incomplete_programs = [
        {
            "sheet_name": program["sheet_name"],
            "warnings": program["extraction"]["warnings"],
            "course_count": program["course_count"],
            "has_completion_requirement": bool(
                program["completion_requirement"]["text"]
            ),
        }
        for program in university["programs"]
        if program["extraction"]["status"] != "success"
    ]
    return {
        "year": year,
        "university_name": university["university_name"],
        "source_file": university["source_file"],
        "status": university["extraction"]["status"],
        "warnings": university["extraction"]["warnings"],
        "incomplete_programs": incomplete_programs,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--input-root", type=Path, default=Path("literacy_application")
    )
    parser.add_argument(
        "--output-dir", type=Path, default=Path("outputs") / "curriculum"
    )
    parser.add_argument("--progress", action="store_true")
    parser.add_argument(
        "--combine-only",
        action="store_true",
        help="Rebuild combined reports from existing year JSON files without rereading Excel files.",
    )
    args = parser.parse_args()

    by_year: dict[str, dict[str, Any]] = {}
    failures: list[dict[str, Any]] = []
    for year in SUPPORTED_YEARS:
        if args.combine_only:
            payload = json.loads(
                (args.output_dir / f"{year}.json").read_text(encoding="utf-8")
            )
        else:
            payload = extract_year(
                year,
                args.input_root / year,
                progress=args.progress,
            )
        by_year[year] = payload
        if not args.combine_only:
            write_json(args.output_dir / f"{year}.json", payload)
        failures.extend(
            failure_record(year, university)
            for university in payload["universities"]
            if university["extraction"]["status"] != "success"
        )

    combined = {
        "schema_version": next(iter(by_year.values()))["schema_version"],
        "years": by_year,
    }
    summary = {
        "schema_version": combined["schema_version"],
        "years": {year: payload["summary"] for year, payload in by_year.items()},
        "totals": {
            "university_count": sum(
                payload["summary"]["university_count"] for payload in by_year.values()
            ),
            "program_count": sum(
                payload["summary"]["program_count"] for payload in by_year.values()
            ),
            "course_count": sum(
                payload["summary"]["course_count"] for payload in by_year.values()
            ),
            "success": sum(
                payload["summary"]["status_counts"]["success"]
                for payload in by_year.values()
            ),
            "partial": sum(
                payload["summary"]["status_counts"]["partial"]
                for payload in by_year.values()
            ),
            "failed": sum(
                payload["summary"]["status_counts"]["failed"]
                for payload in by_year.values()
            ),
        },
        "incomplete_university_count": len(failures),
    }
    failure_payload = {
        "schema_version": combined["schema_version"],
        "summary": {
            "count": len(failures),
            "by_year": {
                year: sum(item["year"] == year for item in failures)
                for year in SUPPORTED_YEARS
            },
        },
        "universities": failures,
    }
    write_json(args.output_dir / "all_years.json", combined)
    write_json(args.output_dir / "summary.json", summary)
    write_json(args.output_dir / "incomplete_universities.json", failure_payload)
    print(json.dumps(summary, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
