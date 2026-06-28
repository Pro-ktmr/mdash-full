#!/usr/bin/env python3
"""Validate generated curriculum JSON files and write a compact QA report."""

from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path
from typing import Any

from curriculum_extractor import SUPPORTED_YEARS, canonical_course_name, write_json


def load_json(path: Path) -> Any:
    return json.loads(path.read_text(encoding="utf-8"))


def validate_year(path: Path) -> tuple[dict[str, Any], list[str]]:
    payload = load_json(path)
    errors: list[str] = []
    statuses = Counter()
    program_count = 0
    course_count = 0
    for university in payload["universities"]:
        status = university["extraction"]["status"]
        statuses[status] += 1
        source_file = Path(university["source_file"])
        if not source_file.exists():
            errors.append(f"missing_source_file:{source_file}")
        if status == "success" and not university["programs"]:
            errors.append(f"success_without_program:{university['university_name']}")
        for program in university["programs"]:
            program_count += 1
            if program["course_count"] != len(program["courses"]):
                errors.append(
                    f"course_count_mismatch:{university['university_name']}:{program['sheet_name']}"
                )
            course_count += len(program["courses"])
            if program["extraction"]["status"] == "success":
                if not program["completion_requirement"]["text"]:
                    errors.append(
                        f"success_without_requirement:{university['university_name']}:{program['sheet_name']}"
                    )
                if not program["courses"]:
                    errors.append(
                        f"success_without_courses:{university['university_name']}:{program['sheet_name']}"
                    )
            seen_names: set[str] = set()
            for course in program["courses"]:
                key = canonical_course_name(course["name"])
                if not key:
                    errors.append(
                        f"blank_course_name:{university['university_name']}:{program['sheet_name']}"
                    )
                elif key in seen_names:
                    errors.append(
                        f"duplicate_course:{university['university_name']}:{program['sheet_name']}:{course['name']}"
                    )
                seen_names.add(key)
                if course["name"].startswith("=") or course["name"].startswith("※"):
                    errors.append(
                        f"invalid_course_name:{university['university_name']}:{course['name'][:80]}"
                    )
                credits = course["credits"]
                if isinstance(credits, (int, float)) and not (0 <= credits <= 100):
                    errors.append(
                        f"invalid_credits:{university['university_name']}:{course['name']}:{credits}"
                    )
                if not course["source_occurrences"]:
                    errors.append(
                        f"course_without_source:{university['university_name']}:{course['name']}"
                    )
                for occurrence in course["source_occurrences"]:
                    if not occurrence.get("sheet_name") or not occurrence.get("source_cell"):
                        errors.append(
                            f"incomplete_course_source:{university['university_name']}:{course['name']}"
                        )

    summary = payload["summary"]
    expected_statuses = {
        "success": statuses.get("success", 0),
        "partial": statuses.get("partial", 0),
        "failed": statuses.get("failed", 0),
    }
    if summary["university_count"] != len(payload["universities"]):
        errors.append("summary_university_count_mismatch")
    if summary["program_count"] != program_count:
        errors.append("summary_program_count_mismatch")
    if summary["course_count"] != course_count:
        errors.append("summary_course_count_mismatch")
    if summary["status_counts"] != expected_statuses:
        errors.append("summary_status_counts_mismatch")
    return {
        "university_count": len(payload["universities"]),
        "program_count": program_count,
        "course_count": course_count,
        "status_counts": expected_statuses,
        "error_count": len(errors),
    }, errors


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--output-dir", type=Path, default=Path("outputs") / "curriculum"
    )
    args = parser.parse_args()

    years: dict[str, Any] = {}
    errors: list[str] = []
    for year in SUPPORTED_YEARS:
        result, year_errors = validate_year(args.output_dir / f"{year}.json")
        years[year] = result
        errors.extend(f"{year}:{error}" for error in year_errors)

    incomplete = load_json(args.output_dir / "incomplete_universities.json")
    actual_incomplete = sum(
        result["status_counts"]["partial"] + result["status_counts"]["failed"]
        for result in years.values()
    )
    if incomplete["summary"]["count"] != actual_incomplete:
        errors.append("incomplete_report_count_mismatch")

    report = {
        "valid": not errors,
        "error_count": len(errors),
        "errors": errors,
        "years": years,
        "incomplete_university_count": actual_incomplete,
    }
    write_json(args.output_dir / "validation_report.json", report)
    print(json.dumps(report, ensure_ascii=False, indent=2))
    return 0 if not errors else 1


if __name__ == "__main__":
    raise SystemExit(main())
