#!/usr/bin/env python3
"""Build a self-contained HTML browser for the extracted curriculum JSON."""

from __future__ import annotations

import argparse
import json
from pathlib import Path
from typing import Any


YEARS = ("r3", "r4", "r5", "r6", "r7", "r7_pre")


def unique_source_refs(course: dict[str, Any]) -> list[str]:
    refs: list[str] = []
    for occurrence in course.get("source_occurrences", []):
        sheet = occurrence.get("sheet_name")
        cell = occurrence.get("source_cell")
        if not sheet or not cell:
            continue
        ref = f"{sheet}!{cell}"
        if ref not in refs:
            refs.append(ref)
    return refs


def course_view(course: dict[str, Any]) -> dict[str, Any]:
    return {
        "name": course["name"],
        "name_variants": course.get("name_variants", []),
        "credits": course.get("credits"),
        "is_required": course.get("is_required"),
        "requirement_types": course.get("requirement_types", []),
        "categories": course.get("categories", []),
        "codes": course.get("model_curriculum_codes", []),
        "sources": unique_source_refs(course),
    }


def program_view(program: dict[str, Any]) -> dict[str, Any]:
    requirement = program["completion_requirement"]
    extraction = program["extraction"]
    return {
        "sheet_name": program["sheet_name"],
        "source_sheets": program.get("source_sheets", [program["sheet_name"]]),
        "program_name": program.get("program_name"),
        "departments": program.get("target_departments", []),
        "requirement": {
            "text": requirement.get("text"),
            "minimum_courses": requirement.get("minimum_courses"),
            "minimum_credits": requirement.get("minimum_credits"),
            "scope": requirement.get("scope", {}),
            "required_status": requirement.get("program_required_status"),
        },
        "courses": [course_view(course) for course in program.get("courses", [])],
        "course_count": program.get("course_count", len(program.get("courses", []))),
        "status": extraction.get("status", "unknown"),
        "warnings": extraction.get("warnings", []),
        "notes": extraction.get("notes", []),
    }


def build_view_model(payload: dict[str, Any]) -> dict[str, Any]:
    universities: list[dict[str, Any]] = []
    year_summaries: dict[str, Any] = {}
    for year in YEARS:
        year_payload = payload["years"][year]
        year_summaries[year] = year_payload["summary"]
        for university in year_payload["universities"]:
            programs = [program_view(program) for program in university.get("programs", [])]
            universities.append(
                {
                    "id": f"{year}-{len(universities) + 1}",
                    "year": year,
                    "name": university["university_name"],
                    "source_file": university["source_file"],
                    "status": university["extraction"].get("status", "unknown"),
                    "warnings": university["extraction"].get("warnings", []),
                    "programs": programs,
                    "program_count": len(programs),
                    "course_count": sum(program["course_count"] for program in programs),
                }
            )

    statuses = {"success": 0, "partial": 0, "failed": 0}
    for university in universities:
        status = university["status"]
        statuses[status] = statuses.get(status, 0) + 1
    return {
        "schema_version": payload.get("schema_version"),
        "summary": {
            "university_count": len(universities),
            "program_count": sum(item["program_count"] for item in universities),
            "course_count": sum(item["course_count"] for item in universities),
            "status_counts": statuses,
        },
        "year_summaries": year_summaries,
        "universities": universities,
    }


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--input",
        type=Path,
        default=Path("outputs") / "curriculum" / "all_years.json",
    )
    parser.add_argument(
        "--template",
        type=Path,
        default=Path(__file__).with_name("curriculum_viewer_template.html"),
    )
    parser.add_argument(
        "--output",
        type=Path,
        default=Path("outputs") / "curriculum" / "curriculum_browser.html",
    )
    args = parser.parse_args()

    payload = json.loads(args.input.read_text(encoding="utf-8"))
    view_model = build_view_model(payload)
    data_json = json.dumps(
        view_model,
        ensure_ascii=False,
        separators=(",", ":"),
    ).replace("</", "<\\/")
    template = args.template.read_text(encoding="utf-8")
    if "{{DATA_JSON}}" not in template:
        raise ValueError("Template does not contain {{DATA_JSON}}")
    html = template.replace("{{DATA_JSON}}", data_json)
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(html, encoding="utf-8")

    summary = view_model["summary"]
    print(
        json.dumps(
            {
                "output": str(args.output),
                "bytes": args.output.stat().st_size,
                **summary,
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
