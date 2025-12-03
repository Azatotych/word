"""Утилита проверки оформления научных статей в формате DOCX (CLI)."""
import argparse
import json
from typing import List

from format_checker_core import Issue, annotate_document, check_document, format_report


def parse_arguments() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Проверка оформления статей в DOCX")
    parser.add_argument("files", nargs="+", help="Пути к файлам .docx для проверки")
    parser.add_argument("--json", action="store_true", dest="json_mode", help="Вывод отчёта в формате JSON")
    parser.add_argument(
        "--no-annotate",
        action="store_true",
        dest="skip_annotate",
        help="Не создавать аннотированную копию документа с подсветкой проблем",
    )
    return parser.parse_args()


def print_report(file_name: str, issues: List[Issue], json_mode: bool) -> List[dict]:
    json_items: List[dict] = []
    if json_mode:
        for issue in issues:
            json_items.append(issue.to_dict(file_name))
    else:
        _, report_text = format_report(issues, file_name)
        print(report_text)
    return json_items


def main() -> None:
    args = parse_arguments()
    all_json_items: List[dict] = []

    for file_path in args.files:
        try:
            issues = check_document(file_path)
        except Exception as exc:  # pragma: no cover - защита от падения на некорректных файлов
            issue = Issue(rule="LOAD", level="ERROR", message=f"Не удалось открыть файл: {exc}")
            if args.json_mode:
                all_json_items.append(issue.to_dict(file_path))
            else:
                print(f"==== {file_path} ====")
                print(f"[{issue.level}] {issue.rule}: {issue.message}\n")
            continue

        json_items = print_report(file_path, issues, args.json_mode)
        all_json_items.extend(json_items)

        if not args.json_mode and not args.skip_annotate:
            annotated_path = annotate_document(file_path, issues)
            print(f"Аннотированный файл с подсветкой проблемных абзацев: {annotated_path}\n")

    if args.json_mode:
        print(json.dumps(all_json_items, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
