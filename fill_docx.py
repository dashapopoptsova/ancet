#!/usr/bin/env python3
"""Fill a DOCX template with values from JSON while preserving formatting.

Expected template placeholders are standard Jinja2/DocxTPL tags, e.g.:
  {{ full_name }}
  {{ value(address) }}
  {{ checkbox("basis", "charter") }}

Conditional blocks can be done with Jinja2:
  {% if is_foreigner %}...{% endif %}
"""
from __future__ import annotations

import argparse
import json
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Iterable

from docxtpl import DocxTemplate, InlineImage
from jinja2 import Environment
from docx.shared import Mm


CHECKED_DEFAULT = "☑"
UNCHECKED_DEFAULT = "☐"


@dataclass
class CheckboxConfig:
    checked: str = CHECKED_DEFAULT
    unchecked: str = UNCHECKED_DEFAULT


@dataclass
class FillConfig:
    empty_placeholder: str = "—"
    checkbox: CheckboxConfig = field(default_factory=CheckboxConfig)


def load_json(path: Path) -> dict[str, Any]:
    with path.open("r", encoding="utf-8") as handle:
        return json.load(handle)


def to_inline_image(doc: DocxTemplate, image_data: dict[str, Any]) -> InlineImage:
    path = Path(image_data["path"]).expanduser()
    width_mm = image_data.get("width_mm")
    height_mm = image_data.get("height_mm")
    width = Mm(width_mm) if width_mm else None
    height = Mm(height_mm) if height_mm else None
    return InlineImage(doc, str(path), width=width, height=height)


def normalize_choice(value: Any) -> set[str]:
    if value is None:
        return set()
    if isinstance(value, str):
        return {value}
    if isinstance(value, Iterable):
        return {str(item) for item in value}
    return {str(value)}


def resolve_path(data: Any, path: str) -> Any:
    current: Any = data
    for raw_part in path.split("."):
        if raw_part == "":
            continue
        if "[" in raw_part and raw_part.endswith("]"):
            key, index_part = raw_part[:-1].split("[", 1)
            if key:
                if not isinstance(current, dict):
                    return None
                current = current.get(key)
            if current is None:
                return None
            try:
                index = int(index_part)
            except ValueError:
                return None
            if not isinstance(current, list):
                return None
            if index < 0 or index >= len(current):
                return None
            current = current[index]
            continue

        if not isinstance(current, dict):
            return None
        current = current.get(raw_part)
    return current


def build_context(data: dict[str, Any], config: FillConfig, doc: DocxTemplate) -> dict[str, Any]:
    choices = data.get("choices", {})
    images = data.get("images", {})

    def checkbox(group: str, option: str) -> str:
        selected = normalize_choice(choices.get(group))
        return config.checkbox.checked if option in selected else config.checkbox.unchecked

    def value(item: Any) -> Any:
        if item is None:
            return config.empty_placeholder
        if isinstance(item, str) and not item.strip():
            return config.empty_placeholder
        return item

    def field(path: str, default: Any | None = None) -> Any:
        resolved = resolve_path(data, path)
        if resolved is None:
            return default if default is not None else config.empty_placeholder
        return value(resolved)

    context: dict[str, Any] = {}
    context.update(data)
    context["checkbox"] = checkbox
    context["value"] = value
    context["field"] = field

    for key, image_data in images.items():
        context[key] = to_inline_image(doc, image_data)

    return context


def build_config(data: dict[str, Any]) -> FillConfig:
    empty_placeholder = data.get("empty_placeholder", "—")
    checkbox_symbols = data.get("checkbox_symbols", {})
    checkbox = CheckboxConfig(
        checked=checkbox_symbols.get("checked", CHECKED_DEFAULT),
        unchecked=checkbox_symbols.get("unchecked", UNCHECKED_DEFAULT),
    )
    return FillConfig(empty_placeholder=empty_placeholder, checkbox=checkbox)


def render_docx(template_path: Path, data_path: Path, output_path: Path) -> None:
    data = load_json(data_path)
    config = build_config(data)

    doc = DocxTemplate(str(template_path))
    env = Environment(autoescape=False)
    doc.render(build_context(data, config, doc), jinja_env=env)
    doc.save(str(output_path))


def main() -> int:
    parser = argparse.ArgumentParser(description="Fill a DOCX template with JSON data.")
    parser.add_argument("template", type=Path, help="Path to template.docx")
    parser.add_argument("data", type=Path, help="Path to filled_fields.json")
    parser.add_argument("output", type=Path, help="Path to result.docx")
    args = parser.parse_args()

    render_docx(args.template, args.data, args.output)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
