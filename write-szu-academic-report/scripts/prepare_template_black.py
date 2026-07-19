#!/usr/bin/env python3
"""Convert red DOCX text to black without changing document wording."""

from __future__ import annotations

import argparse
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

COLOR_RE = re.compile(r'(<w:color\b[^>]*\bw:val=")([0-9A-Fa-f]{6})("[^>]*/?>)')


def is_red(hex_color: str) -> bool:
    red, green, blue = (int(hex_color[i : i + 2], 16) for i in (0, 2, 4))
    return red >= 160 and red >= green * 1.8 and red >= blue * 1.8


def replace_red(xml: str) -> tuple[str, int]:
    count = 0

    def replacement(match: re.Match[str]) -> str:
        nonlocal count
        if is_red(match.group(2)):
            count += 1
            return f'{match.group(1)}000000{match.group(3)}'
        return match.group(0)

    return COLOR_RE.sub(replacement, xml), count


def prepare(input_path: Path, output_path: Path) -> int:
    changed = 0
    output_path.parent.mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as temp_dir:
        temp = Path(temp_dir)
        with zipfile.ZipFile(input_path) as source:
            source.extractall(temp)
        for xml_path in (temp / "word").rglob("*.xml"):
            text = xml_path.read_text(encoding="utf-8")
            updated, count = replace_red(text)
            if count:
                xml_path.write_text(updated, encoding="utf-8")
                changed += count
        archive = shutil.make_archive(str(output_path.with_suffix("")), "zip", temp)
        Path(archive).replace(output_path)
    return changed


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("input", type=Path)
    parser.add_argument("output", type=Path)
    args = parser.parse_args()
    changed = prepare(args.input, args.output)
    print(f"{args.output} red_colors_changed={changed}")


if __name__ == "__main__":
    main()
