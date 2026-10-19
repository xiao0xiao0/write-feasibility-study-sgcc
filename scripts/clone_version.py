from __future__ import annotations

import argparse
import json
import re
import shutil
from pathlib import Path
from zipfile import BadZipFile, ZipFile


def next_version_path(source: Path) -> Path:
    # Choose the final version token; keep descriptive/copy suffixes intact.
    # A filename version identifies an output name, never the authoritative source.
    tokens = list(re.finditer(r"(?i)(?P<prefix>[_-]?v)(?P<number>\d+(?:\.\d+)*)", source.stem))
    suffix_pattern = r"(?:\s*[（(][^()（）]*[)）]|\s*[-_][^()（）]*)*"
    for match in reversed(tokens):
        suffix = source.stem[match.end():]
        if not re.fullmatch(suffix_pattern, suffix):
            continue
        numbers = match.group("number").split(".")
        width = len(numbers[-1])
        numbers[-1] = str(int(numbers[-1]) + 1).zfill(width)
        stem = source.stem[:match.start()] + match.group("prefix") + ".".join(numbers) + suffix
        return source.with_name(stem + source.suffix)
    return source.with_name(source.stem + "_V1" + source.suffix)


def validate_docx(path: Path) -> None:
    if path.suffix.lower() != ".docx":
        raise ValueError(f"source is not a .docx file: {path}")
    try:
        with ZipFile(path) as package:
            names = set(package.namelist())
            required = {"[Content_Types].xml", "word/document.xml", "word/styles.xml"}
            missing = sorted(required - names)
            if missing:
                raise ValueError(f"invalid DOCX; missing parts: {missing}")
            bad = package.testzip()
            if bad:
                raise ValueError(f"invalid ZIP member: {bad}")
    except BadZipFile as exc:
        raise ValueError(f"invalid DOCX ZIP: {path}") from exc


def main() -> int:
    parser = argparse.ArgumentParser(description="Clone a DOCX to the next Vn version without overwriting the source.")
    parser.add_argument("source", type=Path)
    parser.add_argument("--output", type=Path)
    parser.add_argument("--dry-run", action="store_true")
    args = parser.parse_args()

    source = args.source.expanduser().resolve()
    if not source.is_file():
        raise FileNotFoundError(source)
    validate_docx(source)

    output = (args.output.expanduser().resolve() if args.output else next_version_path(source))
    if output == source:
        raise ValueError("output must differ from source")
    if output.exists():
        raise FileExistsError(f"refusing to overwrite: {output}")

    if not args.dry_run:
        output.parent.mkdir(parents=True, exist_ok=True)
        # Another conversation may create this version after the existence check.
        # Exclusive creation must still refuse to overwrite its result.
        with source.open("rb") as original, output.open("xb") as destination:
            shutil.copyfileobj(original, destination)
        shutil.copystat(source, output)
        validate_docx(output)

    print(json.dumps({
        "source": str(source),
        "output": str(output),
        "created": not args.dry_run,
    }, ensure_ascii=False, indent=2))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
