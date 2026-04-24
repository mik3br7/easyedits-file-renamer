from __future__ import annotations

import argparse
from pathlib import Path
from typing import Iterable


def build_new_name(
    path: Path,
    prefix: str,
    suffix: str,
    find_text: str | None,
    replace_text: str,
    start_number: int | None,
    index: int,
) -> str:
    stem = path.stem
    if find_text is not None:
        stem = stem.replace(find_text, replace_text)
    if start_number is not None:
        stem = f"{start_number + index:03d}_{stem}"
    return f"{prefix}{stem}{suffix}{path.suffix}"


def collect_files(folder: Path, pattern: str, recursive: bool) -> list[Path]:
    iterator: Iterable[Path]
    if recursive:
        iterator = folder.rglob(pattern)
    else:
        iterator = folder.glob(pattern)
    return sorted(path for path in iterator if path.is_file())


def rename_files(
    files: list[Path],
    prefix: str,
    suffix: str,
    find_text: str | None,
    replace_text: str,
    start_number: int | None,
    dry_run: bool,
) -> int:
    planned_names: dict[Path, Path] = {}

    for index, path in enumerate(files):
        new_name = build_new_name(
            path=path,
            prefix=prefix,
            suffix=suffix,
            find_text=find_text,
            replace_text=replace_text,
            start_number=start_number,
            index=index,
        )
        target = path.with_name(new_name)
        planned_names[path] = target

    targets = list(planned_names.values())
    if len({target.name for target in targets}) != len(targets):
        raise ValueError("Two or more files would end up with the same name.")

    conflicts = [
        target
        for source, target in planned_names.items()
        if target.exists() and target != source
    ]
    if conflicts:
        conflict_list = "\n".join(f"  - {path.name}" for path in conflicts)
        raise FileExistsError(
            "Rename stopped because these target files already exist:\n"
            f"{conflict_list}"
        )

    for source, target in planned_names.items():
        print(f"{source.name} -> {target.name}")
        if not dry_run and source != target:
            source.rename(target)

    return sum(1 for source, target in planned_names.items() if source != target)


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Rename files in a folder using prefixes, suffixes, text replacement, and numbering."
    )
    parser.add_argument("folder", nargs="?", default=".", help="Folder containing files to rename.")
    parser.add_argument("--pattern", default="*", help="Glob pattern to match files. Default: *")
    parser.add_argument("--prefix", default="", help="Text to add to the beginning of each file name.")
    parser.add_argument("--suffix", default="", help="Text to add before each file extension.")
    parser.add_argument("--find", help="Text to find in each file name.")
    parser.add_argument("--replace", default="", help="Replacement text used with --find.")
    parser.add_argument(
        "--number-from",
        type=int,
        help="Add sequential numbers starting from this value, like 001_filename.ext.",
    )
    parser.add_argument(
        "--recursive",
        action="store_true",
        help="Include files from subfolders.",
    )
    parser.add_argument(
        "--apply",
        action="store_true",
        help="Actually rename files. Without this flag, the program only previews changes.",
    )
    return parser.parse_args()


def main() -> int:
    args = parse_args()
    folder = Path(args.folder).resolve()

    if not folder.exists():
        print(f"Folder not found: {folder}")
        return 1

    files = collect_files(folder, args.pattern, args.recursive)
    if not files:
        print("No files matched the given pattern.")
        return 0

    changed_count = rename_files(
        files=files,
        prefix=args.prefix,
        suffix=args.suffix,
        find_text=args.find,
        replace_text=args.replace,
        start_number=args.number_from,
        dry_run=not args.apply,
    )

    mode = "Previewed" if not args.apply else "Renamed"
    print(f"\n{mode} {changed_count} file(s).")
    if not args.apply:
        print("Run again with --apply to make the changes.")

    return 0


if __name__ == "__main__":
    raise SystemExit(main())
