"""COM harness: open documents in invisible Word and fail on any repair prompt.

Run at phase gates (slow — launches real Word):
    .venv/Scripts/python.exe -X utf8 tests/word_validator.py <file-or-dir> [...]

A file passes only if Word opens it with no repair/recovery dialog and reports
it clean. DisplayAlerts is disabled so a corrupt file raises a COM error
instead of hanging on a modal dialog.
"""

from __future__ import annotations

import sys
from pathlib import Path


def validate_files(paths: list[Path]) -> int:
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    word = win32com.client.DispatchEx("Word.Application")
    word.Visible = False
    word.DisplayAlerts = 0  # wdAlertsNone
    failures = 0
    try:
        for path in paths:
            abs_path = str(path.resolve())
            try:
                doc = word.Documents.Open(
                    abs_path,
                    ReadOnly=True,
                    AddToRecentFiles=False,
                    OpenAndRepair=False,
                )
                # Accessing content forces full load; corruption surfaces here.
                _ = doc.Content.End
                para_count = doc.Paragraphs.Count
                doc.Close(SaveChanges=0)
                print(f"PASS  {path.name}  ({para_count} paragraphs)")
            except Exception as exc:  # noqa: BLE001 - COM errors are opaque
                failures += 1
                print(f"FAIL  {path.name}  {exc}")
    finally:
        word.Quit()
        pythoncom.CoUninitialize()
    return failures


def main() -> None:
    targets: list[Path] = []
    for arg in sys.argv[1:]:
        p = Path(arg)
        if p.is_dir():
            targets.extend(sorted(p.glob("*.docx")))
        elif p.suffix.lower() == ".docx":
            targets.append(p)
    if not targets:
        print("usage: word_validator.py <file-or-dir> [...]")
        sys.exit(2)
    failures = validate_files(targets)
    print(f"\n{len(targets) - failures}/{len(targets)} passed")
    sys.exit(1 if failures else 0)


if __name__ == "__main__":
    main()
