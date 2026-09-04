"""Modal-dialog detection via the OS window layer, NOT via COM.

The 2026-09-03 stress test showed com_word_status reporting "ready" while
the user stared at a three-dialog cascade: COM cannot see Word's modal
dialogs because the dialogs are exactly what blocks COM. The Win32 window
layer can — a modal Word alert is a visible top-level (or owned) window
of a dialog window class belonging to the WINWORD.EXE process. This
module enumerates those windows READ-ONLY (titles and cheap static text);
it never closes, clicks, or messages a dialog — dismissal is a human's
decision.

Pure ctypes (no pywin32, no COM apartment), so it works precisely when
COM is hung, and is trivially safe when Word is not running (empty list).
"""

from __future__ import annotations

import contextlib
import sys

# Window classes Word uses for alerts and dialogs:
# - "#32770": the standard Windows dialog class (file-permission errors,
#   same-name conflicts, save-as prompts — the report's cascade)
# - "NUIDialog": Office's modern alert/dialog class (Word 2013+)
# - "bosa_sdm_msword": Word's classic internal dialog class
DIALOG_CLASSES = {"#32770", "NUIDialog", "bosa_sdm_msword"}

_MAX_TEXT = 512
_MAX_STATICS = 6


def _user32():
    if sys.platform != "win32":  # pragma: no cover
        return None
    import ctypes

    return ctypes, ctypes.windll.user32


def _window_text(ctypes, user32, hwnd) -> str:
    buf = ctypes.create_unicode_buffer(_MAX_TEXT)
    user32.GetWindowTextW(hwnd, buf, _MAX_TEXT)
    return buf.value


def _class_name(ctypes, user32, hwnd) -> str:
    buf = ctypes.create_unicode_buffer(256)
    user32.GetClassNameW(hwnd, buf, 256)
    return buf.value


def _static_texts(ctypes, user32, hwnd) -> list[str]:
    """Visible text of the dialog's Static children (the message body of a
    standard alert), cheap and read-only."""
    texts: list[str] = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def child_cb(child, _lp):
        if len(texts) >= _MAX_STATICS:
            return False
        cls = _class_name(ctypes, user32, child)
        if cls.lower() in ("static", "richedit20w"):
            t = _window_text(ctypes, user32, child).strip()
            if t:
                texts.append(t)
        return True

    with contextlib.suppress(Exception):
        user32.EnumChildWindows(hwnd, child_cb, 0)
    return texts


def pending_dialogs(pids: set | None = None) -> list[dict]:
    """Visible dialog-class windows belonging to the given process ids
    (default: every running WINWORD.EXE). Returns [{title, text?, class}]
    — read-only enumeration; nothing is dismissed or touched. Empty list
    when Word is absent or no dialogs are up."""
    mods = _user32()
    if mods is None:
        return []
    ctypes, user32 = mods
    if pids is None:
        from .bridge import _winword_pids

        pids = _winword_pids()
    if not pids:
        return []
    found: list[dict] = []

    @ctypes.WINFUNCTYPE(ctypes.c_bool, ctypes.c_void_p, ctypes.c_void_p)
    def enum_cb(hwnd, _lp):
        try:
            if not user32.IsWindowVisible(hwnd):
                return True
            pid = ctypes.c_ulong(0)
            user32.GetWindowThreadProcessId(hwnd, ctypes.byref(pid))
            if pid.value not in pids:
                return True
            cls = _class_name(ctypes, user32, hwnd)
            if cls not in DIALOG_CLASSES:
                return True
            entry: dict = {
                "title": _window_text(ctypes, user32, hwnd),
                "class": cls,
            }
            statics = _static_texts(ctypes, user32, hwnd)
            if statics:
                entry["text"] = " | ".join(statics)
            found.append(entry)
        except Exception:
            pass
        return True

    with contextlib.suppress(Exception):
        user32.EnumWindows(enum_cb, 0)
    return found
