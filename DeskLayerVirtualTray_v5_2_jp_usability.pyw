# -*- coding: utf-8 -*-
"""
Desk Layer Virtual Tray v5.1 - Rescue Hardened / Python 3.14 stdlib build

Windows clipboard helper for plain text, files/folders, and DIB images.
Design goals:
- Standard library only.
- No global keyboard hooks. Ctrl+V / Win+V are never intercepted.
- Event-driven clipboard listener on Windows; low-rate polling only as fallback.
- Minimal memory churn for always-on use.
- Clipboard data is not persisted to disk unless the user explicitly exports or pins it. Logs never include clipboard contents.
- Pro Hardened build: repair tools, diagnostics, safer paste engine, responsive UI, text transforms, undo, and corruption recovery.
"""
from __future__ import annotations

import ctypes
import hashlib
import json
import os
import queue
import platform
import re
import struct
import subprocess
import sys
import tempfile
import threading
import time
import traceback
import uuid
import urllib.parse
import html
import datetime
import unicodedata
from ctypes import wintypes
from dataclasses import dataclass, field
from typing import Optional, Sequence
import tkinter as tk
from tkinter import messagebox, simpledialog

APP_NAME = "仮想トレー"
VERSION = "5.2-jp-usability-py314-stdlib"
COMPACT_DEFAULT_SIZE = 102  # v5.1の34pxから約3倍。待機中でも見える・掴める大きさ。
COMPACT_MIN_SIZE = 72
COMPACT_MAX_SIZE = 180
IS_WINDOWS = platform.system().lower().startswith("win")

APPDATA = os.environ.get("APPDATA") or os.path.expanduser("~")
DATA_DIR = os.path.join(APPDATA, "DeskLayerVirtualTray")
CONFIG_PATH = os.path.join(DATA_DIR, "config_v2_2.json")
PINNED_PATH = os.path.join(DATA_DIR, "pinned_v2_2.json")
OLD_CONFIG_PATH = os.path.join(DATA_DIR, "config_v2_1.json")
OLD_PINNED_PATH = os.path.join(DATA_DIR, "pinned_v2_1.json")
OLDER_CONFIG_PATH = os.path.join(DATA_DIR, "config_v2_0.json")
OLDER_PINNED_PATH = os.path.join(DATA_DIR, "pinned_v2_0.json")
EVENT_LOG = os.path.join(DATA_DIR, "event.log")
ERROR_LOG = os.path.join(DATA_DIR, "error.log")
STOP_REQUEST_PATH = os.path.join(DATA_DIR, "stop.request")
os.makedirs(DATA_DIR, exist_ok=True)

LOG_MAX_BYTES = 512_000
DEFAULT_CONFIG = {
    "x": None,
    "y": None,
    "max_items": 80,
    "max_total_bytes": 96 * 1024 * 1024,
    "max_text_chars": 500_000,
    "max_image_bytes": 32 * 1024 * 1024,
    "main_render_limit": 32,
    "auto_capture": True,
    # 常駐時の軽さを優先。小さい待機ボタン中は既定で読み取りません。
    "capture_when_compact": False,
    "topmost": True,
    "alpha": 0.97,
    # 実務用の安全弁。自動取り込み時だけ効き、手動ADD CLIPでは取り込めます。
    "privacy_guard": True,
    "auto_skip_sensitive": True,
    "privacy_exclude_processes": [
        "1password", "bitwarden", "keepass", "lastpass", "dashlane", "authy", "authenticator"
    ],
    "privacy_exclude_titles": [
        "password", "passkey", "otp", "one-time", "secret", "token", "api key",
        "パスワード", "認証コード", "秘密鍵"
    ],
    "large_confirm_bytes": 16 * 1024 * 1024,
    "confirm_all_operations": True,
    "auto_clear_minutes": 0,
    # 個人用便利設定。Excel/Word/Outlookでは通常テキストを優先した方が実務で扱いやすい。
    "clipboard_priority": "text_first",
    "paste_advance": True,
    "confirm_delete": False,
    "compact_badge_pinned": True,
    "compact_size": COMPACT_DEFAULT_SIZE,
    "global_hotkeys": True,
    "paste_delay_ms": 100,
    "confirm_exit": True,
    "diagnostic_tail_lines": 80,
}

# Clipboard formats / Win32 constants.
CF_UNICODETEXT = 13
CF_HDROP = 15
CF_DIB = 8
CF_DIBV5 = 17
GMEM_MOVEABLE = 0x0002
GMEM_ZEROINIT = 0x0040
GHND = GMEM_MOVEABLE | GMEM_ZEROINIT
DROPEFFECT_COPY = 1
KEYEVENTF_KEYUP = 0x0002
VK_CONTROL = 0x11
VK_C = 0x43
VK_V = 0x56
SW_RESTORE = 9
SW_SHOW = 5
INPUT_KEYBOARD = 1
PROCESS_QUERY_LIMITED_INFORMATION = 0x1000
SM_XVIRTUALSCREEN = 76
SM_YVIRTUALSCREEN = 77
SM_CXVIRTUALSCREEN = 78
SM_CYVIRTUALSCREEN = 79
WM_CLIPBOARDUPDATE = 0x031D
WM_HOTKEY = 0x0312
WM_CLOSE = 0x0010
WM_DESTROY = 0x0002
WM_APP = 0x8000
WM_DESKLAYER_REFRESH_HOTKEYS = WM_APP + 0x0522
MOD_ALT = 0x0001
MOD_CONTROL = 0x0002
MOD_SHIFT = 0x0004
MOD_WIN = 0x0008
VK_RETURN = 0x0D
GWL_WNDPROC = -4
ERROR_ALREADY_EXISTS = 183
MUTEX_NAME = "Local\\DeskLayerVirtualTray_v20_single_instance"

HOTKEY_OPEN_PICKER = 0x5221
HOTKEY_ADD_CLIP = 0x5222
HOTKEY_SET_ACTIVE = 0x5223
HOTKEY_PASTE_ACTIVE = 0x5224
HOTKEY_GET_SELECTION = 0x5225
HOTKEY_NOTE = 0x5226
HOTKEY_EXIT = 0x5227
HOTKEY_DEFS = (
    (HOTKEY_OPEN_PICKER, MOD_WIN | MOD_ALT, ord("V"), "Win+Alt+V", "一覧"),
    (HOTKEY_ADD_CLIP, MOD_WIN | MOD_ALT, ord("C"), "Win+Alt+C", "追加"),
    (HOTKEY_SET_ACTIVE, MOD_WIN | MOD_ALT, ord("S"), "Win+Alt+S", "選択をセット"),
    (HOTKEY_PASTE_ACTIVE, MOD_WIN | MOD_ALT, ord("P"), "Win+Alt+P", "選択を貼付"),
    (HOTKEY_GET_SELECTION, MOD_WIN | MOD_ALT, ord("G"), "Win+Alt+G", "選択範囲取得"),
    (HOTKEY_NOTE, MOD_WIN | MOD_ALT, ord("N"), "Win+Alt+N", "メモ"),
    (HOTKEY_EXIT, MOD_WIN | MOD_ALT, ord("Q"), "Win+Alt+Q", "緊急終了"),
)

if IS_WINDOWS:
    user32 = ctypes.windll.user32
    kernel32 = ctypes.windll.kernel32
    shell32 = ctypes.windll.shell32
else:
    user32 = kernel32 = shell32 = None


class POINT(ctypes.Structure):
    _fields_ = [("x", wintypes.LONG), ("y", wintypes.LONG)]


class DROPFILES(ctypes.Structure):
    _fields_ = [
        ("pFiles", wintypes.DWORD),
        ("pt", POINT),
        ("fNC", wintypes.BOOL),
        ("fWide", wintypes.BOOL),
    ]


LONG_PTR = ctypes.c_longlong if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_long


class WNDCLASSW(ctypes.Structure):
    _fields_ = [
        ("style", wintypes.UINT),
        ("lpfnWndProc", ctypes.c_void_p),
        ("cbClsExtra", ctypes.c_int),
        ("cbWndExtra", ctypes.c_int),
        ("hInstance", wintypes.HINSTANCE),
        ("hIcon", wintypes.HICON),
        ("hCursor", ctypes.c_void_p),
        ("hbrBackground", wintypes.HBRUSH),
        ("lpszMenuName", wintypes.LPCWSTR),
        ("lpszClassName", wintypes.LPCWSTR),
    ]


ULONG_PTR = ctypes.c_ulonglong if ctypes.sizeof(ctypes.c_void_p) == 8 else ctypes.c_ulong


class MOUSEINPUT(ctypes.Structure):
    _fields_ = [
        ("dx", wintypes.LONG),
        ("dy", wintypes.LONG),
        ("mouseData", wintypes.DWORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]


class KEYBDINPUT(ctypes.Structure):
    _fields_ = [
        ("wVk", wintypes.WORD),
        ("wScan", wintypes.WORD),
        ("dwFlags", wintypes.DWORD),
        ("time", wintypes.DWORD),
        ("dwExtraInfo", ULONG_PTR),
    ]


class HARDWAREINPUT(ctypes.Structure):
    _fields_ = [
        ("uMsg", wintypes.DWORD),
        ("wParamL", wintypes.WORD),
        ("wParamH", wintypes.WORD),
    ]


class INPUTUNION(ctypes.Union):
    _fields_ = [("mi", MOUSEINPUT), ("ki", KEYBDINPUT), ("hi", HARDWAREINPUT)]


class INPUT(ctypes.Structure):
    _fields_ = [("type", wintypes.DWORD), ("union", INPUTUNION)]


def _setup_winapi() -> None:
    if not IS_WINDOWS:
        return
    # DPI is configured before Tk creates its first HWND.  The UI below avoids
    # fixed geometry wherever practical, so per-monitor awareness improves pointer
    # accuracy without clipping controls.  Set DESKLAYER_DPI_AWARE=0 to opt out.
    if os.environ.get("DESKLAYER_DPI_AWARE", "1").strip().lower() not in {"0", "false", "off", "no"}:
        try:
            if hasattr(user32, "SetProcessDpiAwarenessContext"):
                user32.SetProcessDpiAwarenessContext.argtypes = [wintypes.HANDLE]
                user32.SetProcessDpiAwarenessContext.restype = wintypes.BOOL
                user32.SetProcessDpiAwarenessContext(wintypes.HANDLE(-4))  # PER_MONITOR_AWARE_V2
            else:
                raise AttributeError
        except Exception:
            try:
                shcore = ctypes.windll.shcore
                shcore.SetProcessDpiAwareness.argtypes = [ctypes.c_int]
                shcore.SetProcessDpiAwareness.restype = ctypes.c_long
                shcore.SetProcessDpiAwareness(2)  # PROCESS_PER_MONITOR_DPI_AWARE
            except Exception:
                try:
                    user32.SetProcessDPIAware.argtypes = []
                    user32.SetProcessDPIAware.restype = wintypes.BOOL
                    user32.SetProcessDPIAware()
                except Exception:
                    pass

    user32.OpenClipboard.argtypes = [wintypes.HWND]
    user32.OpenClipboard.restype = wintypes.BOOL
    user32.CloseClipboard.argtypes = []
    user32.CloseClipboard.restype = wintypes.BOOL
    try:
        user32.GetOpenClipboardWindow.argtypes = []
        user32.GetOpenClipboardWindow.restype = wintypes.HWND
    except Exception:
        pass
    user32.EmptyClipboard.argtypes = []
    user32.EmptyClipboard.restype = wintypes.BOOL
    user32.GetClipboardData.argtypes = [wintypes.UINT]
    user32.GetClipboardData.restype = wintypes.HANDLE
    user32.SetClipboardData.argtypes = [wintypes.UINT, wintypes.HANDLE]
    user32.SetClipboardData.restype = wintypes.HANDLE
    user32.IsClipboardFormatAvailable.argtypes = [wintypes.UINT]
    user32.IsClipboardFormatAvailable.restype = wintypes.BOOL
    user32.GetClipboardSequenceNumber.argtypes = []
    user32.GetClipboardSequenceNumber.restype = wintypes.DWORD
    user32.RegisterClipboardFormatW.argtypes = [wintypes.LPCWSTR]
    user32.RegisterClipboardFormatW.restype = wintypes.UINT
    user32.GetForegroundWindow.argtypes = []
    user32.GetForegroundWindow.restype = wintypes.HWND
    user32.GetWindowTextLengthW.argtypes = [wintypes.HWND]
    user32.GetWindowTextLengthW.restype = ctypes.c_int
    user32.GetWindowTextW.argtypes = [wintypes.HWND, wintypes.LPWSTR, ctypes.c_int]
    user32.GetWindowTextW.restype = ctypes.c_int
    user32.GetSystemMetrics.argtypes = [ctypes.c_int]
    user32.GetSystemMetrics.restype = ctypes.c_int
    user32.MoveWindow.argtypes = [wintypes.HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, wintypes.BOOL]
    user32.MoveWindow.restype = wintypes.BOOL
    user32.SetForegroundWindow.argtypes = [wintypes.HWND]
    user32.SetForegroundWindow.restype = wintypes.BOOL
    user32.ShowWindow.argtypes = [wintypes.HWND, ctypes.c_int]
    user32.ShowWindow.restype = wintypes.BOOL
    user32.IsIconic.argtypes = [wintypes.HWND]
    user32.IsIconic.restype = wintypes.BOOL
    user32.IsWindow.argtypes = [wintypes.HWND]
    user32.IsWindow.restype = wintypes.BOOL
    user32.GetWindowThreadProcessId.argtypes = [wintypes.HWND, ctypes.POINTER(wintypes.DWORD)]
    user32.GetWindowThreadProcessId.restype = wintypes.DWORD
    try:
        user32.BringWindowToTop.argtypes = [wintypes.HWND]
        user32.BringWindowToTop.restype = wintypes.BOOL
        user32.SetActiveWindow.argtypes = [wintypes.HWND]
        user32.SetActiveWindow.restype = wintypes.HWND
        user32.AttachThreadInput.argtypes = [wintypes.DWORD, wintypes.DWORD, wintypes.BOOL]
        user32.AttachThreadInput.restype = wintypes.BOOL
    except Exception:
        pass
    user32.keybd_event.argtypes = [wintypes.BYTE, wintypes.BYTE, wintypes.DWORD, wintypes.ULONG]
    user32.keybd_event.restype = None
    try:
        user32.SendInput.argtypes = [wintypes.UINT, ctypes.POINTER(INPUT), ctypes.c_int]
        user32.SendInput.restype = wintypes.UINT
    except Exception:
        pass
    try:
        user32.AddClipboardFormatListener.argtypes = [wintypes.HWND]
        user32.AddClipboardFormatListener.restype = wintypes.BOOL
        user32.RemoveClipboardFormatListener.argtypes = [wintypes.HWND]
        user32.RemoveClipboardFormatListener.restype = wintypes.BOOL
    except Exception:
        pass
    try:
        user32.RegisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int, wintypes.UINT, wintypes.UINT]
        user32.RegisterHotKey.restype = wintypes.BOOL
        user32.UnregisterHotKey.argtypes = [wintypes.HWND, ctypes.c_int]
        user32.UnregisterHotKey.restype = wintypes.BOOL
    except Exception:
        pass
    try:
        user32.RegisterClassW.argtypes = [ctypes.POINTER(WNDCLASSW)]
        user32.RegisterClassW.restype = wintypes.ATOM
        user32.UnregisterClassW.argtypes = [wintypes.LPCWSTR, wintypes.HINSTANCE]
        user32.UnregisterClassW.restype = wintypes.BOOL
        user32.CreateWindowExW.argtypes = [
            wintypes.DWORD, wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.DWORD,
            ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int,
            wintypes.HWND, wintypes.HMENU, wintypes.HINSTANCE, wintypes.LPVOID,
        ]
        user32.CreateWindowExW.restype = wintypes.HWND
        user32.DefWindowProcW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
        user32.DefWindowProcW.restype = LONG_PTR
        user32.DestroyWindow.argtypes = [wintypes.HWND]
        user32.DestroyWindow.restype = wintypes.BOOL
        user32.PostMessageW.argtypes = [wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM]
        user32.PostMessageW.restype = wintypes.BOOL
        user32.PostQuitMessage.argtypes = [ctypes.c_int]
        user32.PostQuitMessage.restype = None
        user32.GetMessageW.argtypes = [ctypes.POINTER(wintypes.MSG), wintypes.HWND, wintypes.UINT, wintypes.UINT]
        user32.GetMessageW.restype = ctypes.c_int
        user32.TranslateMessage.argtypes = [ctypes.POINTER(wintypes.MSG)]
        user32.TranslateMessage.restype = wintypes.BOOL
        user32.DispatchMessageW.argtypes = [ctypes.POINTER(wintypes.MSG)]
        user32.DispatchMessageW.restype = LONG_PTR
    except Exception:
        pass

    kernel32.GetCurrentProcessId.argtypes = []
    kernel32.GetCurrentProcessId.restype = wintypes.DWORD
    try:
        kernel32.GetCurrentThreadId.argtypes = []
        kernel32.GetCurrentThreadId.restype = wintypes.DWORD
        kernel32.SetLastError.argtypes = [wintypes.DWORD]
        kernel32.SetLastError.restype = None
    except Exception:
        pass
    try:
        kernel32.GetModuleHandleW.argtypes = [wintypes.LPCWSTR]
        kernel32.GetModuleHandleW.restype = wintypes.HINSTANCE
    except Exception:
        pass
    kernel32.OpenProcess.argtypes = [wintypes.DWORD, wintypes.BOOL, wintypes.DWORD]
    kernel32.OpenProcess.restype = wintypes.HANDLE
    try:
        kernel32.QueryFullProcessImageNameW.argtypes = [wintypes.HANDLE, wintypes.DWORD, wintypes.LPWSTR, ctypes.POINTER(wintypes.DWORD)]
        kernel32.QueryFullProcessImageNameW.restype = wintypes.BOOL
    except Exception:
        pass
    kernel32.GlobalAlloc.argtypes = [wintypes.UINT, ctypes.c_size_t]
    kernel32.GlobalAlloc.restype = wintypes.HGLOBAL
    kernel32.GlobalLock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalLock.restype = wintypes.LPVOID
    kernel32.GlobalUnlock.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalUnlock.restype = wintypes.BOOL
    kernel32.GlobalFree.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalFree.restype = wintypes.HGLOBAL
    kernel32.GlobalSize.argtypes = [wintypes.HGLOBAL]
    kernel32.GlobalSize.restype = ctypes.c_size_t
    kernel32.CreateMutexW.argtypes = [wintypes.LPVOID, wintypes.BOOL, wintypes.LPCWSTR]
    kernel32.CreateMutexW.restype = wintypes.HANDLE
    kernel32.GetLastError.argtypes = []
    kernel32.GetLastError.restype = wintypes.DWORD
    kernel32.CloseHandle.argtypes = [wintypes.HANDLE]
    kernel32.CloseHandle.restype = wintypes.BOOL

    shell32.DragQueryFileW.argtypes = [wintypes.HANDLE, wintypes.UINT, wintypes.LPWSTR, wintypes.UINT]
    shell32.DragQueryFileW.restype = wintypes.UINT


_setup_winapi()


def rotate_log(path: str) -> None:
    try:
        if os.path.exists(path) and os.path.getsize(path) > LOG_MAX_BYTES:
            bak = path + ".1"
            try:
                if os.path.exists(bak):
                    os.remove(bak)
            except Exception:
                pass
            try:
                os.replace(path, bak)
            except Exception:
                pass
    except Exception:
        pass


def log_event(msg: str) -> None:
    try:
        rotate_log(EVENT_LOG)
        with open(EVENT_LOG, "a", encoding="utf-8") as f:
            f.write(time.strftime("%Y-%m-%d %H:%M:%S") + "  " + msg + "\n")
    except Exception:
        pass


def log_error(exc: object) -> None:
    try:
        rotate_log(ERROR_LOG)
        with open(ERROR_LOG, "a", encoding="utf-8") as f:
            f.write("\n" + "=" * 80 + "\n")
            f.write(time.strftime("%Y-%m-%d %H:%M:%S") + "\n")
            if isinstance(exc, BaseException):
                traceback.print_exception(type(exc), exc, exc.__traceback__, file=f)
            else:
                f.write(str(exc) + "\n")
    except Exception:
        pass


def install_exception_logging() -> None:
    """Keep pythonw/.pyw failures visible in error.log."""
    try:
        def sys_hook(exc_type, exc_value, exc_tb):
            try:
                rotate_log(ERROR_LOG)
                with open(ERROR_LOG, "a", encoding="utf-8") as f:
                    f.write("\n" + "=" * 80 + "\n")
                    f.write(time.strftime("%Y-%m-%d %H:%M:%S") + "  uncaught exception\n")
                    traceback.print_exception(exc_type, exc_value, exc_tb, file=f)
            except Exception:
                pass
        sys.excepthook = sys_hook
    except Exception:
        pass
    try:
        def thread_hook(args):
            try:
                rotate_log(ERROR_LOG)
                with open(ERROR_LOG, "a", encoding="utf-8") as f:
                    f.write("\n" + "=" * 80 + "\n")
                    f.write(time.strftime("%Y-%m-%d %H:%M:%S") + f"  thread exception: {getattr(args, 'thread', None)}\n")
                    traceback.print_exception(args.exc_type, args.exc_value, args.exc_traceback, file=f)
            except Exception:
                pass
        threading.excepthook = thread_hook  # type: ignore[attr-defined]
    except Exception:
        pass


def log_environment(root: Optional[tk.Misc] = None, *, safe_mode: bool = False) -> None:
    try:
        parts = [
            f"start env version={VERSION}",
            f"python={sys.version.split()[0]}",
            f"executable={sys.executable}",
            f"platform={platform.platform()}",
            f"windows={IS_WINDOWS}",
            f"safe={int(bool(safe_mode))}",
        ]
        if root is not None:
            try:
                parts.append(f"tk={root.tk.call('info', 'patchlevel')}")
            except Exception:
                pass
            try:
                sx, sy, sw, sh = virtual_screen_bounds(root)
                parts.append(f"screen={sx},{sy},{sw}x{sh}")
            except Exception:
                pass
        log_event(" ".join(parts))
    except Exception:
        pass


def tail_file(path: str, max_lines: int = 80) -> str:
    try:
        if not os.path.exists(path):
            return ""
        max_lines = max(1, min(1000, int(max_lines)))
        with open(path, "r", encoding="utf-8", errors="replace") as f:
            lines = f.readlines()[-max_lines:]
        return "".join(lines)
    except Exception as exc:
        return f"<tail failed: {exc}>\n"


def backup_or_quarantine_file(path: str, suffix: str) -> str:
    """Move a broken/obsolete data file aside without deleting user data."""
    try:
        if not path or not os.path.exists(path):
            return ""
        stamp = time.strftime("%Y%m%d_%H%M%S")
        dst = f"{path}.{suffix}_{stamp}"
        n = 1
        final = dst
        while os.path.exists(final):
            n += 1
            final = f"{dst}_{n}"
        os.replace(path, final)
        log_event(f"moved data file {os.path.basename(path)} -> {os.path.basename(final)}")
        return final
    except Exception as exc:
        log_error(exc)
        return ""


def load_json_file(path: str, expected_type: type, *, quarantine: bool = True) -> object | None:
    """Load JSON defensively. Corrupt config/pins must never break startup."""
    try:
        with open(path, "r", encoding="utf-8") as f:
            data = json.load(f)
        if not isinstance(data, expected_type):
            raise ValueError(f"expected {expected_type.__name__}, got {type(data).__name__}")
        return data
    except FileNotFoundError:
        return None
    except (json.JSONDecodeError, UnicodeError, ValueError) as exc:
        log_error(RuntimeError(f"Invalid JSON data file: {path}: {exc}"))
        if quarantine:
            backup_or_quarantine_file(path, "corrupt")
        return None
    except Exception as exc:
        log_error(exc)
        return None


def atomic_write_json(path: str, data: object, prefix: str) -> None:
    """Write JSON atomically so power loss or crash cannot leave a half file."""
    tmp = ""
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        fd, tmp = tempfile.mkstemp(prefix=prefix, suffix=".tmp", dir=os.path.dirname(path), text=True)
        with os.fdopen(fd, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
            f.write("\n")
        os.replace(tmp, path)
    except Exception as exc:
        log_error(exc)
        if tmp:
            try:
                os.remove(tmp)
            except Exception:
                pass


def clipboard_busy_owner() -> str:
    if not IS_WINDOWS:
        return ""
    try:
        hwnd = int(user32.GetOpenClipboardWindow() or 0) if hasattr(user32, "GetOpenClipboardWindow") else 0
        if not hwnd:
            return ""
        title = get_window_title(hwnd) if "get_window_title" in globals() else ""
        proc = get_window_process_name(hwnd) if "get_window_process_name" in globals() else ""
        parts = [x for x in (proc, title) if x]
        return " / ".join(parts)[:160]
    except Exception:
        return ""


def now_label() -> str:
    return time.strftime("%H:%M:%S")


def short_text(s: object, limit: int = 54) -> str:
    s2 = " ".join(str(s).replace("\r", " ").replace("\n", " ").split())
    if not s2:
        s2 = "(empty)"
    return s2 if len(s2) <= limit else s2[: max(1, limit - 1)] + "…"


def human_bytes(n: int) -> str:
    try:
        f = float(int(n))
    except Exception:
        return "0 B"
    units = ("B", "KB", "MB", "GB")
    for unit in units:
        if f < 1024 or unit == units[-1]:
            return f"{int(f)} B" if unit == "B" else f"{f:.1f} {unit}"
        f /= 1024
    return f"{int(n)} B"


def as_int(value: object, default: int, min_value: Optional[int] = None, max_value: Optional[int] = None) -> int:
    """Config-safe int parser. Invalid/corrupt JSON must never abort GUI startup."""
    try:
        if isinstance(value, bool):
            out = int(value)
        elif isinstance(value, (int, float)):
            out = int(value)
        else:
            out = int(str(value).strip())
    except Exception:
        out = int(default)
    if min_value is not None:
        out = max(int(min_value), out)
    if max_value is not None:
        out = min(int(max_value), out)
    return out


def as_float(value: object, default: float, min_value: Optional[float] = None, max_value: Optional[float] = None) -> float:
    try:
        out = float(value)
    except Exception:
        out = float(default)
    if min_value is not None:
        out = max(float(min_value), out)
    if max_value is not None:
        out = min(float(max_value), out)
    return out


def as_bool(value: object, default: bool = False) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        return bool(value)
    if isinstance(value, str):
        s = value.strip().lower()
        if s in {"1", "true", "yes", "y", "on", "enable", "enabled", "はい", "有効"}:
            return True
        if s in {"0", "false", "no", "n", "off", "disable", "disabled", "いいえ", "無効"}:
            return False
    return bool(default)


def norm_path(p: str) -> str:
    try:
        return os.path.normcase(os.path.abspath(p)) if p else ""
    except Exception:
        return p or ""


def guess_path_kind(p: str) -> str:
    """Classify file/folder accurately when cheap, without hanging on network paths."""
    if not p:
        return "file"
    try:
        # Avoid probing UNC/network paths on the clipboard hot path; network drives can block.
        if not str(p).startswith(("\\\\", "//")):
            if os.path.isdir(p):
                return "folder"
            if os.path.isfile(p):
                return "file"
    except Exception:
        pass
    tail = os.path.basename(p.rstrip("\\/"))
    if p.endswith(("\\", "/")):
        return "folder"
    if tail and not os.path.splitext(tail)[1]:
        return "folder"
    return "file"


_SECRET_WORD_RE = re.compile(
    r"(?i)(password|passwd|pwd|secret|token|api[_-]?key|access[_-]?key|bearer\s+[a-z0-9._\-]+|authorization:|private\s+key)"
)
_LONG_TOKEN_RE = re.compile(r"(?=.{24,})(?=.*[A-Za-z])(?=.*\d)[A-Za-z0-9_\-./+=]{24,}")
_OTP_RE = re.compile(r"^\s*\d{6,8}\s*$")


def looks_sensitive(text: str) -> bool:
    s = str(text or "").strip()
    if not s:
        return False
    if _SECRET_WORD_RE.search(s):
        return True
    if len(s) <= 12 and _OTP_RE.match(s):
        return True
    if " " not in s and "\n" not in s and _LONG_TOKEN_RE.search(s):
        return True
    if "-----BEGIN" in s and "KEY-----" in s:
        return True
    return False


def dib_dimensions(dib: bytes) -> tuple[int, int]:
    if not dib or len(dib) < 16:
        return 0, 0
    try:
        header_size = int.from_bytes(dib[0:4], "little", signed=False)
        if header_size == 12 and len(dib) >= 8:  # BITMAPCOREHEADER
            w = int.from_bytes(dib[4:6], "little", signed=False)
            h = int.from_bytes(dib[6:8], "little", signed=False)
            return int(w), int(h)
        if header_size >= 40 and len(dib) >= 12:
            w = int.from_bytes(dib[4:8], "little", signed=True)
            h = int.from_bytes(dib[8:12], "little", signed=True)
            return abs(int(w)), abs(int(h))
    except Exception:
        pass
    return 0, 0


def make_sig(kind: str, *, text: str = "", path: str = "", dib: bytes = b"", dib_format: int = CF_DIB) -> str:
    h = hashlib.sha256()
    h.update(kind.encode("ascii", "replace"))
    h.update(b"\0")
    if kind == "text":
        h.update(text.encode("utf-8", "replace"))
    elif kind == "image":
        h.update(str(int(dib_format or CF_DIB)).encode("ascii"))
        h.update(b"\0")
        h.update(dib)
    else:
        h.update(norm_path(path).encode("utf-8", "replace"))
    return h.hexdigest()


@dataclass(slots=True)
class TrayItem:
    id: str
    kind: str  # text / file / folder / image
    title: str
    text: str = ""
    path: str = ""
    detail: str = ""
    added_at: str = field(default_factory=now_label)
    sig: str = ""
    size: int = 0
    sensitive: bool = False
    dib: bytes = b""
    dib_format: int = CF_DIB
    search: str = ""
    pinned: bool = False

    def signature(self) -> str:
        return self.sig

    def preview(self, limit: int = 60) -> str:
        if self.kind == "text":
            return f"🔒 非表示テキスト（{len(self.text):,}文字）" if self.sensitive else short_text(self.text, limit)
        if self.kind == "image":
            return self.detail or self.title
        return short_text(self.path or self.title, limit)


def make_text_item(text: str) -> TrayItem:
    sensitive = looks_sensitive(text)
    title = f"🔒 非表示テキスト（{len(text):,}文字）" if sensitive else short_text(text, 70)
    detail = f"{len(text):,}文字" + (" / 表示は非表示" if sensitive else "")
    search = "" if sensitive else (title + " " + text[:8192]).lower()
    return TrayItem(
        id=str(uuid.uuid4()),
        kind="text",
        title=title,
        text=text,
        detail=detail,
        sig=make_sig("text", text=text),
        size=len(text.encode("utf-16le", "replace")),
        sensitive=sensitive,
        search=search,
    )


def make_path_item(path: str) -> TrayItem:
    kind = guess_path_kind(path)
    title = os.path.basename(path.rstrip("\\/")) or path
    return TrayItem(
        id=str(uuid.uuid4()),
        kind=kind,
        title=title,
        path=path,
        detail=path,
        sig=make_sig(kind, path=path),
        size=len(path.encode("utf-16le", "replace")) + 64,
        search=(title + " " + path).lower(),
    )


def make_image_item(dib: bytes, dib_format: int = CF_DIB) -> TrayItem:
    dib = bytes(dib or b"")
    w, h = dib_dimensions(dib)
    fmt_label = "DIBV5" if int(dib_format or CF_DIB) == CF_DIBV5 else "DIB"
    title = f"画像 {w}x{h}" if w and h else "クリップボード画像"
    detail = f"{human_bytes(len(dib))} / {fmt_label}"
    return TrayItem(
        id=str(uuid.uuid4()),
        kind="image",
        title=title,
        detail=detail,
        sig=make_sig("image", dib=dib, dib_format=dib_format),
        size=len(dib),
        dib=dib,
        dib_format=int(dib_format or CF_DIB),
        search=(title + " " + detail).lower(),
    )


class TrayStore:
    def __init__(self, max_items: int = 60, max_total_bytes: int = 96 * 1024 * 1024) -> None:
        self.items: list[TrayItem] = []
        self._sig: dict[str, TrayItem] = {}
        self.max_items = max(1, int(max_items))
        self.max_total_bytes = max(1_000_000, int(max_total_bytes))
        self.total_bytes = 0

    def item_by_signature(self, sig: str) -> Optional[TrayItem]:
        return self._sig.get(sig)

    def item_by_id(self, item_id: Optional[str]) -> Optional[TrayItem]:
        if not item_id:
            return None
        for it in self.items:
            if it.id == item_id:
                return it
        return None

    def _recalc(self) -> None:
        self._sig = {it.signature(): it for it in self.items}
        self.total_bytes = sum(max(0, int(it.size)) for it in self.items)

    def add(self, item: TrayItem, bump_duplicate: bool = True) -> tuple[bool, str, TrayItem]:
        existing = self._sig.get(item.signature())
        if existing:
            if bump_duplicate and self.items and self.items[0] is not existing:
                try:
                    self.items.remove(existing)
                    self.items.insert(0, existing)
                except ValueError:
                    pass
            if not existing.pinned:
                existing.added_at = now_label()
            self._recalc()
            return False, "DUPLICATE", existing
        self.items.insert(0, item)
        self._sig[item.signature()] = item
        self.total_bytes += max(0, int(item.size))
        self.trim()
        return True, "ADDED", item

    def add_many_preserve_order(self, items: Sequence[TrayItem]) -> tuple[int, int, list[TrayItem]]:
        added = dup = 0
        resolved_reversed: list[TrayItem] = []
        for item in reversed(list(items)):
            ok, reason, resolved = self.add(item, bump_duplicate=True)
            resolved_reversed.append(resolved)
            if ok:
                added += 1
            elif reason == "DUPLICATE":
                dup += 1
        resolved_reversed.reverse()
        return added, dup, resolved_reversed

    def _pop_last_unpinned(self) -> Optional[TrayItem]:
        for i in range(len(self.items) - 1, -1, -1):
            if not self.items[i].pinned:
                return self.items.pop(i)
        return None

    def trim(self) -> None:
        changed = False
        # ピン留めはユーザーが明示保存したものなので、通常のトリムでは消さない。
        while len(self.items) > self.max_items:
            old = self._pop_last_unpinned()
            if old is None:
                break
            self._sig.pop(old.signature(), None)
            self.total_bytes -= max(0, int(old.size))
            changed = True
        while len(self.items) > 1 and self.total_bytes > self.max_total_bytes:
            old = self._pop_last_unpinned()
            if old is None:
                break
            self._sig.pop(old.signature(), None)
            self.total_bytes -= max(0, int(old.size))
            changed = True
        if changed:
            self._recalc()

    def clear_unpinned(self) -> int:
        before = len(self.items)
        self.items = [it for it in self.items if it.pinned]
        self._recalc()
        return before - len(self.items)

    def remove_ids(self, ids: set[str]) -> int:
        before = len(self.items)
        self.items = [it for it in self.items if it.id not in ids]
        self._recalc()
        return before - len(self.items)

    def clear(self) -> None:
        self.items.clear()
        self._sig.clear()
        self.total_bytes = 0


class ClipboardError(RuntimeError):
    pass


class ClipboardBusy(ClipboardError):
    pass


class ClipboardTooLarge(ClipboardError):
    pass


class WinClipboard:
    def __init__(self) -> None:
        self.drop_effect_format: Optional[int] = None
        if IS_WINDOWS:
            try:
                self.drop_effect_format = int(user32.RegisterClipboardFormatW("Preferred DropEffect"))
            except Exception:
                self.drop_effect_format = None

    def sequence(self) -> int:
        if not IS_WINDOWS:
            return 0
        try:
            return int(user32.GetClipboardSequenceNumber())
        except Exception:
            return 0

    def open(self, timeout_ms: int = 300) -> None:
        if not IS_WINDOWS:
            raise ClipboardError("Windows専用機能です")
        deadline = time.monotonic() + max(1, int(timeout_ms)) / 1000.0
        sleep_s = 0.006
        while True:
            if user32.OpenClipboard(None):
                return
            if time.monotonic() >= deadline:
                owner = clipboard_busy_owner()
                detail = f" / 使用中: {owner}" if owner else ""
                raise ClipboardBusy("クリップボードを開けません。他アプリが使用中です。" + detail)
            time.sleep(sleep_s)
            sleep_s = min(0.025, sleep_s * 1.45)

    def close(self) -> None:
        if IS_WINDOWS:
            try:
                user32.CloseClipboard()
            except Exception:
                pass

    def read_text_open(self, max_chars: int) -> str:
        if not user32.IsClipboardFormatAvailable(CF_UNICODETEXT):
            return ""
        handle = user32.GetClipboardData(CF_UNICODETEXT)
        if not handle:
            return ""
        # CF_UNICODETEXT is UTF-16LE. Never use wstring_at with max_chars here:
        # it reads exactly that many wchar_t units and can over-read a small HGLOBAL,
        # which is a plausible cause of sudden .pyw process disappearance.
        size = 0
        try:
            size = int(kernel32.GlobalSize(handle) or 0)
        except Exception:
            size = 0
        if size <= 0:
            return ""
        max_bytes = (int(max_chars) + 1) * 2
        if size > max_bytes:
            raise ClipboardTooLarge(f"テキストが大きすぎます: {human_bytes(size)}")
        ptr = kernel32.GlobalLock(handle)
        if not ptr:
            return ""
        try:
            read_bytes = min(size, max_bytes)
            read_bytes -= read_bytes % 2
            if read_bytes <= 0:
                return ""
            raw = ctypes.string_at(ptr, read_bytes)
        finally:
            kernel32.GlobalUnlock(handle)
        text = raw.decode("utf-16le", "replace") if raw else ""
        nul = text.find("\0")
        if nul >= 0:
            text = text[:nul]
        if len(text) > max_chars:
            raise ClipboardTooLarge(f"テキストが大きすぎます: {len(text):,} chars")
        return text

    def read_paths_open(self, max_paths: int = 500) -> list[str]:
        if not user32.IsClipboardFormatAvailable(CF_HDROP):
            return []
        handle = user32.GetClipboardData(CF_HDROP)
        if not handle:
            return []
        count = int(shell32.DragQueryFileW(handle, 0xFFFFFFFF, None, 0))
        if count <= 0:
            return []
        out: list[str] = []
        for i in range(min(count, max_paths)):
            n = int(shell32.DragQueryFileW(handle, i, None, 0))
            if n <= 0:
                continue
            buf = ctypes.create_unicode_buffer(n + 1)
            shell32.DragQueryFileW(handle, i, buf, n + 1)
            if buf.value:
                out.append(buf.value)
        return out

    def read_dib_open(self, max_bytes: int) -> tuple[bytes, int]:
        for fmt in (CF_DIB, CF_DIBV5):
            if not user32.IsClipboardFormatAvailable(fmt):
                continue
            handle = user32.GetClipboardData(fmt)
            if not handle:
                continue
            size = int(kernel32.GlobalSize(handle) or 0)
            if size <= 0:
                continue
            if size > max_bytes:
                raise ClipboardTooLarge(f"画像が大きすぎます: {human_bytes(size)} > {human_bytes(max_bytes)}")
            ptr = kernel32.GlobalLock(handle)
            if not ptr:
                continue
            try:
                data = ctypes.string_at(ptr, size)
            finally:
                kernel32.GlobalUnlock(handle)
            if data:
                return data, int(fmt)
        return b"", 0

    def snapshot_items(self, *, timeout_ms: int = 300, max_text_chars: int = 500_000, max_image_bytes: int = 32 * 1024 * 1024, include_images: bool = True, priority: str = "text_first") -> list[TrayItem]:
        if not IS_WINDOWS:
            return []
        priority = str(priority or "text_first").lower().strip()
        self.open(timeout_ms=timeout_ms)
        try:
            if priority not in ("text_only", "image_only"):
                paths = self.read_paths_open()
                if paths:
                    return [make_path_item(p) for p in paths]

            def text_item() -> list[TrayItem]:
                text = self.read_text_open(max_text_chars)
                if text and text.strip():
                    return [make_text_item(text)]
                return []

            def image_item() -> list[TrayItem]:
                if not include_images:
                    return []
                dib, fmt = self.read_dib_open(max_image_bytes)
                if dib:
                    return [make_image_item(dib, fmt)]
                return []

            if priority == "text_only":
                return text_item()
            if priority == "image_only":
                return image_item()
            if priority == "image_first":
                return image_item() or text_item()
            # text_first: avoids costly DIB reads when Office/browser copies also expose plain text.
            return text_item() or image_item()
        finally:
            self.close()

    def read_text(self) -> str:
        self.open(timeout_ms=300)
        try:
            return self.read_text_open(max_chars=2_000_000)
        finally:
            self.close()

    def _alloc_unicode(self, text: str) -> int:
        data = (text + "\0").encode("utf-16le", "replace")
        h = kernel32.GlobalAlloc(GHND, len(data))
        if not h:
            raise ClipboardError("GlobalAlloc text failed")
        ptr = kernel32.GlobalLock(h)
        if not ptr:
            kernel32.GlobalFree(h)
            raise ClipboardError("GlobalLock text failed")
        try:
            ctypes.memmove(int(ptr), data, len(data))
        finally:
            kernel32.GlobalUnlock(h)
        return int(h)

    def _alloc_bytes(self, data: bytes) -> int:
        if not data:
            raise ClipboardError("empty binary clipboard payload")
        h = kernel32.GlobalAlloc(GHND, len(data))
        if not h:
            raise ClipboardError("GlobalAlloc binary failed")
        ptr = kernel32.GlobalLock(h)
        if not ptr:
            kernel32.GlobalFree(h)
            raise ClipboardError("GlobalLock binary failed")
        try:
            ctypes.memmove(int(ptr), data, len(data))
        finally:
            kernel32.GlobalUnlock(h)
        return int(h)

    def _alloc_hdrop(self, paths: Sequence[str]) -> int:
        clean = [str(p) for p in paths if p]
        block = ("\0".join(clean) + "\0\0").encode("utf-16le", "replace")
        header = DROPFILES()
        header.pFiles = ctypes.sizeof(DROPFILES)
        header.pt = POINT(0, 0)
        header.fNC = False
        header.fWide = True
        size = ctypes.sizeof(DROPFILES) + len(block)
        h = kernel32.GlobalAlloc(GHND, size)
        if not h:
            raise ClipboardError("GlobalAlloc hdrop failed")
        ptr = kernel32.GlobalLock(h)
        if not ptr:
            kernel32.GlobalFree(h)
            raise ClipboardError("GlobalLock hdrop failed")
        try:
            base = int(ptr)
            ctypes.memmove(base, ctypes.byref(header), ctypes.sizeof(DROPFILES))
            ctypes.memmove(base + ctypes.sizeof(DROPFILES), block, len(block))
        finally:
            kernel32.GlobalUnlock(h)
        return int(h)

    def _set_drop_effect(self, allocated: list[int]) -> None:
        if not self.drop_effect_format:
            return
        h = 0
        try:
            h = int(kernel32.GlobalAlloc(GHND, ctypes.sizeof(wintypes.DWORD)) or 0)
            if not h:
                return
            allocated.append(h)
            ptr = kernel32.GlobalLock(h)
            if not ptr:
                return
            try:
                val = wintypes.DWORD(DROPEFFECT_COPY)
                ctypes.memmove(int(ptr), ctypes.byref(val), ctypes.sizeof(val))
            finally:
                kernel32.GlobalUnlock(h)
            if user32.SetClipboardData(self.drop_effect_format, h):
                allocated.remove(h)
        except Exception:
            pass

    def _write_raw(self, *, text: str = "", paths: Sequence[str] = (), dib: bytes = b"", dib_format: int = CF_DIB) -> None:
        if not IS_WINDOWS:
            raise ClipboardError("Windows専用機能です")
        clean_paths = [str(p) for p in paths if p]
        dib = bytes(dib or b"")
        if not text and not clean_paths and not dib:
            raise ClipboardError("送る内容がありません。")

        # 先に全データをGlobalAllocしておく。
        # これにより、メモリ確保失敗では既存のWindowsクリップボードを空にしない。
        prepared: list[tuple[int, int]] = []
        transferred: set[int] = set()
        try:
            if text:
                prepared.append((CF_UNICODETEXT, self._alloc_unicode(text)))
            if clean_paths:
                prepared.append((CF_HDROP, self._alloc_hdrop(clean_paths)))
                if self.drop_effect_format:
                    h_eff = int(kernel32.GlobalAlloc(GHND, ctypes.sizeof(wintypes.DWORD)) or 0)
                    if h_eff:
                        ptr = kernel32.GlobalLock(h_eff)
                        if ptr:
                            try:
                                val = wintypes.DWORD(DROPEFFECT_COPY)
                                ctypes.memmove(int(ptr), ctypes.byref(val), ctypes.sizeof(val))
                            finally:
                                kernel32.GlobalUnlock(h_eff)
                            prepared.append((int(self.drop_effect_format), h_eff))
                        else:
                            kernel32.GlobalFree(h_eff)
            if dib:
                fmt = int(dib_format or CF_DIB)
                if fmt not in (CF_DIB, CF_DIBV5):
                    fmt = CF_DIB
                prepared.append((fmt, self._alloc_bytes(dib)))

            self.open(timeout_ms=500)
            try:
                if not user32.EmptyClipboard():
                    raise ClipboardError("クリップボード初期化に失敗しました。")
                for fmt, h in prepared:
                    if not user32.SetClipboardData(fmt, h):
                        raise ClipboardError(f"クリップボード形式 {fmt} を設定できません。")
                    transferred.add(h)
            finally:
                self.close()
        finally:
            # SetClipboardData成功後のハンドルはWindowsが所有するため解放しない。
            for _fmt, h in prepared:
                if h not in transferred:
                    try:
                        kernel32.GlobalFree(h)
                    except Exception:
                        pass

    def write(self, *, text: str = "", paths: Sequence[str] = (), dib: bytes = b"", dib_format: int = CF_DIB, backup: bool = True) -> None:
        backup_items: list[TrayItem] = []
        if backup:
            try:
                backup_items = self.snapshot_items(timeout_ms=100, max_text_chars=500_000, max_image_bytes=16 * 1024 * 1024)
            except (ClipboardBusy, ClipboardTooLarge):
                backup_items = []
            except Exception as exc:
                log_error(exc)
        try:
            self._write_raw(text=text, paths=paths, dib=dib, dib_format=dib_format)
        except Exception:
            if backup_items:
                try:
                    b_text, b_paths, b_dib, b_fmt, _ = collect_payload(backup_items)
                    if b_text or b_paths or b_dib:
                        self._write_raw(text=b_text, paths=b_paths, dib=b_dib, dib_format=b_fmt)
                except Exception as restore_exc:
                    log_error(restore_exc)
            raise

    def clear(self) -> None:
        if not IS_WINDOWS:
            raise ClipboardError("Windows専用機能です")
        self.open(timeout_ms=500)
        try:
            user32.EmptyClipboard()
        finally:
            self.close()


def collect_payload(items: Sequence[TrayItem]) -> tuple[str, list[str], bytes, int, int]:
    texts: list[str] = []
    paths: list[str] = []
    dib = b""
    dib_format = CF_DIB
    image_count = 0
    for it in items:
        if it.kind == "text":
            texts.append(it.text)
        elif it.kind in ("file", "folder") and it.path:
            paths.append(it.path)
            texts.append(it.path)
        elif it.kind == "image":
            image_count += 1
            if not dib:
                dib = it.dib
                dib_format = int(it.dib_format or CF_DIB)
    return "\r\n".join(texts), paths, dib, dib_format, image_count


def kind_label(it: TrayItem) -> str:
    if it.kind == "text":
        return "文字"
    if it.kind == "folder":
        return "フォルダ"
    if it.kind == "image":
        return "画像"
    return "ファイル"

def action_label_ja(label: str) -> str:
    label = str(label or "")
    replacements = {
        "SLOT": "候補",
        "ALL": "全項目",
        "PICK": "一覧",
        "SELECTED": "選択項目",
        "TEXT VARIANT": "変換テキスト",
        "TEXT": "文字",
        "IMAGE": "画像",
        "PATH": "パス",
        "QUOTED PATH": "引用符付きパス",
        "BASENAME": "ファイル名",
        "PARENT PATH": "親フォルダ",
        "MARKDOWN LINK": "Markdownリンク",
        "VARIANT": "変換テキスト",
        "SELF TEST": "自己診断",
        "RESTORED": "復元",
    }
    # Longer phrases first.
    for src in sorted(replacements, key=len, reverse=True):
        label = label.replace(src, replacements[src])
    return label.strip()


def ja_status(msg: str) -> str:
    msg = str(msg or "")
    direct = {
        "IDLE": "待機中",
        "READY": "準備完了",
        "MENU FAILED / Ctrl+Qで終了": "メニューを開けません。Ctrl+Qで終了できます。",
        "ERROR / error.log を確認": "エラーが発生しました。error.logを確認してください。",
        "PIN LOAD ERROR / 既存ピン保存は保留": "ピン留めの読み込みに失敗しました。既存ピンの保存は保留中です。",
        "CAPTURE READY": "自動取り込みできます",
        "CAPTURE PAUSED": "自動取り込み停止中",
        "CAPTURE ON": "自動取り込みON",
        "BACKGROUND CAPTURE ON": "格納中の裏取り込みON",
        "BACKGROUND CAPTURE OFF": "格納中の裏取り込みOFF",
        "PRIVACY GUARD ON": "プライバシー保護ON",
        "PRIVACY GUARD OFF": "プライバシー保護OFF",
        "PRIVACY GUARD SKIP": "プライバシー保護により取り込みをスキップしました",
        "SENSITIVE TEXT SKIPPED": "機密情報らしいテキストの自動取り込みをスキップしました",
        "GLOBAL HOTKEYS ON": "全体ホットキーON",
        "GLOBAL HOTKEYS OFF": "全体ホットキーOFF",
        "SAFE MODE: HOTKEYS OFF": "セーフモードのためホットキーはOFFです",
        "IMAGE FIRST": "画像優先",
        "TEXT FIRST": "文字優先",
        "PASTE ADVANCE ON": "貼って次へON",
        "PASTE ADVANCE OFF": "貼って次へOFF",
        "AUTO CLEAR OFF": "自動消去OFF",
        "CLIPBOARD EMPTY / UNSUPPORTED": "クリップボードが空、または未対応の形式です",
        "CLIPBOARD BUSY": "クリップボードが使用中です",
        "ADD FAILED": "追加に失敗しました",
        "ALREADY IN TRAY / MARKED": "既にトレーにあります。貼付候補として印を付けました",
        "NO ACTIVE SLOT": "選択中の候補がありません",
        "TRAY EMPTY": "トレーは空です",
        "COPY FAILED / RESTORED IF POSSIBLE": "コピーに失敗しました。可能な場合は元の内容を復元しました",
        "PASTED": "貼り付けました",
        "CAPTURING…": "選択範囲を取得中…",
        "WINDOWS CLIPBOARD WIPED": "Windowsクリップボードを空にしました",
        "WIPE FAILED": "消去に失敗しました",
        "ALL WIPED": "すべて消去しました",
        "PINNED ONLY": "ピン留め項目だけ残っています",
        "ALL TRAY ITEMS CLEARED": "仮想トレーを全消去しました",
        "UNDO EMPTY": "戻せる操作がありません",
        "UNDO FAILED": "戻す操作に失敗しました",
        "MARKDOWN LINK FAILED": "Markdownリンクの作成に失敗しました",
        "REVEAL FAILED": "Explorerでの表示に失敗しました",
        "MOVED TO TOP": "一番上へ移動しました",
        "MOVED": "移動しました",
        "NO TEXT/PATH ITEM": "文字またはパスの項目がありません",
        "TEXT ONLY": "文字項目だけで使えます",
        "EMPTY TEXT": "文字が空です",
        "SAME TEXT ALREADY EXISTS": "同じ文字が既にあります",
        "TEXT UPDATED": "文字を更新しました",
        "OPENED TEMP TEXT": "一時テキストを開きました",
        "OPEN TEXT FAILED": "テキストを開けませんでした",
        "EMPTY NOTE": "メモが空です",
        "NOTE ADDED": "メモを追加しました",
        "NOTE ALREADY EXISTS": "同じメモが既にあります",
        "NOTE FAILED": "メモの追加に失敗しました",
        "TEXT VARIANT READY": "変換した文字をセットしました",
        "VARIANT COPY FAILED": "変換テキストのコピーに失敗しました",
        "PATH EMPTY": "パスが空です",
        "OPEN FAILED": "開けませんでした",
        "OPEN FOLDER FAILED": "フォルダを開けませんでした",
        "EXPORTED TXT": "テキストを書き出しました",
        "EXPORT FAILED": "書き出しに失敗しました",
        "DIAGNOSTIC READY": "診断レポートを作成しました",
        "DIAGNOSTIC FAILED": "診断レポートの作成に失敗しました",
        "POSITION RESET": "待機ボタン位置を右上へ戻しました",
        "POSITION RESET FAILED": "位置のリセットに失敗しました",
        "WINDOWS ONLY": "Windows専用機能です",
        "STARTUP ON": "Windows起動時の自動起動を登録しました",
        "STARTUP OFF": "Windows起動時の自動起動を解除しました",
        "STARTUP FAILED": "自動起動設定に失敗しました",
        "SELF TEST PASS": "自己診断に成功しました",
        "SELF TEST PASS / CLIP RESTORED": "自己診断に成功し、クリップボードを復元しました",
        "SELF TEST FAIL": "自己診断に失敗しました",
        "IMAGE PIN is session-only": "画像のピン留めはこのセッション中だけ有効です",
        "PINNED": "ピン留めしました",
        "UNPINNED": "ピン留めを解除しました",
    }
    if msg in direct:
        return direct[msg]
    m = re.fullmatch(r"ACTIVE SLOT \[(.*?)\]", msg)
    if m:
        return f"選択 [{m.group(1)}]"
    m = re.fullmatch(r"NO SLOT \[(.*?)\]", msg)
    if m:
        return f"候補 [{m.group(1)}] はありません"
    m = re.fullmatch(r"ADDED (\d+)(?: / DUP (\d+))?", msg)
    if m:
        return f"{m.group(1)}件追加" + (f" / 重複{m.group(2)}件" if m.group(2) else "")
    m = re.fullmatch(r"AUTO CLEARED (\d+) SESSION ITEM\(S\)", msg)
    if m:
        return f"セッション項目を{m.group(1)}件自動消去しました"
    m = re.fullmatch(r"AUTO CLEAR (\d+) MIN", msg)
    if m:
        return f"自動消去：{m.group(1)}分"
    m = re.fullmatch(r"SELECTED (\d+)", msg)
    if m:
        return f"{m.group(1)}件選択しました"
    m = re.fullmatch(r"REMOVED (\d+)", msg)
    if m:
        return f"{m.group(1)}件削除しました"
    m = re.fullmatch(r"SESSION CLEARED (\d+)", msg)
    if m:
        return f"セッション項目を{m.group(1)}件消去しました"
    m = re.fullmatch(r"UNDO (.*?): RESTORED (\d+)(?: / DUP (\d+))?", msg)
    if m:
        label = action_label_ja(m.group(1))
        return f"戻しました：{label} / 復元{m.group(2)}件" + (f" / 重複{m.group(3)}件" if m.group(3) else "")
    m = re.fullmatch(r"CLIPBOARD READY\s*(.*?)\s*(\d+) item\(s\)(.*)", msg)
    if m:
        label = action_label_ja(m.group(1))
        extra = m.group(3) or ""
        if "first image only" in extra:
            extra = " / 画像は先頭1枚のみ"
        elif "image" in extra:
            extra = " / 画像"
        else:
            extra = ""
        prefix = f"{label} " if label else ""
        return f"クリップボードへセットしました：{prefix}{m.group(2)}件{extra}"
    m = re.fullmatch(r"(.*?) READY", msg)
    if m:
        label = action_label_ja(m.group(1))
        return f"{label} をセットしました" if label else "セットしました"
    m = re.fullmatch(r"(.*?) FAILED", msg)
    if m:
        label = action_label_ja(m.group(1))
        return f"{label} に失敗しました" if label else "失敗しました"
    return action_label_ja(msg)


def text_variant(text: str, mode: str) -> str:
    s = str(text or "")
    if mode == "one_line":
        return " ".join(s.split())
    if mode == "trim":
        return s.strip()
    if mode == "rstrip_newline":
        return s.rstrip("\r\n")
    if mode == "lower":
        return s.lower()
    if mode == "upper":
        return s.upper()
    if mode == "bullet":
        lines = [line.rstrip() for line in s.splitlines() if line.strip()]
        return "\r\n".join("- " + line.lstrip("-•* ").strip() for line in lines)
    if mode == "quote":
        return "\r\n".join("> " + line for line in s.splitlines())
    if mode == "codeblock":
        return "```\r\n" + s.strip("\r\n") + "\r\n```"
    if mode == "url_encode":
        return urllib.parse.quote(s, safe="")
    if mode == "url_decode":
        return urllib.parse.unquote(s)
    if mode == "html_escape":
        return html.escape(s)
    if mode == "html_unescape":
        return html.unescape(s)
    if mode == "json_string":
        return json.dumps(s, ensure_ascii=False)
    if mode == "json_pretty":
        try:
            return json.dumps(json.loads(s), ensure_ascii=False, indent=2)
        except Exception:
            return s
    if mode == "nfkc":
        return unicodedata.normalize("NFKC", s)
    if mode == "nfc":
        return unicodedata.normalize("NFC", s)
    if mode == "dedupe_lines":
        seen: set[str] = set()
        out: list[str] = []
        for line in s.splitlines():
            key = line.strip()
            if key in seen:
                continue
            seen.add(key)
            out.append(line.rstrip())
        return "\r\n".join(out)
    if mode == "sort_lines":
        lines = [line.rstrip() for line in s.splitlines() if line.strip()]
        return "\r\n".join(sorted(lines, key=lambda x: x.casefold()))
    if mode == "collapse_blank_lines":
        out: list[str] = []
        blank = False
        for line in s.splitlines():
            if line.strip():
                out.append(line.rstrip())
                blank = False
            elif not blank:
                out.append("")
                blank = True
        return "\r\n".join(out).strip("\r\n")
    if mode == "tabs_to_spaces":
        return s.expandtabs(4)
    if mode == "remove_blank_lines":
        return "\r\n".join(line.rstrip() for line in s.splitlines() if line.strip())
    if mode == "join_comma":
        return ", ".join(line.strip() for line in s.splitlines() if line.strip())
    if mode == "join_tab":
        return "\t".join(line.strip() for line in s.splitlines() if line.strip())
    if mode == "markdown_escape_table":
        return s.replace("\\", "\\\\").replace("|", "\\|")
    return s


def refresh_text_item_fields(it: TrayItem, text: str) -> None:
    new = make_text_item(text)
    it.kind = "text"
    it.title = new.title
    it.text = new.text
    it.path = ""
    it.detail = new.detail
    it.sig = new.sig
    it.size = new.size
    it.sensitive = new.sensitive
    it.dib = b""
    it.dib_format = CF_DIB
    it.search = (new.search + (" pinned pin 固定 ピン" if it.pinned else "")).lower()


def is_own_process_window(hwnd: Optional[int]) -> bool:
    if not IS_WINDOWS or not hwnd:
        return False
    try:
        pid = wintypes.DWORD(0)
        user32.GetWindowThreadProcessId(wintypes.HWND(int(hwnd)), ctypes.byref(pid))
        return int(pid.value) == int(kernel32.GetCurrentProcessId())
    except Exception:
        return False


def valid_external_window(hwnd: Optional[int]) -> bool:
    if not IS_WINDOWS or not hwnd:
        return False
    try:
        h = wintypes.HWND(int(hwnd))
        if not user32.IsWindow(h):
            return False
        if is_own_process_window(hwnd):
            return False
        return True
    except Exception:
        return False



def get_window_title(hwnd: Optional[int]) -> str:
    if not IS_WINDOWS or not hwnd:
        return ""
    try:
        h = wintypes.HWND(int(hwnd))
        n = int(user32.GetWindowTextLengthW(h) or 0)
        if n <= 0:
            return ""
        buf = ctypes.create_unicode_buffer(n + 1)
        user32.GetWindowTextW(h, buf, n + 1)
        return buf.value or ""
    except Exception:
        return ""


def get_window_process_name(hwnd: Optional[int]) -> str:
    if not IS_WINDOWS or not hwnd:
        return ""
    handle = None
    try:
        pid = wintypes.DWORD(0)
        user32.GetWindowThreadProcessId(wintypes.HWND(int(hwnd)), ctypes.byref(pid))
        if not pid.value:
            return ""
        handle = kernel32.OpenProcess(PROCESS_QUERY_LIMITED_INFORMATION, False, pid.value)
        if not handle:
            return ""
        size = wintypes.DWORD(32768)
        buf = ctypes.create_unicode_buffer(size.value)
        if not kernel32.QueryFullProcessImageNameW(handle, 0, buf, ctypes.byref(size)):
            return ""
        return os.path.basename(buf.value or "")
    except Exception:
        return ""
    finally:
        if handle:
            try:
                kernel32.CloseHandle(handle)
            except Exception:
                pass


def virtual_screen_bounds(root: Optional[tk.Misc] = None) -> tuple[int, int, int, int]:
    if IS_WINDOWS:
        try:
            x = int(user32.GetSystemMetrics(SM_XVIRTUALSCREEN))
            y = int(user32.GetSystemMetrics(SM_YVIRTUALSCREEN))
            w = int(user32.GetSystemMetrics(SM_CXVIRTUALSCREEN))
            h = int(user32.GetSystemMetrics(SM_CYVIRTUALSCREEN))
            if w > 0 and h > 0:
                return x, y, w, h
        except Exception:
            pass
    try:
        if root is not None:
            return 0, 0, max(1, root.winfo_screenwidth()), max(1, root.winfo_screenheight())
    except Exception:
        pass
    return 0, 0, 1280, 720


def tk_geometry_spec(w: int, h: int, x: int, y: int) -> str:
    """Tk geometry string that handles negative monitor coordinates correctly."""
    return f"{int(w)}x{int(h)}{int(x):+d}{int(y):+d}"


def move_tk_window(win: tk.Misc, w: int, h: int, x: int, y: int) -> None:
    """Move using Win32 absolute coordinates where possible, so negative-monitor layouts behave better."""
    try:
        win.geometry(f"{int(w)}x{int(h)}")
        win.update_idletasks()
        if IS_WINDOWS:
            hwnd = int(win.winfo_id())
            if hwnd:
                user32.MoveWindow(wintypes.HWND(hwnd), int(x), int(y), int(w), int(h), True)
                return
        # Non-Windows fallback.
        win.geometry(tk_geometry_spec(w, h, x, y))
    except Exception:
        try:
            win.geometry(tk_geometry_spec(w, h, x, y))
        except Exception:
            pass

def window_thread_id(hwnd: Optional[int]) -> int:
    if not IS_WINDOWS or not hwnd:
        return 0
    try:
        pid = wintypes.DWORD(0)
        return int(user32.GetWindowThreadProcessId(wintypes.HWND(int(hwnd)), ctypes.byref(pid)) or 0)
    except Exception:
        return 0


def focus_external_window(hwnd: Optional[int], timeout_s: float = 0.32) -> bool:
    """Bring the last real target window back before sending Ctrl+C/Ctrl+V."""
    if not valid_external_window(hwnd):
        return False
    attached: list[tuple[int, int]] = []
    try:
        h = wintypes.HWND(int(hwnd))
        if user32.IsIconic(h):
            user32.ShowWindow(h, SW_RESTORE)
        else:
            try:
                user32.ShowWindow(h, SW_SHOW)
            except Exception:
                pass

        current_tid = int(kernel32.GetCurrentThreadId()) if hasattr(kernel32, "GetCurrentThreadId") else 0
        target_tid = window_thread_id(hwnd)
        foreground = int(user32.GetForegroundWindow() or 0)
        foreground_tid = window_thread_id(foreground)
        for tid in {target_tid, foreground_tid}:
            if current_tid and tid and tid != current_tid and hasattr(user32, "AttachThreadInput"):
                try:
                    if user32.AttachThreadInput(wintypes.DWORD(current_tid), wintypes.DWORD(tid), True):
                        attached.append((current_tid, tid))
                except Exception:
                    pass
        try:
            user32.BringWindowToTop(h)
        except Exception:
            pass
        try:
            user32.SetActiveWindow(h)
        except Exception:
            pass
        try:
            user32.SetForegroundWindow(h)
        except Exception:
            pass
        deadline = time.monotonic() + max(0.05, timeout_s)
        while time.monotonic() < deadline:
            if int(user32.GetForegroundWindow() or 0) == int(hwnd):
                return True
            time.sleep(0.025)
        return int(user32.GetForegroundWindow() or 0) == int(hwnd)
    except Exception as exc:
        log_error(exc)
        return False
    finally:
        for cur, tid in reversed(attached):
            try:
                user32.AttachThreadInput(wintypes.DWORD(cur), wintypes.DWORD(tid), False)
            except Exception:
                pass


def send_key_chord_to(hwnd: Optional[int], keys: Sequence[int]) -> bool:
    if not focus_external_window(hwnd):
        return False
    keys = [int(k) for k in keys if int(k) > 0]
    if not keys:
        return False
    try:
        n = len(keys) * 2
        arr = (INPUT * n)()
        j = 0
        for vk in keys:
            arr[j].type = INPUT_KEYBOARD
            arr[j].union.ki = KEYBDINPUT(wVk=wintypes.WORD(vk), wScan=0, dwFlags=0, time=0, dwExtraInfo=0)
            j += 1
        for vk in reversed(keys):
            arr[j].type = INPUT_KEYBOARD
            arr[j].union.ki = KEYBDINPUT(wVk=wintypes.WORD(vk), wScan=0, dwFlags=KEYEVENTF_KEYUP, time=0, dwExtraInfo=0)
            j += 1
        sent = int(user32.SendInput(wintypes.UINT(n), arr, ctypes.sizeof(INPUT)) or 0)
        if sent == n:
            return True
    except Exception as exc:
        log_error(exc)
    return False


def _send_ctrl_key_to(hwnd: Optional[int], vk: int) -> bool:
    if not valid_external_window(hwnd):
        return False
    try:
        if hasattr(user32, "SendInput") and send_key_chord_to(hwnd, [VK_CONTROL, int(vk)]):
            return True
    except Exception as exc:
        log_error(exc)

    ctrl_down = False
    key_down = False
    try:
        if not focus_external_window(hwnd):
            return False
        user32.keybd_event(VK_CONTROL, 0, 0, 0)
        ctrl_down = True
        user32.keybd_event(vk, 0, 0, 0)
        key_down = True
        time.sleep(0.015)
        user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)
        key_down = False
        user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)
        ctrl_down = False
        return True
    except Exception as exc:
        log_error(exc)
        return False
    finally:
        try:
            if key_down:
                user32.keybd_event(vk, 0, KEYEVENTF_KEYUP, 0)
            if ctrl_down:
                user32.keybd_event(VK_CONTROL, 0, KEYEVENTF_KEYUP, 0)
        except Exception:
            pass

def send_ctrl_v_to(hwnd: Optional[int]) -> bool:
    return _send_ctrl_key_to(hwnd, VK_V)


def send_ctrl_c_to(hwnd: Optional[int]) -> bool:
    return _send_ctrl_key_to(hwnd, VK_C)


def load_config() -> dict:
    cfg = dict(DEFAULT_CONFIG)
    for cfg_path in (CONFIG_PATH, OLD_CONFIG_PATH, OLDER_CONFIG_PATH):
        if not os.path.exists(cfg_path):
            continue
        data = load_json_file(cfg_path, dict, quarantine=True)
        if isinstance(data, dict):
            cfg.update(data)
            break
    return cfg


def save_config(data: dict) -> None:
    atomic_write_json(CONFIG_PATH, data, "config_v2_2_")


def pinned_to_dict(it: TrayItem) -> Optional[dict]:
    if it.kind == "text":
        return {"kind": "text", "text": it.text, "added_at": it.added_at}
    if it.kind in ("file", "folder") and it.path:
        return {"kind": it.kind, "path": it.path, "added_at": it.added_at}
    return None


def load_pinned_items() -> list[TrayItem]:
    out: list[TrayItem] = []
    pin_data = None
    for pin_path in (PINNED_PATH, OLD_PINNED_PATH, OLDER_PINNED_PATH):
        if not os.path.exists(pin_path):
            continue
        pin_data = load_json_file(pin_path, list, quarantine=True)
        if isinstance(pin_data, list):
            break
    if not isinstance(pin_data, list):
        return out
    try:
        seen: set[str] = set()
        for row in pin_data[:500]:
            if not isinstance(row, dict):
                continue
            kind = str(row.get("kind") or "")
            if kind == "text":
                text = str(row.get("text") or "")
                if not text:
                    continue
                it = make_text_item(text)
            elif kind in ("file", "folder"):
                path = str(row.get("path") or "")
                if not path:
                    continue
                it = make_path_item(path)
                it.kind = kind
                it.sig = make_sig(kind, path=path)
            else:
                continue
            if it.signature() in seen:
                continue
            seen.add(it.signature())
            it.pinned = True
            it.added_at = str(row.get("added_at") or "PIN")[:16]
            it.search = (it.search + " pinned pin 固定 ピン").lower()
            out.append(it)
    except Exception as exc:
        log_error(exc)
    return out


def save_pinned_items(items: Sequence[TrayItem]) -> None:
    rows = []
    try:
        for it in items:
            if not it.pinned:
                continue
            row = pinned_to_dict(it)
            if row is not None:
                rows.append(row)
        atomic_write_json(PINNED_PATH, rows, "pinned_v2_2_")
    except Exception as exc:
        log_error(exc)

class ClipboardUpdateListener:
    """Native Windows message bridge for clipboard updates and global hotkeys.

    The previous build subclassed Tk's HWND. That works on many machines, but it is
    the riskiest part of a .pyw utility because any WNDPROC pointer/signature issue
    can terminate the process without a Python traceback. This bridge instead owns
    a tiny hidden Win32 window on a daemon thread and forwards events to Tk through
    a Queue; all Tk work still runs on the Tk main thread.
    """

    def __init__(self, app: "DeskLayerApp") -> None:
        self.app = app
        self.enabled = False
        self.hwnd = 0
        self._thread: Optional[threading.Thread] = None
        self._ready = threading.Event()
        self._stop_requested = False
        self._wndproc = None
        self._hinstance = 0
        self._class_registered = False
        self._clipboard_registered = False
        self._start_error: Optional[BaseException] = None
        self._events: "queue.Queue[tuple[str, int]]" = queue.Queue(maxsize=256)
        self._hotkey_lock = threading.RLock()
        self._message_thread_ident = 0
        self.registered_hotkeys: list[int] = []
        self.class_name = f"DeskLayerVirtualTrayMsg_{os.getpid()}_{id(self):x}"

    def start(self) -> bool:
        if not IS_WINDOWS:
            return False
        try:
            self._thread = threading.Thread(target=self._thread_main, name="DeskLayerNativeEvents", daemon=True)
            self._thread.start()
            self._ready.wait(1.2)
            # Start draining even if the clipboard listener failed but the thread/window exists;
            # this preserves global hotkeys where possible while fallback polling handles clipboard.
            if self._thread.is_alive() or self.hwnd:
                self.app.root.after(60, self.drain_events)
            if self._start_error is not None:
                log_error(self._start_error)
            return bool(self._clipboard_registered)
        except Exception as exc:
            self._start_error = exc
            log_error(exc)
            self.stop()
            return False

    def _push_event(self, name: str, value: int = 0) -> None:
        try:
            self._events.put_nowait((name, int(value)))
        except queue.Full:
            # Clipboard updates can be coalesced; dropping is better than blocking WNDPROC.
            pass
        except Exception as exc:
            log_error(exc)

    def drain_events(self) -> None:
        if getattr(self.app, "_shutting_down", False):
            return
        try:
            for _ in range(128):
                try:
                    name, value = self._events.get_nowait()
                except queue.Empty:
                    break
                if name == "clipboard":
                    self.app.schedule_clipboard_event()
                elif name == "hotkey":
                    self.app.handle_hotkey(int(value))
        except Exception as exc:
            log_error(exc)
        finally:
            alive = self.enabled or (self._thread is not None and self._thread.is_alive())
            if alive and not getattr(self.app, "_shutting_down", False):
                self.app.root.after(60, self.drain_events)

    def _thread_main(self) -> None:
        if not IS_WINDOWS:
            self._ready.set()
            return
        try:
            self._message_thread_ident = threading.get_ident()
            self._create_message_window()
            if not self.hwnd:
                self._ready.set()
                return
            try:
                self._clipboard_registered = bool(user32.AddClipboardFormatListener(wintypes.HWND(self.hwnd)))
                if not self._clipboard_registered:
                    log_event("clipboard listener registration failed; polling fallback will be used")
            except Exception as exc:
                self._clipboard_registered = False
                log_error(exc)
            try:
                self.register_hotkeys()
            except Exception as exc:
                log_error(exc)
            self.enabled = True
            self._ready.set()
            msg = wintypes.MSG()
            while not self._stop_requested:
                ret = int(user32.GetMessageW(ctypes.byref(msg), None, 0, 0))
                if ret == -1:
                    log_event("GetMessageW failed in native event bridge")
                    break
                if ret == 0:
                    break
                user32.TranslateMessage(ctypes.byref(msg))
                user32.DispatchMessageW(ctypes.byref(msg))
        except BaseException as exc:
            self._start_error = exc if isinstance(exc, BaseException) else RuntimeError(str(exc))
            log_error(exc)
            self._ready.set()
        finally:
            self.enabled = False
            self._cleanup_native()
            self._ready.set()

    def _create_message_window(self) -> None:
        self._hinstance = int(kernel32.GetModuleHandleW(None) or 0)
        wndproc_type = ctypes.WINFUNCTYPE(LONG_PTR, wintypes.HWND, wintypes.UINT, wintypes.WPARAM, wintypes.LPARAM)

        def wndproc(hwnd, msg, wparam, lparam):
            try:
                imsg = int(msg)
                if imsg == WM_CLIPBOARDUPDATE:
                    self._push_event("clipboard", 0)
                    return 0
                if imsg == WM_HOTKEY:
                    self._push_event("hotkey", int(wparam))
                    return 0
                if imsg == WM_DESKLAYER_REFRESH_HOTKEYS:
                    self._register_hotkeys_now()
                    return 0
                if imsg == WM_CLOSE:
                    user32.DestroyWindow(hwnd)
                    return 0
                if imsg == WM_DESTROY:
                    self._cleanup_listener_and_hotkeys()
                    self.hwnd = 0
                    user32.PostQuitMessage(0)
                    return 0
            except Exception as exc:
                log_error(exc)
            return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

        self._wndproc = wndproc_type(wndproc)
        wc = WNDCLASSW()
        wc.style = 0
        wc.lpfnWndProc = ctypes.cast(self._wndproc, ctypes.c_void_p).value
        wc.cbClsExtra = 0
        wc.cbWndExtra = 0
        wc.hInstance = wintypes.HINSTANCE(self._hinstance)
        wc.hIcon = None
        wc.hCursor = None
        wc.hbrBackground = None
        wc.lpszMenuName = None
        wc.lpszClassName = self.class_name
        atom = user32.RegisterClassW(ctypes.byref(wc))
        if not atom:
            raise ctypes.WinError()
        self._class_registered = True
        hwnd = user32.CreateWindowExW(
            0,
            self.class_name,
            self.class_name,
            0,
            0,
            0,
            0,
            0,
            None,
            None,
            wintypes.HINSTANCE(self._hinstance),
            None,
        )
        if not hwnd:
            raise ctypes.WinError()
        self.hwnd = int(hwnd)

    def _cleanup_listener_and_hotkeys(self) -> None:
        if not IS_WINDOWS:
            return
        with self._hotkey_lock:
            if self.hwnd and self._clipboard_registered:
                try:
                    user32.RemoveClipboardFormatListener(wintypes.HWND(self.hwnd))
                except Exception:
                    pass
                self._clipboard_registered = False
            self._unregister_hotkeys_no_lock()

    def _cleanup_native(self) -> None:
        if not IS_WINDOWS:
            return
        try:
            self._cleanup_listener_and_hotkeys()
            if self.hwnd:
                try:
                    user32.DestroyWindow(wintypes.HWND(self.hwnd))
                except Exception:
                    pass
                self.hwnd = 0
        finally:
            if self._class_registered:
                try:
                    user32.UnregisterClassW(self.class_name, wintypes.HINSTANCE(self._hinstance))
                except Exception:
                    pass
                self._class_registered = False
            self._wndproc = None

    def register_hotkeys(self) -> None:
        if not IS_WINDOWS or not self.hwnd:
            return
        # RegisterHotKey is safest from the thread that owns the receiving HWND.
        if self._message_thread_ident and threading.get_ident() != self._message_thread_ident:
            try:
                user32.PostMessageW(wintypes.HWND(self.hwnd), WM_DESKLAYER_REFRESH_HOTKEYS, 0, 0)
            except Exception as exc:
                log_error(exc)
            return
        self._register_hotkeys_now()

    def _register_hotkeys_now(self) -> None:
        if not IS_WINDOWS or not self.hwnd:
            return
        with self._hotkey_lock:
            self._unregister_hotkeys_no_lock()
            if not bool(getattr(self.app, "global_hotkeys", False)):
                return
            for hotkey_id, mods, vk, label, action in HOTKEY_DEFS:
                try:
                    if user32.RegisterHotKey(wintypes.HWND(self.hwnd), int(hotkey_id), int(mods), int(vk)):
                        self.registered_hotkeys.append(int(hotkey_id))
                    else:
                        log_event(f"hotkey failed {label} {action}")
                except Exception as exc:
                    log_error(exc)

    def _unregister_hotkeys_no_lock(self) -> None:
        if not IS_WINDOWS or not self.hwnd:
            self.registered_hotkeys.clear()
            return
        for hotkey_id in list(self.registered_hotkeys):
            try:
                user32.UnregisterHotKey(wintypes.HWND(self.hwnd), int(hotkey_id))
            except Exception:
                pass
        self.registered_hotkeys.clear()

    def unregister_hotkeys(self) -> None:
        if IS_WINDOWS and self.hwnd and self._message_thread_ident and threading.get_ident() != self._message_thread_ident:
            try:
                user32.PostMessageW(wintypes.HWND(self.hwnd), WM_DESKLAYER_REFRESH_HOTKEYS, 0, 0)
            except Exception as exc:
                log_error(exc)
            return
        with self._hotkey_lock:
            self._unregister_hotkeys_no_lock()

    def stop(self) -> None:
        self._stop_requested = True
        if IS_WINDOWS and self.hwnd:
            try:
                user32.PostMessageW(wintypes.HWND(self.hwnd), WM_CLOSE, 0, 0)
            except Exception:
                pass
        if self._thread and self._thread.is_alive():
            try:
                self._thread.join(timeout=0.8)
            except RuntimeError:
                pass
        self.enabled = False


class DeskLayerApp:
    def __init__(self) -> None:
        self.config = load_config()
        self.compact_size = as_int(self.config.get("compact_size", COMPACT_DEFAULT_SIZE), COMPACT_DEFAULT_SIZE, COMPACT_MIN_SIZE, COMPACT_MAX_SIZE)
        self.safe_mode = "--safe" in sys.argv
        self._shutting_down = False
        self.root = tk.Tk()
        self.root.title(APP_NAME)
        self.root.report_callback_exception = self.report_callback_exception
        self.root.overrideredirect(True)
        try:
            self.root.minsize(self.compact_size, self.compact_size)
        except Exception:
            pass
        try:
            self.root.attributes("-topmost", as_bool(self.config.get("topmost", True), True))
        except Exception:
            pass
        try:
            self.root.attributes("-alpha", as_float(self.config.get("alpha", 0.97), 0.97, 0.35, 1.0))
        except Exception:
            pass

        self.colors = {
            "bg": "#030711",
            "panel": "#071225",
            "panel2": "#09182e",
            "line": "#16445f",
            "cyan": "#16f6ff",
            "mag": "#ff35d1",
            "violet": "#9a4dff",
            "text": "#ecfbff",
            "muted": "#8498ad",
            "warn": "#ff5d89",
            "green": "#60ffba",
            "dark": "#01040a",
            "slot_active": "#42164c",
        }
        self.clip = WinClipboard()
        self.store = TrayStore(
            max_items=as_int(self.config.get("max_items", 80), 80, 1, 500),
            max_total_bytes=as_int(self.config.get("max_total_bytes", 96 * 1024 * 1024), 96 * 1024 * 1024, 1_000_000, 1024 * 1024 * 1024),
        )
        self.max_text_chars = as_int(self.config.get("max_text_chars", 500_000), 500_000, 1_000, 5_000_000)
        self.max_image_bytes = as_int(self.config.get("max_image_bytes", 32 * 1024 * 1024), 32 * 1024 * 1024, 1_000_000, 256 * 1024 * 1024)
        self.main_render_limit = as_int(self.config.get("main_render_limit", 32), 32, 5, 120)
        self.auto_capture = as_bool(self.config.get("auto_capture", True), True)
        self.capture_when_compact = as_bool(self.config.get("capture_when_compact", False), False)
        self.privacy_guard = as_bool(self.config.get("privacy_guard", True), True)
        self.auto_skip_sensitive = as_bool(self.config.get("auto_skip_sensitive", True), True)
        self.large_confirm_bytes = as_int(self.config.get("large_confirm_bytes", 16 * 1024 * 1024), 16 * 1024 * 1024, 1_000_000, 512 * 1024 * 1024)
        self.confirm_all_operations = as_bool(self.config.get("confirm_all_operations", True), True)
        self.auto_clear_minutes = as_int(self.config.get("auto_clear_minutes", 0), 0, 0, 24 * 60)
        self.clipboard_priority = str(self.config.get("clipboard_priority", "text_first") or "text_first")
        if self.clipboard_priority not in ("text_first", "image_first"):
            self.clipboard_priority = "text_first"
        self.paste_advance = as_bool(self.config.get("paste_advance", True), True)
        self.confirm_delete = as_bool(self.config.get("confirm_delete", False), False)
        self.compact_badge_pinned = as_bool(self.config.get("compact_badge_pinned", True), True)
        self.global_hotkeys = as_bool(self.config.get("global_hotkeys", True), True)
        self.paste_delay_ms = as_int(self.config.get("paste_delay_ms", 100), 100, 60, 1500)
        self.confirm_exit = as_bool(self.config.get("confirm_exit", True), True)
        self.diagnostic_tail_lines = as_int(self.config.get("diagnostic_tail_lines", 80), 80, 20, 500)
        if self.safe_mode:
            self.auto_capture = False
            self.capture_when_compact = False
            self.global_hotkeys = False

        self.compact = True
        self.panel_open = False
        self.selected: set[str] = set()
        self.active_id: Optional[str] = None
        self.clipboard_item_ids: tuple[str, ...] = ()
        self.clipboard_label = ""
        self.last_seq = self.clip.sequence() if IS_WINDOWS else 0
        self.own_write_seq: Optional[int] = None
        self.last_target_hwnd: Optional[int] = None
        self.drag_offset: Optional[tuple[int, int]] = None
        self.compact_press_xy: Optional[tuple[int, int]] = None
        self.compact_dragged = False
        self._compact_hold_after_id: Optional[str] = None
        self._compact_hold_menu_opened = False
        self.picker: Optional[tk.Toplevel] = None
        self.toast_id = 0
        self._clip_event_pending = False
        self._refresh_pending = False
        self._refresh_header_only = False
        self._last_skip_msg = ""
        self._last_skip_time = 0.0
        self.compact_anchor: Optional[tuple[int, int]] = None
        self.last_tray_change = time.monotonic()
        self._pinned_load_queue = queue.Queue(maxsize=1)
        self._pinned_loaded = False
        self._pins_dirty = False
        self._discard_startup_pins = False
        self._undo_stack: list[tuple[str, list[TrayItem]]] = []
        self._undo_limit = 10

        self.status_var = tk.StringVar(value="待機中")
        self.active_slot_var = tk.StringVar(value="選択 --")
        self.ready_var = tk.StringVar(value="項目が選択されていません")
        self.clip_slot_var = tk.StringVar(value="貼付候補 --")
        self.clip_ready_var = tk.StringVar(value="セットするとCtrl+Vで貼り付けできます")

        self.load_position()
        self.build_compact()
        self.start_pinned_loader()
        self.listener = ClipboardUpdateListener(self)
        self.listener_ok = False
        # Paint the tiny launcher first.  Native listener/hotkey work starts just
        # after the first frame, so a bad clipboard owner cannot make startup look dead.
        self.root.after(10, self.start_runtime_services)
        self.root.after(3_000, self.ensure_visible_loop)
        self.root.after(1_000, self.stop_request_loop)
        self.root.after(60_000, self.housekeeping)
        self.root.protocol("WM_DELETE_WINDOW", lambda: self.stop_panel() if self.panel_open else self.request_exit_app())
        self.install_rescue_bindings()
        log_environment(self.root, safe_mode=self.safe_mode)
        log_event(f"start {VERSION} scheduled")

    # ----- rescue / exit hardening -----
    def install_rescue_bindings(self) -> None:
        """Make the app stoppable even when the tiny right-click target is unreliable."""
        def normal_exit(e=None):
            self.request_exit_app()
            return "break"

        def emergency_exit(e=None):
            self.emergency_exit_app("keyboard")
            return "break"

        def popup(e=None):
            self.app_menu(e)
            return "break"

        for seq in ("<Control-KeyPress-q>", "<Control-KeyPress-Q>"):
            try:
                self.root.bind_all(seq, normal_exit, add="+")
            except Exception:
                pass
        for seq in ("<Control-Shift-KeyPress-Q>", "<Control-Shift-KeyPress-q>"):
            try:
                self.root.bind_all(seq, emergency_exit, add="+")
            except Exception:
                pass
        for seq in ("<Alt-KeyPress-F4>", "<Alt-F4>"):
            try:
                self.root.bind_all(seq, normal_exit, add="+")
            except Exception:
                pass
        for seq in ("<Shift-F10>", "<KeyPress-Menu>", "<KeyPress-App>"):
            try:
                self.root.bind_all(seq, popup, add="+")
            except Exception:
                pass

    def bind_app_menu(self, widget: tk.Widget) -> None:
        for seq in ("<Button-3>", "<Button-2>", "<Control-Button-1>"):
            try:
                widget.bind(seq, self.app_menu, add="+")
            except Exception:
                pass

    def event_screen_xy(self, e: Optional[tk.Event] = None) -> tuple[int, int]:
        try:
            if e is not None and hasattr(e, "x_root") and hasattr(e, "y_root"):
                return int(e.x_root), int(e.y_root)
        except Exception:
            pass
        try:
            return int(self.root.winfo_rootx() + max(10, self.root.winfo_width() // 2)), int(self.root.winfo_rooty() + max(10, self.root.winfo_height() // 2))
        except Exception:
            return 100, 100

    def popup_menu(self, menu: tk.Menu, e: Optional[tk.Event] = None, x: Optional[int] = None, y: Optional[int] = None):
        """Show a Tk popup menu and always release Tk's implicit grab.

        Tk popup handling can be touchy with overrideredirect/topmost windows; this
        helper prevents a failed popup from leaving the UI in a grabbed-looking state.
        """
        try:
            if x is None or y is None:
                x, y = self.event_screen_xy(e)
            self.root.update_idletasks()
            menu.tk_popup(int(x), int(y))
            return "break"
        except Exception as exc:
            log_error(exc)
            try:
                self.toast("MENU FAILED / Ctrl+Qで終了")
            except Exception:
                pass
            return "break"
        finally:
            try:
                menu.grab_release()
            except Exception:
                pass

    def app_menu(self, e: Optional[tk.Event] = None):
        m = tk.Menu(self.root, tearoff=0)
        if self.compact:
            m.add_command(label="開く", command=self.start_active)
        else:
            m.add_command(label="格納", command=self.stop_panel)
        m.add_command(label="一覧を開く", command=self.open_picker_from_hotkey)
        m.add_command(label="クリップボードから追加", command=lambda: self.add_from_clipboard(False))
        m.add_command(label="メモを追加", command=self.add_manual_text)
        m.add_separator()
        m.add_command(label="自動取り込み：ON" if self.auto_capture else "自動取り込み：停止中", command=self.toggle_capture)
        m.add_command(label="格納中の裏取り込み：ON" if self.capture_when_compact else "格納中の裏取り込み：OFF", command=self.toggle_background_capture)
        m.add_command(label="全体ホットキー：ON" if self.global_hotkeys else "全体ホットキー：OFF", command=self.toggle_global_hotkeys)
        m.add_separator()
        m.add_command(label="診断レポートを作成", command=self.open_diagnostic_report)
        m.add_command(label="設定・ログフォルダを開く", command=self.open_log_folder)
        m.add_command(label="待機ボタン位置を右上へ戻す", command=self.reset_window_position)
        m.add_command(label=f"待機ボタンサイズ：{self.compact_size}px", command=self.cycle_compact_size)
        m.add_command(label="使い方 / 停止方法", command=self.show_quick_help)
        m.add_separator()
        m.add_command(label="終了（確認あり）  Ctrl+Q", command=self.request_exit_app)
        m.add_command(label="緊急終了（確認なし）  Ctrl+Shift+Q / Win+Alt+Q", command=lambda: self.emergency_exit_app("app-menu"))
        return self.popup_menu(m, e)

    def stop_request_loop(self) -> None:
        if getattr(self, "_shutting_down", False):
            return
        try:
            if os.path.exists(STOP_REQUEST_PATH):
                try:
                    os.remove(STOP_REQUEST_PATH)
                except Exception:
                    pass
                self.emergency_exit_app("stop-request-file")
                return
        except Exception as exc:
            log_error(exc)
        finally:
            if not getattr(self, "_shutting_down", False):
                self.root.after(1_000, self.stop_request_loop)

    def cancel_compact_hold_timer(self) -> None:
        aid = getattr(self, "_compact_hold_after_id", None)
        if aid:
            try:
                self.root.after_cancel(aid)
            except Exception:
                pass
        self._compact_hold_after_id = None

    def compact_long_press_menu(self) -> None:
        self._compact_hold_after_id = None
        try:
            if self.compact and self.compact_press_xy and not self.compact_dragged:
                x, y = self.compact_press_xy
                self._compact_hold_menu_opened = True
                self.drag_offset = None
                self.compact_menu(None, x=x, y=y)
        except Exception as exc:
            log_error(exc)

    def start_runtime_services(self) -> None:
        if getattr(self, "_shutting_down", False):
            return
        try:
            if not self.safe_mode:
                self.listener_ok = self.listener.start()
            else:
                self.listener_ok = False
            if not self.listener_ok:
                self.root.after(1200, self.poll_clipboard_fallback)
            self.root.after(1200, self.track_foreground)
            log_event(f"runtime listener={'on' if self.listener_ok else 'fallback'} hotkeys={'on' if self.global_hotkeys else 'off'} safe={int(self.safe_mode)}")
        except Exception as exc:
            log_error(exc)
            if not getattr(self, "_shutting_down", False):
                self.root.after(1200, self.poll_clipboard_fallback)
                self.root.after(1200, self.track_foreground)

    def report_callback_exception(self, exc_type, exc_value, exc_tb) -> None:
        """Tk callback exceptions are invisible under .pyw; log them instead of failing silently."""
        try:
            rotate_log(ERROR_LOG)
            with open(ERROR_LOG, "a", encoding="utf-8") as f:
                f.write("\n" + "=" * 80 + "\n")
                f.write(time.strftime("%Y-%m-%d %H:%M:%S") + "  Tk callback error\n")
                traceback.print_exception(exc_type, exc_value, exc_tb, file=f)
        except Exception:
            pass
        try:
            self.toast("ERROR / error.log を確認")
        except Exception:
            pass

    def start_pinned_loader(self) -> None:
        """Load pinned items after the tiny launcher is already visible."""
        def worker() -> None:
            err: Optional[BaseException] = None
            items: list[TrayItem] = []
            try:
                items = load_pinned_items()
            except BaseException as exc:  # keep startup alive even if pinned JSON is corrupt/huge
                err = exc
                log_error(exc)
            try:
                self._pinned_load_queue.put_nowait((items, err))
            except Exception:
                pass

        try:
            threading.Thread(target=worker, name="DeskLayerPinnedLoader", daemon=True).start()
            self.root.after(80, self.finish_pinned_loader)
        except Exception as exc:
            log_error(exc)
            self._pinned_loaded = False

    def finish_pinned_loader(self) -> None:
        if getattr(self, "_shutting_down", False):
            return
        try:
            items, err = self._pinned_load_queue.get_nowait()
        except queue.Empty:
            self.root.after(80, self.finish_pinned_loader)
            return
        except Exception as exc:
            log_error(exc)
            return

        if err is not None:
            self._pinned_loaded = False
            self.toast("PIN LOAD ERROR / 既存ピン保存は保留")
            return

        self._pinned_loaded = True
        if self._discard_startup_pins:
            items = []

        added = 0
        try:
            # Keep original startup order while avoiding a blank/slow first paint.
            for pinned in reversed(items):
                ok, _reason, _resolved = self.store.add(pinned, bump_duplicate=False)
                if ok:
                    added += 1
            if self.store.items and not self.active_id:
                self.active_id = self.store.items[0].id
                self.selected = {self.active_id}
            self.clear_markers_if_missing()
            self.request_refresh()
            if added:
                log_event(f"pinned loaded {added}")
            if self._pins_dirty:
                self.save_pinned_safe()
        except Exception as exc:
            self._pinned_loaded = False
            log_error(exc)

    def save_pinned_safe(self) -> None:
        if getattr(self, "_pinned_loaded", False):
            save_pinned_items(self.store.items)
            self._pins_dirty = False
        else:
            self._pins_dirty = True
            log_event("pin save deferred until startup pinned load finishes")

    # ----- position / config -----
    def load_position(self) -> None:
        sx, sy, sw, sh = virtual_screen_bounds(self.root)
        default_x = sx + max(20, sw - self.compact_size - 24)
        default_y = sy + max(20, min(170, sh - self.compact_size - 24))
        x = self.config.get("x")
        y = self.config.get("y")
        self.x = default_x if x is None else as_int(x, default_x)
        self.y = default_y if y is None else as_int(y, default_y)
        self.x, self.y = self.clamp_xy(self.x, self.y, self.compact_size, self.compact_size, margin=8)

    def persist_config(self) -> None:
        self.config.update({
            "x": int(self.x),
            "y": int(self.y),
            "max_items": int(self.store.max_items),
            "max_total_bytes": int(self.store.max_total_bytes),
            "max_text_chars": int(self.max_text_chars),
            "max_image_bytes": int(self.max_image_bytes),
            "main_render_limit": int(self.main_render_limit),
            "auto_capture": bool(self.auto_capture),
            "capture_when_compact": bool(self.capture_when_compact),
            "privacy_guard": bool(self.privacy_guard),
            "auto_skip_sensitive": bool(self.auto_skip_sensitive),
            "large_confirm_bytes": int(self.large_confirm_bytes),
            "confirm_all_operations": bool(self.confirm_all_operations),
            "auto_clear_minutes": int(self.auto_clear_minutes),
            "clipboard_priority": str(self.clipboard_priority),
            "paste_advance": bool(self.paste_advance),
            "confirm_delete": bool(self.confirm_delete),
            "compact_badge_pinned": bool(self.compact_badge_pinned),
            "compact_size": int(self.compact_size),
            "global_hotkeys": bool(self.global_hotkeys),
            "paste_delay_ms": int(self.paste_delay_ms),
            "confirm_exit": bool(self.confirm_exit),
            "diagnostic_tail_lines": int(self.diagnostic_tail_lines),
        })
        save_config(self.config)

    def clamp_xy(self, x: int, y: int, w: int, h: int, margin: int = 8) -> tuple[int, int]:
        sx, sy, sw, sh = virtual_screen_bounds(self.root)
        min_x = sx + margin
        min_y = sy + margin
        max_x = sx + max(margin, sw - w - margin)
        max_y = sy + max(margin, sh - h - margin)
        return max(min_x, min(int(x), max_x)), max(min_y, min(int(y), max_y))

    def set_root_geometry(self, w: int, h: int, x: int, y: int) -> None:
        margin = 8
        x, y = self.clamp_xy(int(x), int(y), int(w), int(h), margin=margin)
        move_tk_window(self.root, w, h, x, y)

    def active_geometry_from_compact(self) -> tuple[int, int, int, int]:
        _sx, _sy, sw, sh = virtual_screen_bounds(self.root)
        # Keep enough logical pixels for Japanese labels and high-DPI fonts.
        # On very small screens, fit to the visible area instead of forcing off-screen.
        min_w = 640 if sw >= 700 else max(360, sw - 24)
        min_h = 560 if sh >= 640 else max(360, sh - 72)
        w = min(760, max(min_w, sw - 24))
        h = min(820, max(min_h, sh - 72))
        desired_x = self.root.winfo_x() + self.compact_size - w
        desired_y = self.root.winfo_y()
        x, y = self.clamp_xy(desired_x, desired_y, w, h, margin=8)
        return w, h, x, y

    def picker_geometry(self, w: int = 940, h: int = 680) -> str:
        _sx, _sy, sw, sh = virtual_screen_bounds(self.root)
        w = min(w, max(760 if sw >= 820 else 360, sw - 24))
        h = min(h, max(500 if sh >= 580 else 360, sh - 72))
        desired_x = self.root.winfo_x() - 260
        desired_y = self.root.winfo_y() + 35
        x, y = self.clamp_xy(desired_x, desired_y, w, h, margin=8)
        return tk_geometry_spec(w, h, x, y)

    def ensure_visible_loop(self) -> None:
        if getattr(self, "_shutting_down", False):
            return
        try:
            self.ensure_visible()
        finally:
            if not getattr(self, "_shutting_down", False):
                self.root.after(15_000, self.ensure_visible_loop)

    def ensure_visible(self) -> None:
        if self.drag_offset:
            return
        try:
            w = max(self.compact_size if self.compact else 580, int(self.root.winfo_width() or (self.compact_size if self.compact else 580)))
            h = max(self.compact_size if self.compact else 540, int(self.root.winfo_height() or (self.compact_size if self.compact else 540)))
            margin = 8
            x = int(self.root.winfo_x())
            y = int(self.root.winfo_y())
            nx, ny = self.clamp_xy(x, y, w, h, margin=margin)
            if (nx, ny) != (x, y):
                move_tk_window(self.root, w, h, nx, ny)
                if self.compact:
                    self.x, self.y = nx, ny
            try:
                self.root.attributes("-topmost", as_bool(self.config.get("topmost", True), True))
            except Exception:
                pass
        except Exception as exc:
            log_error(exc)

    # ----- clipboard event pipeline -----
    def schedule_clipboard_event(self) -> None:
        if getattr(self, "_shutting_down", False) or self._clip_event_pending:
            return
        self._clip_event_pending = True
        self.root.after(120, self.handle_clipboard_event)

    def poll_clipboard_fallback(self) -> None:
        if getattr(self, "_shutting_down", False):
            return
        try:
            self.schedule_clipboard_event()
        finally:
            if not getattr(self, "_shutting_down", False):
                self.root.after(1200 if self.compact else 800, self.poll_clipboard_fallback)

    def handle_clipboard_event(self) -> None:
        self._clip_event_pending = False
        if not IS_WINDOWS:
            return
        try:
            seq = self.clip.sequence()
            if seq == self.last_seq:
                return
            self.last_seq = seq
            if self.own_write_seq is not None and seq == self.own_write_seq:
                self.own_write_seq = None
                return
            if self.should_capture_now():
                self.add_from_clipboard(auto=True)
            else:
                # External clipboard changed while paused/idle. Do not read heavy formats;
                # just clear the HUD so it never lies about what Ctrl+V will paste.
                self.clipboard_item_ids = ()
                self.clipboard_label = ""
                self.request_refresh(header_only=True)
        except ClipboardBusy:
            return
        except Exception as exc:
            log_error(exc)

    def should_capture_now(self) -> bool:
        if not self.auto_capture:
            return False
        if self.panel_open:
            return True
        return bool(self.capture_when_compact)

    def should_skip_for_privacy(self) -> bool:
        if not IS_WINDOWS or not self.privacy_guard:
            return False
        hwnd = int(user32.GetForegroundWindow() or 0) if IS_WINDOWS else 0
        if not valid_external_window(hwnd):
            hwnd = self.last_target_hwnd or 0
        title = get_window_title(hwnd).lower()
        proc = get_window_process_name(hwnd).lower()
        processes = [str(x).lower() for x in self.config.get("privacy_exclude_processes", []) if str(x).strip()]
        titles = [str(x).lower() for x in self.config.get("privacy_exclude_titles", []) if str(x).strip()]
        for kw in processes:
            if kw and kw in proc:
                return True
        for kw in titles:
            if kw and kw in title:
                return True
        return False

    def mark_tray_changed(self) -> None:
        self.last_tray_change = time.monotonic()

    def housekeeping(self) -> None:
        try:
            minutes = int(self.auto_clear_minutes or 0)
            if minutes > 0 and self.store.items:
                age = time.monotonic() - self.last_tray_change
                if age >= minutes * 60:
                    removed = self.store.clear_unpinned()
                    if removed:
                        self.selected = {self.store.items[0].id} if self.store.items else set()
                        self.active_id = self.store.items[0].id if self.store.items else None
                        self.clipboard_item_ids = tuple(i for i in self.clipboard_item_ids if self.store.item_by_id(i))
                        self.request_refresh()
                        self.toast(f"AUTO CLEARED {removed} SESSION ITEM(S)")
                    self.mark_tray_changed()
        except Exception as exc:
            log_error(exc)
        finally:
            if not getattr(self, "_shutting_down", False):
                self.root.after(60_000, self.housekeeping)

    # ----- foreground tracking / paste helpers -----
    def update_last_target_from_foreground(self) -> Optional[int]:
        if not IS_WINDOWS:
            return self.last_target_hwnd
        try:
            hwnd = int(user32.GetForegroundWindow() or 0)
            if hwnd and valid_external_window(hwnd):
                self.last_target_hwnd = hwnd
                return hwnd
        except Exception:
            pass
        return self.last_target_hwnd

    def track_foreground(self) -> None:
        if getattr(self, "_shutting_down", False):
            return
        self.update_last_target_from_foreground()
        if not getattr(self, "_shutting_down", False):
            self.root.after(1200, self.track_foreground)

    # ----- item selection -----
    def slot_index(self, item_or_id: object) -> Optional[int]:
        target = item_or_id.id if isinstance(item_or_id, TrayItem) else item_or_id
        for i, it in enumerate(self.store.items, start=1):
            if it.id == target:
                return i
        return None

    def active_item(self) -> Optional[TrayItem]:
        it = self.store.item_by_id(self.active_id)
        if it:
            return it
        if self.store.items:
            self.active_id = self.store.items[0].id
            return self.store.items[0]
        self.active_id = None
        return None

    def active_items(self) -> list[TrayItem]:
        it = self.active_item()
        return [it] if it else []

    def set_active_item(self, it: TrayItem, refresh: bool = True) -> None:
        self.active_id = it.id
        self.selected = {it.id}
        if refresh:
            self.request_refresh()
        self.toast(f"ACTIVE SLOT [{self.slot_index(it) or '?'}]")

    def set_active_by_slot(self, slot: int) -> bool:
        if 1 <= slot <= len(self.store.items):
            self.set_active_item(self.store.items[slot - 1])
            return True
        self.toast(f"NO SLOT [{slot}]")
        return False

    def items_for_ids(self, ids: Sequence[str]) -> list[TrayItem]:
        idset = set(ids)
        return [it for it in self.store.items if it.id in idset]

    def selected_items_or_all(self) -> list[TrayItem]:
        return [it for it in self.store.items if it.id in self.selected] if self.selected else list(self.store.items)

    # ----- read/write operations -----
    def add_from_clipboard(self, auto: bool = False, mode: str = "auto") -> None:
        if not IS_WINDOWS:
            self.toast("Windowsで実行してください")
            return
        try:
            if auto and self.should_skip_for_privacy():
                self.clipboard_item_ids = ()
                self.clipboard_label = ""
                self.request_refresh(header_only=True)
                self.throttled_toast("PRIVACY GUARD SKIP", auto=True)
                return
            mode = str(mode or "auto").lower().strip()
            priority = self.clipboard_priority
            if mode in ("text", "text_only"):
                priority = "text_only"
            elif mode in ("image", "image_only"):
                priority = "image_only"
            items = self.clip.snapshot_items(
                timeout_ms=90 if auto else 400,
                max_text_chars=self.max_text_chars,
                max_image_bytes=self.max_image_bytes,
                include_images=True,
                priority=priority,
            )
            if auto and self.auto_skip_sensitive and items and all(it.kind == "text" and it.sensitive for it in items):
                self.clipboard_item_ids = ()
                self.clipboard_label = ""
                self.request_refresh(header_only=True)
                self.throttled_toast("SENSITIVE TEXT SKIPPED", auto=True)
                return
            if not items:
                if not auto:
                    self.toast("CLIPBOARD EMPTY / UNSUPPORTED")
                self.clipboard_item_ids = ()
                self.clipboard_label = ""
                self.request_refresh(header_only=True)
                return
            added, dup, resolved = self.store.add_many_preserve_order(items)
            ids = [it.id for it in resolved]
            if ids:
                self.clipboard_item_ids = tuple(ids)
                self.clipboard_label = "CAPTURED"
                self.active_id = ids[0]
                self.selected = {ids[0]}
            self.clear_markers_if_missing()
            self.mark_tray_changed()
            self.request_refresh()
            if not auto or self.panel_open:
                if added:
                    self.toast(f"ADDED {added}" + (f" / DUP {dup}" if dup else ""))
                elif dup:
                    self.toast("ALREADY IN TRAY / MARKED")
        except ClipboardBusy:
            if not auto:
                self.toast("CLIPBOARD BUSY")
        except ClipboardTooLarge as exc:
            self.clipboard_item_ids = ()
            self.clipboard_label = ""
            self.request_refresh(header_only=True)
            self.throttled_toast(str(exc), auto=auto)
        except Exception as exc:
            if not auto:
                self.toast("ADD FAILED")
            log_error(exc)

    def throttled_toast(self, msg: str, auto: bool) -> None:
        now = time.monotonic()
        if auto and msg == self._last_skip_msg and now - self._last_skip_time < 6.0:
            return
        self._last_skip_msg = msg
        self._last_skip_time = now
        if self.panel_open or not auto:
            self.toast(msg)

    def mark_clipboard_items(self, items: Sequence[TrayItem], label: str = "") -> None:
        self.clipboard_item_ids = tuple(it.id for it in items)
        self.clipboard_label = label
        self.request_refresh(header_only=True)

    def copy_active(self) -> None:
        items = self.active_items()
        if not items:
            self.toast("NO ACTIVE SLOT")
            return
        self.copy_items(items, label=f"SLOT [{self.slot_index(items[0]) or '?'}]")

    def paste_active(self) -> None:
        items = self.active_items()
        if not items:
            self.toast("NO ACTIVE SLOT")
            return
        self.paste_items(items, label=f"SLOT [{self.slot_index(items[0]) or '?'}]")

    def confirm_clipboard_operation(self, items: Sequence[TrayItem], text: str, paths: Sequence[str], dib: bytes, image_count: int, label: str = "") -> bool:
        if not items:
            return False
        total = len(text.encode("utf-16le", "replace")) + sum(len(p.encode("utf-16le", "replace")) + 64 for p in paths) + len(dib or b"")
        needs_confirm = total > self.large_confirm_bytes or len(items) > 20 or image_count > 1
        if self.confirm_all_operations and label.upper() in ("ALL", "PICK ALL") and len(items) > 1:
            needs_confirm = True
        if not needs_confirm:
            return True
        msg = f"{len(items)}件 / 約{human_bytes(total)} をWindowsクリップボードへセットします。"
        if image_count > 1:
            msg += "\n\n画像はWindowsクリップボード仕様上、先頭1枚だけが貼付対象になります。"
        msg += "\n\n続行しますか？"
        try:
            return bool(messagebox.askyesno(APP_NAME, msg))
        except Exception:
            return True

    def copy_items(self, items: Sequence[TrayItem], label: str = "") -> bool:
        items = list(items)
        if not items:
            self.toast("TRAY EMPTY")
            return False
        text, paths, dib, dib_format, image_count = collect_payload(items)
        if not self.confirm_clipboard_operation(items, text, paths, dib, image_count, label=label):
            return False
        marked_items: list[TrayItem] = []
        image_marked = False
        for it in items:
            if it.kind == "image":
                if image_marked:
                    continue
                image_marked = True
            marked_items.append(it)
        try:
            self.clip.write(text=text, paths=paths, dib=dib, dib_format=dib_format, backup=True)
            self.own_write_seq = self.clip.sequence() if IS_WINDOWS else None
            if self.own_write_seq is not None:
                self.last_seq = self.own_write_seq
            # Windows clipboard can represent many text/file entries, but only one normal bitmap.
            # Keep the HUD honest when multiple image items were selected.
            self.mark_clipboard_items(marked_items, label=label)
            extra = ""
            if image_count > 1:
                extra = " / first image only"
            elif image_count == 1 and not text and not paths:
                extra = " / image"
            prefix = f"{label} " if label else ""
            self.toast(f"CLIPBOARD READY {prefix}{len(marked_items)} item(s){extra}")
            return True
        except ClipboardBusy:
            self.toast("CLIPBOARD BUSY")
        except Exception as exc:
            self.toast("COPY FAILED / RESTORED IF POSSIBLE")
            log_error(exc)
        return False

    def paste_items(self, items: Sequence[TrayItem], label: str = "") -> bool:
        items = list(items)
        if not items:
            self.toast("TRAY EMPTY")
            return False
        if self.copy_items(items, label=label):
            self.root.after(max(60, int(self.paste_delay_ms)), self._paste_after_copy)
            return True
        return False

    def _paste_after_copy(self) -> None:
        if send_ctrl_v_to(self.last_target_hwnd):
            self.toast("PASTED")
        else:
            self.toast("クリップボードへセット済み。手動Ctrl+V。")

    def capture_selection_now(self) -> None:
        if not IS_WINDOWS:
            self.toast("Windowsで実行してください")
            return
        before = self.clip.sequence()
        if not send_ctrl_c_to(self.last_target_hwnd):
            self.toast("対象に戻れません。手動Ctrl+C後に『追加』を押してください。")
            return
        self.toast("CAPTURING…")
        self.root.after(230, lambda: self.wait_clip(before, 0))

    def wait_clip(self, before: int, attempt: int) -> None:
        seq = self.clip.sequence() if IS_WINDOWS else before
        if seq != before:
            self.last_seq = seq
            self.add_from_clipboard(auto=False)
            return
        if attempt >= 3:
            self.toast("選択コピーを検出できません。手動Ctrl+C後にADD。")
            return
        self.root.after([220, 360, 620][attempt], lambda: self.wait_clip(before, attempt + 1))

    def wipe_clipboard(self) -> None:
        if not IS_WINDOWS:
            self.toast("Windowsで実行してください")
            return
        if not messagebox.askyesno(APP_NAME, "Windowsクリップボードを空にしますか？\n\n仮想トレーの候補は残ります。"):
            return
        try:
            self.clip.clear()
            self.own_write_seq = self.clip.sequence()
            self.last_seq = self.own_write_seq
            self.clipboard_item_ids = ()
            self.clipboard_label = ""
            self.request_refresh()
            self.toast("WINDOWS CLIPBOARD WIPED")
        except Exception as exc:
            self.toast("WIPE FAILED")
            log_error(exc)

    def wipe_all(self) -> None:
        if not messagebox.askyesno(APP_NAME, "仮想トレーとWindowsクリップボードを両方空にしますか？"):
            return
        try:
            self.remember_undo("WIPE ALL", list(self.store.items))
            self.store.clear()
            if not self._pinned_loaded:
                self._discard_startup_pins = True
            self.save_pinned_safe()
            self.selected.clear()
            self.active_id = None
            self.clipboard_item_ids = ()
            self.clipboard_label = ""
            if IS_WINDOWS:
                self.clip.clear()
                self.own_write_seq = self.clip.sequence()
                self.last_seq = self.own_write_seq
            self.request_refresh()
            self.toast("ALL WIPED")
        except Exception as exc:
            self.toast("WIPE FAILED")
            log_error(exc)

    def clear_session(self) -> None:
        if not self.store.items:
            self.toast("TRAY EMPTY")
            return
        removed_items = [it for it in self.store.items if not it.pinned]
        self.remember_undo("SESSION CLEAR", removed_items)
        removed = self.store.clear_unpinned()
        self.selected = {self.store.items[0].id} if self.store.items else set()
        self.active_id = self.store.items[0].id if self.store.items else None
        self.clipboard_item_ids = tuple(i for i in self.clipboard_item_ids if self.store.item_by_id(i))
        self.clipboard_label = "" if not self.clipboard_item_ids else self.clipboard_label
        self.mark_tray_changed()
        self.request_refresh()
        self.toast(f"SESSION CLEARED {removed}" if removed else "PINNED ONLY")

    def clear_all(self) -> None:
        if not self.store.items:
            self.toast("TRAY EMPTY")
            return
        if messagebox.askyesno(APP_NAME, "ピン留めも含めて仮想トレーを全消去しますか？\n\nWindowsクリップボードは変更しません。"):
            self.remember_undo("CLEAR ALL", list(self.store.items))
            self.store.clear()
            if not self._pinned_loaded:
                self._discard_startup_pins = True
            self.selected.clear()
            self.active_id = None
            self.clipboard_item_ids = ()
            self.clipboard_label = ""
            self.save_pinned_safe()
            self.mark_tray_changed()
            self.request_refresh()
            self.toast("ALL TRAY ITEMS CLEARED")

    def delete_ids(self, ids: set[str]) -> None:
        ids = {i for i in ids if i}
        if not ids:
            return
        if self.confirm_delete and not messagebox.askyesno(APP_NAME, f"選択した {len(ids)} 件を削除しますか？"):
            return
        removed_items = [it for it in self.store.items if it.id in ids]
        self.remember_undo("DELETE", removed_items)
        if ids.intersection(set(self.clipboard_item_ids)):
            self.clipboard_item_ids = ()
            self.clipboard_label = ""
        n = self.store.remove_ids(ids)
        self.selected.difference_update(ids)
        if self.active_id in ids:
            self.active_id = self.store.items[0].id if self.store.items else None
            self.selected = {self.active_id} if self.active_id else set()
        self.clear_markers_if_missing()
        self.save_pinned_safe()
        self.mark_tray_changed()
        self.request_refresh()
        self.toast(f"REMOVED {n}")

    def clear_markers_if_missing(self) -> None:
        alive = {it.id for it in self.store.items}
        if self.active_id and self.active_id not in alive:
            self.active_id = self.store.items[0].id if self.store.items else None
        if any(i not in alive for i in self.clipboard_item_ids):
            self.clipboard_item_ids = ()
            self.clipboard_label = ""

    def remember_undo(self, label: str, items: Sequence[TrayItem]) -> None:
        saved = [it for it in items if it is not None]
        if not saved:
            return
        self._undo_stack.append((str(label or "UNDO"), list(saved)))
        if len(self._undo_stack) > self._undo_limit:
            del self._undo_stack[:len(self._undo_stack) - self._undo_limit]

    def restore_last_removed(self) -> None:
        if not self._undo_stack:
            self.toast("UNDO EMPTY")
            return
        label, items = self._undo_stack.pop()
        try:
            added, dup, resolved = self.store.add_many_preserve_order(items)
            if resolved:
                self.active_id = resolved[0].id
                self.selected = {resolved[0].id}
            self.save_pinned_safe()
            self.mark_tray_changed()
            self.request_refresh()
            self.toast(f"UNDO {label}: RESTORED {added}" + (f" / DUP {dup}" if dup else ""))
        except Exception as exc:
            log_error(exc)
            self.toast("UNDO FAILED")

    # ----- UI build / refresh -----
    def clear_root(self) -> None:
        for child in self.root.winfo_children():
            child.destroy()

    def request_refresh(self, header_only: bool = False) -> None:
        if getattr(self, "_shutting_down", False):
            return
        if self._refresh_pending:
            # Do not let a pending header-only update swallow a later full redraw.
            if not header_only:
                self._refresh_header_only = False
            return
        self._refresh_pending = True
        self._refresh_header_only = bool(header_only)
        self.root.after_idle(self._do_refresh)

    def _do_refresh(self) -> None:
        header_only = bool(self._refresh_header_only)
        self._refresh_pending = False
        self._refresh_header_only = False
        try:
            if self.compact:
                self.update_compact_count()
            elif header_only:
                self.update_header()
                self.refresh_slot_strip()
            else:
                self.refresh_cards()
        except Exception as exc:
            log_error(exc)

    def build_compact(self) -> None:
        self.compact = True
        self.panel_open = False
        self.root.unbind("<KeyPress>")
        self.root.unbind("<MouseWheel>")
        self.root.unbind("<Button-4>")
        self.root.unbind("<Button-5>")
        self.root.bind_all("<Alt-ButtonPress-1>", self.begin_drag)
        self.root.bind_all("<Alt-B1-Motion>", self.on_drag)
        self.root.bind_all("<Alt-ButtonRelease-1>", self.end_drag)
        self.clear_root()
        c = self.colors
        self.root.configure(bg=c["dark"])
        self.x, self.y = self.clamp_xy(self.x, self.y, self.compact_size, self.compact_size, margin=8)
        self.set_root_geometry(self.compact_size, self.compact_size, self.x, self.y)
        try:
            self.root.deiconify()
            self.root.lift()
        except Exception:
            pass
        frame = tk.Frame(self.root, bg=c["dark"], highlightthickness=1, highlightbackground=c["cyan"], cursor="fleur")
        frame.pack(fill="both", expand=True, padx=1, pady=1)
        label_text = self.compact_badge_text()
        self.compact_label = tk.Label(
            frame,
            text=label_text,
            fg=c["cyan"],
            bg=c["dark"],
            font=("Meiryo UI", 10, "bold"),
            justify="center",
            cursor="fleur",
        )
        self.compact_label.pack(fill="both", expand=True)
        for w in (frame, self.compact_label):
            w.bind("<ButtonPress-1>", self.compact_press)
            w.bind("<B1-Motion>", self.compact_motion)
            w.bind("<ButtonRelease-1>", self.compact_release)
            w.bind("<Button-3>", self.compact_menu)
            w.bind("<Button-2>", self.compact_menu)
            w.bind("<Control-Button-1>", self.compact_menu)
            w.bind("<Shift-Button-1>", self.compact_menu)
        self.status_var.set("待機中")

    def compact_badge_text(self) -> str:
        count = len(self.store.items)
        if count <= 0:
            count_text = "空"
        else:
            mark = "★" if self.compact_badge_pinned and any(it.pinned for it in self.store.items) else ""
            count_text = f"{mark}{min(count, 99)}件"
        return f"待機中\n{count_text}\nクリックで開く"

    def update_compact_count(self) -> None:
        if hasattr(self, "compact_label"):
            self.compact_label.configure(text=self.compact_badge_text())

    def start_active(self) -> None:
        self.compact = False
        self.panel_open = True
        self.compact_anchor = (self.root.winfo_x(), self.root.winfo_y())
        w, h, px, py = self.active_geometry_from_compact()
        self.set_root_geometry(w, h, px, py)
        if self.store.items and not self.active_item():
            self.active_id = self.store.items[0].id
            self.selected = {self.active_id}
        self.build_active()
        try:
            self.root.lift()
        except Exception:
            pass
        self.root.after(80, self.root.focus_force)
        self.toast("CAPTURE READY" if self.auto_capture else "CAPTURE PAUSED")

    def stop_panel(self) -> None:
        self.panel_open = False
        self.compact = True
        self.x = self.root.winfo_x() + max(self.compact_size, self.root.winfo_width()) - self.compact_size
        self.y = self.root.winfo_y()
        try:
            if self.picker and self.picker.winfo_exists():
                self.picker.destroy()
        except Exception:
            pass
        self.picker = None
        self.persist_config()
        self.build_compact()

    def build_active(self) -> None:
        self.clear_root()
        c = self.colors
        self.root.configure(bg=c["bg"])
        self.root.bind("<KeyPress>", self.on_main_key)
        self.root.bind("<MouseWheel>", self.on_canvas_mousewheel)
        self.root.bind("<Button-4>", self.on_canvas_mousewheel)
        self.root.bind("<Button-5>", self.on_canvas_mousewheel)
        self.root.bind_all("<Alt-ButtonPress-1>", self.begin_drag)
        self.root.bind_all("<Alt-B1-Motion>", self.on_drag)
        self.root.bind_all("<Alt-ButtonRelease-1>", self.end_drag)

        outer = tk.Frame(self.root, bg=c["bg"], highlightthickness=1, highlightbackground=c["cyan"])
        outer.pack(fill="both", expand=True, padx=3, pady=3)
        self.bind_drag(outer)
        self.bind_app_menu(outer)

        header = tk.Frame(outer, bg=c["panel"], highlightthickness=1, highlightbackground=c["cyan"])
        header.pack(fill="x", padx=6, pady=(6, 0))
        self.bind_drag(header)
        self.bind_app_menu(header)

        drag_handle = tk.Label(
            header,
            text="ここをドラッグして移動できます　/　右クリックでメニュー",
            fg=c["text"],
            bg="#051628",
            font=("Meiryo UI", 8, "bold"),
            anchor="w",
            padx=10,
            cursor="fleur",
        )
        drag_handle.pack(fill="x", padx=6, pady=(6, 0), ipady=5)
        self.bind_drag(drag_handle)
        self.bind_app_menu(drag_handle)

        def hover_label(lbl: tk.Label, normal_bg: str = "#081426", hover_bg: str = "#142b46") -> None:
            lbl.bind("<Enter>", lambda e, w=lbl, bg=hover_bg: w.configure(bg=bg))
            lbl.bind("<Leave>", lambda e, w=lbl, bg=normal_bg: w.configure(bg=bg))

        def header_btn(parent: tk.Misc, text: str, cmd, color: str, bg: str = "#081426", width: Optional[int] = None) -> tk.Label:
            lbl = tk.Label(parent, text=text, fg=color, bg=bg, font=("Segoe UI", 8, "bold"), padx=8, pady=4,
                           highlightthickness=1, highlightbackground=color, cursor="hand2")
            lbl.pack(side="left", padx=3, pady=2)
            if width:
                lbl.configure(width=width)
            lbl.bind("<Button-1>", lambda e: cmd())
            hover_label(lbl, normal_bg=bg)
            return lbl

        title_row = tk.Frame(header, bg=c["panel"])
        title_row.pack(fill="x", padx=12, pady=(8, 4))
        self.bind_app_menu(title_row)
        title_block = tk.Frame(title_row, bg=c["panel"])
        title_block.pack(side="left", fill="x", expand=True)
        self.bind_app_menu(title_block)
        title_line = tk.Frame(title_block, bg=c["panel"])
        title_line.pack(anchor="w")
        tk.Label(title_line, text="仮想トレー", fg=c["cyan"], bg=c["panel"], font=("Segoe UI", 17, "bold")).pack(side="left")
        tk.Label(title_line, text=" // クリップ管理", fg=c["muted"], bg=c["panel"], font=("Segoe UI", 10, "bold")).pack(side="left", padx=(6, 0), pady=(7, 0))
        sub = "日本語UI / 使いやすさ改善版" + (" / セーフ" if self.safe_mode else "")
        tk.Label(title_block, text=sub, fg=c["mag"], bg=c["panel"], font=("Segoe UI", 7, "bold")).pack(anchor="w")

        top_controls = tk.Frame(title_row, bg=c["panel"])
        top_controls.pack(side="right", anchor="e")
        self.bind_app_menu(top_controls)
        self.priority_btn = header_btn(top_controls, "文字優先", self.toggle_clipboard_priority, c["cyan"], width=11)
        self.capture_btn = header_btn(top_controls, "● 取込中", self.toggle_capture, c["green"], bg="#08221d", width=10)
        self.bg_btn = header_btn(top_controls, "裏取込OFF", self.toggle_background_capture, c["muted"], width=7)
        header_btn(top_controls, "終了", self.request_exit_app, c["warn"], bg="#240914", width=5)
        header_btn(top_controls, "—", self.stop_panel, c["muted"], bg=c["panel"], width=2)
        header_btn(top_controls, "×", self.stop_panel, c["muted"], bg=c["panel"], width=2)

        panels = tk.Frame(header, bg=c["panel"])
        panels.pack(fill="x", padx=12, pady=(4, 10))
        self.bind_app_menu(panels)
        panels.grid_columnconfigure(0, weight=1, uniform="status")
        panels.grid_columnconfigure(1, weight=1, uniform="status")
        slot_panel = tk.Frame(panels, bg="#07111f", highlightthickness=1, highlightbackground=c["mag"])
        clip_panel = tk.Frame(panels, bg="#061722", highlightthickness=1, highlightbackground=c["cyan"])
        slot_panel.grid(row=0, column=0, sticky="ew", padx=(0, 6))
        clip_panel.grid(row=0, column=1, sticky="ew", padx=(6, 0))
        tk.Label(slot_panel, text="選択中", fg=c["muted"], bg="#07111f", font=("Segoe UI", 7, "bold")).pack(anchor="w", padx=10, pady=(6, 0))
        tk.Label(slot_panel, textvariable=self.active_slot_var, fg=c["mag"], bg="#07111f", font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=10)
        tk.Label(slot_panel, textvariable=self.ready_var, fg=c["text"], bg="#07111f", font=("Meiryo UI", 8), anchor="w").pack(fill="x", padx=10, pady=(0, 7))
        tk.Label(clip_panel, text="Windowsクリップボード", fg=c["muted"], bg="#061722", font=("Segoe UI", 7, "bold")).pack(anchor="w", padx=10, pady=(6, 0))
        tk.Label(clip_panel, textvariable=self.clip_slot_var, fg=c["cyan"], bg="#061722", font=("Segoe UI", 13, "bold")).pack(anchor="w", padx=10)
        tk.Label(clip_panel, textvariable=self.clip_ready_var, fg=c["text"], bg="#061722", font=("Meiryo UI", 8), anchor="w").pack(fill="x", padx=10, pady=(0, 7))

        hint_text = "Enter=セット / P=貼る / J=貼って次へ / ↑↓=選択移動 / F=一覧 / A=追加 / T=文字だけ / I=画像だけ / U=戻す / ?=ヘルプ / Win+Alt+V=一覧"
        hint = tk.Label(outer, text=hint_text, fg=c["muted"], bg=c["bg"], font=("Meiryo UI", 8), anchor="w", justify="left")
        hint.pack(fill="x", padx=12, pady=(7, 4))
        self.bind_app_menu(hint)

        wrap = tk.Frame(outer, bg=c["bg"], highlightthickness=1, highlightbackground=c["line"])
        wrap.pack(fill="both", expand=True, padx=12, pady=4)
        self.canvas = tk.Canvas(wrap, bg=c["bg"], highlightthickness=0)
        sb = tk.Scrollbar(wrap, command=self.canvas.yview)
        sb.pack(side="right", fill="y")
        self.canvas.pack(side="left", fill="both", expand=True)
        self.canvas.configure(yscrollcommand=sb.set)
        self.card_area = tk.Frame(self.canvas, bg=c["bg"])
        self.canvas_window = self.canvas.create_window((0, 0), window=self.card_area, anchor="nw")
        self.card_area.bind("<Configure>", lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all")))
        self.canvas.bind("<Configure>", lambda e: self.canvas.itemconfigure(self.canvas_window, width=e.width))

        action = tk.Frame(outer, bg=c["panel"], highlightthickness=1, highlightbackground=c["cyan"])
        action.pack(fill="x", padx=6, pady=(4, 6))
        self.bind_app_menu(action)
        self.slot_strip = tk.Frame(action, bg=c["panel"])
        self.slot_strip.pack(fill="x", padx=12, pady=(8, 2))

        row1 = tk.Frame(action, bg=c["panel"])
        row1.pack(fill="x", padx=8)
        for text, cmd, color in (
            ("セット", self.copy_active, c["mag"]),
            ("貼る", self.paste_active, c["green"]),
            ("貼って次", self.paste_active_then_next, c["green"]),
            ("一覧", self.open_picker, c["cyan"]),
            ("追加", lambda: self.add_from_clipboard(False), c["cyan"]),
            ("文字", lambda: self.add_from_clipboard(False, mode="text"), c["green"]),
            ("画像", lambda: self.add_from_clipboard(False, mode="image"), c["violet"]),
            ("メモ", self.add_manual_text, c["green"]),
            ("全セット", lambda: self.copy_items(self.store.items, label="ALL"), c["cyan"]),
        ):
            self.pack_button(row1, text, cmd, color)

        row2 = tk.Frame(action, bg=c["panel"])
        row2.pack(fill="x", padx=8)
        for text, cmd, color in (
            ("選択取得", self.capture_selection_now, c["green"]),
            ("全貼付", lambda: self.paste_items(self.store.items, label="ALL"), c["violet"]),
            ("1行化", lambda: self.copy_selected_as_text_variant("one_line"), c["cyan"]),
            ("空行整理", lambda: self.copy_selected_as_text_variant("collapse_blank_lines"), c["cyan"]),
            ("書出", self.export_selected_as_text, c["green"]),
            ("戻す", self.restore_last_removed, c["green"]),
            ("CB消去", self.wipe_clipboard, c["warn"]),
            ("整理", self.clear_session, c["warn"]),
            ("格納", self.stop_panel, c["violet"]),
            ("終了", self.request_exit_app, c["warn"]),
        ):
            self.pack_button(row2, text, cmd, color)

        tk.Label(action, textvariable=self.status_var, fg=c["green"], bg=c["panel"], anchor="w", font=("Segoe UI", 8, "bold")).pack(fill="x", padx=12, pady=(0, 8))
        self.refresh_cards()

    def button(self, parent, text: str, x: int, y: int, w: int, cmd, color: str, h: int = 28):
        lbl = tk.Label(parent, text=text, fg=color, bg="#081426", font=("Segoe UI", 8, "bold"), highlightthickness=1, highlightbackground=color, cursor="hand2")
        lbl.place(x=x, y=y, width=w, height=h)
        lbl.bind("<Button-1>", lambda e: cmd())
        lbl.bind("<Enter>", lambda e: lbl.configure(bg="#142b46"))
        lbl.bind("<Leave>", lambda e: lbl.configure(bg="#081426"))
        return lbl

    def pack_button(self, parent, text: str, cmd, color: str):
        lbl = tk.Label(parent, text=text, fg=color, bg="#081426", font=("Segoe UI", 8, "bold"), highlightthickness=1, highlightbackground=color, cursor="hand2", padx=10, pady=7)
        lbl.pack(side="left", padx=4, pady=8)
        lbl.bind("<Button-1>", lambda e: cmd())
        lbl.bind("<Enter>", lambda e: lbl.configure(bg="#142b46"))
        lbl.bind("<Leave>", lambda e: lbl.configure(bg="#081426"))
        return lbl

    def bind_drag(self, widget: tk.Widget) -> None:
        widget.bind("<ButtonPress-1>", self.begin_drag, add="+")
        widget.bind("<B1-Motion>", self.on_drag, add="+")
        widget.bind("<ButtonRelease-1>", self.end_drag, add="+")

    def on_canvas_mousewheel(self, e: tk.Event):
        if self.compact or not hasattr(self, "canvas"):
            return None
        try:
            if getattr(e, "num", None) == 4:
                self.canvas.yview_scroll(-3, "units")
            elif getattr(e, "num", None) == 5:
                self.canvas.yview_scroll(3, "units")
            else:
                delta = int(getattr(e, "delta", 0) or 0)
                if delta:
                    self.canvas.yview_scroll(-max(-3, min(3, delta // 120)), "units")
            return "break"
        except Exception:
            return None

    def begin_drag(self, e: tk.Event) -> None:
        self.drag_offset = (e.x_root - self.root.winfo_x(), e.y_root - self.root.winfo_y())

    def on_drag(self, e: tk.Event) -> None:
        if self.drag_offset:
            dx, dy = self.drag_offset
            move_tk_window(self.root, self.root.winfo_width(), self.root.winfo_height(), e.x_root - dx, e.y_root - dy)

    def end_drag(self, e: tk.Event) -> None:
        if self.drag_offset and self.compact:
            try:
                w = max(self.compact_size, int(self.root.winfo_width() or self.compact_size))
                h = max(self.compact_size, int(self.root.winfo_height() or self.compact_size))
                self.x, self.y = self.clamp_xy(self.root.winfo_x(), self.root.winfo_y(), w, h, margin=8)
                move_tk_window(self.root, w, h, self.x, self.y)
                self.persist_config()
            except Exception as exc:
                log_error(exc)
        self.drag_offset = None
        self.ensure_visible()

    def compact_press(self, e: tk.Event) -> None:
        self.cancel_compact_hold_timer()
        self._compact_hold_menu_opened = False
        self.drag_offset = (e.x_root - self.root.winfo_x(), e.y_root - self.root.winfo_y())
        self.compact_press_xy = (e.x_root, e.y_root)
        self.compact_dragged = False
        try:
            self._compact_hold_after_id = self.root.after(1100, self.compact_long_press_menu)
        except Exception:
            self._compact_hold_after_id = None

    def compact_motion(self, e: tk.Event) -> None:
        if not self.drag_offset:
            return
        if self.compact_press_xy:
            sx, sy = self.compact_press_xy
            if abs(e.x_root - sx) + abs(e.y_root - sy) > 4:
                self.compact_dragged = True
                self.cancel_compact_hold_timer()
        dx, dy = self.drag_offset
        move_tk_window(self.root, self.root.winfo_width(), self.root.winfo_height(), e.x_root - dx, e.y_root - dy)

    def compact_release(self, e: tk.Event) -> None:
        self.cancel_compact_hold_timer()
        if self._compact_hold_menu_opened:
            self.drag_offset = None
            self.compact_press_xy = None
            self.compact_dragged = False
            self._compact_hold_menu_opened = False
            return
        if self.compact_dragged:
            w = max(self.compact_size, int(self.root.winfo_width() or self.compact_size))
            h = max(self.compact_size, int(self.root.winfo_height() or self.compact_size))
            self.x, self.y = self.clamp_xy(self.root.winfo_x(), self.root.winfo_y(), w, h, margin=8)
            move_tk_window(self.root, w, h, self.x, self.y)
            self.persist_config()
        else:
            self.start_active()
        self.drag_offset = None
        self.compact_press_xy = None
        self.compact_dragged = False

    def update_header(self) -> None:
        if hasattr(self, "capture_btn"):
            if self.auto_capture:
                self.capture_btn.configure(text="● 取込中", fg=self.colors["green"], bg="#08221d")
            else:
                self.capture_btn.configure(text="○ 停止中", fg=self.colors["warn"], bg="#240914")
        if hasattr(self, "bg_btn"):
            if self.capture_when_compact:
                self.bg_btn.configure(text="裏取込ON", fg=self.colors["green"], bg="#08221d")
            else:
                self.bg_btn.configure(text="裏取込OFF", fg=self.colors["muted"], bg="#081426")
        if hasattr(self, "priority_btn"):
            if self.clipboard_priority == "image_first":
                self.priority_btn.configure(text="画像優先", fg=self.colors["violet"], bg="#180f2f")
            else:
                self.priority_btn.configure(text="文字優先", fg=self.colors["cyan"], bg="#081426")
        active = self.active_item()
        if active:
            slot = self.slot_index(active) or 0
            self.active_slot_var.set(f"選択 [{slot}]")
            self.ready_var.set(f"{kind_label(active)}  {active.preview(44)}")
        else:
            self.active_slot_var.set("選択 --")
            self.ready_var.set("項目が選択されていません")
        clip_slot, clip_ready = self.clipboard_description()
        self.clip_slot_var.set(clip_slot)
        self.clip_ready_var.set(clip_ready)

    def clipboard_description(self) -> tuple[str, str]:
        self.clear_markers_if_missing()
        ids = list(self.clipboard_item_ids)
        if not ids:
            return "貼付候補 --", "セットするとCtrl+Vで貼り付けできます"
        items = self.items_for_ids(ids)
        if not items:
            return "貼付候補 --", "貼付候補なし"
        if len(items) == len(self.store.items) and len(items) > 1:
            return "貼付候補 [全項目]", f"{len(items)}件を貼付できます"
        slots = [str(self.slot_index(it)) for it in items if self.slot_index(it)]
        slot_text = "+".join(slots[:4]) + ("…" if len(slots) > 4 else "")
        if len(items) == 1:
            it = items[0]
            return f"貼付候補 [{slot_text}]", f"{kind_label(it)}  {short_text(it.title, 36)}"
        kinds = "/".join(sorted({kind_label(it) for it in items}))
        return f"貼付候補 [{slot_text}]", f"{len(items)}件選択 / {kinds}"

    def refresh_slot_strip(self) -> None:
        if self.compact or not hasattr(self, "slot_strip"):
            return
        for w in self.slot_strip.winfo_children():
            w.destroy()
        c = self.colors
        if not self.store.items:
            tk.Label(self.slot_strip, text="候補なし", fg=c["muted"], bg=c["panel"], font=("Segoe UI", 8, "bold")).pack(side="left")
            return
        clip_ids = set(self.clipboard_item_ids)
        pin_count = sum(1 for x in self.store.items if x.pinned)
        tk.Label(self.slot_strip, text=f"候補 {len(self.store.items)} / ピン {pin_count} / {human_bytes(self.store.total_bytes)}", fg=c["muted"], bg=c["panel"], font=("Segoe UI", 8, "bold")).pack(side="left", padx=(0, 6))
        for slot, it in enumerate(self.store.items[:12], start=1):
            accent = c["green"] if it.id in clip_ids else (c["mag"] if it.id == self.active_id else c["cyan"])
            bg = "#143326" if it.id in clip_ids else ("#381346" if it.id == self.active_id else "#081426")
            lbl = tk.Label(self.slot_strip, text=str(slot), fg=accent, bg=bg, font=("Segoe UI", 9, "bold"), width=3, highlightthickness=1, highlightbackground=accent, cursor="hand2")
            lbl.pack(side="left", padx=2, ipady=2)
            lbl.bind("<Button-1>", lambda e, n=slot: self.set_active_by_slot(n))
            lbl.bind("<Double-Button-1>", lambda e, item=it, n=slot: self.copy_items([item], label=f"SLOT [{n}]"))
        if len(self.store.items) > 12:
            tk.Label(self.slot_strip, text="…", fg=c["muted"], bg=c["panel"], font=("Segoe UI", 9, "bold")).pack(side="left", padx=4)

    def refresh_cards(self) -> None:
        if self.compact or not hasattr(self, "card_area"):
            return
        self.clear_markers_if_missing()
        if self.active_id and not any(it.id == self.active_id for it in self.store.items):
            self.active_id = self.store.items[0].id if self.store.items else None
            self.selected = {self.active_id} if self.active_id else set()
        elif not self.active_id and self.store.items:
            self.active_id = self.store.items[0].id
            self.selected = {self.active_id}
        for w in self.card_area.winfo_children():
            w.destroy()
        if not self.store.items:
            tk.Label(self.card_area, text="トレーは空です\n文字・ファイル・画像をCtrl+Cしてください", fg=self.colors["muted"], bg=self.colors["bg"], font=("Meiryo UI", 10), pady=70).pack(fill="both", expand=True)
        else:
            for idx, it in enumerate(self.store.items[: self.main_render_limit], start=1):
                self.draw_card(it, idx)
            more = len(self.store.items) - self.main_render_limit
            if more > 0:
                tk.Label(self.card_area, text=f"他 {more} 件は一覧で表示できます。常駐時の軽さ優先でメイン表示を制限中。", fg=self.colors["muted"], bg=self.colors["bg"], font=("Meiryo UI", 8), pady=8).pack(fill="x")
        self.update_header()
        self.refresh_slot_strip()

    def draw_card(self, it: TrayItem, slot: int) -> None:
        c = self.colors
        is_active = it.id == self.active_id
        selected = it.id in self.selected
        on_clip = it.id in set(self.clipboard_item_ids)
        accent = c["green"] if on_clip else (c["mag"] if it.kind == "text" else (c["violet"] if it.kind == "folder" else c["cyan"]))
        bg = "#123426" if on_clip else (c["slot_active"] if is_active else ("#111b35" if selected else c["panel2"]))
        frame = tk.Frame(self.card_area, bg=bg, highlightthickness=2 if (is_active or on_clip) else 1, highlightbackground=accent)
        frame.pack(fill="x", padx=8, pady=6)
        left = tk.Frame(frame, bg=bg, width=58)
        left.pack(side="left", fill="y", padx=(8, 2), pady=8)
        left.pack_propagate(False)
        tk.Label(left, text=f"[{slot}]", fg=c["text"], bg=accent, font=("Segoe UI", 10, "bold")).pack(fill="x", pady=(0, 6))
        tk.Label(left, text=("ピン" if it.pinned else kind_label(it)), fg=accent, bg=bg, font=("Segoe UI", 7, "bold")).pack(fill="x")
        mid = tk.Frame(frame, bg=bg)
        mid.pack(side="left", fill="x", expand=True, pady=8, padx=6)
        tag = "貼付候補" if on_clip else ("選択中" if is_active else ("ピン留め" if it.pinned else kind_label(it)))
        tk.Label(mid, text=tag, fg=accent, bg=bg, font=("Segoe UI", 7, "bold")).pack(anchor="w")
        tk.Label(mid, text=short_text(it.title, 48), fg=c["text"], bg=bg, font=("Meiryo UI", 9, "bold" if is_active else "normal")).pack(anchor="w")
        tk.Label(mid, text=short_text(it.detail, 54), fg=c["muted"], bg=bg, font=("Segoe UI", 7)).pack(anchor="w")
        right = tk.Frame(frame, bg=bg, width=70)
        right.pack(side="right", fill="y", padx=6, pady=8)
        right.pack_propagate(False)
        tk.Label(right, text=it.added_at, fg=accent, bg=bg, font=("Segoe UI", 8)).pack(anchor="e")
        act = tk.Label(right, text="候補" if on_clip else ("セット" if is_active else "操作"), fg=accent, bg="#0a1424", font=("Segoe UI", 7, "bold"), cursor="hand2")
        act.pack(anchor="e", pady=(9, 0), ipadx=7, ipady=2)
        act.bind("<Button-1>", lambda e, item=it: self.card_quick_action(e, item))

        def bind_card_widget(w: tk.Widget) -> None:
            if w is act:
                return
            w.bind("<Button-1>", lambda e, item=it: self.card_click(e, item), add="+")
            w.bind("<Double-Button-1>", lambda e, item=it: self.copy_items([item], label=f"SLOT [{self.slot_index(item) or '?'}]"), add="+")
            w.bind("<Button-3>", lambda e, item=it: self.card_menu(e, item), add="+")
            for child in w.winfo_children():
                bind_card_widget(child)

        bind_card_widget(frame)

    def card_click(self, e: tk.Event, it: TrayItem) -> None:
        self.active_id = it.id
        if e.state & 0x0004 or e.state & 0x0001:  # Ctrl or Shift-ish
            if it.id in self.selected:
                self.selected.remove(it.id)
            else:
                self.selected.add(it.id)
            if not self.selected:
                self.selected = {it.id}
        else:
            self.selected = {it.id}
        self.request_refresh()

    def card_quick_action(self, e: tk.Event, it: TrayItem):
        if it.id == self.active_id:
            self.copy_items([it], label=f"SLOT [{self.slot_index(it) or '?'}]")
        else:
            self.card_menu(e, it)
        return "break"

    def card_menu(self, e: tk.Event, it: TrayItem) -> None:
        m = tk.Menu(self.root, tearoff=0)
        slot = self.slot_index(it) or "?"
        m.add_command(label=f"[{slot}] この項目をクリップボードへ", command=lambda: self.copy_items([it], label=f"SLOT [{slot}]"))
        m.add_command(label=f"[{slot}] この項目を即貼り", command=lambda: self.paste_items([it], label=f"SLOT [{slot}]"))
        if it.kind == "text":
            m.add_separator()
            m.add_command(label="テキストを編集", command=lambda item=it: self.edit_text_item(item))
            m.add_command(label="テキストを一時ファイルで開く", command=lambda item=it: self.open_text_temp(item))
            m.add_separator()
            m.add_command(label="テキストを1行化してセット", command=lambda item=it: self.copy_text_variant(item, "one_line"))
            m.add_command(label="前後の空白を除いてセット", command=lambda item=it: self.copy_text_variant(item, "trim"))
            m.add_command(label="末尾改行だけ除いてセット", command=lambda item=it: self.copy_text_variant(item, "rstrip_newline"))
            m.add_command(label="Markdown引用にしてセット", command=lambda item=it: self.copy_text_variant(item, "quote"))
            m.add_command(label="箇条書きにしてセット", command=lambda item=it: self.copy_text_variant(item, "bullet"))
            m.add_command(label="コードブロックにしてセット", command=lambda item=it: self.copy_text_variant(item, "codeblock"))
            m.add_command(label="URLエンコードしてセット", command=lambda item=it: self.copy_text_variant(item, "url_encode"))
            m.add_command(label="URLデコードしてセット", command=lambda item=it: self.copy_text_variant(item, "url_decode"))
            m.add_command(label="JSON文字列としてセット", command=lambda item=it: self.copy_text_variant(item, "json_string"))
            m.add_command(label="JSONを整形してセット", command=lambda item=it: self.copy_text_variant(item, "json_pretty"))
            m.add_command(label="全角/半角などをNFKC正規化してセット", command=lambda item=it: self.copy_text_variant(item, "nfkc"))
            m.add_command(label="重複行を除いてセット", command=lambda item=it: self.copy_text_variant(item, "dedupe_lines"))
            m.add_command(label="行をソートしてセット", command=lambda item=it: self.copy_text_variant(item, "sort_lines"))
            m.add_command(label="空行を整理してセット", command=lambda item=it: self.copy_text_variant(item, "collapse_blank_lines"))
            m.add_command(label="空行を削除してセット", command=lambda item=it: self.copy_text_variant(item, "remove_blank_lines"))
            m.add_command(label="行をカンマ区切りにしてセット", command=lambda item=it: self.copy_text_variant(item, "join_comma"))
            m.add_command(label="行をタブ区切りにしてセット", command=lambda item=it: self.copy_text_variant(item, "join_tab"))
            m.add_command(label="Markdown表用に | をエスケープしてセット", command=lambda item=it: self.copy_text_variant(item, "markdown_escape_table"))
        if it.kind in ("file", "folder") and it.path:
            m.add_separator()
            m.add_command(label="ファイル/フォルダを開く", command=lambda p=it.path: self.open_path(p))
            m.add_command(label="親フォルダを開く", command=lambda p=it.path: self.open_parent(p))
            m.add_command(label="エクスプローラーで選択", command=lambda p=it.path: self.reveal_path(p))
            m.add_separator()
            m.add_command(label="パスをテキストとしてセット", command=lambda p=it.path: self.copy_plain_text(p, "PATH"))
            m.add_command(label="引用符付きパスをセット", command=lambda p=it.path: self.copy_plain_text('"' + p + '"', "QUOTED PATH"))
            m.add_command(label="ファイル名だけセット", command=lambda p=it.path: self.copy_plain_text(os.path.basename(p.rstrip("\\/")) or p, "BASENAME"))
            m.add_command(label="親フォルダパスをセット", command=lambda p=it.path: self.copy_plain_text(os.path.dirname(p.rstrip("\\/")) or p, "PARENT PATH"))
            m.add_command(label="Markdownリンクとしてセット", command=lambda p=it.path: self.copy_path_as_markdown(p))
        m.add_separator()
        m.add_command(label="一番上へ移動", command=lambda item=it: self.move_item_to_top(item))
        m.add_command(label="1つ上へ移動", command=lambda item=it: self.move_item_delta(item, -1))
        m.add_command(label="1つ下へ移動", command=lambda item=it: self.move_item_delta(item, 1))
        if it.kind != "image":
            m.add_separator()
            m.add_command(label=("ピン留めを解除" if it.pinned else "ピン留めして明示保存"), command=lambda item=it: self.toggle_pin(item))
        m.add_separator()
        m.add_command(label="選択項目をクリップボードへ", command=lambda: self.copy_items(self.selected_items_or_all(), label="SELECTED"))
        m.add_command(label="選択項目を1行化してセット", command=lambda: self.copy_selected_as_text_variant("one_line"))
        m.add_command(label="選択項目を箇条書きにしてセット", command=lambda: self.copy_selected_as_text_variant("bullet"))
        m.add_command(label="選択項目を空行整理してセット", command=lambda: self.copy_selected_as_text_variant("collapse_blank_lines"))
        m.add_command(label="選択項目をカンマ区切りにしてセット", command=lambda: self.copy_selected_as_text_variant("join_comma"))
        m.add_command(label="選択項目をタブ区切りにしてセット", command=lambda: self.copy_selected_as_text_variant("join_tab"))
        m.add_command(label="選択項目をテキスト書き出し", command=self.export_selected_as_text)
        m.add_command(label="選択項目を削除", command=lambda: self.delete_ids(set(self.selected) or {it.id}))
        m.add_command(label="直前の削除/消去を戻す", command=self.restore_last_removed)
        m.add_separator()
        m.add_command(label="Windowsクリップボードだけ空にする", command=self.wipe_clipboard)
        m.add_command(label="セッション項目だけ消去（ピン留め保持）", command=self.clear_session)
        m.add_command(label="仮想トレーを全消去（ピン留め含む）", command=self.clear_all)
        m.add_command(label="仮想トレー + Windowsクリップボードを全消去", command=self.wipe_all)
        m.add_separator()
        m.add_command(label="選択/全項目をテキスト書き出し", command=self.export_selected_as_text)
        m.add_command(label="診断レポートを作成", command=self.open_diagnostic_report)
        m.add_command(label="待機ボタン位置を右上へ戻す", command=self.reset_window_position)
        m.add_command(label=f"待機ボタンサイズ：{self.compact_size}px", command=self.cycle_compact_size)
        m.add_separator()
        m.add_command(label="使い方 / ホットキー", command=self.show_quick_help)
        m.add_command(label="格納", command=self.stop_panel)
        m.add_command(label="終了（確認あり）", command=self.request_exit_app)
        m.add_command(label="緊急終了（確認なし）", command=lambda: self.emergency_exit_app("card-menu"))
        return self.popup_menu(m, e)

    def compact_menu(self, e: Optional[tk.Event] = None, x: Optional[int] = None, y: Optional[int] = None):
        m = tk.Menu(self.root, tearoff=0)
        m.add_command(label="開く", command=self.start_active)
        m.add_command(label="一覧を開く", command=self.open_picker_from_hotkey)
        m.add_command(label="クリップボードから追加", command=lambda: self.add_from_clipboard(False))
        m.add_command(label="文字だけ追加", command=lambda: self.add_from_clipboard(False, mode="text"))
        m.add_command(label="画像だけ追加", command=lambda: self.add_from_clipboard(False, mode="image"))
        m.add_command(label="メモを追加", command=self.add_manual_text)
        m.add_command(label="格納中の裏取り込み：ON" if self.capture_when_compact else "格納中の裏取り込み：OFF", command=self.toggle_background_capture)
        m.add_command(label="自動取り込み：ON" if self.auto_capture else "自動取り込み：停止中", command=self.toggle_capture)
        m.add_command(label="優先：画像" if self.clipboard_priority == "image_first" else "優先：文字", command=self.toggle_clipboard_priority)
        m.add_command(label="貼って次へ：ON" if self.paste_advance else "貼って次へ：OFF", command=self.toggle_paste_advance)
        m.add_command(label="全体ホットキー：ON" if self.global_hotkeys else "全体ホットキー：OFF", command=self.toggle_global_hotkeys)
        m.add_command(label="プライバシー保護：ON" if self.privacy_guard else "プライバシー保護：OFF", command=self.toggle_privacy_guard)
        m.add_command(label=f"自動消去：{self.auto_clear_minutes}分" if self.auto_clear_minutes else "自動消去：OFF", command=self.cycle_auto_clear)
        m.add_separator()
        m.add_command(label="Windowsクリップボードだけ空にする", command=self.wipe_clipboard)
        m.add_command(label="セッション項目だけ消去（ピン留め保持）", command=self.clear_session)
        m.add_command(label="仮想トレー + Windowsクリップボードを全消去", command=self.wipe_all)
        m.add_separator()
        m.add_command(label="選択/全項目をテキスト書き出し", command=self.export_selected_as_text)
        m.add_command(label="直前の削除/消去を戻す", command=self.restore_last_removed)
        m.add_command(label="診断レポートを作成", command=self.open_diagnostic_report)
        m.add_command(label="待機ボタン位置を右上へ戻す", command=self.reset_window_position)
        m.add_command(label=f"待機ボタンサイズ：{self.compact_size}px", command=self.cycle_compact_size)
        m.add_separator()
        m.add_command(label="使い方 / ホットキー", command=self.show_quick_help)
        startup_label = "Windows起動時の自動起動を解除" if self.is_startup_registered() else "Windows起動時に自動起動する"
        m.add_command(label=startup_label, command=self.toggle_startup_registered)
        m.add_command(label="設定・ログフォルダを開く", command=self.open_log_folder)
        m.add_command(label="自己診断テスト", command=self.self_test)
        m.add_separator()
        m.add_command(label="終了（確認あり）", command=self.request_exit_app)
        m.add_command(label="緊急終了（確認なし）", command=lambda: self.emergency_exit_app("compact-menu"))
        return self.popup_menu(m, e, x=x, y=y)

    def open_picker(self) -> None:
        try:
            if self.picker is not None and self.picker.winfo_exists():
                self.picker.lift()
                self.picker.focus_force()
                return
        except Exception:
            self.picker = None
        c = self.colors
        p = tk.Toplevel(self.root)
        self.picker = p
        p.title("トレー一覧")
        p.geometry(self.picker_geometry())
        try:
            _sx, _sy, sw, sh = virtual_screen_bounds(self.root)
            p.minsize(min(760, max(360, sw - 24)), min(500, max(320, sh - 72)))
        except Exception:
            pass
        p.configure(bg=c["bg"])
        try:
            p.attributes("-topmost", True)
        except Exception:
            pass

        def close_picker() -> None:
            try:
                p.destroy()
            except Exception:
                pass
            if self.picker is p:
                self.picker = None

        p.protocol("WM_DELETE_WINDOW", close_picker)
        p.bind("<Escape>", lambda e: close_picker())

        header = tk.Frame(p, bg=c["panel"], highlightthickness=1, highlightbackground=c["cyan"])
        header.pack(fill="x", padx=8, pady=(8, 0))
        top_row = tk.Frame(header, bg=c["panel"])
        top_row.pack(fill="x", padx=14, pady=(8, 2))
        tk.Label(top_row, text="トレー一覧", fg=c["cyan"], bg=c["panel"], font=("Segoe UI", 22, "bold")).pack(side="left")
        xbtn = tk.Label(top_row, text="×", fg=c["muted"], bg=c["panel"], font=("Segoe UI", 18), cursor="hand2")
        xbtn.pack(side="right", padx=(8, 0))
        xbtn.bind("<Button-1>", lambda e: close_picker())
        tk.Label(top_row, text="Ctrl+Enter 貼る", fg=c["cyan"], bg="#071e2c", font=("Segoe UI", 8, "bold"), highlightthickness=1, highlightbackground=c["cyan"], padx=8, pady=4).pack(side="right", padx=5)
        tk.Label(top_row, text="Enter セット", fg=c["green"], bg="#08221d", font=("Segoe UI", 8, "bold"), highlightthickness=1, highlightbackground=c["green"], padx=8, pady=4).pack(side="right", padx=5)
        tk.Label(header, text="番号で選ぶ / Enter=セット / Ctrl+Enter=即貼り / J=貼って次へ / ●が実際のWindowsクリップボード", fg=c["muted"], bg=c["panel"], font=("Meiryo UI", 9), anchor="w", wraplength=760).pack(fill="x", padx=16, pady=(0, 8))

        tools = tk.Frame(p, bg=c["bg"])
        tools.pack(fill="x", padx=14, pady=10)
        qvar = tk.StringVar()
        entry = tk.Entry(tools, textvariable=qvar, bg="#07111f", fg=c["text"], insertbackground=c["cyan"], highlightthickness=1, highlightbackground=c["cyan"], font=("Meiryo UI", 10))
        entry.pack(side="left", fill="x", expand=True, ipady=7)
        fvar = tk.StringVar(value="all")
        for label, val in (("すべて", "all"), ("ピン", "pinned"), ("文字", "text"), ("画像", "image"), ("ファイル", "file"), ("フォルダ", "folder")):
            tk.Radiobutton(tools, text=label, value=val, variable=fvar, indicatoron=False, fg=c["cyan"], bg="#0a1424", selectcolor="#183556", font=("Segoe UI", 8, "bold"), padx=10, pady=6).pack(side="left", padx=(7, 0))

        main = tk.Frame(p, bg=c["bg"])
        main.pack(fill="both", expand=True, padx=14, pady=(0, 8))
        body = tk.Frame(main, bg=c["bg"], highlightthickness=1, highlightbackground=c["line"])
        body.pack(side="left", fill="both", expand=True)
        lb = tk.Listbox(body, selectmode="extended", bg="#07111f", fg=c["text"], selectbackground="#4b1a55", selectforeground="white", font=("Meiryo UI", 10), activestyle="none", highlightthickness=0)
        lb.pack(side="left", fill="both", expand=True)
        sb = tk.Scrollbar(body, command=lb.yview)
        sb.pack(side="right", fill="y")
        lb.configure(yscrollcommand=sb.set)

        preview = tk.Frame(main, bg=c["panel"], width=286, highlightthickness=1, highlightbackground=c["mag"])
        preview.pack(side="right", fill="y", padx=(10, 0))
        preview.pack_propagate(False)
        pv_slot = tk.StringVar(value="選択 --")
        pv_type = tk.StringVar(value="種類 --")
        pv_text = tk.StringVar(value="番号を選ぶとここにプレビュー")
        tk.Label(preview, textvariable=pv_slot, fg=c["mag"], bg=c["panel"], font=("Segoe UI", 16, "bold")).pack(anchor="w", padx=14, pady=(16, 2))
        tk.Label(preview, textvariable=pv_type, fg=c["cyan"], bg=c["panel"], font=("Segoe UI", 8, "bold")).pack(anchor="w", padx=15)
        tk.Label(preview, textvariable=pv_text, fg=c["text"], bg=c["panel"], justify="left", wraplength=252, font=("Meiryo UI", 9)).pack(anchor="nw", padx=14, pady=14, fill="both", expand=True)

        shown: list[TrayItem] = []

        def populate(*_):
            shown.clear()
            lb.delete(0, "end")
            q = qvar.get().lower().strip()
            flt = fvar.get()
            clip_ids = set(self.clipboard_item_ids)
            for slot, it in enumerate(self.store.items, start=1):
                if flt == "pinned":
                    if not it.pinned:
                        continue
                elif flt != "all" and it.kind != flt:
                    continue
                if q and q not in it.search:
                    continue
                shown.append(it)
                active_mark = "●" if it.id in clip_ids else ("◆" if it.id == self.active_id else ("★" if it.pinned else " "))
                lb.insert("end", f"{active_mark} [{slot:02d}] {kind_label(it):<5}  {short_text(it.title, 48):<48}  {it.added_at}")
            update_preview()

        def selected_items() -> list[TrayItem]:
            return [shown[i] for i in lb.curselection() if 0 <= i < len(shown)]

        def update_preview(*_):
            sel = selected_items()
            it = sel[0] if sel else self.active_item()
            if not it:
                pv_slot.set("選択 --")
                pv_type.set("種類 --")
                pv_text.set("候補がありません")
                return
            slot = self.slot_index(it) or 0
            pv_slot.set(f"候補 [{slot}]")
            state = "● 貼付候補" if it.id in set(self.clipboard_item_ids) else ("◆ 選択中" if it.id == self.active_id else ("★ ピン留め" if it.pinned else "候補"))
            pv_type.set(f"{kind_label(it)}  /  {state}")
            pv_text.set(it.preview(360))

        def copy_sel():
            items = selected_items() or self.active_items()
            self.copy_items(items, label="PICK")
            populate()

        def paste_sel():
            items = selected_items() or self.active_items()
            self.paste_items(items, label="PICK")

        def paste_next_sel():
            items = selected_items()
            if items:
                self.set_active_item(items[0], refresh=False)
            self.paste_active_then_next()
            populate()

        def remove_sel():
            items = selected_items()
            if not items:
                self.toast("選択されていません")
                return
            self.delete_ids({it.id for it in items})
            populate()

        def activate_selection():
            items = selected_items()
            if items:
                self.set_active_item(items[0], refresh=False)
                populate()
                self.request_refresh()
                update_preview()

        def export_sel():
            if selected_items():
                old = set(self.selected)
                self.selected = {it.id for it in selected_items()}
                try:
                    self.export_selected_as_text()
                finally:
                    self.selected = old
            else:
                self.export_selected_as_text()

        def select_all() -> str:
            lb.selection_set(0, "end")
            update_preview()
            return "break"

        def is_entry_event(e) -> bool:
            return isinstance(e.widget, tk.Entry)

        def on_number(e):
            if is_entry_event(e):
                return None
            ch = e.char or ""
            if ch.isdigit():
                slot = 10 if ch == "0" else int(ch)
                if self.set_active_by_slot(slot):
                    populate()
                    for i, it in enumerate(shown):
                        if self.slot_index(it) == slot:
                            lb.selection_clear(0, "end")
                            lb.selection_set(i)
                            lb.see(i)
                            update_preview()
                            break
                return "break"
            return None

        def on_return(e):
            if is_entry_event(e):
                return None
            copy_sel()
            return "break"

        def on_ctrl_return(e):
            if is_entry_event(e):
                return None
            paste_sel()
            return "break"

        def on_delete(e):
            if is_entry_event(e):
                return None
            remove_sel()
            return "break"

        lb.bind("<<ListboxSelect>>", update_preview)
        lb.bind("<Double-Button-1>", lambda e: (activate_selection(), copy_sel()))
        p.bind("<KeyPress>", on_number)
        p.bind("<Return>", on_return)
        p.bind("<Control-Return>", on_ctrl_return)
        p.bind("<Delete>", on_delete)
        p.bind("<Control-a>", lambda e: None if is_entry_event(e) else select_all())
        lb.bind("<Control-a>", lambda e: select_all())
        p.bind("j", lambda e: None if is_entry_event(e) else (paste_next_sel() or "break"))
        p.bind("J", lambda e: None if is_entry_event(e) else (paste_next_sel() or "break"))

        bottom = tk.Frame(p, bg=c["panel"], highlightthickness=1, highlightbackground=c["line"])
        bottom.pack(fill="x", padx=8, pady=(0, 8))
        btn_row = tk.Frame(bottom, bg=c["panel"])
        btn_row.pack(anchor="w", padx=10)
        self.pack_button(btn_row, "選択にする", activate_selection, c["cyan"])
        self.pack_button(btn_row, "セット", copy_sel, c["mag"])
        self.pack_button(btn_row, "貼る", paste_sel, c["green"])
        self.pack_button(btn_row, "貼って次", paste_next_sel, c["green"])
        self.pack_button(btn_row, "削除", remove_sel, c["warn"])
        self.pack_button(btn_row, "書き出し", export_sel, c["green"])
        self.pack_button(btn_row, "CB消去", self.wipe_clipboard, c["warn"])
        self.pack_button(btn_row, "閉じる", close_picker, c["cyan"])
        qvar.trace_add("write", populate)
        fvar.trace_add("write", populate)
        populate()
        lb.focus_set()
        p.after(80, p.focus_force)

    # ----- personal power actions -----
    def open_picker_from_hotkey(self) -> None:
        if self.compact:
            self.start_active()
            self.root.after(120, self.open_picker)
        else:
            self.open_picker()

    def handle_hotkey(self, hotkey_id: int) -> None:
        if not self.global_hotkeys:
            return
        self.update_last_target_from_foreground()
        if hotkey_id == HOTKEY_OPEN_PICKER:
            self.open_picker_from_hotkey()
        elif hotkey_id == HOTKEY_ADD_CLIP:
            self.add_from_clipboard(False)
        elif hotkey_id == HOTKEY_SET_ACTIVE:
            self.copy_active()
        elif hotkey_id == HOTKEY_PASTE_ACTIVE:
            self.paste_active()
        elif hotkey_id == HOTKEY_GET_SELECTION:
            self.capture_selection_now()
        elif hotkey_id == HOTKEY_NOTE:
            if self.compact:
                self.start_active()
            self.root.after(80, self.add_manual_text)
        elif hotkey_id == HOTKEY_EXIT:
            self.emergency_exit_app("Win+Alt+Q")

    def copy_plain_text(self, text: str, label: str = "TEXT") -> None:
        try:
            self.clip.write(text=str(text or ""), backup=True)
            self.own_write_seq = self.clip.sequence() if IS_WINDOWS else None
            if self.own_write_seq is not None:
                self.last_seq = self.own_write_seq
            self.clipboard_item_ids = ()
            self.clipboard_label = label
            self.request_refresh(header_only=True)
            self.toast(f"{label} READY")
        except Exception as exc:
            self.toast(f"{label} FAILED")
            log_error(exc)

    def copy_path_as_markdown(self, path: str) -> None:
        try:
            name = os.path.basename(path.rstrip("\\/")) or path
            uri = "file:///" + os.path.abspath(path).replace("\\", "/").replace(" ", "%20")
            self.copy_plain_text(f"[{name}]({uri})", "MARKDOWN LINK")
        except Exception as exc:
            self.toast("MARKDOWN LINK FAILED")
            log_error(exc)

    def reveal_path(self, path: str) -> None:
        try:
            if not path:
                return
            if IS_WINDOWS and os.path.exists(path) and not os.path.isdir(path):
                subprocess.Popen(["explorer", "/select," + path])
            else:
                self.open_parent(path)
        except Exception as exc:
            self.toast("REVEAL FAILED")
            log_error(exc)

    def set_active_delta(self, delta: int) -> bool:
        if not self.store.items:
            self.toast("TRAY EMPTY")
            return False
        cur = self.active_item()
        idx = self.store.items.index(cur) if cur in self.store.items else 0
        idx = max(0, min(len(self.store.items) - 1, idx + int(delta)))
        self.set_active_item(self.store.items[idx])
        try:
            if hasattr(self, "canvas"):
                # Rough, cheap scroll hint. Avoid querying every card on the hot path.
                frac = 0.0 if len(self.store.items) <= 1 else idx / max(1, len(self.store.items) - 1)
                self.canvas.yview_moveto(max(0.0, min(1.0, frac)))
        except Exception:
            pass
        return True

    def paste_active_then_next(self) -> None:
        it = self.active_item()
        if not it:
            self.toast("NO ACTIVE SLOT")
            return
        idx = self.store.items.index(it) if it in self.store.items else 0
        pasted = self.paste_items([it], label=f"SLOT [{self.slot_index(it) or '?'}]")
        if pasted and self.paste_advance and idx + 1 < len(self.store.items):
            self.root.after(280, lambda target=idx + 1: self.set_active_item(self.store.items[target]) if target < len(self.store.items) else None)

    def move_item_to_top(self, it: TrayItem) -> None:
        try:
            self.store.items.remove(it)
            self.store.items.insert(0, it)
            self.store._recalc()
            self.active_id = it.id
            self.selected = {it.id}
            self.save_pinned_safe()
            self.mark_tray_changed()
            self.request_refresh()
            self.toast("MOVED TO TOP")
        except ValueError:
            pass

    def move_item_delta(self, it: TrayItem, delta: int) -> None:
        try:
            i = self.store.items.index(it)
            j = max(0, min(len(self.store.items) - 1, i + int(delta)))
            if i == j:
                return
            self.store.items.pop(i)
            self.store.items.insert(j, it)
            self.store._recalc()
            self.save_pinned_safe()
            self.mark_tray_changed()
            self.request_refresh()
            self.toast("MOVED")
        except ValueError:
            pass

    def copy_selected_as_text_variant(self, mode: str) -> None:
        items = self.selected_items_or_all()
        texts = []
        for it in items:
            if it.kind == "text":
                texts.append(it.text)
            elif it.kind in ("file", "folder") and it.path:
                texts.append(it.path)
        if not texts:
            self.toast("NO TEXT/PATH ITEM")
            return
        joined = "\r\n".join(texts)
        tmp = make_text_item(text_variant(joined, mode))
        self.copy_items([tmp], label="TEXT VARIANT")

    def show_text_dialog(self, title: str, initial: str = "") -> Optional[str]:
        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.configure(bg=self.colors["bg"])
        try:
            dlg.attributes("-topmost", True)
        except Exception:
            pass
        w, h = 640, 460
        _sx, _sy, sw, sh = virtual_screen_bounds(self.root)
        x, y = self.clamp_xy(self.root.winfo_x() - 40, self.root.winfo_y() + 40, min(w, sw - 24), min(h, sh - 48), margin=10)
        dlg.geometry(tk_geometry_spec(min(w, sw - 24), min(h, sh - 48), x, y))
        header = tk.Label(dlg, text=title, fg=self.colors["cyan"], bg=self.colors["panel"], font=("Segoe UI", 12, "bold"), anchor="w", padx=12)
        header.pack(fill="x", padx=8, pady=(8, 0), ipady=8)
        txt = tk.Text(dlg, bg="#07111f", fg=self.colors["text"], insertbackground=self.colors["cyan"], undo=True, wrap="word", font=("Meiryo UI", 10))
        txt.pack(fill="both", expand=True, padx=8, pady=8)
        if initial:
            txt.insert("1.0", initial)
        result: dict[str, Optional[str]] = {"value": None}
        bottom = tk.Frame(dlg, bg=self.colors["panel"])
        bottom.pack(fill="x", padx=8, pady=(0, 8))
        def ok():
            result["value"] = txt.get("1.0", "end-1c")
            dlg.destroy()
        def cancel():
            result["value"] = None
            dlg.destroy()
        tk.Button(bottom, text="追加 / 保存", command=ok).pack(side="right", padx=8, pady=8)
        tk.Button(bottom, text="キャンセル", command=cancel).pack(side="right", padx=4, pady=8)
        dlg.bind("<Control-Return>", lambda e: (ok(), "break"))
        dlg.bind("<Escape>", lambda e: (cancel(), "break"))
        txt.focus_set()
        dlg.grab_set()
        self.root.wait_window(dlg)
        return result["value"]

    def edit_text_item(self, it: TrayItem) -> None:
        if it.kind != "text":
            self.toast("TEXT ONLY")
            return
        new_text = self.show_text_dialog("文字項目を編集  /  Ctrl+Enterで保存", it.text)
        if new_text is None:
            return
        if not new_text.strip():
            self.toast("EMPTY TEXT")
            return
        new_sig = make_sig("text", text=new_text)
        other = self.store.item_by_signature(new_sig)
        if other and other.id != it.id:
            self.toast("SAME TEXT ALREADY EXISTS")
            self.active_id = other.id
            self.selected = {other.id}
            self.request_refresh()
            return
        refresh_text_item_fields(it, new_text)
        self.store._recalc()
        self.save_pinned_safe()
        self.mark_tray_changed()
        self.request_refresh()
        self.toast("TEXT UPDATED")

    def open_text_temp(self, it: TrayItem) -> None:
        if it.kind != "text":
            self.toast("TEXT ONLY")
            return
        try:
            tmp_dir = os.path.join(DATA_DIR, "temp_preview")
            os.makedirs(tmp_dir, exist_ok=True)
            fname = "text_" + time.strftime("%Y%m%d_%H%M%S") + ".txt"
            path = os.path.join(tmp_dir, fname)
            with open(path, "w", encoding="utf-8", newline="") as f:
                f.write(it.text)
            if IS_WINDOWS:
                os.startfile(path)  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["xdg-open", path])
            self.toast("OPENED TEMP TEXT")
        except Exception as exc:
            self.toast("OPEN TEXT FAILED")
            log_error(exc)

    def toggle_clipboard_priority(self) -> None:
        self.clipboard_priority = "image_first" if self.clipboard_priority == "text_first" else "text_first"
        self.persist_config()
        self.update_header()
        self.toast("画像優先" if self.clipboard_priority == "image_first" else "文字優先")

    def toggle_paste_advance(self) -> None:
        self.paste_advance = not self.paste_advance
        self.persist_config()
        self.toast("PASTE ADVANCE ON" if self.paste_advance else "PASTE ADVANCE OFF")

    def toggle_global_hotkeys(self) -> None:
        if self.safe_mode:
            self.toast("SAFE MODE: HOTKEYS OFF")
            return
        self.global_hotkeys = not self.global_hotkeys
        self.persist_config()
        try:
            self.listener.register_hotkeys()
        except Exception as exc:
            log_error(exc)
        self.toast("GLOBAL HOTKEYS ON" if self.global_hotkeys else "GLOBAL HOTKEYS OFF")

    # ----- menus / toggles / keys -----
    def toggle_capture(self) -> None:
        self.auto_capture = not self.auto_capture
        self.persist_config()
        self.update_header()
        self.toast("CAPTURE ON" if self.auto_capture else "CAPTURE PAUSED")

    def toggle_background_capture(self) -> None:
        self.capture_when_compact = not self.capture_when_compact
        self.persist_config()
        self.update_header()
        self.toast("BACKGROUND CAPTURE ON" if self.capture_when_compact else "BACKGROUND CAPTURE OFF")

    def toggle_privacy_guard(self) -> None:
        self.privacy_guard = not self.privacy_guard
        self.persist_config()
        self.toast("PRIVACY GUARD ON" if self.privacy_guard else "PRIVACY GUARD OFF")

    def cycle_auto_clear(self) -> None:
        vals = [0, 30, 60, 240]
        try:
            idx = vals.index(int(self.auto_clear_minutes))
        except ValueError:
            idx = 0
        self.auto_clear_minutes = vals[(idx + 1) % len(vals)]
        self.persist_config()
        self.toast(f"AUTO CLEAR {self.auto_clear_minutes} MIN" if self.auto_clear_minutes else "AUTO CLEAR OFF")

    def on_main_key(self, e: tk.Event):
        if self.compact:
            return None
        ch = e.char or ""
        key = e.keysym or ""
        if (getattr(e, "state", 0) & 0x0004) and str(key).lower() == "a":
            self.selected = {it.id for it in self.store.items}
            self.request_refresh()
            self.toast(f"SELECTED {len(self.selected)}")
            return "break"
        if ch.isdigit():
            slot = 10 if ch == "0" else int(ch)
            self.set_active_by_slot(slot)
            return "break"
        if key in ("Return", "KP_Enter"):
            self.copy_active()
            return "break"
        if key == "Down":
            self.set_active_delta(1)
            return "break"
        if key == "Up":
            self.set_active_delta(-1)
            return "break"
        if key == "Escape":
            self.stop_panel()
            return "break"
        if key == "Delete":
            ids = set(self.selected) or ({self.active_id} if self.active_id else set())
            if ids:
                self.delete_ids(ids)
            return "break"
        low = ch.lower()
        if low == "p":
            self.paste_active()
            return "break"
        if low == "j":
            self.paste_active_then_next()
            return "break"
        if low == "k":
            self.set_active_delta(-1)
            return "break"
        if low == "m":
            it = self.active_item()
            if it:
                self.move_item_to_top(it)
            return "break"
        if low == "r":
            self.toggle_clipboard_priority()
            return "break"
        if low == "a":
            self.add_from_clipboard(False)
            return "break"
        if low == "t":
            self.add_from_clipboard(False, mode="text")
            return "break"
        if low == "i":
            self.add_from_clipboard(False, mode="image")
            return "break"
        if low == "g":
            self.capture_selection_now()
            return "break"
        if low == "n":
            self.add_manual_text()
            return "break"
        if low == "f":
            self.open_picker()
            return "break"
        if low == "w":
            self.wipe_clipboard()
            return "break"
        if low == "u":
            self.restore_last_removed()
            return "break"
        if low == "e":
            self.export_selected_as_text()
            return "break"
        if low in ("?", "h"):
            self.show_quick_help()
            return "break"
        return None

    def toggle_pin(self, it: TrayItem) -> None:
        if it.kind == "image":
            self.toast("IMAGE PIN is session-only")
            return
        if not it.pinned and it.sensitive:
            ok = messagebox.askyesno(APP_NAME, "このテキストは機密情報の可能性があります。\n\nピン留めすると内容がローカル設定ファイルに明示保存されます。続行しますか？")
            if not ok:
                return
        it.pinned = not it.pinned
        if it.pinned and it.added_at != "PIN":
            it.added_at = "PIN"
            it.search = (it.search + " pinned pin 固定 ピン").lower()
        self.save_pinned_safe()
        self.mark_tray_changed()
        self.request_refresh()
        self.toast("PINNED" if it.pinned else "UNPINNED")

    def add_manual_text(self) -> None:
        try:
            text = self.show_text_dialog("メモを追加  /  Ctrl+Enterで追加", "")
            if text is None:
                return
            if not text.strip():
                self.toast("EMPTY NOTE")
                return
            it = make_text_item(text)
            ok, _reason, resolved = self.store.add(it)
            self.active_id = resolved.id
            self.selected = {resolved.id}
            self.mark_tray_changed()
            self.request_refresh()
            self.toast("NOTE ADDED" if ok else "NOTE ALREADY EXISTS")
        except Exception as exc:
            self.toast("NOTE FAILED")
            log_error(exc)

    def copy_text_variant(self, it: TrayItem, mode: str) -> None:
        if it.kind != "text":
            self.toast("TEXT ONLY")
            return
        text = text_variant(it.text, mode)
        try:
            self.clip.write(text=text, backup=True)
            self.own_write_seq = self.clip.sequence() if IS_WINDOWS else None
            if self.own_write_seq is not None:
                self.last_seq = self.own_write_seq
            self.clipboard_item_ids = ()
            self.clipboard_label = "VARIANT"
            self.request_refresh(header_only=True)
            self.toast("TEXT VARIANT READY")
        except Exception as exc:
            self.toast("VARIANT COPY FAILED")
            log_error(exc)

    # ----- filesystem / misc -----
    def open_path(self, path: str) -> None:
        try:
            if not path:
                self.toast("PATH EMPTY")
                return
            if IS_WINDOWS:
                os.startfile(path)  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["xdg-open", path])
        except Exception as exc:
            self.toast("OPEN FAILED")
            log_error(exc)

    def open_parent(self, path: str) -> None:
        try:
            if not path:
                return
            # os.path.isdir is only used on explicit user action, not the clipboard hot path.
            target = os.path.dirname(path) if not os.path.isdir(path) else path
            if target:
                self.open_path(target)
        except Exception as exc:
            self.toast("OPEN FOLDER FAILED")
            log_error(exc)

    def export_selected_as_text(self) -> None:
        items = self.selected_items_or_all()
        if not items:
            self.toast("TRAY EMPTY")
            return
        sensitive = sum(1 for it in items if it.kind == "text" and it.sensitive)
        if sensitive:
            try:
                if not messagebox.askyesno(APP_NAME, f"選択中に機密情報らしいテキストが {sensitive} 件あります。\n\nエクスポートすると内容が平文の .txt に保存されます。続行しますか？"):
                    return
            except Exception:
                pass
        try:
            export_dir = os.path.join(DATA_DIR, "exports")
            os.makedirs(export_dir, exist_ok=True)
            path = os.path.join(export_dir, "tray_export_" + time.strftime("%Y%m%d_%H%M%S") + ".txt")
            with open(path, "w", encoding="utf-8", newline="\n") as f:
                f.write(f"{APP_NAME} export\n")
                f.write(f"version: {VERSION}\n")
                f.write(f"created: {datetime.datetime.now().isoformat(timespec='seconds')}\n")
                f.write(f"items: {len(items)}\n")
                f.write("\n")
                for idx, it in enumerate(items, start=1):
                    slot = self.slot_index(it) or idx
                    f.write("=" * 72 + "\n")
                    f.write(f"[{slot}] {kind_label(it)}  pinned={int(bool(it.pinned))}  added={it.added_at}\n")
                    f.write(f"title: {it.title}\n")
                    if it.detail:
                        f.write(f"detail: {it.detail}\n")
                    f.write("-" * 72 + "\n")
                    if it.kind == "text":
                        f.write(it.text)
                        if not it.text.endswith("\n"):
                            f.write("\n")
                    elif it.kind in ("file", "folder"):
                        f.write(it.path + "\n")
                    elif it.kind == "image":
                        f.write(f"<image binary is not exported here> {it.detail}\n")
                    f.write("\n")
            self.toast("EXPORTED TXT")
            self.open_path(path)
        except Exception as exc:
            self.toast("EXPORT FAILED")
            log_error(exc)

    def create_diagnostic_report(self) -> str:
        os.makedirs(DATA_DIR, exist_ok=True)
        path = os.path.join(DATA_DIR, "diagnostic_" + time.strftime("%Y%m%d_%H%M%S") + ".txt")
        try:
            cfg = dict(self.config)
            # Keep the report useful but do not dump clipboard contents or pinned text.
            for key in ("privacy_exclude_processes", "privacy_exclude_titles"):
                if key in cfg and isinstance(cfg[key], list):
                    cfg[key] = f"<{len(cfg[key])} entries>"
            counts: dict[str, int] = {}
            pinned = 0
            sensitive = 0
            total = 0
            for it in self.store.items:
                counts[it.kind] = counts.get(it.kind, 0) + 1
                pinned += int(bool(it.pinned))
                sensitive += int(bool(it.sensitive))
                total += max(0, int(it.size))
            with open(path, "w", encoding="utf-8", newline="\n") as f:
                f.write(f"{APP_NAME} diagnostic report\n")
                f.write(f"created: {datetime.datetime.now().isoformat(timespec='seconds')}\n")
                f.write(f"version: {VERSION}\n")
                f.write(f"python: {sys.version}\n")
                f.write(f"executable: {sys.executable}\n")
                f.write(f"platform: {platform.platform()}\n")
                f.write(f"windows: {IS_WINDOWS}\n")
                try:
                    f.write(f"tk: {self.root.tk.call('info', 'patchlevel')}\n")
                except Exception:
                    pass
                try:
                    sx, sy, sw, sh = virtual_screen_bounds(self.root)
                    f.write(f"virtual_screen: {sx},{sy},{sw}x{sh}\n")
                    f.write(f"root_geometry: {self.root.winfo_width()}x{self.root.winfo_height()}+{self.root.winfo_x()}+{self.root.winfo_y()} compact={self.compact}\n")
                except Exception:
                    pass
                f.write(f"data_dir: {DATA_DIR}\n")
                f.write(f"config_path: {CONFIG_PATH}\n")
                f.write(f"pinned_path: {PINNED_PATH}\n")
                f.write(f"listener_ok: {self.listener_ok}\n")
                f.write(f"safe_mode: {self.safe_mode}\n")
                f.write(f"global_hotkeys: {self.global_hotkeys}\n")
                f.write(f"auto_capture: {self.auto_capture}\n")
                f.write(f"background_capture: {self.capture_when_compact}\n")
                f.write(f"clipboard_priority: {self.clipboard_priority}\n")
                f.write(f"items_count: {len(self.store.items)}\n")
                f.write(f"items_by_kind: {counts}\n")
                f.write(f"pinned_count: {pinned}\n")
                f.write(f"sensitive_masked_count: {sensitive}\n")
                f.write(f"store_total_bytes: {total}\n")
                f.write("\n-- sanitized config --\n")
                f.write(json.dumps(cfg, ensure_ascii=False, indent=2, default=str))
                f.write("\n\n-- file sizes --\n")
                for fp in (CONFIG_PATH, PINNED_PATH, EVENT_LOG, ERROR_LOG):
                    try:
                        f.write(f"{fp}: {os.path.getsize(fp) if os.path.exists(fp) else 0} bytes\n")
                    except Exception as exc:
                        f.write(f"{fp}: <stat failed {exc}>\n")
                f.write("\n-- event.log tail --\n")
                f.write(tail_file(EVENT_LOG, self.diagnostic_tail_lines) or "<empty>\n")
                f.write("\n-- error.log tail --\n")
                f.write(tail_file(ERROR_LOG, self.diagnostic_tail_lines) or "<empty>\n")
            return path
        except Exception as exc:
            log_error(exc)
            raise

    def open_diagnostic_report(self) -> None:
        try:
            path = self.create_diagnostic_report()
            self.toast("DIAGNOSTIC READY")
            self.open_path(path)
        except Exception as exc:
            self.toast("DIAGNOSTIC FAILED")
            log_error(exc)

    def cycle_compact_size(self) -> None:
        sizes = [84, 102, 126, 150]
        cur = int(self.compact_size or COMPACT_DEFAULT_SIZE)
        next_size = sizes[0]
        for size in sizes:
            if cur < size:
                next_size = size
                break
        else:
            next_size = sizes[0]
        self.compact_size = max(COMPACT_MIN_SIZE, min(COMPACT_MAX_SIZE, int(next_size)))
        try:
            self.root.minsize(self.compact_size, self.compact_size)
        except Exception:
            pass
        self.persist_config()
        if self.compact:
            self.build_compact()
        self.toast(f"待機ボタンサイズ：{self.compact_size}px")

    def reset_window_position(self) -> None:
        try:
            sx, sy, sw, sh = virtual_screen_bounds(self.root)
            self.x = sx + max(20, sw - self.compact_size - 24)
            self.y = sy + max(20, min(170, sh - self.compact_size - 24))
            self.x, self.y = self.clamp_xy(self.x, self.y, self.compact_size, self.compact_size, margin=8)
            self.compact_anchor = (self.x, self.y)
            self.panel_open = False
            self.compact = True
            self.persist_config()
            self.build_compact()
            self.toast("POSITION RESET")
        except Exception as exc:
            self.toast("POSITION RESET FAILED")
            log_error(exc)

    def open_log_folder(self) -> None:
        try:
            if IS_WINDOWS:
                os.startfile(DATA_DIR)  # type: ignore[attr-defined]
            else:
                subprocess.Popen(["xdg-open", DATA_DIR])
        except Exception as exc:
            log_error(exc)

    def show_quick_help(self) -> None:
        help_text = (
            f"{APP_NAME}\n\n"
            "起動に必要なのはこの .pyw ファイルだけです。\n"
            "設定・ピン留め・ログは %APPDATA%\\DeskLayerVirtualTray に自動保存されます。\n\n"
            "■ 普段の使い方\n"
            "1. 待機ボタンをクリックしてパネルを開きます。\n"
            "2. 待機ボタン、またはパネル上部の『ここをドラッグして移動できます』をドラッグすると場所を動かせます。\n"
            "3. コピー済みの内容は『追加』『文字』『画像』でトレーへ入れます。\n"
            "4. 1〜9/0、または一覧から候補を選びます。\n"
            "5. Enterまたは『セット』でWindowsクリップボードへ入れ、貼り付け先でCtrl+Vします。\n"
            "6. 『貼る』は直前に使っていた外部ウィンドウへ戻って貼り付けを試します。失敗時は手動Ctrl+Vしてください。\n\n"
            "■ ホットキー\n"
            "Win+Alt+V : 一覧を開く\n"
            "Win+Alt+C : クリップボードから追加\n"
            "Win+Alt+S : 選択中の候補をセット\n"
            "Win+Alt+P : 選択中の候補を貼り付け\n"
            "Win+Alt+G : 外部アプリの選択範囲を取得\n"
            "Win+Alt+N : メモを追加\n"
            "Win+Alt+Q : 緊急終了（確認なし）\n\n"
            "■ パネル内キー\n"
            "Enter=セット / P=貼る / J=貼って次へ / A=追加 / T=文字だけ / I=画像だけ / "
            "R=文字優先・画像優先の切替 / F=一覧 / G=選択取得 / N=メモ / U=戻す / E=書き出し / "
            "W=Windowsクリップボード消去 / Esc=格納 / Alt+左ドラッグ=移動 / Ctrl+Q=終了 / Ctrl+Shift+Q=緊急終了\n\n"
            "■ 停止できない時\n"
            "右クリック/中クリック/Shift+F10メニュー、Ctrl+Q、Ctrl+Shift+Q、Win+Alt+Q、"
            "または コマンドプロンプトで python このファイル.pyw --stop-running を使えます。\n\n"
            "■ 推奨設定\n"
            "自動取り込み：ON / 格納中の裏取り込み：OFF / 優先：文字 / プライバシー保護：ON\n\n"
            "右クリックメニューから、画像優先、格納中の裏取り込み、プライバシー保護、自動消去、"
            "Windows起動時の自動起動を切り替えられます。"
        )
        try:
            messagebox.showinfo(f"{APP_NAME} - 使い方", help_text)
        except Exception as exc:
            log_error(exc)

    def startup_command(self) -> str:
        script_path = os.path.abspath(__file__)
        exe = sys.executable or "pythonw.exe"
        if IS_WINDOWS:
            try:
                pyw = os.path.join(os.path.dirname(exe), "pythonw.exe")
                if os.path.exists(pyw):
                    exe = pyw
            except Exception:
                pass
        return f'"{exe}" "{script_path}"'

    def is_startup_registered(self) -> bool:
        if not IS_WINDOWS:
            return False
        try:
            import winreg
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_READ) as key:
                value, _typ = winreg.QueryValueEx(key, "DeskLayerVirtualTray")
            return str(value).strip() == self.startup_command()
        except FileNotFoundError:
            return False
        except OSError:
            return False
        except Exception as exc:
            log_error(exc)
            return False

    def set_startup_registered(self, enabled: bool) -> None:
        if not IS_WINDOWS:
            self.toast("WINDOWS ONLY")
            return
        try:
            import winreg
            with winreg.CreateKeyEx(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Run", 0, winreg.KEY_SET_VALUE) as key:
                if enabled:
                    winreg.SetValueEx(key, "DeskLayerVirtualTray", 0, winreg.REG_SZ, self.startup_command())
                else:
                    try:
                        winreg.DeleteValue(key, "DeskLayerVirtualTray")
                    except FileNotFoundError:
                        pass
            self.toast("STARTUP ON" if enabled else "STARTUP OFF")
            try:
                messagebox.showinfo(APP_NAME, "Windows起動時の自動起動を登録しました。" if enabled else "Windows起動時の自動起動を解除しました。")
            except Exception:
                pass
        except Exception as exc:
            log_error(exc)
            self.toast("STARTUP FAILED")
            try:
                messagebox.showerror(APP_NAME, f"自動起動設定に失敗しました。\n\n{exc}\n\n{ERROR_LOG}")
            except Exception:
                pass

    def toggle_startup_registered(self) -> None:
        self.set_startup_registered(not self.is_startup_registered())

    def toast(self, msg: str) -> None:
        self.toast_id += 1
        token = self.toast_id
        msg = ja_status(msg)
        self.status_var.set(msg)
        if self.panel_open:
            self.root.after(2600, lambda t=token: self.status_var.set("準備完了") if self.panel_open and t == self.toast_id else None)

    def self_test(self) -> None:
        if not IS_WINDOWS:
            self.toast("Windows専用です")
            return
        backup_items: list[TrayItem] = []
        try:
            try:
                backup_items = self.clip.snapshot_items(timeout_ms=200, max_text_chars=500_000, max_image_bytes=16 * 1024 * 1024, priority=self.clipboard_priority)
            except Exception:
                backup_items = []
            test_text = "DeskLayer self test " + now_label()
            self.clip.write(text=test_text, backup=False)
            self.own_write_seq = self.clip.sequence()
            self.last_seq = self.own_write_seq
            if test_text not in self.clip.read_text():
                raise RuntimeError("clipboard read/write mismatch")
            item = make_text_item(test_text)
            item.title = "SELF TEST"
            item.detail = "test item"
            self.store.add(item)
            header = struct.pack("<IiiHHIIiiII", 40, 1, 1, 1, 32, 0, 4, 0, 0, 0, 0)
            img = make_image_item(header + b"\x00\x00\xff\xff", CF_DIB)
            self.store.add(img)
            self.copy_items([item], label="自己診断テスト")
            if test_text not in self.clip.read_text():
                raise RuntimeError("copy item mismatch")
            if backup_items:
                b_text, b_paths, b_dib, b_fmt, _ = collect_payload(backup_items)
                self.clip.write(text=b_text, paths=b_paths, dib=b_dib, dib_format=b_fmt, backup=False)
                self.own_write_seq = self.clip.sequence()
                self.last_seq = self.own_write_seq
                self.clipboard_item_ids = ()
                self.clipboard_label = "RESTORED"
            self.request_refresh()
            self.toast("SELF TEST PASS / CLIP RESTORED" if backup_items else "SELF TEST PASS")
        except Exception as exc:
            log_error(exc)
            self.toast("自己診断に失敗しました")
            messagebox.showerror(APP_NAME, f"自己診断に失敗しました\n\n{exc}\n\n{ERROR_LOG}")

    def emergency_exit_app(self, reason: str = "emergency") -> None:
        """Stop without confirmation.  This is the last-resort escape hatch."""
        try:
            log_event(f"emergency_exit {reason}")
        except Exception:
            pass
        try:
            self.exit_app()
        except Exception as exc:
            log_error(exc)
            try:
                self._shutting_down = True
                self.root.destroy()
            except Exception:
                os._exit(0)

    def request_exit_app(self) -> None:
        try:
            if self.confirm_exit:
                if not messagebox.askyesno(APP_NAME, "仮想トレーを終了しますか？\n\n常駐を続ける場合は『格納/×』で待機ボタンへ戻してください。"):
                    return
        except Exception:
            pass
        self.exit_app()

    def exit_app(self) -> None:
        self._shutting_down = True
        try:
            self.listener.stop()
        except Exception:
            pass
        if self.panel_open and self.compact_anchor:
            self.x, self.y = self.compact_anchor
        else:
            self.x, self.y = self.root.winfo_x(), self.root.winfo_y()
        self.persist_config()
        self.save_pinned_safe()
        log_event("exit")
        self.root.destroy()

    def run(self) -> None:
        self.root.mainloop()



def run_doctor() -> None:
    """Non-GUI sanity check: python DeskLayerVirtualTray_v5_0_pro.pyw --doctor"""
    print(f"{APP_NAME} {VERSION}")
    print(f"Python: {sys.version.split()[0]}  Platform: {platform.platform()}")
    checks: list[tuple[str, bool, str]] = []

    def record(name: str, ok: bool, detail: str = "") -> None:
        checks.append((name, bool(ok), detail))

    try:
        it = make_text_item("hello")
        record("text item", it.kind == "text" and it.text == "hello" and not it.sensitive)
        record("sensitive mask", make_text_item("password=abc123").sensitive)
        pth = make_path_item(os.getcwd())
        record("path item", pth.kind in {"file", "folder"}, pth.kind)
        header = struct.pack("<IiiHHIIiiII", 40, 1, 1, 1, 32, 0, 4, 0, 0, 0, 0)
        img = make_image_item(header + b"\x00\x00\xff\xff", CF_DIB)
        record("image item", img.kind == "image" and img.size >= 44, img.detail)
        store = TrayStore(max_items=2, max_total_bytes=10_000)
        store.add(make_text_item("a"))
        store.add(make_text_item("b"))
        store.add(make_text_item("c"))
        record("store trim", [x.text for x in store.items] == ["c", "b"])
        record("text one_line", text_variant(" a\n b ", "one_line") == "a b")
        record("text dedupe", text_variant("a\na\nb", "dedupe_lines") == "a\r\nb")
        record("text join", text_variant("a\nb\n", "join_comma") == "a, b")
        record("json pretty", text_variant('{"b":2,"a":1}', "json_pretty").startswith("{"))
        record("geometry negative", tk_geometry_spec(100, 50, -10, 20) == "100x50-10+20")
        tcl = tk.Tcl()
        record("tk/tcl", bool(tcl.eval("info patchlevel")), tcl.eval("info patchlevel"))
        record("stop request path", STOP_REQUEST_PATH.endswith("stop.request"), STOP_REQUEST_PATH)
        record("rescue hotkey", any(h[0] == HOTKEY_EXIT for h in HOTKEY_DEFS), "Win+Alt+Q")
    except Exception as exc:
        record("doctor exception", False, repr(exc))

    failed = 0
    for name, ok, detail in checks:
        if not ok:
            failed += 1
        print(("PASS" if ok else "FAIL") + f"  {name}" + (f"  {detail}" if detail else ""))
    if failed:
        raise SystemExit(1)
    print("ALL PASS")


def reset_position_config() -> None:
    cfg = load_config()
    cfg["x"] = None
    cfg["y"] = None
    save_config(cfg)
    print(f"Position reset in {CONFIG_PATH}")


def reset_config_file() -> None:
    if os.path.exists(CONFIG_PATH):
        bak = CONFIG_PATH + ".bak_" + time.strftime("%Y%m%d_%H%M%S")
        os.replace(CONFIG_PATH, bak)
        print(f"Config moved to {bak}")
    else:
        print("Config file does not exist; nothing to reset")


_MUTEX_HANDLE = None


def acquire_single_instance() -> bool:
    global _MUTEX_HANDLE
    if not IS_WINDOWS:
        return True
    try:
        _MUTEX_HANDLE = kernel32.CreateMutexW(None, False, MUTEX_NAME)
        if not _MUTEX_HANDLE:
            return True
        return int(kernel32.GetLastError()) != ERROR_ALREADY_EXISTS
    except Exception:
        return True


def release_single_instance() -> None:
    global _MUTEX_HANDLE
    if IS_WINDOWS and _MUTEX_HANDLE:
        try:
            kernel32.CloseHandle(_MUTEX_HANDLE)
        except Exception:
            pass
        


def reset_pins_file() -> None:
    if os.path.exists(PINNED_PATH):
        bak = backup_or_quarantine_file(PINNED_PATH, "bak")
        print(f"Pinned file moved to {bak}" if bak else "Pinned file could not be moved")
    else:
        print("Pinned file does not exist; nothing to reset")


def print_data_dir() -> None:
    print(DATA_DIR)
    print(CONFIG_PATH)
    print(PINNED_PATH)
    print(EVENT_LOG)
    print(ERROR_LOG)
    print(STOP_REQUEST_PATH)


def request_running_instance_stop() -> None:
    os.makedirs(DATA_DIR, exist_ok=True)
    with open(STOP_REQUEST_PATH, "w", encoding="utf-8") as f:
        f.write(str(os.getpid()) + "\n" + time.strftime("%Y-%m-%d %H:%M:%S") + "\n")
    print(f"Stop request written: {STOP_REQUEST_PATH}")

def show_already_running_message() -> None:
    try:
        r = tk.Tk()
        r.withdraw()
        messagebox.showinfo(APP_NAME, "仮想トレーは既に起動しています。")
        r.destroy()
    except Exception:
        pass


def main() -> None:
    install_exception_logging()
    if "--doctor" in sys.argv:
        run_doctor()
        return
    if "--reset-position" in sys.argv:
        reset_position_config()
        return
    if "--reset-config" in sys.argv:
        reset_config_file()
        return
    if "--reset-pins" in sys.argv:
        reset_pins_file()
        return
    if "--data-dir" in sys.argv:
        print_data_dir()
        return
    if "--stop-running" in sys.argv or "--stop" in sys.argv:
        request_running_instance_stop()
        return
    if not acquire_single_instance():
        show_already_running_message()
        return
    try:
        app = DeskLayerApp()
        app.run()
    finally:
        release_single_instance()


if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        log_error(exc)
        try:
            r = tk.Tk()
            r.withdraw()
            messagebox.showerror(APP_NAME, f"起動エラー: {exc}\n\n{ERROR_LOG}")
            r.destroy()
        except Exception:
            pass
        # Under pythonw/.pyw there is no console; keeping the exception logged and
        # avoiding a second unhandled raise prevents a silent-looking crash loop.
