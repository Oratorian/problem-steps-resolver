"""
PSR (Problem Steps Recorder) Clone for Windows
================================================
Recreates the classic Windows PSR functionality with a faithful GUI:
- Compact toolbar window matching the original Windows Steps Recorder
- Captures a screenshot on every mouse click
- Logs the active window title and click coordinates
- Detects UI element under cursor when possible
- Generates an HTML report (like the original PSR .mht files)
- Can also export as a ZIP archive
- PSR window is click-transparent to recording (won't capture its own clicks)

Requirements:
    pip install pynput mss Pillow pygetwindow

Optional (for richer UI element detection):
    pip install pywinauto comtypes

Usage:
    python psr.py                  # Launch GUI (default)
    python psr.py --no-gui         # CLI-only mode (Ctrl+Shift+F9 to stop)
    python psr.py --output report  # Custom output name
    python psr.py --delay 0.5      # Min delay between captures (seconds)
    python psr.py --no-keyboard    # Don't log keyboard actions
"""

import os
import sys
import time
import datetime
import argparse
import threading
import base64
import json
import zipfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from io import BytesIO
from pathlib import Path
from dataclasses import dataclass, field
from typing import Callable, Optional

try:
    from pynput import mouse, keyboard
except ImportError:
    sys.exit("Missing dependency: pip install pynput")

try:
    import mss
    import mss.tools
except ImportError:
    sys.exit("Missing dependency: pip install mss")

try:
    from PIL import Image, ImageDraw, ImageFont
except ImportError:
    sys.exit("Missing dependency: pip install Pillow")

try:
    import pygetwindow as gw
    HAS_GETWINDOW = True
except ImportError:
    HAS_GETWINDOW = False

# Optional: richer UI element detection on Windows
HAS_WINAUTO = False
HAS_WIN32 = False

if sys.platform == "win32":
    try:
        import ctypes
        from ctypes import wintypes
        HAS_WIN32 = True
        # Set correct return types for HWND-returning functions.
        # Default restype is c_int (32-bit) which truncates 64-bit HWNDs.
        ctypes.windll.user32.WindowFromPoint.restype = wintypes.HWND
        ctypes.windll.user32.GetAncestor.restype = wintypes.HWND
        ctypes.windll.user32.ChildWindowFromPointEx.restype = wintypes.HWND
        ctypes.windll.user32.GetForegroundWindow.restype = wintypes.HWND
    except ImportError:
        pass

    try:
        from pywinauto import Desktop
        HAS_WINAUTO = True
    except ImportError:
        pass


# ---------------------------------------------------------------------------
# Configuration persistence
# ---------------------------------------------------------------------------

CONFIG_DIR = Path.home() / ".psr-clone-by-oratorian"
CONFIG_FILE = CONFIG_DIR / "settings.json"

DEFAULT_SETTINGS = {
    "output": "steps_report",
    "delay": 0.3,
    "record_keyboard": True,
    "fullscreen": False,
    "format": "html",
    "stop_hotkey": "ctrl+shift+f9",
}


def load_settings() -> dict:
    """Load settings from the config file, falling back to defaults."""
    settings = dict(DEFAULT_SETTINGS)
    if CONFIG_FILE.exists():
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
            settings.update(saved)
        except Exception:
            pass
    return settings


def save_settings(settings: dict):
    """Persist settings to the config file."""
    try:
        CONFIG_DIR.mkdir(parents=True, exist_ok=True)
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
    except Exception:
        pass


# ---------------------------------------------------------------------------
# Data structures
# ---------------------------------------------------------------------------

@dataclass
class Step:
    """One recorded step (click or key action)."""
    step_number: int
    timestamp: str
    action: str              # e.g. "Left Click", "Right Click", "Keyboard Input"
    window_title: str
    position: Optional[tuple] = None   # (x, y) screen coords
    ui_element: Optional[str] = None   # element name/class under cursor
    screenshot_png: Optional[bytes] = None  # PNG data
    details: str = ""


# ---------------------------------------------------------------------------
# UI Element Detection
# ---------------------------------------------------------------------------

def get_active_window_title() -> str:
    """Get the title of the currently active/foreground window."""
    if HAS_WIN32:
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        length = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
        buf = ctypes.create_unicode_buffer(length + 1)
        ctypes.windll.user32.GetWindowTextW(hwnd, buf, length + 1)
        return buf.value or "(unknown)"
    elif HAS_GETWINDOW:
        try:
            win = gw.getActiveWindow()
            return win.title if win else "(unknown)"
        except Exception:
            return "(unknown)"
    return "(unknown)"


def get_ui_element_at(x: int, y: int) -> str:
    """Try to identify the UI element at screen coordinates (x, y)."""
    if HAS_WINAUTO:
        try:
            desktop = Desktop(backend="uia")
            elem = desktop.from_point(x, y)
            name = elem.window_text() or ""
            ctrl_type = elem.element_info.control_type or ""
            class_name = elem.element_info.class_name or ""
            parts = [p for p in [ctrl_type, f'"{name}"' if name else "", class_name] if p]
            return " — ".join(parts) if parts else "(unknown element)"
        except Exception:
            pass

    if HAS_WIN32:
        try:
            point = wintypes.POINT(x, y)
            hwnd = ctypes.windll.user32.WindowFromPoint(point)
            if hwnd:
                class_buf = ctypes.create_unicode_buffer(256)
                ctypes.windll.user32.GetClassNameW(hwnd, class_buf, 256)
                text_len = ctypes.windll.user32.GetWindowTextLengthW(hwnd)
                text_buf = ctypes.create_unicode_buffer(text_len + 1)
                ctypes.windll.user32.GetWindowTextW(hwnd, text_buf, text_len + 1)
                parts = []
                if text_buf.value:
                    parts.append(f'"{text_buf.value}"')
                if class_buf.value:
                    parts.append(f"[{class_buf.value}]")
                return " ".join(parts) if parts else "(unknown element)"
        except Exception:
            pass

    return ""


def _get_deepest_child_hwnd(x: int, y: int):
    """Walk down the window hierarchy to find the deepest child control at (x, y)."""
    if not HAS_WIN32:
        return None
    try:
        point = wintypes.POINT(x, y)
        hwnd = ctypes.windll.user32.WindowFromPoint(point)
        if not hwnd:
            return None
        # Drill into children: convert screen coords to client coords and
        # repeatedly find child windows until we reach the deepest one.
        CWP_SKIPINVISIBLE = 0x0001
        CWP_SKIPDISABLED = 0x0002
        CWP_SKIPTRANSPARENT = 0x0004
        flags = CWP_SKIPINVISIBLE | CWP_SKIPDISABLED | CWP_SKIPTRANSPARENT
        while True:
            client_pt = wintypes.POINT(x, y)
            ctypes.windll.user32.ScreenToClient(hwnd, ctypes.byref(client_pt))
            child = ctypes.windll.user32.ChildWindowFromPointEx(
                hwnd, client_pt, flags
            )
            if not child or child == hwnd:
                break
            hwnd = child
        return hwnd
    except Exception:
        return None


def get_element_rect_at(x: int, y: int) -> Optional[tuple]:
    """
    Get the bounding rectangle (left, top, right, bottom) of the UI element
    at screen coordinates (x, y). Used to draw a highlight box like the
    original Windows PSR.
    """
    # Try pywinauto UIA first — gives the tightest element bounds
    if HAS_WINAUTO:
        try:
            desktop = Desktop(backend="uia")
            elem = desktop.from_point(x, y)
            rect = elem.element_info.rectangle
            if rect and rect.width() > 0 and rect.height() > 0:
                r = (rect.left, rect.top, rect.right, rect.bottom)
                # Reject if the rect is unreasonably large (probably a
                # top-level window instead of a specific control)
                if (rect.right - rect.left) < 800 and (rect.bottom - rect.top) < 600:
                    return r
        except Exception:
            pass

    # Fallback: drill into child windows to find the deepest control
    if HAS_WIN32:
        try:
            hwnd = _get_deepest_child_hwnd(x, y)
            if hwnd:
                rect = wintypes.RECT()
                if ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect)):
                    w = rect.right - rect.left
                    h = rect.bottom - rect.top
                    if 0 < w < 800 and 0 < h < 600:
                        return (rect.left, rect.top, rect.right, rect.bottom)
        except Exception:
            pass

    return None


# ---------------------------------------------------------------------------
# Screenshot capture
# ---------------------------------------------------------------------------

def _get_hwnd_rect(hwnd) -> Optional[tuple]:
    """Get the bounding rect for a window handle, or None if invalid/minimized."""
    # Try DwmGetWindowAttribute for accurate bounds (accounts for shadows/DWM)
    rect = wintypes.RECT()
    try:
        dwmapi = ctypes.windll.dwmapi
        DWMWA_EXTENDED_FRAME_BOUNDS = 9
        hr = dwmapi.DwmGetWindowAttribute(
            hwnd, DWMWA_EXTENDED_FRAME_BOUNDS,
            ctypes.byref(rect), ctypes.sizeof(rect),
        )
        if hr == 0:
            w, h = rect.right - rect.left, rect.bottom - rect.top
            if w > 0 and h > 0:
                return (rect.left, rect.top, rect.right, rect.bottom)
    except Exception:
        pass

    # Fallback to GetWindowRect
    rect = wintypes.RECT()
    if ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect)):
        w, h = rect.right - rect.left, rect.bottom - rect.top
        if w > 0 and h > 0:
            return (rect.left, rect.top, rect.right, rect.bottom)

    return None


def get_foreground_window_rect(exclude_hwnd: Optional[int] = None) -> Optional[tuple]:
    """Get the bounding rectangle (left, top, right, bottom) of the foreground window.
    If exclude_hwnd is set and the foreground window matches it, try the window
    at the last known click position instead."""
    if not HAS_WIN32:
        return None

    try:
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        if not hwnd:
            return None

        # Skip our own (minimized) PSR window
        if exclude_hwnd and int(hwnd) == exclude_hwnd:
            return None

        return _get_hwnd_rect(hwnd)
    except Exception:
        pass

    return None


def capture_screenshot(highlight_pos: Optional[tuple] = None,
                       element_rect: Optional[tuple] = None,
                       fullscreen: bool = False,
                       exclude_hwnd: Optional[int] = None) -> bytes:
    """
    Capture a screenshot of the active/foreground window.
    Falls back to full screen if the window rect can't be determined.

    If element_rect is provided (left, top, right, bottom in screen coords),
    a red border is drawn around the element — matching the original Windows PSR.
    Otherwise falls back to a small red box around the click position.
    """
    window_rect = None if fullscreen else get_foreground_window_rect(exclude_hwnd)

    with mss.mss() as sct:
        if window_rect:
            left, top, right, bottom = window_rect
            # Clamp to screen bounds
            monitor_info = sct.monitors[0]
            left = max(left, monitor_info["left"])
            top = max(top, monitor_info["top"])
            right = min(right, monitor_info["left"] + monitor_info["width"])
            bottom = min(bottom, monitor_info["top"] + monitor_info["height"])

            region = {
                "left": left,
                "top": top,
                "width": right - left,
                "height": bottom - top,
            }
            raw = sct.grab(region)
        else:
            # Fallback: full screen
            raw = sct.grab(sct.monitors[0])
            left, top = 0, 0

        img = Image.frombytes("RGB", raw.size, raw.bgra, "raw", "BGRX")

    draw = ImageDraw.Draw(img)

    if element_rect:
        # Draw a red border around the clicked element (like original PSR)
        el, et, er, eb = element_rect
        # Convert screen coords to image-relative coords
        rx1 = el - left
        ry1 = et - top
        rx2 = er - left
        ry2 = eb - top
        # Clamp to image bounds
        rx1 = max(0, rx1)
        ry1 = max(0, ry1)
        rx2 = min(img.width - 1, rx2)
        ry2 = min(img.height - 1, ry2)
        if rx2 > rx1 and ry2 > ry1:
            draw.rectangle([rx1, ry1, rx2, ry2], outline="red", width=3)
    elif highlight_pos:
        # Fallback: small red box centered on click position
        x = highlight_pos[0] - left
        y = highlight_pos[1] - top
        if 0 <= x <= img.width and 0 <= y <= img.height:
            r = 12
            draw.rectangle([x - r, y - r, x + r, y + r], outline="red", width=3)

    buf = BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Recorder
# ---------------------------------------------------------------------------

class StepsRecorder:
    def __init__(self, min_delay: float = 0.3, record_keyboard: bool = True,
                 fullscreen: bool = False,
                 stop_hotkey: str = "ctrl+shift+f9"):
        self.steps: list[Step] = []
        self.step_counter = 0
        self.min_delay = min_delay
        self.record_keyboard = record_keyboard
        self.fullscreen = fullscreen
        self.stop_hotkey = stop_hotkey
        self.last_capture_time = 0.0
        self.recording = False
        self.paused = False
        self._lock = threading.Lock()
        self._mouse_listener = None
        self._kb_listener = None
        self._stop_event = threading.Event()
        self._kb_buffer = []
        self._kb_timer = None
        self._kb_flush_delay = 1.0  # seconds before flushing kb buffer
        self._gui_hwnd = None  # HWND of the PSR GUI window (clicks on it are ignored)
        self._on_step_callback: Optional[Callable] = None
        # Build the set of key names that are part of the stop hotkey
        self._hotkey_keys = self._build_hotkey_keys(stop_hotkey)

    @staticmethod
    def _build_hotkey_keys(hotkey_str: str) -> set:
        """Expand a hotkey string into the set of individual key names to filter."""
        keys = set()
        for part in hotkey_str.lower().split("+"):
            p = part.strip()
            if p in ("ctrl", "control"):
                keys.update(("ctrl_l", "ctrl_r"))
            elif p == "shift":
                keys.update(("shift", "shift_l", "shift_r"))
            elif p == "alt":
                keys.update(("alt_l", "alt_r", "alt_gr"))
            else:
                keys.add(p)
        return keys

    # -- GUI integration ------------------------------------------------------

    def set_gui_hwnd(self, hwnd: int):
        """Set the HWND of the PSR window so clicks on it are ignored."""
        self._gui_hwnd = hwnd

    def _is_own_window_click(self, x: int, y: int) -> bool:
        """Check if a click landed on the PSR GUI window (should be ignored)."""
        if not HAS_WIN32 or not self._gui_hwnd:
            return False
        try:
            ix, iy = int(x), int(y)

            # Primary: geometry-based check — is the click inside our window rect?
            # This is the most reliable method (works for title bar, borders, etc.)
            rect = wintypes.RECT()
            if ctypes.windll.user32.GetWindowRect(
                wintypes.HWND(self._gui_hwnd), ctypes.byref(rect)
            ):
                if rect.left <= ix <= rect.right and rect.top <= iy <= rect.bottom:
                    return True

            # Secondary: HWND ancestry check
            point = wintypes.POINT(ix, iy)
            clicked_hwnd = ctypes.windll.user32.WindowFromPoint(point)
            if not clicked_hwnd:
                return False
            GA_ROOT = 2
            root_hwnd = ctypes.windll.user32.GetAncestor(clicked_hwnd, GA_ROOT)
            return int(root_hwnd) == self._gui_hwnd if root_hwnd else False
        except Exception:
            return False

    def _resolve_window_title(self, x: int, y: int, fallback: str) -> str:
        """If the foreground window is our own PSR window, get the title
        from the root window at the click coordinates instead."""
        try:
            fg = ctypes.windll.user32.GetForegroundWindow()
            if fg and int(fg) == self._gui_hwnd:
                point = wintypes.POINT(x, y)
                hwnd = ctypes.windll.user32.WindowFromPoint(point)
                if hwnd:
                    GA_ROOT = 2
                    root = ctypes.windll.user32.GetAncestor(hwnd, GA_ROOT)
                    target = root if root else hwnd
                    length = ctypes.windll.user32.GetWindowTextLengthW(target)
                    buf = ctypes.create_unicode_buffer(length + 1)
                    ctypes.windll.user32.GetWindowTextW(target, buf, length + 1)
                    if buf.value:
                        return buf.value
        except Exception:
            pass
        return fallback

    # -- Recording control --------------------------------------------------

    def start(self, blocking: bool = True):
        """Begin recording steps. If blocking=False, returns immediately."""
        self.recording = True
        self.paused = False
        self._stop_event.clear()
        hotkey_display = self.stop_hotkey.replace("+", "+").title()
        print("\n Recording started.")
        print("   Perform your steps — each click is captured.")
        print(f"   Press {hotkey_display} to stop recording.\n")

        self._mouse_listener = mouse.Listener(on_click=self._on_click)
        self._mouse_listener.start()

        if self.record_keyboard:
            self._kb_listener = keyboard.Listener(
                on_press=self._on_key_press,
            )
            self._kb_listener.start()

        if blocking:
            # Block until stop signal (CLI mode)
            self._stop_event.wait()

    def stop(self):
        """Stop recording."""
        if not self.recording:
            return
        self.recording = False
        self.paused = False
        self._flush_kb_buffer()

        if self._mouse_listener:
            self._mouse_listener.stop()
        if self._kb_listener:
            self._kb_listener.stop()

        self._stop_event.set()
        print(f"\n Recording stopped. {len(self.steps)} step(s) captured.")

    def pause(self):
        """Pause recording (stop capturing without ending session)."""
        self.paused = True

    def resume(self):
        """Resume a paused recording."""
        self.paused = False

    # -- Mouse handler ------------------------------------------------------

    def _on_click(self, x, y, button, pressed):
        if not pressed or not self.recording or self.paused:
            return  # only capture on press, not release; skip if paused

        # Ignore clicks on the PSR GUI window itself
        if self._is_own_window_click(x, y):
            return

        now = time.time()
        if now - self.last_capture_time < self.min_delay:
            return  # debounce
        self.last_capture_time = now

        # Flush pending keyboard buffer before this click
        self._flush_kb_buffer()

        action = "Left Click" if button == mouse.Button.left else \
                 "Right Click" if button == mouse.Button.right else \
                 "Middle Click"

        # Delay so the clicked window gains focus before we query title/screenshot
        time.sleep(0.15)

        window_title = get_active_window_title()

        # If PSR is still the foreground window (always-on-top), get the title
        # from the window at the click point instead — that's what the user
        # actually clicked.
        if self._gui_hwnd and HAS_WIN32:
            window_title = self._resolve_window_title(
                int(x), int(y), window_title
            )

        ui_elem = get_ui_element_at(int(x), int(y))
        elem_rect = get_element_rect_at(int(x), int(y))

        screenshot = capture_screenshot(
            highlight_pos=(int(x), int(y)),
            element_rect=elem_rect,
            fullscreen=self.fullscreen,
            exclude_hwnd=self._gui_hwnd,
        )

        with self._lock:
            self.step_counter += 1
            step = Step(
                step_number=self.step_counter,
                timestamp=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                action=action,
                window_title=window_title,
                position=(int(x), int(y)),
                ui_element=ui_elem if ui_elem else None,
                screenshot_png=screenshot,
                details=f"{action} at ({x}, {y})"
                         + (f" on {ui_elem}" if ui_elem else ""),
            )
            self.steps.append(step)
            print(f"  Step {step.step_number}: {step.action} in \"{step.window_title}\""
                  f" at ({x},{y})")

        if self._on_step_callback:
            try:
                self._on_step_callback(step)
            except Exception:
                pass

    # -- Keyboard handler ---------------------------------------------------

    def _on_key_press(self, key):
        if not self.recording or self.paused:
            return

        try:
            char = key.char
        except AttributeError:
            char = str(key).replace("Key.", "")

        # Don't buffer keys that are part of the stop hotkey combo
        if char.lower() in self._hotkey_keys:
            return

        with self._lock:
            self._kb_buffer.append(char)

        # Reset the flush timer
        if self._kb_timer:
            self._kb_timer.cancel()
        self._kb_timer = threading.Timer(self._kb_flush_delay, self._flush_kb_buffer)
        self._kb_timer.start()

    def _flush_kb_buffer(self):
        if not self.recording:
            return

        with self._lock:
            if not self._kb_buffer:
                return
            keys = self._kb_buffer.copy()
            self._kb_buffer.clear()

        # Mask the actual characters for privacy (like original PSR)
        typed_summary = f"Typed {len(keys)} character(s)"
        window_title = get_active_window_title()

        screenshot = capture_screenshot(fullscreen=self.fullscreen,
                                        exclude_hwnd=self._gui_hwnd)

        with self._lock:
            self.step_counter += 1
            step = Step(
                step_number=self.step_counter,
                timestamp=datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                action="Keyboard Input",
                window_title=window_title,
                screenshot_png=screenshot,
                details=typed_summary,
            )
            self.steps.append(step)
            print(f"  Step {step.step_number}: Keyboard input in \"{step.window_title}\"")

        if self._on_step_callback:
            try:
                self._on_step_callback(step)
            except Exception:
                pass

    # -- Report generation --------------------------------------------------

    def generate_html_report(self, output_path: str):
        """Generate an HTML report similar to original PSR output."""
        html = self._build_html()
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"📄 Report saved: {output_path}")

    def generate_zip_report(self, output_path: str):
        """Generate a ZIP containing the HTML report and individual screenshots."""
        html = self._build_html(embed_images=False, image_prefix="screenshots/")

        with zipfile.ZipFile(output_path, "w", zipfile.ZIP_DEFLATED) as zf:
            zf.writestr("report.html", html)
            for step in self.steps:
                if step.screenshot_png:
                    zf.writestr(
                        f"screenshots/step_{step.step_number:03d}.png",
                        step.screenshot_png,
                    )
        print(f"📦 ZIP report saved: {output_path}")

    def _build_html(self, embed_images: bool = True, image_prefix: str = "") -> str:
        """Build the HTML report content."""
        now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        steps_html = []
        for step in self.steps:
            img_tag = ""
            if step.screenshot_png:
                if embed_images:
                    b64 = base64.b64encode(step.screenshot_png).decode("ascii")
                    img_tag = f'<img src="data:image/png;base64,{b64}" alt="Step {step.step_number}" />'
                else:
                    img_tag = f'<img src="{image_prefix}step_{step.step_number:03d}.png" alt="Step {step.step_number}" />'

            pos_info = f" at ({step.position[0]}, {step.position[1]})" if step.position else ""
            ui_info = f'<div class="ui-element">UI Element: {step.ui_element}</div>' if step.ui_element else ""

            steps_html.append(f"""
            <div class="step">
                <div class="step-header">
                    <span class="step-num">Step {step.step_number}</span>
                    <span class="step-time">{step.timestamp}</span>
                </div>
                <div class="step-action">
                    <strong>{step.action}</strong>{pos_info}
                    — Window: <em>"{step.window_title}"</em>
                </div>
                {ui_info}
                {f'<div class="step-details">{step.details}</div>' if step.details else ""}
                <div class="screenshot">{img_tag}</div>
            </div>
            """)

        return f"""<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="utf-8" />
<title>Problem Steps Recorder — Report</title>
<style>
    * {{ margin: 0; padding: 0; box-sizing: border-box; }}
    body {{
        font-family: "Segoe UI", Tahoma, Geneva, Verdana, sans-serif;
        background: #1a1a2e; color: #e0e0e0;
        padding: 2rem; line-height: 1.6;
    }}
    h1 {{
        text-align: center; color: #e94560;
        margin-bottom: 0.5rem; font-size: 1.8rem;
    }}
    .meta {{
        text-align: center; color: #888;
        margin-bottom: 2rem; font-size: 0.9rem;
    }}
    .step {{
        background: #16213e; border-radius: 12px;
        padding: 1.5rem; margin-bottom: 1.5rem;
        border-left: 4px solid #e94560;
        box-shadow: 0 2px 8px rgba(0,0,0,0.3);
    }}
    .step-header {{
        display: flex; justify-content: space-between;
        margin-bottom: 0.5rem;
    }}
    .step-num {{
        font-weight: bold; color: #e94560;
        font-size: 1.1rem;
    }}
    .step-time {{ color: #888; font-size: 0.85rem; }}
    .step-action {{ margin-bottom: 0.5rem; }}
    .step-action strong {{ color: #0f3460; color: #53d8fb; }}
    .step-action em {{ color: #aaa; }}
    .ui-element {{
        font-size: 0.85rem; color: #a0a0a0;
        margin-bottom: 0.4rem; font-style: italic;
    }}
    .step-details {{
        font-size: 0.85rem; color: #999;
        margin-bottom: 0.75rem;
    }}
    .screenshot img {{
        max-width: 100%; border-radius: 8px;
        border: 1px solid #333; margin-top: 0.5rem;
        cursor: pointer; transition: transform 0.2s;
    }}
    .screenshot img:hover {{ transform: scale(1.02); }}
    .summary {{
        background: #16213e; border-radius: 12px;
        padding: 1.5rem; margin-bottom: 2rem;
        border: 1px solid #333;
    }}
    .summary h2 {{ color: #53d8fb; margin-bottom: 0.5rem; }}
</style>
</head>
<body>
<h1>🔴 Problem Steps Recorder</h1>
<div class="meta">Generated: {now} — {len(self.steps)} step(s) recorded</div>

<div class="summary">
    <h2>Recording Summary</h2>
    <p>This report contains {len(self.steps)} recorded step(s).
       Click on any screenshot to view it at full size.</p>
    <p>Steps include mouse clicks and keyboard input with corresponding
       screenshots of the screen state at each action.</p>
</div>

{"".join(steps_html)}

<script>
    // Lightbox overlay for screenshot viewing
    (function() {{
        const overlay = document.createElement('div');
        overlay.id = 'lightbox';
        overlay.innerHTML = `
            <div class="lb-close">&times;</div>
            <img src="" alt="Screenshot" />
        `;
        overlay.style.cssText = `
            display:none; position:fixed; inset:0; z-index:9999;
            background:rgba(0,0,0,0.92); cursor:pointer;
            justify-content:center; align-items:center;
        `;
        const closeBtn = overlay.querySelector('.lb-close');
        closeBtn.style.cssText = `
            position:absolute; top:1rem; right:1.5rem;
            font-size:2.5rem; color:#fff; cursor:pointer;
            line-height:1; font-weight:bold; z-index:10000;
            opacity:0.8; transition:opacity 0.2s;
        `;
        closeBtn.onmouseenter = () => closeBtn.style.opacity = '1';
        closeBtn.onmouseleave = () => closeBtn.style.opacity = '0.8';

        const lbImg = overlay.querySelector('img');
        lbImg.style.cssText = `
            max-width:95vw; max-height:92vh;
            border-radius:8px; box-shadow:0 4px 30px rgba(0,0,0,0.6);
            cursor:default; object-fit:contain;
        `;

        document.body.appendChild(overlay);

        function openLightbox(src) {{
            lbImg.src = src;
            overlay.style.display = 'flex';
            document.body.style.overflow = 'hidden';
        }}
        function closeLightbox() {{
            overlay.style.display = 'none';
            lbImg.src = '';
            document.body.style.overflow = '';
        }}

        overlay.addEventListener('click', (e) => {{
            if (e.target !== lbImg) closeLightbox();
        }});
        document.addEventListener('keydown', (e) => {{
            if (e.key === 'Escape') closeLightbox();
        }});

        document.querySelectorAll('.screenshot img').forEach(img => {{
            img.style.cursor = 'pointer';
            img.addEventListener('click', () => openLightbox(img.src));
        }});
    }})();
</script>
</body>
</html>"""


# ---------------------------------------------------------------------------
# Settings Dialog
# ---------------------------------------------------------------------------

class SettingsDialog(tk.Toplevel):
    """Settings dialog matching the original PSR settings window."""

    def __init__(self, parent, settings: dict):
        super().__init__(parent)
        self.title("Settings")
        self.resizable(False, False)
        self.transient(parent)
        self.grab_set()
        self.configure(bg="#2b2b2b")

        self.result = None

        pad = {"padx": 10, "pady": 4}

        style = ttk.Style(self)
        style.configure("Dlg.TLabel", background="#2b2b2b", foreground="#cccccc",
                         font=("Segoe UI", 9))
        style.configure("Dlg.TCheckbutton", background="#2b2b2b", foreground="#cccccc",
                         font=("Segoe UI", 9))
        style.configure("Dlg.TRadiobutton", background="#2b2b2b", foreground="#cccccc",
                         font=("Segoe UI", 9))
        style.configure("Dlg.TFrame", background="#2b2b2b")

        row = 0
        ttk.Label(self, text="Output File:", style="Dlg.TLabel").grid(
            row=row, column=0, sticky="w", **pad)
        self.output_var = tk.StringVar(value=settings.get("output", "steps_report"))
        ttk.Entry(self, textvariable=self.output_var, width=30).grid(
            row=row, column=1, **pad)
        ttk.Button(self, text="Browse...", command=self._browse).grid(
            row=row, column=2, **pad)

        row += 1
        ttk.Label(self, text="Capture Delay (s):", style="Dlg.TLabel").grid(
            row=row, column=0, sticky="w", **pad)
        self.delay_var = tk.StringVar(value=str(settings.get("delay", 0.3)))
        ttk.Entry(self, textvariable=self.delay_var, width=10).grid(
            row=row, column=1, sticky="w", **pad)

        row += 1
        self.kb_var = tk.BooleanVar(value=settings.get("record_keyboard", True))
        ttk.Checkbutton(self, text="Record keyboard input",
                         variable=self.kb_var, style="Dlg.TCheckbutton").grid(
            row=row, column=0, columnspan=3, sticky="w", **pad)

        row += 1
        self.fs_var = tk.BooleanVar(value=settings.get("fullscreen", False))
        ttk.Checkbutton(self, text="Capture full desktop (instead of active window)",
                         variable=self.fs_var, style="Dlg.TCheckbutton").grid(
            row=row, column=0, columnspan=3, sticky="w", **pad)

        row += 1
        ttk.Label(self, text="Output Format:", style="Dlg.TLabel").grid(
            row=row, column=0, sticky="w", **pad)
        self.format_var = tk.StringVar(value=settings.get("format", "html"))
        fmt_frame = ttk.Frame(self, style="Dlg.TFrame")
        fmt_frame.grid(row=row, column=1, columnspan=2, sticky="w", **pad)
        for val, label in [("html", "HTML"), ("zip", "ZIP"), ("both", "Both")]:
            ttk.Radiobutton(fmt_frame, text=label, variable=self.format_var,
                            value=val, style="Dlg.TRadiobutton").pack(
                side="left", padx=4)

        row += 1
        ttk.Label(self, text="Stop Hotkey:", style="Dlg.TLabel").grid(
            row=row, column=0, sticky="w", **pad)
        self.hotkey_var = tk.StringVar(
            value=settings.get("stop_hotkey", DEFAULT_SETTINGS["stop_hotkey"]))
        self._hotkey_entry = ttk.Entry(self, textvariable=self.hotkey_var, width=20)
        self._hotkey_entry.grid(row=row, column=1, sticky="w", **pad)
        self._hotkey_btn = ttk.Button(
            self, text="Capture...", command=self._capture_hotkey)
        self._hotkey_btn.grid(row=row, column=2, **pad)

        row += 1
        btn_frame = ttk.Frame(self, style="Dlg.TFrame")
        btn_frame.grid(row=row, column=0, columnspan=3, pady=10)
        ttk.Button(btn_frame, text="OK", command=self._ok, width=10).pack(
            side="left", padx=4)
        ttk.Button(btn_frame, text="Cancel", command=self.destroy, width=10).pack(
            side="left", padx=4)

        self.update_idletasks()
        # Position below parent window
        x = parent.winfo_x() + (parent.winfo_width() - self.winfo_width()) // 2
        y = parent.winfo_y() + parent.winfo_height() + 4
        self.geometry(f"+{x}+{y}")

    def _browse(self):
        path = filedialog.asksaveasfilename(
            defaultextension=".html",
            filetypes=[("HTML files", "*.html"), ("ZIP files", "*.zip"),
                       ("All files", "*.*")],
            initialfile=self.output_var.get(),
        )
        if path:
            self.output_var.set(Path(path).stem)

    def _capture_hotkey(self):
        """Open a small dialog that captures a key combination."""
        self._hotkey_btn.configure(text="Press keys...", state="disabled")
        captured = {"modifiers": set(), "key": None}

        def on_press(key):
            try:
                name = key.char
            except AttributeError:
                name = str(key).replace("Key.", "")
            low = name.lower() if name else ""

            if low in ("ctrl_l", "ctrl_r"):
                captured["modifiers"].add("ctrl")
            elif low in ("shift", "shift_l", "shift_r"):
                captured["modifiers"].add("shift")
            elif low in ("alt_l", "alt_r", "alt_gr"):
                captured["modifiers"].add("alt")
            else:
                captured["key"] = low

            if captured["key"]:
                parts = sorted(captured["modifiers"]) + [captured["key"]]
                self.hotkey_var.set("+".join(parts))
                self._hotkey_btn.configure(text="Capture...", state="normal")
                return False  # stop listener

        listener = keyboard.Listener(on_press=on_press)
        listener.daemon = True
        listener.start()

    def _ok(self):
        try:
            delay = float(self.delay_var.get())
            if delay < 0.05:
                delay = 0.05
        except ValueError:
            delay = 0.3

        self.result = {
            "output": self.output_var.get(),
            "delay": delay,
            "record_keyboard": self.kb_var.get(),
            "fullscreen": self.fs_var.get(),
            "format": self.format_var.get(),
            "stop_hotkey": self.hotkey_var.get().strip().lower(),
        }
        self.destroy()


# ---------------------------------------------------------------------------
# PSR GUI Application
# ---------------------------------------------------------------------------

class PSRApp:
    """
    Tkinter GUI matching the Windows Problem Steps Recorder toolbar.
    The window is always-on-top and its clicks are invisible to the recorder.
    """

    def __init__(self, settings: dict):
        self.settings = settings
        self.recorder = None
        self._record_thread = None

        # -- Window setup ---------------------------------------------------
        self.root = tk.Tk()
        self.root.title("Problem Steps Recorder")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

        self.root.configure(bg="#2b2b2b")
        self._set_icon()
        self._setup_style()
        self._build_ui()

        # Get the native HWND for click filtering (Windows only)
        self._hwnd = None
        if HAS_WIN32:
            self.root.update_idletasks()
            self._hwnd = self._get_root_hwnd()

    def _set_icon(self):
        """Set the window icon from psr.ico if available."""
        # Look for ico next to the script / exe
        for base in [Path(__file__).parent, Path(sys.executable).parent]:
            ico = base / "psr.ico"
            if ico.exists():
                try:
                    self.root.iconbitmap(str(ico))
                    return
                except Exception:
                    pass

    def _get_root_hwnd(self) -> Optional[int]:
        """Get the top-level HWND of the tkinter window."""
        try:
            frame_hwnd = self.root.winfo_id()
            GA_ROOT = 2
            root_hwnd = ctypes.windll.user32.GetAncestor(frame_hwnd, GA_ROOT)
            # Normalize to plain int for reliable comparisons
            hwnd = root_hwnd if root_hwnd else frame_hwnd
            return int(hwnd) if hwnd else None
        except Exception:
            return None

    def _setup_style(self):
        """Configure ttk style to approximate the Windows PSR look."""
        style = ttk.Style(self.root)
        style.theme_use("clam")

        style.configure("Toolbar.TFrame", background="#2b2b2b")
        style.configure("Status.TFrame", background="#333333")

        style.configure("Toolbar.TButton", padding=(10, 4),
                         font=("Segoe UI", 9))
        style.configure("Record.TButton", padding=(10, 4),
                         font=("Segoe UI", 9, "bold"))

        style.configure("Status.TLabel", background="#333333",
                         foreground="#cccccc", font=("Segoe UI", 9),
                         padding=(8, 4))

    def _build_ui(self):
        """Build the compact toolbar interface."""
        # -- Button toolbar -------------------------------------------------
        toolbar = ttk.Frame(self.root, style="Toolbar.TFrame")
        toolbar.pack(fill="x", padx=6, pady=(6, 2))

        self.btn_start = ttk.Button(
            toolbar, text="\u25cf Start Record",
            style="Record.TButton", command=self._start_recording)
        self.btn_start.pack(side="left", padx=(0, 3))

        self.btn_pause = ttk.Button(
            toolbar, text="\u23f8 Pause Record",
            style="Toolbar.TButton", command=self._toggle_pause,
            state="disabled")
        self.btn_pause.pack(side="left", padx=3)

        self.btn_stop = ttk.Button(
            toolbar, text="\u25a0 Stop Record",
            style="Toolbar.TButton", command=self._stop_recording,
            state="disabled")
        self.btn_stop.pack(side="left", padx=3)

        self.btn_settings = ttk.Button(
            toolbar, text="\u2699 Settings",
            style="Toolbar.TButton", command=self._open_settings)
        self.btn_settings.pack(side="left", padx=(3, 0))

        # -- Status bar -----------------------------------------------------
        status_frame = ttk.Frame(self.root, style="Status.TFrame")
        status_frame.pack(fill="x", padx=6, pady=(2, 6))

        self.status_label = ttk.Label(
            status_frame,
            text="Ready. Click 'Start Record' to begin.",
            style="Status.TLabel")
        self.status_label.pack(fill="x")

    # -- Recording controls ------------------------------------------------

    def _start_recording(self):
        """Start a new recording session."""
        self.recorder = StepsRecorder(
            min_delay=self.settings.get("delay", 0.3),
            record_keyboard=self.settings.get("record_keyboard", True),
            fullscreen=self.settings.get("fullscreen", False),
            stop_hotkey=self.settings.get("stop_hotkey", DEFAULT_SETTINGS["stop_hotkey"]),
        )
        self.recorder._on_step_callback = self._on_step_recorded

        # Pass our HWND so the recorder ignores clicks on this window
        if self._hwnd:
            self.recorder.set_gui_hwnd(self._hwnd)

        # Setup the global stop hotkey
        hotkey = self.settings.get("stop_hotkey", DEFAULT_SETTINGS["stop_hotkey"])
        setup_stop_hotkey(self.recorder, on_stop_callback=self._on_hotkey_stop,
                          hotkey_str=hotkey)

        # Update UI state
        self.btn_start.configure(state="disabled")
        self.btn_pause.configure(state="normal")
        self.btn_stop.configure(state="normal")
        self.btn_settings.configure(state="disabled")
        self.status_label.configure(text="\u25cf Recording... Step 0 captured")

        # Start recording in background thread (non-blocking)
        self._record_thread = threading.Thread(
            target=self.recorder.start, kwargs={"blocking": True},
            daemon=True)
        self._record_thread.start()

        # Minimize PSR so it's out of the way while recording
        self.root.iconify()

    def _stop_recording(self):
        """Stop recording and save the report."""
        if not self.recorder or not self.recorder.recording:
            return
        self.recorder.stop()
        self._finalize()

    def _toggle_pause(self):
        """Toggle pause/resume."""
        if not self.recorder or not self.recorder.recording:
            return

        if self.recorder.paused:
            self.recorder.resume()
            self.btn_pause.configure(text="\u23f8 Pause Record")
            n = len(self.recorder.steps)
            self.status_label.configure(
                text=f"\u25cf Recording... Step {n} captured")
        else:
            self.recorder.pause()
            self.btn_pause.configure(text="\u25b6 Resume Record")
            self.status_label.configure(text="\u23f8 Recording paused")

    def _on_step_recorded(self, step: Step):
        """Callback from recorder thread when a new step is captured."""
        try:
            self.root.after_idle(self._update_step_count)
        except Exception:
            pass

    def _update_step_count(self):
        """Update the status bar with current step count."""
        if self.recorder and self.recorder.recording:
            n = len(self.recorder.steps)
            self.status_label.configure(
                text=f"\u25cf Recording... Step {n} captured")

    def _on_hotkey_stop(self):
        """Called from the hotkey listener thread when Ctrl+Shift+F9 is pressed."""
        try:
            self.root.after_idle(self._finalize)
        except Exception:
            pass

    def _finalize(self):
        """Save reports and reset UI after recording stops."""
        # Restore from taskbar so the user sees results
        self.root.deiconify()
        self.root.lift()

        self.btn_start.configure(state="normal")
        self.btn_pause.configure(state="disabled", text="\u23f8 Pause Record")
        self.btn_stop.configure(state="disabled")
        self.btn_settings.configure(state="normal")

        if not self.recorder or not self.recorder.steps:
            self.status_label.configure(text="No steps recorded.")
            return

        n = len(self.recorder.steps)
        output_name = self.settings.get("output", "steps_report")
        fmt = self.settings.get("format", "html")

        saved_files = []
        if fmt in ("html", "both"):
            path = f"{output_name}.html"
            self.recorder.generate_html_report(path)
            saved_files.append(path)
        if fmt in ("zip", "both"):
            path = f"{output_name}.zip"
            self.recorder.generate_zip_report(path)
            saved_files.append(path)

        files_str = ", ".join(saved_files)
        self.status_label.configure(
            text=f"Done! {n} step(s) saved to {files_str}")

        if saved_files:
            self.root.after(200, lambda: self._offer_open(saved_files[0]))

    def _offer_open(self, path: str):
        """Ask user if they want to open the report."""
        if messagebox.askyesno(
                "Recording Complete",
                f"{len(self.recorder.steps)} step(s) recorded.\n\n"
                f"Open {path}?"):
            try:
                os.startfile(os.path.abspath(path))
            except Exception:
                pass

    # -- Settings -----------------------------------------------------------

    def _open_settings(self):
        """Open the settings dialog."""
        dlg = SettingsDialog(self.root, self.settings)
        self.root.wait_window(dlg)
        if dlg.result:
            self.settings.update(dlg.result)
            save_settings(self.settings)

    # -- Window lifecycle ---------------------------------------------------

    def _on_close(self):
        """Handle window close — stop recording first if active."""
        if self.recorder and self.recorder.recording:
            self.recorder.stop()
        self.root.destroy()

    def run(self):
        """Start the tkinter main loop."""
        self.root.update_idletasks()
        w = self.root.winfo_width()
        sw = self.root.winfo_screenwidth()
        x = (sw - w) // 2
        self.root.geometry(f"+{x}+40")
        self.root.mainloop()


# ---------------------------------------------------------------------------
# Global hotkey to stop recording
# ---------------------------------------------------------------------------

def _parse_hotkey(hotkey_str: str) -> tuple:
    """Parse a hotkey string like 'ctrl+shift+f9' into (modifiers_set, key_name)."""
    parts = [p.strip().lower() for p in hotkey_str.split("+") if p.strip()]
    modifiers = set()
    key_name = None
    for p in parts:
        if p in ("ctrl", "control"):
            modifiers.add("ctrl")
        elif p in ("shift",):
            modifiers.add("shift")
        elif p in ("alt",):
            modifiers.add("alt")
        else:
            key_name = p
    return modifiers, key_name


def _key_to_name(key) -> str:
    """Convert a pynput key to a comparable name string."""
    try:
        return key.char.lower() if key.char else ""
    except AttributeError:
        return str(key).replace("Key.", "").lower()


def setup_stop_hotkey(recorder: StepsRecorder, on_stop_callback=None,
                      hotkey_str: str = "ctrl+shift+f9"):
    """Listen for a configurable hotkey to stop recording."""
    required_mods, required_key = _parse_hotkey(hotkey_str)
    active_mods = set()

    def on_press(key):
        name = _key_to_name(key)
        if name in ("ctrl_l", "ctrl_r"):
            active_mods.add("ctrl")
        elif name in ("shift", "shift_l", "shift_r"):
            active_mods.add("shift")
        elif name in ("alt_l", "alt_r", "alt_gr"):
            active_mods.add("alt")

        if name == required_key and required_mods.issubset(active_mods):
            recorder.stop()
            if on_stop_callback:
                on_stop_callback()
            return False  # stop this listener

    def on_release(key):
        name = _key_to_name(key)
        if name in ("ctrl_l", "ctrl_r"):
            active_mods.discard("ctrl")
        elif name in ("shift", "shift_l", "shift_r"):
            active_mods.discard("shift")
        elif name in ("alt_l", "alt_r", "alt_gr"):
            active_mods.discard("alt")

    listener = keyboard.Listener(on_press=on_press, on_release=on_release)
    listener.daemon = True
    listener.start()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    # Load persisted settings as defaults
    saved = load_settings()

    parser = argparse.ArgumentParser(
        description="PSR Clone — Problem Steps Recorder for Python",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=f"Press {saved.get('stop_hotkey', 'Ctrl+Shift+F9').replace('+', '+')} to stop recording.",
    )
    parser.add_argument("--output", "-o", default=None,
                        help="Output filename (without extension)")
    parser.add_argument("--delay", "-d", type=float, default=None,
                        help="Minimum delay between captures in seconds")
    parser.add_argument("--no-keyboard", action="store_true",
                        help="Don't record keyboard actions")
    parser.add_argument("--zip", action="store_true",
                        help="Also generate a ZIP archive with separate screenshots")
    parser.add_argument("--fullscreen", action="store_true",
                        help="Capture full desktop instead of just the active window")
    parser.add_argument("--format", choices=["html", "zip", "both"], default=None,
                        help="Output format (default: html)")
    parser.add_argument("--no-gui", action="store_true",
                        help="Run in CLI-only mode (no GUI window)")
    args = parser.parse_args()

    # Merge: CLI args override saved settings, saved settings override defaults
    settings = dict(saved)
    if args.output is not None:
        settings["output"] = args.output
    if args.delay is not None:
        settings["delay"] = args.delay
    if args.no_keyboard:
        settings["record_keyboard"] = False
    if args.fullscreen:
        settings["fullscreen"] = True
    if args.format is not None:
        settings["format"] = args.format
    if args.zip:
        settings["format"] = "both"

    if args.no_gui:
        # Original CLI-only mode
        hotkey = settings.get("stop_hotkey", DEFAULT_SETTINGS["stop_hotkey"])
        recorder = StepsRecorder(
            min_delay=settings["delay"],
            record_keyboard=settings["record_keyboard"],
            fullscreen=settings["fullscreen"],
            stop_hotkey=hotkey,
        )
        setup_stop_hotkey(recorder, hotkey_str=hotkey)

        try:
            recorder.start(blocking=True)
        except KeyboardInterrupt:
            recorder.stop()

        if not recorder.steps:
            print("No steps recorded.")
            return

        fmt = settings["format"]
        if fmt in ("html", "both"):
            recorder.generate_html_report(f"{settings['output']}.html")
        if fmt in ("zip", "both"):
            recorder.generate_zip_report(f"{settings['output']}.zip")
    else:
        # GUI mode (default)
        app = PSRApp(settings)
        app.run()


if __name__ == "__main__":
    main()