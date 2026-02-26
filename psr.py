"""
PSR (Problem Steps Recorder) Clone for Windows
================================================
Recreates the classic Windows PSR functionality:
- Captures a screenshot on every mouse click
- Logs the active window title and click coordinates
- Detects UI element under cursor when possible
- Generates an HTML report (like the original PSR .mht files)
- Can also export as a ZIP archive

Requirements:
    pip install pynput mss Pillow pygetwindow

Optional (for richer UI element detection):
    pip install pywinauto comtypes

Usage:
    python psr.py                  # Start recording (Ctrl+Shift+F9 to stop)
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
from io import BytesIO
from pathlib import Path
from dataclasses import dataclass, field
from typing import Optional

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
    except ImportError:
        pass

    try:
        from pywinauto import Desktop
        HAS_WINAUTO = True
    except ImportError:
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


# ---------------------------------------------------------------------------
# Screenshot capture
# ---------------------------------------------------------------------------

def get_foreground_window_rect() -> Optional[tuple]:
    """Get the bounding rectangle (left, top, right, bottom) of the foreground window."""
    if not HAS_WIN32:
        return None

    try:
        hwnd = ctypes.windll.user32.GetForegroundWindow()
        if not hwnd:
            return None

        # Use DwmGetWindowAttribute for accurate bounds (accounts for shadows/DWM)
        # DWMWA_EXTENDED_FRAME_BOUNDS = 9
        rect = wintypes.RECT()
        try:
            dwmapi = ctypes.windll.dwmapi
            DWMWA_EXTENDED_FRAME_BOUNDS = 9
            hr = dwmapi.DwmGetWindowAttribute(
                hwnd,
                DWMWA_EXTENDED_FRAME_BOUNDS,
                ctypes.byref(rect),
                ctypes.sizeof(rect),
            )
            if hr == 0:  # S_OK
                return (rect.left, rect.top, rect.right, rect.bottom)
        except Exception:
            pass

        # Fallback to GetWindowRect
        rect = wintypes.RECT()
        if ctypes.windll.user32.GetWindowRect(hwnd, ctypes.byref(rect)):
            return (rect.left, rect.top, rect.right, rect.bottom)
    except Exception:
        pass

    return None


def capture_screenshot(highlight_pos: Optional[tuple] = None, fullscreen: bool = False) -> bytes:
    """
    Capture a screenshot of the active/foreground window.
    Falls back to full screen if the window rect can't be determined.
    The highlight position is adjusted to be relative to the window.
    """
    window_rect = None if fullscreen else get_foreground_window_rect()

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

    # Draw a highlight circle at the click position (adjusted to window-relative coords)
    if highlight_pos:
        draw = ImageDraw.Draw(img)
        x = highlight_pos[0] - left
        y = highlight_pos[1] - top

        # Only draw if the click is within the captured image
        if 0 <= x <= img.width and 0 <= y <= img.height:
            r = 18
            # Red circle with border
            draw.ellipse([x - r, y - r, x + r, y + r], outline="red", width=3)
            # Inner crosshair
            draw.line([x - r + 4, y, x + r - 4, y], fill="red", width=2)
            draw.line([x, y - r + 4, x, y + r - 4], fill="red", width=2)

    buf = BytesIO()
    img.save(buf, format="PNG", optimize=True)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Recorder
# ---------------------------------------------------------------------------

class StepsRecorder:
    def __init__(self, min_delay: float = 0.3, record_keyboard: bool = True, fullscreen: bool = False):
        self.steps: list[Step] = []
        self.step_counter = 0
        self.min_delay = min_delay
        self.record_keyboard = record_keyboard
        self.fullscreen = fullscreen
        self.last_capture_time = 0.0
        self.recording = False
        self._lock = threading.Lock()
        self._mouse_listener = None
        self._kb_listener = None
        self._stop_event = threading.Event()
        self._kb_buffer = []
        self._kb_timer = None
        self._kb_flush_delay = 1.0  # seconds before flushing kb buffer

    # -- Recording control --------------------------------------------------

    def start(self):
        """Begin recording steps."""
        self.recording = True
        self._stop_event.clear()
        print("\n🔴 Recording started.")
        print("   Perform your steps — each click is captured.")
        print("   Press Ctrl+Shift+F9 to stop recording.\n")

        self._mouse_listener = mouse.Listener(on_click=self._on_click)
        self._mouse_listener.start()

        if self.record_keyboard:
            self._kb_listener = keyboard.Listener(
                on_press=self._on_key_press,
            )
            self._kb_listener.start()

        # Block until stop signal
        self._stop_event.wait()

    def stop(self):
        """Stop recording."""
        if not self.recording:
            return
        self.recording = False
        self._flush_kb_buffer()

        if self._mouse_listener:
            self._mouse_listener.stop()
        if self._kb_listener:
            self._kb_listener.stop()

        self._stop_event.set()
        print(f"\n⏹ Recording stopped. {len(self.steps)} step(s) captured.")

    # -- Mouse handler ------------------------------------------------------

    def _on_click(self, x, y, button, pressed):
        if not pressed or not self.recording:
            return  # only capture on press, not release

        now = time.time()
        if now - self.last_capture_time < self.min_delay:
            return  # debounce
        self.last_capture_time = now

        # Flush pending keyboard buffer before this click
        self._flush_kb_buffer()

        action = "Left Click" if button == mouse.Button.left else \
                 "Right Click" if button == mouse.Button.right else \
                 "Middle Click"

        window_title = get_active_window_title()
        ui_elem = get_ui_element_at(int(x), int(y))

        # Small delay so the UI state reflects the click
        time.sleep(0.05)

        screenshot = capture_screenshot(highlight_pos=(int(x), int(y)), fullscreen=self.fullscreen)

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

    # -- Keyboard handler ---------------------------------------------------

    # Keys that are part of the stop hotkey — ignore these in the keyboard buffer
    _HOTKEY_KEYS = {"ctrl_l", "ctrl_r", "shift", "shift_l", "shift_r", "f9"}

    def _on_key_press(self, key):
        if not self.recording:
            return

        try:
            char = key.char
        except AttributeError:
            char = str(key).replace("Key.", "")

        # Don't buffer keys that are part of the stop hotkey combo
        if char.lower() in self._HOTKEY_KEYS:
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

        screenshot = capture_screenshot(fullscreen=self.fullscreen)

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
# Global hotkey to stop recording
# ---------------------------------------------------------------------------

def setup_stop_hotkey(recorder: StepsRecorder):
    """Listen for Ctrl+Shift+F9 to stop recording."""
    pressed_keys = set()

    def on_press(key):
        pressed_keys.add(key)
        if (keyboard.Key.ctrl_l in pressed_keys or keyboard.Key.ctrl_r in pressed_keys) \
                and keyboard.Key.shift in pressed_keys \
                and keyboard.Key.f9 in pressed_keys:
            recorder.stop()
            return False  # stop this listener

    def on_release(key):
        pressed_keys.discard(key)

    listener = keyboard.Listener(on_press=on_press, on_release=on_release)
    listener.daemon = True
    listener.start()


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="PSR Clone — Problem Steps Recorder for Python",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Press Ctrl+Shift+F9 to stop recording.",
    )
    parser.add_argument("--output", "-o", default="steps_report",
                        help="Output filename (without extension)")
    parser.add_argument("--delay", "-d", type=float, default=0.3,
                        help="Minimum delay between captures in seconds (default: 0.3)")
    parser.add_argument("--no-keyboard", action="store_true",
                        help="Don't record keyboard actions")
    parser.add_argument("--zip", action="store_true",
                        help="Also generate a ZIP archive with separate screenshots")
    parser.add_argument("--fullscreen", action="store_true",
                        help="Capture full desktop instead of just the active window")
    parser.add_argument("--format", choices=["html", "zip", "both"], default="html",
                        help="Output format (default: html)")
    args = parser.parse_args()

    recorder = StepsRecorder(
        min_delay=args.delay,
        record_keyboard=not args.no_keyboard,
        fullscreen=args.fullscreen,
    )

    # Setup the stop hotkey listener
    setup_stop_hotkey(recorder)

    try:
        recorder.start()
    except KeyboardInterrupt:
        recorder.stop()

    if not recorder.steps:
        print("No steps recorded.")
        return

    # Generate reports
    fmt = args.format
    if args.zip:
        fmt = "both"

    if fmt in ("html", "both"):
        recorder.generate_html_report(f"{args.output}.html")
    if fmt in ("zip", "both"):
        recorder.generate_zip_report(f"{args.output}.zip")


if __name__ == "__main__":
    main()