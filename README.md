# 🔴 PSR — Problem Steps Recorder

A Python recreation of the classic Windows Problem Steps Recorder (`psr.exe`) that Microsoft removed in favor of the Snipping Tool. This version captures a screenshot of the active window on every mouse click, logs what you did, and generates a clean HTML report you can share or archive.

![Python](https://img.shields.io/badge/Python-3.10%2B-blue)
![Platform](https://img.shields.io/badge/Platform-Crossplatform-0078D6)
![License](https://img.shields.io/badge/License-GPL--3.0-blue)

## Features

- **Automatic screenshot capture** on every mouse click just like the original PSR
- **Active window capture** screenshots are cropped to the foreground window, not the entire desktop
- **Click position highlight** a red crosshair circle marks exactly where you clicked
- **Window title logging** records which application was active for each step
- **UI element detection** identifies the control/element under the cursor (with optional `pywinauto`)
- **Keyboard input tracking** logs typing activity with character counts (content is masked for privacy)
- **HTML report generation** self-contained HTML file with embedded screenshots and a dark theme
- **Lightbox viewer** click any screenshot in the report to view it full-size; close with ×, click outside, or Escape
- **ZIP export** optionally bundle the report with separate screenshot PNGs
- **Configurable** adjust capture delay, output format, and more via CLI flags

## Grab your release from the release page

# [Releases](https://github.com/Oratorian/problem-steps-recorder/releases)

## Requirements ( For development or running as a Python script directly )

**Python 3.10+** on Windows.

### Core dependencies

```
pip install pynput mss Pillow pygetwindow
```

### Optional (recommended)

For richer UI element detection (control type, element name, class):

```
pip install pywinauto comtypes
```

## Usage

```bash
# Start recording — press Ctrl+Shift+F9 to stop
python psr.py

# Custom output filename
python psr.py --output my_report

# Adjust minimum delay between captures (default: 0.3s)
python psr.py --delay 0.5

# Don't record keyboard actions
python psr.py --no-keyboard

# Capture full desktop instead of just the active window
python psr.py --fullscreen

# Export as ZIP with separate screenshot files
python psr.py --format zip

# Export both HTML and ZIP
python psr.py --format both
```

### CLI Reference

| Flag | Short | Default | Description |
|------|-------|---------|-------------|
| `--output` | `-o` | `steps_report` | Output filename (without extension) |
| `--delay` | `-d` | `0.3` | Minimum seconds between captures (debounce) |
| `--no-keyboard` | | `false` | Disable keyboard input logging |
| `--fullscreen` | | `false` | Capture entire desktop instead of active window |
| `--format` | | `html` | Output format: `html`, `zip`, or `both` |
| `--zip` | | `false` | Shorthand for `--format both` |

### Stop Hotkey

Press **Ctrl+Shift+F9** at any time to stop recording. The report is generated automatically.

You can also press **Ctrl+C** in the terminal.

## Output

### HTML Report

A self-contained `.html` file with all screenshots embedded as base64. Open it in any browser - no server or extra files needed.

Each step shows:
- Step number and timestamp
- Action type (Left Click, Right Click, Keyboard Input)
- Active window title
- Click coordinates
- UI element info (when available)
- Screenshot with click position highlighted

### ZIP Archive

When using `--format zip` or `--format both`, a `.zip` is created containing:
- `report.html` — the report referencing external images
- `screenshots/step_001.png`, `step_002.png`, etc.

## How It Works

1. **Mouse listener** (`pynput`) watches for click events
2. On each click, the recorder:
   - Identifies the foreground window using Win32 API (`DwmGetWindowAttribute` for accurate bounds)
   - Captures just that window region using `mss`
   - Draws a red crosshair at the click position
   - Logs the window title and UI element under the cursor
3. **Keyboard listener** buffers keystrokes and flushes them as a single step after a 1-second pause
4. On stop, all steps are compiled into an HTML report with an embedded lightbox viewer

## Comparison with Original PSR

| Feature | Original PSR | This Clone |
|---------|-------------|------------|
| Screenshot on click | ✅ | ✅ |
| Active window capture | ✅ | ✅ |
| Click position highlight | ✅ (green border) | ✅ (red border) |
| Window title logging | ✅ | ✅ |
| UI element detection | ✅ (MSAA) | ✅ (UIA via pywinauto) |
| Keyboard logging | ✅ (masked) | ✅ (masked) |
| Output format | `.mht` (ZIP) | `.html` / `.zip` |
| Fullscreen option | ❌ | ✅ |
| Comment/annotation | ✅ | ❌ (planned) |
| Configurable delay | ❌ | ✅ |

## Troubleshooting

**"Window: (unknown)"** — If `pygetwindow` can't detect the active window, the recorder falls back gracefully. Installing on Windows with the Win32 API available (default with CPython) gives the best results.

**Screenshots are full desktop** — The active window detection requires Win32 APIs (`ctypes`). If you're running in an unusual environment, use `--fullscreen` and crop manually, or install `pywinauto` for better window detection.

**Double/rapid captures** — Increase the debounce delay: `--delay 0.5` or higher.

## License

[GNU GENERAL PUBLIC LICENSE (GPL-3.0)](LICENSE)
