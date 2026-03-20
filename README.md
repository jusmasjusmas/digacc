# Slide A11y Remediator

A local web app for remediating PowerPoint accessibility issues. Upload a `.pptx`, get WCAG 2.1 AA / Title II violations flagged and auto-fixed, review the rest slide-by-slide with live thumbnails, then download the remediated file.

Inspired by [Grackle Slides](https://www.grackledocs.com/) — runs entirely on your machine.

![Python](https://img.shields.io/badge/python-3.10%2B-blue) ![Flask](https://img.shields.io/badge/flask-3.x-lightgrey) ![License](https://img.shields.io/badge/license-MIT-green)

---

## Features

- **14 accessibility checks** covering structure, images, tables, content, and visual contrast
- **Auto-fix engine** — silently repairs language tags, trailing empty paragraphs, empty text boxes, merged table cells, and more
- **AI-powered** — uses Claude to generate alt text for images and titles for untitled slides (requires Anthropic API key)
- **Live slide thumbnails** via LibreOffice — previews update after every fix
- **Human-in-the-loop** — Accept / Edit / Skip each pending issue; bulk "Accept All" for quick review
- **Keyboard navigation** — `←` / `→` to move between slides
- **Dark theme** React UI, no build step required

---

## Setup

### 1. Install system dependencies

**macOS**
```bash
brew install libreoffice poppler
```

**Ubuntu/Debian**
```bash
sudo apt-get install libreoffice poppler-utils
```

### 2. Install Python dependencies

```bash
pip install -r requirements.txt
```

### 3. (Optional) Set your Anthropic API key

```bash
export ANTHROPIC_API_KEY=sk-ant-your-key-here
```

Without a key the app still runs — AI-generated alt text and slide titles are skipped and flagged for manual review instead. All other checks are unaffected.

### 4. Run

```bash
digacc                          # default: http://localhost:5001
digacc --open                   # start and open browser automatically
digacc --port 8080              # custom port
digacc --host 0.0.0.0           # expose on your LAN
digacc --api-key sk-ant-...     # set Anthropic key without exporting env var
digacc --no-debug               # disable Flask reloader
digacc --help                   # full usage
```

> **First-time setup:** add the shell function to your `~/.zshrc`:
> ```bash
> digacc() { python /path/to/digaccapp/app.py "$@"; }
> ```
> Or just run `python app.py [options]` directly from the project folder.

---

## Workflow

1. **Drop a `.pptx`** onto the upload zone (or click to browse)
2. The backend scans for all issues and silently auto-applies the "always auto" fixes
3. You land in a **three-panel workspace**:
   - **Left** — slide strip with issue-count badges
   - **Center** — full-size slide preview (updates live after each fix)
   - **Right** — issues panel with Pending / Auto-fixed / Done tabs
4. For each pending issue, choose:
   - **Accept** — apply the suggested fix as-is
   - **Edit** — modify the suggested value, then confirm
   - **Skip** — leave it for manual follow-up in PowerPoint
5. Use **Accept All** to batch-accept all auto-fixable issues at once
6. **Download** the remediated `.pptx` when done

---

## Checks

| Check | Category | Default |
|---|---|---|
| Presentation metadata title | Structure | Always auto |
| Language tags on all text runs | Structure | Always auto |
| Trailing empty paragraphs | Structure | Always auto |
| Empty text boxes | Structure | Always auto |
| Unmerge merged table cells | Tables | Always auto |
| Fill empty table cells with — | Tables | Always auto |
| Missing / empty slide titles (AI) | Structure | Auto (toggleable) |
| Duplicate slide titles | Structure | Auto (toggleable) |
| Broken list formatting | Content | Auto (toggleable) |
| Image alt text (AI-generated) | Images | Auto (toggleable) |
| Table descriptions (alt text) | Tables | Auto (toggleable) |
| Explicit low-contrast text colors | Visual | Auto (toggleable) |
| Fine print text (< 18pt) | Visual | Manual review only |
| Theme-inherited color contrast | Visual | Manual review only |

Toggleable checks can be switched off per-session in the Settings panel before upload or during review.

---

## Notes

- Sessions are in-memory — restarting the server clears them
- The original file is never modified; all edits go to a working copy
- Thumbnail generation runs LibreOffice headlessly — expect ~1 min for a 60-slide deck
- Tested on macOS with LibreOffice 26.x and Python 3.11
