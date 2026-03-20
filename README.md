# Slide A11y Remediator

Interactive web app for remediating PowerPoint accessibility issues. Checks for WCAG 2.1 AA compliance and Grackle Slides compatibility, with live slide previews and a human-in-the-loop approval workflow.

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

Without a key the app still runs — AI-generated alt text and slide titles are skipped, and those issues are flagged for manual review instead.

### 4. Run

```bash
python app.py
```

Then open **http://localhost:5000** in your browser.

---

## How it works

1. **Drop a .pptx** onto the upload zone
2. The backend scans for all accessibility issues and immediately auto-applies the ones configured to run silently
3. You're shown a **three-panel workspace**: slide strip → live preview → issues panel
4. **Pending issues** each have an Accept / Edit / Skip decision:
   - **Accept** — applies the suggested fix as-is
   - **Edit** — lets you change the suggested value before applying
   - **Skip** — marks it as skipped (you'll handle it in PowerPoint)
5. The slide preview **updates live** after each fix
6. **Download** the remediated .pptx when you're done

---

## Checks performed

| Check | Default |
|---|---|
| Presentation metadata title | Always auto |
| Language tags on all text runs | Always auto |
| Trailing empty paragraphs | Always auto |
| Empty text boxes | Always auto |
| Unmerge merged table cells | Always auto |
| Fill empty table cells with — | Always auto |
| Missing / empty slide titles (AI) | Auto (toggleable) |
| Duplicate slide titles | Auto (toggleable) |
| Broken list formatting | Auto (toggleable) |
| Image alt text (AI-generated) | Auto (toggleable) |
| Table descriptions | Auto (toggleable) |
| Explicit low-contrast text colors | Auto (toggleable) |
| Fine print text (< 18pt) | Manual only |
| Theme-inherited color contrast | Manual only |

---

## Notes

- Sessions are stored in memory — restarting the server clears them
- The original file is never modified; all changes go to a working copy
- For large decks (50+ slides), thumbnail generation may take 15–30 seconds
