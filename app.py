"""
PPTX Accessibility Remediator — Flask Backend

Usage:
  python app.py                        # default: localhost:5001
  python app.py --port 8080            # custom port
  python app.py --host 0.0.0.0         # expose to network
  python app.py --open                 # auto-open browser
  python app.py --api-key sk-ant-...   # set Anthropic key for this session
  python app.py --no-debug             # production-style (no reloader)
"""

import os, uuid, json, tempfile, argparse, webbrowser, threading
from pathlib import Path
from flask import Flask, request, jsonify, send_file, render_template
from accessibility_engine import AccessibilitySession, CHECKS

app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # 100 MB upload limit

# In-memory session store — suitable for local single-user use
sessions: dict[str, AccessibilitySession] = {}


# ── ROUTES ─────────────────────────────────────────────────────────────────────

@app.route("/")
def index():
    return send_file(Path(app.root_path) / "static" / "index.html")


@app.route("/api/checks")
def get_checks():
    """Return the full check registry for building the settings UI."""
    return jsonify(CHECKS)


@app.route("/api/upload", methods=["POST"])
def upload():
    """Accept a PPTX file, scan it, auto-apply configured fixes, return session data."""
    if "file" not in request.files:
        return jsonify({"error": "No file provided"}), 400

    file = request.files["file"]
    if not file.filename.lower().endswith(".pptx"):
        return jsonify({"error": "Only .pptx files are supported"}), 400

    # Parse settings and API key from form data
    raw_settings = request.form.get("settings", "{}")
    try:
        settings = json.loads(raw_settings)
    except json.JSONDecodeError:
        settings = {}

    api_key = request.form.get("api_key") or os.environ.get("ANTHROPIC_API_KEY")

    # Create working directory and save file
    session_id = str(uuid.uuid4())
    temp_dir   = tempfile.mkdtemp(prefix=f"a11y_{session_id[:8]}_")
    pptx_path  = os.path.join(temp_dir, file.filename)
    file.save(pptx_path)

    # Build session, scan, and generate thumbnails
    session = AccessibilitySession(session_id, pptx_path, settings, api_key)
    session.scan_and_auto_fix()
    session.generate_thumbnails()
    sessions[session_id] = session

    return jsonify(session.to_dict())


@app.route("/api/thumbnail/<session_id>/<int:slide_index>")
def thumbnail(session_id, slide_index):
    """Serve a slide thumbnail image."""
    session = sessions.get(session_id)
    if not session:
        return "Session not found", 404
    path = session.get_thumbnail_path(slide_index)
    if not path:
        return "Thumbnail not found", 404
    return send_file(path, mimetype="image/png")


@app.route("/api/apply-fix", methods=["POST"])
def apply_fix():
    """Apply (or skip) a single pending issue."""
    data = request.json or {}
    session = sessions.get(data.get("session_id"))
    if not session:
        return jsonify({"error": "Session not found"}), 404

    result = session.apply_fix(
        issue_id = data["issue_id"],
        action   = data["action"],   # accept | edit | skip
        value    = data.get("value"),
    )

    # Regenerate thumbnail for the affected slide
    slide_index = result.get("slide_index", -1)
    if slide_index >= 0:
        session.regenerate_thumbnail(slide_index)
        result["thumbnail_url"] = (
            f"/api/thumbnail/{session.session_id}/{slide_index}"
            f"?t={uuid.uuid4().hex[:6]}"
        )

    # Return updated full session state so the frontend can re-render
    result["session"] = session.to_dict()
    return jsonify(result)


@app.route("/api/update-settings", methods=["POST"])
def update_settings():
    """Update the autofix toggle settings for an existing session."""
    data = request.json or {}
    session = sessions.get(data.get("session_id"))
    if not session:
        return jsonify({"error": "Session not found"}), 404
    session.settings = data.get("settings", {})
    return jsonify({"ok": True, "session": session.to_dict()})


@app.route("/api/download/<session_id>")
def download(session_id):
    """Download the remediated PPTX file."""
    session = sessions.get(session_id)
    if not session:
        return "Session not found", 404
    stem = session.original_path.stem
    return send_file(
        str(session.pptx_path),
        as_attachment=True,
        download_name=f"{stem}_accessible.pptx",
        mimetype=(
            "application/vnd.openxmlformats-officedocument"
            ".presentationml.presentation"
        ),
    )


# ── ENTRY POINT ────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(
        prog="app.py",
        description="PPTX Accessibility Remediator — local web app",
    )
    parser.add_argument("--port",    type=int, default=5001,        metavar="PORT",
                        help="port to listen on (default: 5001)")
    parser.add_argument("--host",    default="127.0.0.1",           metavar="HOST",
                        help="host to bind to (default: 127.0.0.1; use 0.0.0.0 to expose on LAN)")
    parser.add_argument("--open",    action="store_true",
                        help="open the browser automatically on startup")
    parser.add_argument("--api-key", metavar="KEY",
                        help="Anthropic API key (overrides ANTHROPIC_API_KEY env var)")
    parser.add_argument("--no-debug", action="store_true",
                        help="disable Flask debug mode / reloader")
    args = parser.parse_args()

    if args.api_key:
        os.environ["ANTHROPIC_API_KEY"] = args.api_key

    url = f"http://{'localhost' if args.host == '127.0.0.1' else args.host}:{args.port}"

    print("\n  PPTX Accessibility Remediator")
    print("  ─────────────────────────────")
    print(f"  Open: {url}")
    if args.host == "0.0.0.0":
        import socket
        lan_ip = socket.gethostbyname(socket.gethostname())
        print(f"  LAN:  http://{lan_ip}:{args.port}")
    if os.environ.get("ANTHROPIC_API_KEY"):
        print("  AI:   enabled (Claude alt text + slide titles)")
    else:
        print("  AI:   disabled (no ANTHROPIC_API_KEY set)")
    print()

    if args.open:
        threading.Timer(1.0, lambda: webbrowser.open(url)).start()

    app.run(host=args.host, port=args.port, debug=not args.no_debug)
