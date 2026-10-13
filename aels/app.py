"""Flask web app: upload a report, label new lines, export the result."""

from __future__ import annotations

import os
import uuid
from pathlib import Path

from flask import Flask, jsonify, render_template, request, send_file
from werkzeug.utils import secure_filename

from .db import LabelStore
from .processing import assign_label, export_labeled, process_sof, totals_by_section

MAX_UPLOAD_MB = 25


def create_app(db_path: str | None = None, data_dir: str | None = None) -> Flask:
    app = Flask(__name__)
    app.config["MAX_CONTENT_LENGTH"] = MAX_UPLOAD_MB * 1024 * 1024

    data_dir = Path(data_dir or os.environ.get("AELS_DATA_DIR", "data"))
    upload_dir = data_dir / "uploads"
    output_dir = data_dir / "outputs"
    upload_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)

    store = LabelStore(db_path or os.environ.get("AELS_DB_PATH", str(data_dir / "aels.db")))

    # Parsed reports kept in memory while you work on them, keyed by session id.
    # Labels are written to the database immediately, so nothing learned is
    # lost if the server restarts -- only the in-progress view is.
    sessions: dict[str, dict] = {}

    def _session_or_404(session_id: str | None):
        sess = sessions.get(session_id or "")
        if sess is None:
            return None, (jsonify(error="Session not found. Please upload the file again."), 404)
        return sess, None

    def _summary(sess: dict) -> dict:
        items = sess["items"]
        return {
            "labeled": sum(1 for i in items if i["section"]),
            "unlabeled": sum(1 for i in items if not i["section"]),
            "totals": totals_by_section(items),
        }

    @app.get("/")
    def index():
        return render_template("index.html")

    @app.post("/api/upload")
    def upload():
        file = request.files.get("file")
        if file is None or not file.filename:
            return jsonify(error="No file was uploaded."), 400
        if not file.filename.lower().endswith(".xlsx"):
            return jsonify(error="Please upload an .xlsx file."), 400

        session_id = uuid.uuid4().hex[:12]
        filename = secure_filename(file.filename)
        path = upload_dir / f"{session_id}_{filename}"
        file.save(path)

        try:
            result = process_sof(str(path), store)
        except Exception as exc:  # surface parsing problems to the user
            path.unlink(missing_ok=True)
            return jsonify(error=f"Could not read this file: {exc}"), 422

        sessions[session_id] = {"filename": filename, "items": result["items"]}
        return jsonify(
            session_id=session_id,
            filename=filename,
            items=result["items"],
            stats=result["stats"],
            sections=store.get_sections(),
            summary=_summary(sessions[session_id]),
        )

    @app.post("/api/label")
    def label():
        data = request.get_json(silent=True) or {}
        sess, error = _session_or_404(data.get("session_id"))
        if error:
            return error

        commitment_item = (data.get("commitment_item") or "").strip()
        description = (data.get("description") or "").strip()
        section = (data.get("section") or "").strip()
        if not (commitment_item and section):
            return jsonify(error="commitment_item and section are required."), 400

        store.save_label(commitment_item, description, section)
        updated = assign_label(sess["items"], commitment_item, description, section)
        return jsonify(ok=True, updated=updated, summary=_summary(sess))

    @app.get("/api/sections")
    def get_sections():
        return jsonify(sections=store.get_sections())

    @app.post("/api/sections")
    def add_section():
        name = ((request.get_json(silent=True) or {}).get("name") or "").strip()
        if not name:
            return jsonify(error="Section name is required."), 400
        store.add_section(name)
        return jsonify(sections=store.get_sections())

    @app.post("/api/export")
    def export():
        data = request.get_json(silent=True) or {}
        sess, error = _session_or_404(data.get("session_id"))
        if error:
            return error

        out_name = sess["filename"].rsplit(".", 1)[0] + "_LABELED.xlsx"
        out_path = output_dir / f"{data['session_id']}_{out_name}"
        export_labeled(sess["items"], str(out_path))
        return send_file(out_path, as_attachment=True, download_name=out_name)

    return app


def main() -> None:
    port = int(os.environ.get("PORT", 5050))
    print(f"Automated Expenditure Labeling System running at http://localhost:{port}")
    create_app().run(host="127.0.0.1", port=port, debug=False)


if __name__ == "__main__":
    main()
