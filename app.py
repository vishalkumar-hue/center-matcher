import sys
import os
import importlib.util
_this_dir = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _this_dir)
_spec = importlib.util.spec_from_file_location(
    "matcher_local",
    os.path.join(_this_dir, "matcher.py")
)
_matcher_module = importlib.util.module_from_spec(_spec)
_spec.loader.exec_module(_matcher_module)
run_matching = _matcher_module.run_matching
run_matching_multi = _matcher_module.run_matching_multi   # ← NAYA IMPORT

from flask import Flask, render_template, request, send_file, jsonify
import uuid
from pathlib import Path

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 100 * 1024 * 1024   # 50MB se 100MB kiya (multi-file ke liye)
BASE_DIR = Path(__file__).parent
UPLOAD_DIR = BASE_DIR / "uploads"
RESULT_DIR = BASE_DIR / "results"
UPLOAD_DIR.mkdir(exist_ok=True)
RESULT_DIR.mkdir(exist_ok=True)
ALLOWED_EXT = {'.xlsx', '.xls'}


def allowed_file(filename):
    return Path(filename).suffix.lower() in ALLOWED_EXT


@app.route('/')
def index():
    return render_template('index.html')


# ── PURANA SINGLE-FILE ROUTE — waise hi rehne do (backward compatible) ──
@app.route('/run', methods=['POST'])
def run():
    file_a = request.files.get('file_a')
    file_b = request.files.get('file_b')
    errors = []
    if not file_a or file_a.filename == '':
        errors.append("List A file required hai.")
    elif not allowed_file(file_a.filename):
        errors.append("List A must be .xlsx or .xls")
    if not file_b or file_b.filename == '':
        errors.append("List B file required hai.")
    elif not allowed_file(file_b.filename):
        errors.append("List B must be .xlsx or .xls")
    if errors:
        return jsonify({"success": False, "errors": errors}), 400
    job_id = str(uuid.uuid4())[:8]
    path_a = UPLOAD_DIR / f"{job_id}_A{Path(file_a.filename).suffix}"
    path_b = UPLOAD_DIR / f"{job_id}_B{Path(file_b.filename).suffix}"
    output_path = RESULT_DIR / f"Result_{job_id}.xlsx"
    file_a.save(str(path_a))
    file_b.save(str(path_b))
    try:
        stats = run_matching(str(path_a), str(path_b), str(output_path))
        stats["job_id"] = job_id
        stats["filename"] = output_path.name
        return jsonify({"success": True, **stats})
    except Exception as e:
        return jsonify({"success": False, "errors": [str(e)]}), 500
    finally:
        if path_a.exists(): path_a.unlink()
        if path_b.exists(): path_b.unlink()


# ── NAYA ROUTE — MULTIPLE B FILES KE LIYE ──────────────────────
@app.route('/run-multi', methods=['POST'])
def run_multi():
    file_a = request.files.get('file_a')
    files_b = request.files.getlist('files_b')   # ← multiple files yahan se aayenge

    errors = []
    if not file_a or file_a.filename == '':
        errors.append("List A file required hai.")
    elif not allowed_file(file_a.filename):
        errors.append("List A must be .xlsx or .xls")

    if not files_b or len(files_b) == 0 or files_b[0].filename == '':
        errors.append("Kam se kam ek List B file required hai.")
    else:
        for fb in files_b:
            if not allowed_file(fb.filename):
                errors.append(f"{fb.filename} must be .xlsx or .xls")

    if errors:
        return jsonify({"success": False, "errors": errors}), 400

    job_id = str(uuid.uuid4())[:8]
    path_a = UPLOAD_DIR / f"{job_id}_A{Path(file_a.filename).suffix}"
    file_a.save(str(path_a))

    # Har B file ko save karo aur (tab_name, path) tuple banao
    b_paths_with_names = []
    saved_b_paths = []
    for i, fb in enumerate(files_b):
        suffix = Path(fb.filename).suffix
        path_b = UPLOAD_DIR / f"{job_id}_B{i}{suffix}"
        fb.save(str(path_b))
        saved_b_paths.append(path_b)
        # Tab name file ke naam se banao (extension hata ke, max 31 chars Excel limit ke liye)
        tab_name = Path(fb.filename).stem[:31]
        b_paths_with_names.append((tab_name, str(path_b)))

    output_path = RESULT_DIR / f"Result_Multi_{job_id}.xlsx"

    try:
        stats = run_matching_multi(str(path_a), b_paths_with_names, str(output_path))
        stats["job_id"] = job_id
        stats["filename"] = output_path.name
        return jsonify({"success": True, **stats})
    except Exception as e:
        return jsonify({"success": False, "errors": [str(e)]}), 500
    finally:
        if path_a.exists():
            path_a.unlink()
        for p in saved_b_paths:
            if p.exists():
                p.unlink()


@app.route('/download/<filename>')
def download(filename):
    file_path = RESULT_DIR / filename
    if not file_path.exists():
        return "File not found", 404
    return send_file(str(file_path), as_attachment=True, download_name=filename)


if __name__ == '__main__':
    app.run(debug=True, host='0.0.0.0', port=5000)
