from flask import Flask, render_template, request, redirect, url_for, send_file, flash, jsonify
import os
import tempfile
import uuid
import re
import threading
import time
from queue import Queue, Empty

from converter import process_doc

ALLOWED_EXTENSIONS = {"docx"}

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "dev-secret-key")
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024  # 50MB


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def sanitize_filename_keep_parens(name: str) -> str:
    """
    Sanitize the filename while preserving parentheses. Prevents directory traversal and removes
    characters outside of the allowed set: letters, numbers, spaces, underscore, hyphen, dot, parentheses.
    """
    base = os.path.basename(name)
    # Remove Windows-forbidden characters: <>:"/\|?* and control chars
    base = re.sub(r"[<>:\"/\\|?*\x00-\x1F]", "", base)
    # Remove newlines and tabs explicitly just in case
    base = base.replace("\n", "").replace("\r", "").replace("\t", "")
    # Windows doesn't allow trailing spaces or dots in filenames
    base = base.rstrip(" .")
    # Avoid empty filename
    if not base:
        base = "document.docx"
    # Ensure .docx extension
    if not base.lower().endswith(".docx"):
        base = os.path.splitext(base)[0] + ".docx"
    return base


@app.route("/", methods=["GET"]) 
def index():
    return render_template("index.html")


JOBS = {}
JOBS_LOCK = threading.Lock()
JOB_QUEUE: Queue[str] = Queue()


def _run_job(job_id: str, input_path: str, output_path: str):
    def on_progress(pct: int):
        with JOBS_LOCK:
            job = JOBS.get(job_id)
            if job:
                job["processing_pct"] = max(0, min(100, int(pct)))

    try:
        with JOBS_LOCK:
            JOBS[job_id]["status"] = "processing"
            JOBS[job_id]["processing_pct"] = 0

        final_path = process_doc(input_path, output_path=output_path, visible=False, progress=on_progress)
        with JOBS_LOCK:
            JOBS[job_id]["status"] = "done"
            JOBS[job_id]["final_path"] = final_path
            JOBS[job_id]["processing_pct"] = 100
    except Exception as e:
        with JOBS_LOCK:
            JOBS[job_id]["status"] = "error"
            JOBS[job_id]["error"] = str(e)
        # Log full traceback for debugging
        try:
            app.logger.exception("Job %s failed: %s", job_id, e)
        except Exception:
            pass


def _worker_loop():
    while True:
        try:
            job_id = JOB_QUEUE.get(timeout=0.5)
        except Empty:
            continue
        with JOBS_LOCK:
            job = JOBS.get(job_id)
        if not job:
            continue
        # Each job dict stores paths
        input_path = job.get("input_path")
        output_path = job.get("output_path")
        try:
            _run_job(job_id, input_path, output_path)
        finally:
            JOB_QUEUE.task_done()


@app.route("/start", methods=["POST"]) 
def start():
    if "file" not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files["file"]
    if file.filename == "":
        return jsonify({"error": "No file selected"}), 400
    if not allowed_file(file.filename):
        return jsonify({"error": "Only .docx allowed"}), 400

    filename = sanitize_filename_keep_parens(file.filename)

    # Save upload to temp input; output to convertedDocx with same name
    tmpdir = tempfile.mkdtemp()
    input_path = os.path.join(tmpdir, f"{uuid.uuid4()}_{filename}")
    file.save(input_path)

    project_root = os.path.dirname(os.path.abspath(__file__))
    out_dir = os.path.join(project_root, "convertedDocx")
    os.makedirs(out_dir, exist_ok=True)
    output_path = os.path.join(out_dir, filename)

    job_id = str(uuid.uuid4())
    with JOBS_LOCK:
        JOBS[job_id] = {
            "status": "queued",
            "filename": filename,
            "processing_pct": 0,
            "final_path": None,
            "error": None,
            "input_path": input_path,
            "output_path": output_path,
        }
    JOB_QUEUE.put(job_id)

    return jsonify({"job_id": job_id, "filename": filename})


@app.route("/start-multi", methods=["POST"]) 
def start_multi():
    files = request.files.getlist("files")
    if not files:
        return jsonify({"error": "No files uploaded"}), 400

    jobs = []
    project_root = os.path.dirname(os.path.abspath(__file__))
    out_dir = os.path.join(project_root, "convertedDocx")
    os.makedirs(out_dir, exist_ok=True)

    for f in files:
        if f.filename == "":
            return jsonify({"error": "One of the files has an empty filename"}), 400
        if not allowed_file(f.filename):
            return jsonify({"error": f"Unsupported file: {f.filename}"}), 400

        filename = sanitize_filename_keep_parens(f.filename)
        tmpdir = tempfile.mkdtemp()
        input_path = os.path.join(tmpdir, f"{uuid.uuid4()}_{filename}")
        f.save(input_path)

        output_path = os.path.join(out_dir, filename)
        job_id = str(uuid.uuid4())
        with JOBS_LOCK:
            JOBS[job_id] = {
                "status": "queued",
                "filename": filename,
                "processing_pct": 0,
                "final_path": None,
                "error": None,
                "input_path": input_path,
                "output_path": output_path,
            }
        JOB_QUEUE.put(job_id)
        jobs.append({"job_id": job_id, "filename": filename})

    return jsonify({"jobs": jobs})


@app.route("/progress/<job_id>", methods=["GET"]) 
def progress(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return jsonify({"error": "Invalid job id"}), 404
        return jsonify({
            "status": job["status"],
            "processing_pct": job.get("processing_pct", 0),
            "filename": job.get("filename"),
            "error": job.get("error")
        })


@app.route("/result/<job_id>", methods=["GET"]) 
def result(job_id):
    with JOBS_LOCK:
        job = JOBS.get(job_id)
        if not job:
            return jsonify({"error": "Invalid job id"}), 404
        if job["status"] != "done" or not job.get("final_path"):
            return jsonify({"error": "Not ready"}), 400
        path = job["final_path"]
        name = job.get("filename") or os.path.basename(path)
    # Send file with original filename; parentheses preserved
    return send_file(path, as_attachment=True, download_name=name)


@app.route("/convert", methods=["POST"]) 
def convert():
    if "file" not in request.files:
        flash("No file part in the request.", "error")
        return redirect(url_for("index"))

    file = request.files["file"]
    if file.filename == "":
        flash("No file selected.", "error")
        return redirect(url_for("index"))

    if not allowed_file(file.filename):
        flash("Please upload a .docx file.", "error")
        return redirect(url_for("index"))

    filename = sanitize_filename_keep_parens(file.filename)

    # Temporary input path
    with tempfile.TemporaryDirectory() as tmpdir:
        input_path = os.path.join(tmpdir, f"{uuid.uuid4()}_{filename}")
        file.save(input_path)

        # Ensure output directory exists in project root
        project_root = os.path.dirname(os.path.abspath(__file__))
        out_dir = os.path.join(project_root, "convertedDocx")
        os.makedirs(out_dir, exist_ok=True)

        # Save with the same name as the uploaded file
        out_name = filename
        output_path = os.path.join(out_dir, out_name)

        try:
            final_path = process_doc(input_path, output_path=output_path, visible=False)
        except FileNotFoundError:
            flash("Uploaded file could not be found.", "error")
            return redirect(url_for("index"))
        except Exception as e:
            # Helpful message for COM/Word issues
            flash("Conversion failed. Ensure Microsoft Word is installed and that 'pywin32' is available.", "error")
            app.logger.exception("Conversion error: %s", e)
            return redirect(url_for("index"))

        # Send file as attachment from the permanent convertedDocx folder
        return send_file(final_path, as_attachment=True, download_name=out_name)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 5000))
    # Disable the reloader to avoid spawning multiple processes which interferes with COM
    # Start the single worker thread for serialized COM automation
    worker = threading.Thread(target=_worker_loop, daemon=True)
    worker.start()
    app.run(host="0.0.0.0", port=port, debug=True, use_reloader=False)
