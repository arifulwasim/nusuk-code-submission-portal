"""
Nusuk Code Submission Portal
----------------------------
Install:
    pip install -r requirements.txt

Run locally:
    python app.py

Open:
    http://127.0.0.1:5000

Admin:
    Username: admin
    Password: set ADMIN_PASSWORD environment variable.
    If not set locally, temporary default is ChangeThisAdminPassword123!
    CHANGE IT before public deployment.

For Render:
    Build command: pip install -r requirements.txt
    Start command: gunicorn app:app
    Environment variables:
        SECRET_KEY = a-long-random-secret
        ADMIN_PASSWORD = your-strong-admin-password
        DATA_DIR = /var/data   (when using a Persistent Disk mounted at /var/data)
"""

import os
import hmac
import secrets
import sqlite3
from pathlib import Path
from datetime import datetime
from threading import Lock

from flask import (
    Flask, request, redirect, url_for, render_template_string,
    send_file, flash, session
)
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, PatternFill, Alignment

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "local-dev-secret-change-before-deploy")
ADMIN_PASSWORD = os.environ.get(
    "ADMIN_PASSWORD", "ChangeThisAdminPassword123!"
)
ADMIN_USERNAME = os.environ.get("ADMIN_USERNAME", "admin")

DATA_DIR = Path(os.environ.get("DATA_DIR", Path(__file__).resolve().parent))
DATA_DIR.mkdir(parents=True, exist_ok=True)

DB_FILE = DATA_DIR / "submissions.db"
EXCEL_FILE = DATA_DIR / "submissions.xlsx"
excel_lock = Lock()

PAGE = r"""
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Nusuk Code Submission Portal</title>
<style>
*{box-sizing:border-box}
body{margin:0;padding:28px 16px;font-family:Arial,sans-serif;background:#f3f6fb;color:#172033}
.container{max-width:1200px;margin:auto}
header{text-align:center;margin-bottom:24px}
header h1{margin:0 0 8px;color:#173f8a;font-size:30px}
header p{margin:0;color:#64748b}
.topbar{display:flex;justify-content:flex-end;margin-bottom:12px}
.admin-link{color:#1d4ed8;text-decoration:none;font-size:14px}
.stats{background:white;padding:18px 22px;border-radius:14px;margin-bottom:20px;box-shadow:0 4px 18px #0f172a0b}
.stat-label{color:#64748b;font-size:14px}
.stat-number{color:#1d4ed8;font-size:30px;font-weight:700;margin-top:4px}
.layout{display:grid;grid-template-columns:minmax(280px,.85fr) minmax(0,1.5fr);gap:20px;align-items:start}
.card{background:white;padding:22px;border-radius:14px;box-shadow:0 4px 18px #0f172a0b;min-width:0}
h2{margin:0 0 16px;font-size:20px}
textarea{width:100%;min-height:170px;padding:14px;resize:vertical;border:1px solid #cbd5e1;border-radius:9px;font-size:15px;font-family:monospace;outline:none}
textarea:focus{border-color:#2563eb;box-shadow:0 0 0 3px #2563eb18}
button,.btn{display:inline-flex;align-items:center;justify-content:center;border:0;border-radius:8px;padding:11px 15px;font-size:14px;font-weight:600;cursor:pointer;text-decoration:none}
.primary{background:#2563eb;color:white;width:100%;margin-top:12px}
.primary:hover{background:#1d4ed8}
.copy{background:#0f766e;color:white}
.copy:hover{background:#115e59}
.download{background:#e8efff;color:#1d4ed8}
.message{padding:11px 12px;border-radius:8px;margin-bottom:12px;font-size:14px}
.success{background:#dcfce7;color:#166534}
.error{background:#fee2e2;color:#991b1b}
.list-head{display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap;margin-bottom:12px}
.list-head h2{margin:0}
.actions{display:flex;gap:8px;flex-wrap:wrap}
.table-wrap{overflow-x:auto;border:1px solid #e2e8f0;border-radius:9px}
table{width:100%;border-collapse:collapse;min-width:500px}
th,td{padding:12px;border-bottom:1px solid #e2e8f0;text-align:left;font-size:14px}
th{background:#eff6ff;color:#1d4ed8}
tr:last-child td{border-bottom:0}
td.code{white-space:pre-wrap;overflow-wrap:anywhere;font-family:monospace}
.empty{text-align:center;color:#64748b;padding:25px}
.hint{color:#64748b;font-size:12px;line-height:1.5;margin-top:10px}
@media(max-width:800px){.layout{grid-template-columns:1fr}}
@media(max-width:520px){body{padding:16px 10px}.card{padding:16px}.list-head{align-items:flex-start}.actions{width:100%}.actions button,.actions a{flex:1}}
</style>
</head>
<body>
<main class="container">
<header>
  <h1>Nusuk Code Submission Portal</h1>
  <p>Submit your codes — one code per line.</p>
</header>
<div class="topbar"><a class="admin-link" href="{{ url_for('admin_login') }}">Admin Access</a></div>

<section class="stats">
  <div class="stat-label">Total Codes Submitted</div>
  <div class="stat-number">{{ total }}</div>
</section>

<div class="layout">
  <section class="card">
    <h2>Submit Your Codes</h2>
    {% with messages = get_flashed_messages(with_categories=true) %}
      {% for category, message in messages %}
        <div class="message {{ category }}">{{ message }}</div>
      {% endfor %}
    {% endwith %}
    <form method="POST" action="{{ url_for('index') }}">
      <input type="hidden" name="csrf_token" value="{{ csrf_token }}">
      <textarea name="codes" maxlength="20000" placeholder="Paste your codes here&#10;NG-hFjy15A8&#10;NG-XUr0pLK4&#10;NG-IFsnzGEm" required></textarea>
      <button class="primary" type="submit">Submit Codes</button>
    </form>
    <p class="hint">Each non-empty new line will be counted as a separate code. Blank lines are ignored. Duplicate codes are also counted as separate submissions.</p>
  </section>

  <section class="card">
    <div class="list-head">
      <h2>Submitted Codes</h2>
      <div class="actions">
        <button class="copy" type="button" onclick="copyAllData()">Copy All Data</button>
        <a class="btn download" href="{{ url_for('download_excel') }}">Download Excel</a>
      </div>
    </div>
    <div class="table-wrap">
      <table id="submissionTable">
        <thead><tr><th>Serial</th><th>Code</th><th>Submission Time</th></tr></thead>
        <tbody>
        {% for item in submissions %}
          <tr><td>{{ item["id"] }}</td><td class="code">{{ item["code"] }}</td><td>{{ item["submitted_at"] }}</td></tr>
        {% else %}
          <tr><td class="empty" colspan="3">No codes submitted yet.</td></tr>
        {% endfor %}
        </tbody>
      </table>
    </div>
    <p class="hint">Copy All Data copies the full list, including serial number and submission time.</p>
  </section>
</div>
</main>
<script>
async function copyAllData(){
  const table=document.getElementById("submissionTable");
  const lines=Array.from(table.querySelectorAll("tr")).map(row=>
    Array.from(row.querySelectorAll("th,td"))
      .map(cell=>cell.innerText.trim().replace(/\s+/g," "))
      .join("\\t")
  );
  const text=lines.join("\\n");
  try{
    await navigator.clipboard.writeText(text);
    alert("All submission data copied!");
  }catch(e){
    const box=document.createElement("textarea");
    box.value=text;document.body.appendChild(box);box.select();
    const ok=document.execCommand("copy");box.remove();
    alert(ok?"All submission data copied!":"Copy failed. Please try again.");
  }
}
</script>
</body>
</html>
"""

LOGIN_PAGE = r"""
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Admin Login - Nusuk</title>
<style>
body{font-family:Arial;background:#f3f6fb;margin:0;padding:30px 15px;color:#172033}
.card{max-width:420px;margin:70px auto;background:white;padding:28px;border-radius:14px;box-shadow:0 4px 18px #0001}
h1{font-size:24px;color:#173f8a}
input{width:100%;padding:12px;margin:8px 0 14px;border:1px solid #cbd5e1;border-radius:8px;box-sizing:border-box}
button{width:100%;padding:12px;background:#2563eb;color:white;border:0;border-radius:8px;font-weight:bold;cursor:pointer}
a{display:inline-block;margin-top:16px;color:#2563eb;text-decoration:none}
.message{padding:10px;border-radius:7px;background:#fee2e2;color:#991b1b;margin-bottom:12px}
</style></head><body><div class="card">
<h1>Admin Login</h1>
{% with messages = get_flashed_messages(with_categories=true) %}
{% for category,message in messages %}<div class="message">{{ message }}</div>{% endfor %}
{% endwith %}
<form method="POST">
<input type="hidden" name="csrf_token" value="{{ csrf_token }}">
<label>Username</label><input name="username" required autocomplete="username">
<label>Password</label><input type="password" name="password" required autocomplete="current-password">
<button type="submit">Login</button>
</form><a href="{{ url_for('index') }}">← Back to website</a>
</div></body></html>
"""

ADMIN_PAGE = r"""
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1">
<title>Admin Dashboard - Nusuk</title>
<style>
body{font-family:Arial;background:#f3f6fb;margin:0;padding:25px 14px;color:#172033}
.wrap{max-width:1100px;margin:auto}.head{display:flex;justify-content:space-between;align-items:center;gap:12px;flex-wrap:wrap}
h1{color:#173f8a}.card{background:white;padding:20px;border-radius:12px;margin-top:18px;box-shadow:0 4px 18px #0001}
table{width:100%;border-collapse:collapse;min-width:650px}th,td{padding:11px;border-bottom:1px solid #e2e8f0;text-align:left;font-size:14px}
th{background:#eff6ff;color:#1d4ed8}.tablewrap{overflow-x:auto}
a,button{padding:8px 12px;border:0;border-radius:7px;text-decoration:none;cursor:pointer;font-weight:bold}
.edit{background:#e8efff;color:#1d4ed8}.delete{background:#fee2e2;color:#991b1b}.logout{background:#334155;color:white}
.message{padding:10px;background:#dcfce7;color:#166534;border-radius:7px;margin-top:10px}
.code{font-family:monospace;overflow-wrap:anywhere}
</style></head><body><div class="wrap">
<div class="head"><div><h1>Nusuk Admin Dashboard</h1><p>Total codes: <b>{{ total }}</b></p></div>
<a class="logout" href="{{ url_for('admin_logout') }}">Logout</a></div>
{% with messages = get_flashed_messages(with_categories=true) %}
{% for category,message in messages %}<div class="message">{{ message }}</div>{% endfor %}
{% endwith %}
<div class="card"><div class="tablewrap"><table>
<thead><tr><th>Serial</th><th>Code</th><th>Submitted At</th><th>Actions</th></tr></thead><tbody>
{% for item in submissions %}
<tr><td>{{ item["id"] }}</td><td class="code">{{ item["code"] }}</td><td>{{ item["submitted_at"] }}</td>
<td>
<a class="edit" href="{{ url_for('admin_edit', submission_id=item['id']) }}">Edit</a>
<form method="POST" action="{{ url_for('admin_delete', submission_id=item['id']) }}" style="display:inline" onsubmit="return confirm('Delete this code?')">
<input type="hidden" name="csrf_token" value="{{ csrf_token }}">
<button class="delete" type="submit">Delete</button></form>
</td></tr>
{% else %}<tr><td colspan="4">No submissions yet.</td></tr>{% endfor %}
</tbody></table></div></div>
</div></body></html>
"""

EDIT_PAGE = r"""
<!DOCTYPE html><html lang="en"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">
<title>Edit Code - Nusuk</title><style>
body{font-family:Arial;background:#f3f6fb;padding:25px;color:#172033}.card{max-width:600px;margin:40px auto;background:white;padding:24px;border-radius:12px}
textarea{width:100%;min-height:100px;padding:12px;box-sizing:border-box;border:1px solid #cbd5e1;border-radius:8px;font-family:monospace}
button{margin-top:12px;padding:11px 18px;background:#2563eb;color:white;border:0;border-radius:8px;cursor:pointer}
a{margin-left:10px;color:#2563eb;text-decoration:none}
</style></head><body><div class="card"><h1>Edit Submitted Code</h1>
<form method="POST"><input type="hidden" name="csrf_token" value="{{ csrf_token }}">
<textarea name="code" maxlength="2000" required>{{ item["code"] }}</textarea><br>
<button type="submit">Save Changes</button><a href="{{ url_for('admin_dashboard') }}">Cancel</a>
</form></div></body></html>
"""


def get_db():
    conn = sqlite3.connect(DB_FILE, timeout=30)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_db() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS submissions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                code TEXT NOT NULL,
                submitted_at TEXT NOT NULL
            )
        """)


def csrf_token():
    if "_csrf_token" not in session:
        session["_csrf_token"] = secrets.token_urlsafe(32)
    return session["_csrf_token"]


def valid_csrf():
    submitted = request.form.get("csrf_token", "")
    expected = session.get("_csrf_token", "")
    return bool(submitted and expected and hmac.compare_digest(submitted, expected))


def admin_required():
    return session.get("is_admin") is True


def export_excel():
    with excel_lock:
        with get_db() as conn:
            rows = conn.execute(
                "SELECT id, code, submitted_at FROM submissions ORDER BY id ASC"
            ).fetchall()

        if EXCEL_FILE.exists():
            workbook = load_workbook(EXCEL_FILE)
            sheet = workbook.active
            sheet.delete_rows(1, sheet.max_row)
        else:
            workbook = Workbook()
            sheet = workbook.active

        sheet.title = "Submissions"
        sheet.append(["Serial", "Submitted Code", "Submission Time"])
        for cell in sheet[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="173F8A")
            cell.alignment = Alignment(horizontal="center")

        for row in rows:
            sheet.append([row["id"], row["code"], row["submitted_at"]])
            sheet.cell(sheet.max_row, 2).data_type = "s"

        sheet.column_dimensions["A"].width = 12
        sheet.column_dimensions["B"].width = 40
        sheet.column_dimensions["C"].width = 24
        temp_file = DATA_DIR / "submissions_temp.xlsx"
        workbook.save(temp_file)
        os.replace(temp_file, EXCEL_FILE)


@app.context_processor
def inject_csrf():
    return {"csrf_token": csrf_token()}


@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        if not valid_csrf():
            flash("Session expired. Please try submitting again.", "error")
            return redirect(url_for("index"))

        raw_codes = request.form.get("codes", "")
        # Normalize Windows/Mac line endings. Each non-empty line is a separate code.
        codes = [line.strip() for line in raw_codes.replace("\r\n", "\n").replace("\r", "\n").split("\n") if line.strip()]

        if not codes:
            flash("Please enter at least one code.", "error")
            return redirect(url_for("index"))
        if any(len(code) > 2000 for code in codes):
            flash("Each code must be 2,000 characters or fewer.", "error")
            return redirect(url_for("index"))

        submitted_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        try:
            with get_db() as conn:
                conn.executemany(
                    "INSERT INTO submissions (code, submitted_at) VALUES (?, ?)",
                    [(code, submitted_at) for code in codes]
                )
            export_excel()
            flash(f"Successfully submitted {len(codes)} code(s).", "success")
        except Exception:
            app.logger.exception("Submission failed")
            flash("Submission failed. Please try again.", "error")
        return redirect(url_for("index"))

    with get_db() as conn:
        submissions = conn.execute(
            "SELECT id, code, submitted_at FROM submissions ORDER BY id DESC"
        ).fetchall()
        total = conn.execute("SELECT COUNT(*) FROM submissions").fetchone()[0]
    return render_template_string(PAGE, submissions=submissions, total=total)


@app.route("/download")
def download_excel():
    export_excel()
    return send_file(EXCEL_FILE, as_attachment=True, download_name="Nusuk_Code_Submissions.xlsx")


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if admin_required():
        return redirect(url_for("admin_dashboard"))
    if request.method == "POST":
        if not valid_csrf():
            flash("Session expired. Please try again.", "error")
            return redirect(url_for("admin_login"))
        username = request.form.get("username", "")
        password = request.form.get("password", "")
        if hmac.compare_digest(username, ADMIN_USERNAME) and hmac.compare_digest(password, ADMIN_PASSWORD):
            session.clear()
            session["is_admin"] = True
            session["_csrf_token"] = secrets.token_urlsafe(32)
            return redirect(url_for("admin_dashboard"))
        flash("Incorrect username or password.", "error")
    return render_template_string(LOGIN_PAGE)


@app.route("/admin")
def admin_dashboard():
    if not admin_required():
        return redirect(url_for("admin_login"))
    with get_db() as conn:
        submissions = conn.execute(
            "SELECT id, code, submitted_at FROM submissions ORDER BY id DESC"
        ).fetchall()
        total = conn.execute("SELECT COUNT(*) FROM submissions").fetchone()[0]
    return render_template_string(ADMIN_PAGE, submissions=submissions, total=total)


@app.route("/admin/edit/<int:submission_id>", methods=["GET", "POST"])
def admin_edit(submission_id):
    if not admin_required():
        return redirect(url_for("admin_login"))
    with get_db() as conn:
        item = conn.execute(
            "SELECT id, code, submitted_at FROM submissions WHERE id = ?",
            (submission_id,)
        ).fetchone()
    if item is None:
        flash("Code not found.", "error")
        return redirect(url_for("admin_dashboard"))

    if request.method == "POST":
        if not valid_csrf():
            flash("Session expired. Please try again.", "error")
            return redirect(url_for("admin_dashboard"))
        new_code = request.form.get("code", "").strip()
        if not new_code or len(new_code) > 2000:
            flash("Code cannot be empty or exceed 2,000 characters.", "error")
            return redirect(url_for("admin_edit", submission_id=submission_id))
        with get_db() as conn:
            conn.execute(
                "UPDATE submissions SET code = ? WHERE id = ?",
                (new_code, submission_id)
            )
        export_excel()
        flash("Code updated successfully.", "success")
        return redirect(url_for("admin_dashboard"))

    return render_template_string(EDIT_PAGE, item=item)


@app.route("/admin/delete/<int:submission_id>", methods=["POST"])
def admin_delete(submission_id):
    if not admin_required():
        return redirect(url_for("admin_login"))
    if not valid_csrf():
        flash("Session expired. Please try again.", "error")
        return redirect(url_for("admin_dashboard"))
    with get_db() as conn:
        conn.execute("DELETE FROM submissions WHERE id = ?", (submission_id,))
    export_excel()
    flash("Code deleted successfully.", "success")
    return redirect(url_for("admin_dashboard"))


@app.route("/admin/logout")
def admin_logout():
    session.clear()
    return redirect(url_for("index"))


if __name__ == "__main__":
    init_db()
    export_excel()
    if ADMIN_PASSWORD == "ChangeThisAdminPassword123!":
        print("WARNING: Set a strong ADMIN_PASSWORD environment variable before public deployment.")
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", 5000)), debug=False)
