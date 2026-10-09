"""
Mail Blaster — Web application for bulk personalized email sending.

Features:
- Upload Excel with email addresses
- Rich text editor (Quill) for email body
- Variable replacement ({name}, {email}, {handle}, or any {column_name})
- Random 2-10s delay between sends (anti-spam)
- Real-time progress polling
- SMTP config saved locally
"""

import os
import json
import smtplib
import random
import time
import tempfile
import threading
from email.message import EmailMessage
from email.utils import formataddr
from datetime import datetime
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, Form, Request
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse, Response
from openpyxl import load_workbook

app = FastAPI(title="Mail Blaster")


# ── Access Password (anti-leak) ──
# Set APP_PASSWORD env var (Railway) to require a password before ANY use.
# Link leaks are harmless: anyone without the password cannot open or use the app.
def _check_access(request: Request) -> bool:
    """Return True if the request has valid access credentials."""
    env_pass = os.environ.get("APP_PASSWORD")
    if not env_pass:
        return True  # no password configured → open access (local dev)
    auth = request.headers.get("Authorization", "")
    if auth.startswith("Basic "):
        import base64
        try:
            decoded = base64.b64decode(auth[6:]).decode("utf-8")
            _, _, password = decoded.partition(":")
            return password == env_pass
        except Exception:
            return False
    return False


def _deny(response_class=None):
    """Return a 401 response that triggers the browser native password prompt."""
    if response_class is JSONResponse:
        return JSONResponse(status_code=401, content={"error": "Unauthorized"})
    return Response(
        status_code=401,
        headers={"WWW-Authenticate": 'Basic realm="Mail Blaster", charset="UTF-8"'},
    )


@app.middleware("http")
async def access_guard(request: Request, call_next):
    """Protect every route: /, /api/*, uploads, everything."""
    path = request.url.path
    # Health check must stay open for Railway monitoring
    if path == "/api/health":
        return await call_next(request)
    if not _check_access(request):
        return _deny()
    return await call_next(request)

BASE_DIR = Path(__file__).parent

# Use Railway persistent volume if mounted, otherwise local directory
UPLOAD_DIR = Path(os.environ.get("UPLOAD_DIR", str(BASE_DIR / "uploads")))
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

# SMTP config: prefer environment variables (Railway-safe), fall back to local file
SMTP_CONFIG_FILE = BASE_DIR / ".smtp_config.json"
JOB_STATE_FILE = Path(os.environ.get("JOB_STATE_FILE", str(BASE_DIR / "data" / "job_state.json")))
JOB_STATE_FILE.parent.mkdir(parents=True, exist_ok=True)


def _load_smtp_from_env():
    """Load SMTP config from environment variables (used on Railway/Heroku/etc.)."""
    host = os.environ.get("SMTP_HOST")
    username = os.environ.get("SMTP_USERNAME")
    password = os.environ.get("SMTP_PASSWORD")
    if host and username and password:
        return {
            "host": host,
            "port": int(os.environ.get("SMTP_PORT", 465)),
            "username": username,
            "password": password,
            "sender_name": os.environ.get("SENDER_NAME", "Luna Hei"),
        }
    return None

_job_lock = threading.Lock()
_job_thread = None
_stop_flag = threading.Event()


# ── SMTP Config ──
def save_smtp_config(host, port, username, password, sender_name):
    config = {
        "host": host,
        "port": int(port),
        "username": username,
        "password": password,
        "sender_name": sender_name,
    }
    with open(SMTP_CONFIG_FILE, "w") as f:
        json.dump(config, f, indent=2)
    os.chmod(SMTP_CONFIG_FILE, 0o600)
    return config


def load_smtp_config():
    # 1) Prefer environment variables (Railway / production)
    env_cfg = _load_smtp_from_env()
    if env_cfg:
        return env_cfg
    # 2) Fall back to local file (for local dev)
    if not SMTP_CONFIG_FILE.exists():
        return None
    try:
        with open(SMTP_CONFIG_FILE) as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        return None


# ── Job State ──
def _save_state(state):
    fd, tmp = tempfile.mkstemp(dir=str(JOB_STATE_FILE.parent), suffix=".tmp")
    with os.fdopen(fd, "w") as f:
        json.dump(state, f, indent=2, ensure_ascii=False)
    os.replace(tmp, JOB_STATE_FILE)


def _load_state():
    if not JOB_STATE_FILE.exists():
        return {"status": "idle", "sent": 0, "failed": 0, "total": 0, "log": []}
    try:
        with open(JOB_STATE_FILE) as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        return {"status": "idle", "sent": 0, "failed": 0, "total": 0, "log": []}


def get_state():
    s = _load_state()
    # Check if worker is actually alive
    if s.get("status") == "running" and _job_thread and not _job_thread.is_alive():
        s["status"] = "interrupted"
        _save_state(s)
    return s


# ── Excel Processing ──
excel_columns_cache = []
excel_rows_cache = []
excel_path_cache = None


@app.post("/api/upload_excel")
async def upload_excel(file: UploadFile = File(...)):
    global excel_columns_cache, excel_rows_cache, excel_path_cache
    
    # Save file
    save_path = UPLOAD_DIR / file.filename
    with open(save_path, "wb") as f:
        f.write(await file.read())
    
    # Parse
    wb = load_workbook(save_path, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(values_only=True))
    if not rows:
        return JSONResponse({"ok": False, "error": "Empty Excel"}, status_code=400)
    
    headers = [str(h).strip() if h else f"col_{i+1}" for i, h in enumerate(rows[0])]
    
    # Find email column
    email_col = None
    for i, h in enumerate(headers):
        if "email" in h.lower() or "邮箱" in h or "mail" in h.lower():
            email_col = i
            break
    
    if email_col is None:
        # Return columns for manual selection
        excel_columns_cache = headers
        excel_rows_cache = rows[1:]
        excel_path_cache = str(save_path)
        return JSONResponse({
            "ok": True,
            "headers": headers,
            "row_count": len(rows) - 1,
            "email_column": None,
            "needs_column_select": True,
        })
    
    excel_columns_cache = headers
    excel_rows_cache = rows[1:]
    excel_path_cache = str(save_path)
    wb.close()
    return JSONResponse({
        "ok": True,
        "headers": headers,
        "email_column": email_col,
        "email_column_name": headers[email_col],
        "row_count": len(rows) - 1,
        "needs_column_select": False,
    })


@app.post("/api/set_email_column")
async def set_email_column(request: Request):
    global excel_columns_cache, excel_rows_cache
    body = await request.json()
    col_idx = int(body.get("column_index", -1))
    if col_idx < 0 or col_idx >= len(excel_columns_cache):
        return JSONResponse({"ok": False, "error": "Invalid column"}, status_code=400)
    return JSONResponse({
        "ok": True,
        "email_column": col_idx,
        "email_column_name": excel_columns_cache[col_idx],
    })


# ── SMTP Config ──
@app.post("/api/smtp/configure")
async def smtp_configure(
    host: str = Form(...),
    port: int = Form(default=465),
    username: str = Form(...),
    password: str = Form(...),
    sender_name: str = Form(default=""),
):
    config = save_smtp_config(host, port, username, password, sender_name)
    # Test connection
    try:
        with smtplib.SMTP_SSL(config["host"], config["port"], timeout=10) as srv:
            srv.login(config["username"], config["password"])
        return JSONResponse({"ok": True, "message": "SMTP connected successfully", "connected": True})
    except Exception as e:
        return JSONResponse({"ok": False, "error": str(e), "connected": False}, status_code=400)


@app.get("/api/smtp/status")
def smtp_status():
    config = load_smtp_config()
    if not config:
        return JSONResponse({"configured": False, "source": "none"})
    # Detect config source: env vars (production) or local file (dev)
    from_env = _load_smtp_from_env() is not None
    return JSONResponse({
        "configured": True,
        "host": config["host"],
        "port": config["port"],
        "username": config["username"],
        "sender_name": config.get("sender_name", ""),
        "source": "environment" if from_env else "file",
    })


@app.get("/api/health")
def health():
    """Health check endpoint for Railway / monitoring."""
    config = load_smtp_config()
    return JSONResponse({
        "status": "ok",
        "smtp_configured": config is not None,
        "upload_dir": str(UPLOAD_DIR),
        "platform": os.environ.get("RAILWAY_ENVIRONMENT", "local"),
    })


# ── Variable Preview ──
@app.post("/api/preview_variables")
async def preview_variables(request: Request):
    """Preview which variables are available and a sample replacement."""
    body = await request.json()
    html = body.get("html", "")
    round_params = body.get("round_params", {}) or {}
    if not excel_rows_cache:
        return JSONResponse({"ok": False, "error": "No Excel uploaded"}, status_code=400)
    # Apply round-level fixed params first (same for all emails this round)
    sample_html = html
    for rp_name, rp_val in round_params.items():
        if rp_val is None:
            rp_val = ""
        sample_html = sample_html.replace("{" + rp_name + "}", str(rp_val))
    # Find all {variable} in html
    import re
    variables = set(re.findall(r'\{(\w+)\}', html))
    # Try to match with columns
    matched = {}
    for v in variables:
        for i, h in enumerate(excel_columns_cache):
            if v.lower() in h.lower() or h.lower() in v.lower():
                matched[v] = h
                break
    # Build sample with first row
    sample_row = excel_rows_cache[0] if excel_rows_cache else []
    for v, col_name in matched.items():
        col_idx = excel_columns_cache.index(col_name)
        val = sample_row[col_idx] if col_idx < len(sample_row) else ""
        sample_html = sample_html.replace("{" + v + "}", str(val or ""))
    return JSONResponse({
        "ok": True,
        "variables": list(variables),
        "matched": matched,
        "sample_html": sample_html,
    })


# ── Bulk Sending ──
@app.post("/api/send")
async def start_send(
    request: Request,
):
    global _job_thread, _stop_flag
    
    body = await request.json()
    subject = body.get("subject", "")
    html_body = body.get("html", "")
    email_col = int(body.get("email_column", 0))
    from_name = body.get("from_name", "")
    min_delay = int(body.get("min_delay", 2))
    max_delay = int(body.get("max_delay", 10))
    test_mode = body.get("test_mode", False)
    test_email = body.get("test_email", "")
    round_params = body.get("round_params", {}) or {}
    
    config = load_smtp_config()
    if not config:
        return JSONResponse({"ok": False, "error": "SMTP not configured"}, status_code=400)
    
    if not excel_rows_cache:
        return JSONResponse({"ok": False, "error": "No Excel uploaded"}, status_code=400)
    
    current = get_state()
    if current.get("status") == "running":
        return JSONResponse({"ok": False, "error": "A job is already running"}, status_code=409)
    
    _stop_flag.clear()
    
    # Save job params
    job_data = {
        "subject": subject,
        "html": html_body,
        "email_col": email_col,
        "from_name": from_name or config.get("sender_name", ""),
        "min_delay": min_delay,
        "max_delay": max_delay,
        "test_mode": test_mode,
        "test_email": test_email,
        "round_params": round_params,
        "config": config,
        "rows": [[str(c) if c is not None else "" for c in row] for row in excel_rows_cache],
        "headers": excel_columns_cache,
    }
    
    _job_thread = threading.Thread(
        target=_run_send_job,
        args=(job_data,),
        daemon=True,
    )
    _job_thread.start()
    
    return JSONResponse({"ok": True, "status": "running", "total": len(excel_rows_cache)})


def _run_send_job(job_data):
    subject = job_data["subject"]
    html_body = job_data["html"]
    email_col = job_data["email_col"]
    from_name = job_data["from_name"]
    min_delay = job_data["min_delay"]
    max_delay = job_data["max_delay"]
    test_mode = job_data["test_mode"]
    test_email = job_data["test_email"]
    round_params = job_data.get("round_params", {}) or {}
    config = job_data["config"]
    rows = job_data["rows"]
    headers = job_data["headers"]
    
    total = len(rows)
    sent = 0
    failed = 0
    log = []
    
    _save_state({
        "status": "running",
        "total": total,
        "sent": 0,
        "failed": 0,
        "current_row": 0,
        "log": [],
        "started_at": datetime.now().isoformat(),
    })
    
    for idx, row in enumerate(rows):
        if _stop_flag.is_set():
            _save_state({
                "status": "stopped",
                "total": total,
                "sent": sent,
                "failed": failed,
                "current_row": idx,
                "log": log[-100:],
            })
            return
        
        # Build variable map
        variables = {}
        for i, h in enumerate(headers):
            variables[h] = row[i] if i < len(row) else ""
        
        # Get recipient
        recipient = test_email if test_mode and test_email else (row[email_col] if email_col < len(row) else "")
        
        if not recipient or "@" not in str(recipient):
            failed += 1
            log.append({"row": idx + 1, "email": str(recipient), "ok": False, "error": "No valid email"})
            _save_state({
                "status": "running",
                "total": total,
                "sent": sent,
                "failed": failed,
                "current_row": idx + 1,
                "log": log[-100:],
            })
            continue
        
        # Replace variables in subject and body
        # 1) First apply round-level fixed params (same for all emails this round)
        personalized_subject = subject
        personalized_html = html_body
        for rp_name, rp_val in round_params.items():
            if rp_val is None:
                rp_val = ""
            rp_val = str(rp_val)
            personalized_subject = personalized_subject.replace("{" + rp_name + "}", rp_val)
            personalized_html = personalized_html.replace("{" + rp_name + "}", rp_val)
        
        # 2) Then replace per-row Excel variables (overrides round params if same name)
        import re
        for var_name, var_val in variables.items():
            personalized_subject = personalized_subject.replace("{" + var_name + "}", str(var_val))
            personalized_html = personalized_html.replace("{" + var_name + "}", str(var_val))
        
        # Also replace {name} with email prefix if not in columns
        if "{name}" in personalized_html and "name" not in headers:
            name_part = str(recipient).split("@")[0]
            personalized_html = personalized_html.replace("{name}", name_part)
        if "{name}" in personalized_subject and "name" not in headers:
            name_part = str(recipient).split("@")[0]
            personalized_subject = personalized_subject.replace("{name}", name_part)
        
        # Send email
        try:
            msg = EmailMessage()
            msg["From"] = formataddr((from_name or "", config["username"]))
            msg["To"] = recipient
            msg["Subject"] = personalized_subject
            msg.set_content("Please enable HTML to view this email.")
            msg.add_alternative(personalized_html, subtype="html")
            
            with smtplib.SMTP_SSL(config["host"], config["port"], timeout=30) as srv:
                srv.login(config["username"], config["password"])
                srv.send_message(msg)
            
            sent += 1
            log.append({"row": idx + 1, "email": recipient, "ok": True, "error": ""})
        except Exception as e:
            failed += 1
            log.append({"row": idx + 1, "email": recipient, "ok": False, "error": str(e)})
        
        _save_state({
            "status": "running",
            "total": total,
            "sent": sent,
            "failed": failed,
            "current_row": idx + 1,
            "log": log[-100:],
        })
        
        # Random delay (skip after last email)
        if idx < total - 1 and not _stop_flag.is_set():
            delay = random.uniform(min_delay, max_delay)
            time.sleep(delay)
    
    final_status = "completed" if not _stop_flag.is_set() else "stopped"
    _save_state({
        "status": final_status,
        "total": total,
        "sent": sent,
        "failed": failed,
        "current_row": total,
        "log": log[-100:],
        "finished_at": datetime.now().isoformat(),
    })


@app.post("/api/stop")
def stop_send():
    global _stop_flag
    _stop_flag.set()
    return JSONResponse({"ok": True, "message": "Stop requested"})


@app.get("/api/status")
def status():
    s = get_state()
    return JSONResponse(s)


# ── Frontend ──
@app.get("/", response_class=HTMLResponse)
def index():
    return FRONTEND_HTML


FRONTEND_HTML = r"""<!DOCTYPE html>
<html lang="zh">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>Mail Blaster — 批量个性化邮件发送</title>
<!-- CodeMirror — HTML code editor with syntax highlighting -->
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.16/codemirror.min.css">
<link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.16/theme/material-darker.min.css">
<script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.16/codemirror.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.16/mode/htmlmixed/htmlmixed.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.16/mode/xml/xml.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.16/mode/javascript/javascript.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.16/mode/css/css.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.16/addon/edit/matchbrackets.min.js"></script>
<script src="https://cdnjs.cloudflare.com/ajax/libs/codemirror/5.65.16/addon/edit/closebrackets.min.js"></script>
<style>
*{margin:0;padding:0;box-sizing:border-box}
body{font-family:-apple-system,BlinkMacSystemFont,'Segoe UI',Roboto,sans-serif;background:#f0f2f5;color:#333;font-size:14px}
.header{background:#1a73e8;color:#fff;padding:16px 24px;display:flex;align-items:center;justify-content:space-between}
.header h1{font-size:20px;font-weight:600}
.container{max-width:900px;margin:0 auto;padding:20px}
.card{background:#fff;border-radius:8px;padding:20px;margin-bottom:16px;box-shadow:0 1px 3px rgba(0,0,0,0.08)}
.card h2{font-size:16px;margin-bottom:12px;color:#1a73e8}
label{display:block;font-weight:600;margin-bottom:4px;font-size:13px}
input[type=text],input[type=email],input[type=number],input[type=password],select{width:100%;padding:8px 10px;border:1px solid #ddd;border-radius:6px;font-size:13px;margin-bottom:10px}
input[type=file]{margin-bottom:10px}
.btn{display:inline-block;padding:8px 20px;border:none;border-radius:6px;cursor:pointer;font-size:14px;font-weight:500;text-decoration:none}
.btn-primary{background:#1a73e8;color:#fff}
.btn-success{background:#059669;color:#fff}
.btn-danger{background:#dc2626;color:#fff}
.btn-secondary{background:#6b7280;color:#fff}
.btn:disabled{opacity:0.5;cursor:not-allowed}
.grid2{display:grid;grid-template-columns:1fr 1fr;gap:12px}

/* HTML Editor */
.editor-tabs{display:flex;border-bottom:1px solid #ddd;margin-bottom:8px;gap:4px}
.editor-tab{padding:8px 16px;cursor:pointer;font-size:13px;font-weight:600;color:#6b7280;border-bottom:2px solid transparent;margin-bottom:-1px;user-select:none}
.editor-tab.active{color:#1a73e8;border-bottom-color:#1a73e8}
.editor-tab:hover:not(.active){color:#374151;background:#f3f4f6}
.CodeMirror{border-radius:6px;font-size:13px;font-family:'JetBrains Mono','Fira Code',Consolas,monospace;border:1px solid #ddd}
.CodeMirror-focused{outline:none!important;border-color:#1a73e8!important}
.preview-pane{border:1px solid #ddd;border-radius:6px;padding:16px;background:#fff;min-height:380px;max-height:380px;overflow-y:auto}
.preview-pane img{max-width:100%}
.preview-empty{color:#9ca3af;text-align:center;padding:60px 0;font-size:13px}

/* Image gallery */
.image-gallery{display:grid;grid-template-columns:repeat(auto-fill,minmax(110px,1fr));gap:8px;margin-top:8px}
.image-thumb{border:1px solid #ddd;border-radius:6px;padding:6px;cursor:pointer;text-align:center;background:#fafafa;transition:all 0.15s}
.image-thumb:hover{background:#eef2ff;border-color:#1a73e8;transform:translateY(-1px)}
.image-thumb img{width:100%;height:70px;object-fit:cover;border-radius:4px;display:block;margin-bottom:4px}
.image-thumb .name{font-size:11px;color:#6b7280;word-break:break-all}

.stats{display:flex;gap:20px;margin:12px 0}
.stat{text-align:center;padding:10px 20px;border-radius:8px;background:#f8f9fa}
.stat .num{font-size:24px;font-weight:700}
.stat .label{font-size:12px;color:#6b7280}
.progress-bar{width:100%;height:8px;background:#e5e7eb;border-radius:4px;overflow:hidden;margin:8px 0}
.progress-fill{height:100%;background:#1a73e8;transition:width 0.3s}
.log-box{max-height:200px;overflow-y:auto;border:1px solid #eee;border-radius:6px;padding:8px;font-size:12px;font-family:monospace}
.log-line{padding:2px 0;border-bottom:1px solid #f5f5f5}
.log-ok{color:#059669}
.log-fail{color:#dc2626}
.badge{display:inline-block;padding:2px 8px;border-radius:12px;font-size:11px;font-weight:600}
.badge-ok{background:#d1fae5;color:#065f46}
.badge-no{background:#fee2e2;color:#991b1b}
.hint{font-size:12px;color:#6b7280;margin-top:4px}
</style>
</head>
<body>

<div class="header">
  <h1>📧 Mail Blaster</h1>
  <span style="font-size:13px;opacity:0.9">批量个性化邮件发送 · 随机间隔防垃圾</span>
</div>

<div class="container">

<!-- Step 1: SMTP -->
<div class="card">
  <h2>1️⃣ SMTP 配置</h2>
  <div id="smtp-status" style="margin-bottom:10px;font-size:13px;"></div>
  <div id="smtp-form">
    <div class="grid2">
      <div><label>SMTP Host</label><input type="text" id="smtp-host" placeholder="smtp.gmail.com"></div>
      <div><label>Port</label><input type="number" id="smtp-port" value="465"></div>
      <div><label>用户名 (Email)</label><input type="text" id="smtp-user" placeholder="your@gmail.com"></div>
      <div><label>密码 (App Password)</label><input type="password" id="smtp-pass" placeholder="xxxx xxxx xxxx xxxx"></div>
      <div><label>发件人名称</label><input type="text" id="smtp-name" placeholder="Luna Hei"></div>
    </div>
    <button class="btn btn-primary" onclick="configureSMTP()">连接测试 & 保存</button>
  </div>
</div>

<!-- Step 2: Excel -->
<div class="card">
  <h2>2️⃣ 上传 Excel</h2>
  <input type="file" id="excel-file" accept=".xlsx,.xls">
  <button class="btn btn-primary" onclick="uploadExcel()">上传</button>
  <div id="excel-info" style="margin-top:8px;font-size:13px;"></div>
  <div id="email-col-select" style="display:none;margin-top:8px;">
    <label>选择邮箱所在列</label>
    <select id="email-column"></select>
    <button class="btn btn-secondary" onclick="setEmailColumn()">确认</button>
  </div>
</div>

<!-- Step 3: Round Params (fixed for this batch, e.g. category/URL/picture) -->
<div class="card" id="round-params-card">
  <h2>3️⃣ 本轮统一参数 <span style="font-weight:normal;font-size:12px;color:#888;">(本批所有邮件一致，例如产品分类/链接/图片)</span></h2>
  <div class="grid2">
    <div>
      <label>产品分类 <span class="snippet">{category_type}</span></label>
      <input type="text" id="rp-category_type" placeholder="例如: Holiday Decor">
    </div>
    <div>
      <label>产品链接 <span class="snippet">{product_url}</span></label>
      <input type="text" id="rp-product_url" placeholder="https://...">
    </div>
  </div>
  <div>
    <label>产品图片 URL <span class="snippet">{picture_url}</span></label>
    <input type="text" id="rp-picture_url" placeholder="https://你的CDN图片地址...">
  </div>
  <p class="hint">💡 这三个变量在本轮所有邮件中一致。如果 Excel 中有同名列，会以 Excel 行为准（更灵活）。</p>
</div>

<!-- Step 4: Email Content -->
<div class="card">
  <h2>4️⃣ 邮件内容 (HTML)</h2>
  <label>邮件主题</label>
  <input type="text" id="subject" placeholder="Hi {handle}, TikTok Shop product opportunities from Liveology US" value="Hi {handle}, TikTok Shop product opportunities from Liveology US">
  <p class="hint">
    💡 变量用 <span class="snippet">{"{"}列名{"}"}</span> 替换（自动识别 Excel 列名）：
    <span class="snippet">{"{"}handle{"}"}</span> / <span class="snippet">{"{"}name{"}"}</span> / <span class="snippet">{"{"}email{"}"}</span> 等
  </p>

  <!-- Image gallery (uploaded images for click-to-insert) -->
  <div style="margin-bottom:8px;">
    <label style="display:inline;">图片库 (点缩略图插入)</label>
    <input type="file" id="inline-image-file" accept="image/*" style="display:inline-block;margin-left:8px;margin-bottom:0;width:auto;">
    <button class="btn btn-secondary" style="padding:4px 10px;font-size:12px;" onclick="uploadInlineImage()">上传并转 base64</button>
    <br>
    <input type="text" id="image-url-input" placeholder="或粘贴图片 CDN URL —— https://..." style="display:inline-block;margin-top:6px;width:380px;margin-bottom:6px;">
    <button class="btn btn-secondary" style="padding:4px 10px;font-size:12px;" onclick="addImageByUrl()">+ 通过 URL 添加</button>
    <div id="image-gallery" class="image-gallery"></div>
  </div>

  <!-- HTML Editor (CodeMirror) -->
  <div class="editor-tabs">
    <div class="editor-tab active" data-tab="code" onclick="switchEditorTab('code')">📝 HTML 源码</div>
    <div class="editor-tab" data-tab="preview" onclick="switchEditorTab('preview')">👁️ 实时预览</div>
  </div>
  <div id="code-pane"><textarea id="html-editor"></textarea></div>
  <div id="preview-pane" class="preview-pane" style="display:none;">
    <div id="preview-content" class="preview-empty">切换到 "HTML 源码" 编辑内容后，这里会实时预览。</div>
  </div>

  <button class="btn btn-secondary" onclick="previewVariables()">预览变量替换</button>
  <button class="btn btn-secondary" onclick="formatHtml()">格式化</button>
  <div id="var-preview" style="margin-top:8px;border:1px solid #eee;border-radius:6px;padding:12px;display:none;"></div>
</div>

<!-- Step 4: Send Settings -->
<div class="card">
  <h2>5️⃣ 发送设置</h2>
  <div class="grid2">
    <div><label>最小间隔 (秒)</label><input type="number" id="min-delay" value="2"></div>
    <div><label>最大间隔 (秒)</label><input type="number" id="max-delay" value="10"></div>
  </div>
  <label><input type="checkbox" id="test-mode" checked style="width:auto;margin-right:6px;">TEST MODE（只发到测试邮箱，不发给真实收件人）</label>
  <div id="test-email-box" style="margin-top:8px;">
    <label>测试收件邮箱</label>
    <input type="email" id="test-email" placeholder="your-test@gmail.com">
  </div>
</div>

<!-- Step 5: Send -->
<div class="card">
  <h2>6️⃣ 发送</h2>
  <div style="display:flex;gap:10px;margin-bottom:12px;">
    <button class="btn btn-success" id="btn-send" onclick="startSend()">开始发送</button>
    <button class="btn btn-danger" id="btn-stop" onclick="stopSend()" disabled>停止</button>
  </div>
  
  <div class="stats">
    <div class="stat"><div class="num" id="stat-total">0</div><div class="label">总数</div></div>
    <div class="stat"><div class="num" id="stat-sent" style="color:#059669">0</div><div class="label">已发送</div></div>
    <div class="stat"><div class="num" id="stat-failed" style="color:#dc2626">0</div><div class="label">失败</div></div>
    <div class="stat"><div class="num" id="stat-remain">0</div><div class="label">剩余</div></div>
  </div>
  
  <div class="progress-bar"><div class="progress-fill" id="progress-fill" style="width:0%"></div></div>
  
  <div class="log-box" id="log-box">
    <div class="log-line" style="color:#999;">等待发送...</div>
  </div>
</div>

</div>

<script>
// ── CodeMirror HTML Editor + Preview ──
const DEFAULT_HTML = `<html>
<body style="font-family:'Helvetica Neue',Helvetica,Arial,sans-serif;font-size:14px;line-height:1.6;color:#333333;max-width:600px;margin:0 auto;padding:20px;">

<p style="margin-bottom:10px;">Hi {handle},</p>

<p style="margin-bottom:12px;">I'm Luna from Liveology US, a <strong>TikTok Top Strategic Partner</strong> based in New York City.</p>

<p style="margin-bottom:20px;">We're inviting selected TikTok creators to explore <strong>{category_type}</strong> products with free-sample and exclusive high-commission opportunities.</p>
<!-- 产品图片：使用本轮统一参数 {picture_url}（在"3️⃣ 本轮统一参数"里填 CDN 图片地址） -->
<p style="margin:0 0 6px 0;text-align:center;">
  <img src="{picture_url}" alt="{category_type}" style="display:block;max-width:100%;width:100%;height:auto;border-radius:10px;border:1px solid #f0f0f0;">
</p>

<!-- 产品链接 -->
<p style="margin:0 0 24px 0;text-align:center;">
  <a href="{product_url}" style="color:#0066cc;text-decoration:underline;font-size:15px;">🔗 View the product</a>
</p>

<!-- 核心 Benefit 词组区 -->
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:20px;">
  <tr><td style="background:#f8f9fa;border-radius:8px;padding:16px 20px;">
    <p style="margin:0 0 10px 0;"><strong>What's in it for you:</strong></p>
    <p style="margin:0 0 6px 0;">💯 <strong>100% FREE</strong> sample</p>
    <p style="margin:0 0 6px 0;">📦 Keep the product — it's yours</p>
    <p style="margin:0 0 6px 0;">💰 Earn <strong>high commission</strong> on every sale</p>
    <p style="margin:0 0 0 0;">🎥 Perfect content for your audience</p>
  </td></tr>
</table>

<!-- 3 步行动 -->
<p style="margin:0 0 10px 0;"><strong>Claim in 3 steps:</strong></p>
<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="margin-bottom:20px;">
  <tr><td style="background:#ffffff;border:2px solid #f0f0f0;border-radius:8px;padding:14px 20px;">
    <p style="margin:0 0 4px 0;">1️⃣ Click the link</p>
    <p style="margin:0 0 4px 0;">2️⃣ Browse the picks</p>
    <p style="margin:0 0 0 0;">3️⃣ Apply for free samples</p>
  </td></tr>
</table>

<!-- CTA 按钮（改颜色适配品牌色） -->
<p style="margin:0 0 24px 0;text-align:center;">
  <a href="{product_url}" style="display:inline-block;background:#DC1F26;color:#ffffff;padding:14px 32px;border-radius:6px;text-decoration:none;font-weight:bold;font-size:16px;">👉 Claim My FREE Sample</a>
</p>

<p style="margin:0 0 8px 0;color:#666666;">Want more opportunities like this? Reply <strong>1</strong> — we'll send weekly updates.</p>
<p style="margin:0 0 20px 0;color:#666666;">You can opt out anytime.</p>

<p style="margin:0 0 20px 0;">Looking forward to working with you!</p>

<p style="margin:0;">Best,<br>
Luna Hei<br>
Liveology US<br>
NYC Office<br>
1350 Avenue of the Americas, Floor 2<br>
New York, NY 10019</p>

</body>
</html>`;

const editor = CodeMirror.fromTextArea(document.getElementById('html-editor'), {
  mode: 'htmlmixed',
  theme: 'material-darker',
  lineNumbers: true,
  matchBrackets: true,
  autoCloseBrackets: true,
  indentUnit: 2,
  tabSize: 2,
  lineWrapping: true,
  extraKeys: {
    'Tab': cm => cm.replaceSelection('  ', 'end')
  }
});
editor.setValue(DEFAULT_HTML);

// Tab switching
let currentTab = 'code';
function switchEditorTab(tab) {
  currentTab = tab;
  document.querySelectorAll('.editor-tab').forEach(t => t.classList.toggle('active', t.dataset.tab === tab));
  document.getElementById('code-pane').style.display = tab === 'code' ? '' : 'none';
  document.getElementById('preview-pane').style.display = tab === 'preview' ? '' : 'none';
  if (tab === 'preview') updatePreview();
}

// Live preview (debounced)
let previewTimer = null;
editor.on('change', () => {
  if (currentTab === 'preview') {
    clearTimeout(previewTimer);
    previewTimer = setTimeout(updatePreview, 200);
  }
});

function updatePreview() {
  const html = editor.getValue();
  const previewContent = document.getElementById('preview-content');
  if (!html.trim()) {
    previewContent.className = 'preview-empty';
    previewContent.textContent = 'HTML 内容为空。';
    return;
  }
  previewContent.className = '';
  // Use srcdoc to isolate the HTML (no global style leaks)
  const iframe = document.createElement('iframe');
  iframe.style.cssText = 'width:100%;height:340px;border:none;background:#fff;border-radius:4px;';
  previewContent.innerHTML = '';
  previewContent.appendChild(iframe);
  iframe.srcdoc = html;
}

// Pretty-print HTML (best-effort, browser built-in)
function formatHtml() {
  try {
    const html = editor.getValue();
    // Use DOMParser to re-serialize with indentation
    const doc = new DOMParser().parseFromString('<div>' + html + '</div>', 'text/html');
    const formatted = doc.body.firstChild.outerHTML
      .replace(/></g, '>\n<')
      .replace(/^\s+|\s+$/gm, '');
    editor.setValue(formatted);
  } catch (e) {
    alert('格式化失败：' + e.message);
  }
}

// ── Image upload (converts to base64 for inline embed) ──
async function uploadInlineImage() {
  const f = document.getElementById('inline-image-file').files[0];
  if (!f) { alert('请选择图片'); return; }
  // Convert to base64
  const reader = new FileReader();
  reader.onload = function(e) {
    const dataUrl = e.target.result;
    addImageToGallery(f.name, dataUrl);
    // Auto-insert into editor
    const imgTag = `<img src="${dataUrl}" alt="${f.name}" style="max-width:100%;border-radius:8px;margin:8px 0;">`;
    const cursor = editor.getCursor();
    editor.replaceRange('\n' + imgTag + '\n', cursor);
  };
  reader.readAsDataURL(f);
}

function addImageToGallery(name, url) {
  const gallery = document.getElementById('image-gallery');
  const div = document.createElement('div');
  div.className = 'image-thumb';
  div.innerHTML = `<img src="${url}" alt="${name}"><div class="name">${name}</div>`;
  div.onclick = () => {
    const imgTag = `<img src="${url}" alt="${name}" style="max-width:100%;border-radius:8px;margin:8px 0;">`;
    const cursor = editor.getCursor();
    editor.replaceRange('\n' + imgTag + '\n', cursor);
    editor.focus();
  };
  gallery.appendChild(div);
}

// Add image to gallery via URL (no file upload needed)
function addImageByUrl() {
  const urlInput = document.getElementById('image-url-input');
  const url = (urlInput.value || '').trim();
  if (!url) { alert('请先粘贴图片 URL'); return; }
  // Basic URL validation
  if (!/^https?:\/\/.+/.test(url)) { alert('URL 格式不对，请以 http(s):// 开头'); return; }
  
  // Extract a name from URL (last path segment)
  let name = url.split('/').filter(Boolean).pop() || 'image';
  // Clean query params & truncate
  name = decodeURIComponent(name.split('?')[0]).slice(0, 40) || 'image';
  
  // Preload to verify it's loadable, then add to gallery
  const probe = new Image();
  probe.onload = () => {
    addImageToGallery(name, url);
    urlInput.value = '';  // clear input
    alert('✅ 图片已添加到图片库，点击缩略图即可插入');
  };
  probe.onerror = () => {
    // Still add it — URL may be fine but blocked in preview (e.g. some CDNs)
    addImageToGallery(name, url);
    urlInput.value = '';
    alert('⚠️ 图片 URL 已添加（预览器可能无法加载某些 CDN 图片，但邮件中通常可正常显示）');
  };
  probe.src = url;
}

// Helper to get current HTML content
function getEditorHtml() {
  return editor.getValue();
}

// ── API helpers ──
async function api(path, opts={}) {
  const res = await fetch(path, opts);
  const text = await res.text();
  try { return JSON.parse(text); } catch { return {ok:false, error:text}; }
}

// ── SMTP ──
async function checkSMTP() {
  const d = await api('/api/smtp/status');
  const el = document.getElementById('smtp-status');
  if (d.configured) {
    el.innerHTML = '<span class="badge badge-ok">✅ 已配置</span> ' + d.host + ':' + d.port + ' (' + d.username + ')' +
      (d.sender_name ? ' · ' + d.sender_name : '');
    document.getElementById('smtp-host').value = d.host;
    document.getElementById('smtp-port').value = d.port;
    document.getElementById('smtp-user').value = d.username;
    document.getElementById('smtp-name').value = d.sender_name || '';
  } else {
    el.innerHTML = '<span class="badge badge-no">❌ 未配置</span> 请填写 SMTP 信息';
  }
}

async function configureSMTP() {
  const fd = new FormData();
  fd.append('host', document.getElementById('smtp-host').value);
  fd.append('port', document.getElementById('smtp-port').value);
  fd.append('username', document.getElementById('smtp-user').value);
  fd.append('password', document.getElementById('smtp-pass').value);
  fd.append('sender_name', document.getElementById('smtp-name').value);
  const d = await api('/api/smtp/configure', {method:'POST', body:fd});
  if (d.ok) {
    alert('✅ SMTP 连接成功！');
    checkSMTP();
  } else {
    alert('❌ SMTP 连接失败: ' + (d.error||''));
  }
}

// ── Excel ──
let emailColIdx = null;
let rowCount = 0;

async function uploadExcel() {
  const f = document.getElementById('excel-file').files[0];
  if (!f) { alert('请选择文件'); return; }
  const fd = new FormData();
  fd.append('file', f);
  const d = await api('/api/upload_excel', {method:'POST', body:fd});
  const info = document.getElementById('excel-info');
  if (d.ok) {
    rowCount = d.row_count;
    info.innerHTML = '✅ 共 ' + d.row_count + ' 行 · 列: ' + d.headers.join(', ');
    if (d.needs_column_select) {
      document.getElementById('email-col-select').style.display = 'block';
      const sel = document.getElementById('email-column');
      sel.innerHTML = d.headers.map((h,i) => '<option value="'+i+'">'+(i+1)+': '+h+'</option>').join('');
    } else {
      emailColIdx = d.email_column;
      info.innerHTML += '<br>📬 邮箱列: ' + d.email_column_name;
    }
  } else {
    info.innerHTML = '❌ ' + (d.error||'上传失败');
  }
}

async function setEmailColumn() {
  const idx = parseInt(document.getElementById('email-column').value);
  const d = await api('/api/set_email_column', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({column_index:idx})});
  if (d.ok) {
    emailColIdx = idx;
    document.getElementById('excel-info').innerHTML += '<br>📬 邮箱列: ' + d.email_column_name;
    document.getElementById('email-col-select').style.display = 'none';
  }
}

// ── Variable Preview ──
async function previewVariables() {
  const html = getEditorHtml();
  const rp = collectRoundParams();
  const d = await api('/api/preview_variables', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({html, round_params: rp})});
  const box = document.getElementById('var-preview');
  if (d.ok) {
    box.style.display = 'block';
    let html = '<strong>检测到的变量:</strong> ';
    if (d.variables.length === 0) {
      html += '<span style="color:#999;">无变量</span>';
    } else {
      html += d.variables.map(v => {
        const m = d.matched[v];
        return m ? '<span class="badge badge-ok">{'+'{'+v+'}'+'} → '+m+'</span>' : '<span class="badge badge-no">{'+'{'+v+'}'+'} 未匹配</span>';
      }).join(' ');
    }
    html += '<hr style="margin:8px 0;"><strong>预览 (第一行数据):</strong><br>' + d.sample_html;
    box.innerHTML = html;
  }
}

// ── Send ──
let pollTimer = null;

function collectRoundParams() {
  const params = {};
  ['category_type', 'product_url', 'picture_url'].forEach(key => {
    const el = document.getElementById('rp-' + key);
    if (el && el.value.trim()) params[key] = el.value.trim();
  });
  return params;
}

async function startSend() {
  const subject = document.getElementById('subject').value;
  const html = getEditorHtml();
  if (!subject) { alert('请填写邮件主题'); return; }
  if (!html || !html.trim()) { alert('请填写邮件正文'); return; }
  if (emailColIdx === null) { alert('请先上传 Excel 并选择邮箱列'); return; }
  
  const testMode = document.getElementById('test-mode').checked;
  const testEmail = document.getElementById('test-email').value;
  if (testMode && !testEmail) { alert('TEST MODE 需要填写测试邮箱'); return; }
  
  const body = {
    subject, html,
    email_column: emailColIdx,
    from_name: document.getElementById('smtp-name').value,
    min_delay: parseInt(document.getElementById('min-delay').value),
    max_delay: parseInt(document.getElementById('max-delay').value),
    test_mode: testMode,
    test_email: testEmail,
    round_params: collectRoundParams(),
  };
  
  const d = await api('/api/send', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(body)});
  if (d.ok) {
    document.getElementById('btn-send').disabled = true;
    document.getElementById('btn-stop').disabled = false;
    startPolling();
  } else {
    alert('❌ ' + (d.error||'发送失败'));
  }
}

async function stopSend() {
  await api('/api/stop', {method:'POST'});
  alert('停止请求已发送，正在等待当前邮件发完...');
}

function startPolling() {
  if (pollTimer) clearInterval(pollTimer);
  pollTimer = setInterval(updateProgress, 2000);
  updateProgress();
}

async function updateProgress() {
  const d = await api('/api/status');
  document.getElementById('stat-total').textContent = d.total || 0;
  document.getElementById('stat-sent').textContent = d.sent || 0;
  document.getElementById('stat-failed').textContent = d.failed || 0;
  document.getElementById('stat-remain').textContent = (d.total||0) - (d.sent||0) - (d.failed||0);
  
  const pct = d.total > 0 ? ((d.sent + d.failed) / d.total * 100) : 0;
  document.getElementById('progress-fill').style.width = pct + '%';
  
  // Log
  const lb = document.getElementById('log-box');
  const log = d.log || [];
  lb.innerHTML = log.slice(-20).map(e =>
    '<div class="log-line ' + (e.ok ? 'log-ok' : 'log-fail') + '">' +
    'Row ' + e.row + ' · ' + e.email + ' · ' + (e.ok ? '✅' : '❌ ' + e.error) +
    '</div>'
  ).join('');
  lb.scrollTop = lb.scrollHeight;
  
  // Status check
  if (d.status && d.status !== 'running' && d.status !== 'idle') {
    document.getElementById('btn-send').disabled = false;
    document.getElementById('btn-stop').disabled = true;
    if (d.status === 'completed') {
      if (pollTimer) clearInterval(pollTimer);
    }
  }
}

// ── Test mode toggle ──
document.getElementById('test-mode').addEventListener('change', function() {
  document.getElementById('test-email-box').style.display = this.checked ? 'block' : 'none';
});

// ── Init ──
checkSMTP();
updateProgress();
</script>
</body>
</html>
"""


# ── Start server ──
if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8003))
    uvicorn.run(app, host="0.0.0.0", port=port)

