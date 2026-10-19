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
from fastapi.responses import HTMLResponse, JSONResponse, FileResponse, Response, RedirectResponse
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


def _deny():
    """Return a 401 response that triggers the browser native password prompt."""
    return Response(
        status_code=401,
        headers={"WWW-Authenticate": 'Basic realm="Mail Blaster", charset="UTF-8"'},
    )


@app.middleware("http")
async def access_guard(request: Request, call_next):
    """Protect every route: /, /api/*, uploads, everything."""
    path = request.url.path
    # Health check must stay open for Railway monitoring.
    # SMTP diag is also exempt: it only probes a fixed allowlist of mail hosts
    # (no secrets returned, no email sent), needed for remote debugging.
    if path in ("/api/health", "/api/smtp/diag", "/api/google/oauth/callback"):
        return await call_next(request)
    if not _check_access(request):
        return _deny()
    return await call_next(request)

BASE_DIR = Path(__file__).parent

# Container start time, for /api/health: lets us tell whether a redeploy actually
# happened after Railway variables were edited.
_BOOT_TS = time.time()
_BOOT_AT = datetime.now().isoformat(timespec="seconds")

# Use Railway persistent volume if mounted, otherwise local directory
UPLOAD_DIR = Path(os.environ.get("UPLOAD_DIR", str(BASE_DIR / "uploads")))
UPLOAD_DIR.mkdir(parents=True, exist_ok=True)

# SMTP config: prefer environment variables (Railway-safe), fall back to local file
SMTP_CONFIG_FILE = BASE_DIR / ".smtp_config.json"
JOB_STATE_FILE = Path(os.environ.get("JOB_STATE_FILE", str(BASE_DIR / "data" / "job_state.json")))
JOB_STATE_FILE.parent.mkdir(parents=True, exist_ok=True)


def _open_smtp(host, port, timeout=60):
    """Open an SMTP connection: implicit SSL (typically 465) or STARTTLS (587/25)."""
    port = int(port)
    if port == 465:
        return smtplib.SMTP_SSL(host, port, timeout=timeout)
    srv = smtplib.SMTP(host, port, timeout=timeout)
    srv.ehlo()
    if srv.has_extn("starttls"):
        import ssl
        ctx = ssl.create_default_context()
        srv.starttls(context=ctx)
        srv.ehlo()
    return srv


def _diagnose_smtp_targets():
    """Candidate SMTP endpoints to probe for Lark/Feishu mailboxes."""
    return [
        ("smtp.larksuite.com", 465),
        ("smtp.larksuite.com", 587),
        ("smtp.feishu.cn", 465),
        ("smtp.feishu.cn", 587),
    ]


def _diagnose_https_targets():
    """Fixed allowlist of Lark Open API hosts (443). Used to verify the HTTP API
    path is reachable from this server's network before building on it."""
    return [
        "https://open.larksuite.com",
        "https://open.feishu.cn",
    ]


def _probe_https(url, timeout=20):
    """GET a fixed allowlisted URL; report reachability only (no data exfil)."""
    import urllib.request
    try:
        req = urllib.request.Request(url, method="GET", headers={"User-Agent": "smtp-diag/1.0"})
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return {"url": url, "connect": "OK", "http_status": resp.status}
    except Exception as e:
        status = getattr(e, "code", None)
        if status:  # HTTP error still proves the host is reachable
            return {"url": url, "connect": "OK", "http_status": status}
        return {"url": url, "connect": "FAILED", "error": f"{type(e).__name__}: {str(e)[:100]}"}


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


def _job_alive():
    """True only when a send worker is actually running in THIS process.

    Note: _job_thread is None right after a container restart, which counts as
    "not alive" — this is what lets stale persisted states self-heal.
    """
    return _job_thread is not None and _job_thread.is_alive()


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
    # Heartbeat: lets the UI tell whether these counters are live or leftovers
    state = dict(state)
    state["updated_at"] = datetime.now().isoformat()
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
    # Self-heal stale states. After a container restart/redeploy _job_thread is
    # None, so "no thread running" must also cover the None case — otherwise a
    # persisted 'running' state locks /api/send forever and the counters freeze.
    if s.get("status") == "running" and not _job_alive():
        s["status"] = "interrupted"
        s["interrupted_reason"] = "no live send worker in this process (stale state file)"
        _save_state(s)
    return s


# ── Upload History (persisted to volume, survives restarts) ──
UPLOAD_HISTORY_FILE = Path(os.environ.get("UPLOAD_HISTORY_FILE",
                                          str(BASE_DIR / "data" / "upload_history.json")))


def _load_upload_history():
    """Return list of {upload_id, original_name, saved_name, uploaded_at, headers, row_count, source}."""
    if not UPLOAD_HISTORY_FILE.exists():
        return []
    try:
        with open(UPLOAD_HISTORY_FILE) as f:
            data = json.load(f)
            return data if isinstance(data, list) else []
    except (json.JSONDecodeError, IOError):
        return []


def _append_upload_history(record):
    UPLOAD_HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
    history = _load_upload_history()
    history.append(record)
    if len(history) > 200:
        history = history[-200:]
    try:
        tmp = UPLOAD_HISTORY_FILE.with_suffix(".json.tmp")
        with open(tmp, "w") as f:
            json.dump(history, f, indent=2, ensure_ascii=False)
        os.replace(tmp, UPLOAD_HISTORY_FILE)
    except (IOError, OSError):
        pass


# ── Excel Processing ──
excel_columns_cache = []
excel_rows_cache = []
excel_path_cache = None
excel_history_id_cache = None   # upload record id, propagated to job_data + log rows
excel_original_name_cache = None  # original filename as uploaded by user


@app.post("/api/upload_excel")
async def upload_excel(file: UploadFile = File(...)):
    global excel_columns_cache, excel_rows_cache, excel_path_cache, excel_history_id_cache, excel_original_name_cache

    # Sanitize filename: strip path components to prevent path traversal
    safe_name = os.path.basename(file.filename or "upload.xlsx")
    if not safe_name:
        return JSONResponse({"ok": False, "error": "Invalid filename"}, status_code=400)
    # Add timestamp prefix to avoid collisions
    save_path = UPLOAD_DIR / f"{int(time.time())}_{safe_name}"

    # Save file
    with open(save_path, "wb") as f:
        f.write(await file.read())

    # Parse (workbook is always closed via finally to avoid file handle leak)
    wb = load_workbook(save_path, read_only=True, data_only=True)
    try:
        ws = wb.active
        rows = list(ws.iter_rows(values_only=True))
        if not rows:
            return JSONResponse({"ok": False, "error": "Empty Excel"}, status_code=400)

        headers = [str(h).strip() if h else f"col_{i+1}" for i, h in enumerate(rows[0])]

        # Register in persistent history (survives restarts) so future sends can be
        # traced back to the exact file the emails came from.
        upload_id = f"{time.time_ns()}"  # nanosecond precision: unique even for same-second uploads
        _append_upload_history({
            "upload_id": upload_id,
            "original_name": safe_name,
            "saved_name": save_path.name,
            "uploaded_at": datetime.now().isoformat(),
            "headers": headers,
            "row_count": len(rows) - 1,
        })
        excel_history_id_cache = upload_id
        excel_original_name_cache = safe_name

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
                "source_file": safe_name,
                "upload_id": upload_id,
            })

        excel_columns_cache = headers
        excel_rows_cache = rows[1:]
        excel_path_cache = str(save_path)
        return JSONResponse({
            "ok": True,
            "headers": headers,
            "email_column": email_col,
            "email_column_name": headers[email_col],
            "row_count": len(rows) - 1,
            "needs_column_select": False,
            "source_file": safe_name,
            "upload_id": upload_id,
        })
    finally:
        wb.close()


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
        with _open_smtp(config["host"], config["port"], timeout=25) as srv:
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
def health(request: Request):
    """Health check endpoint for Railway / monitoring."""
    config = load_smtp_config()
    store = _google_load_tokens()
    return JSONResponse({
        "status": "ok",
        "build_sha": os.environ.get("RAILWAY_GIT_COMMIT_SHA", "local")[:7],
        # boot_at / uptime_seconds prove whether a fresh container really started.
        # Without them, "google_client_configured=false" is ambiguous: it could mean
        # the variables were added but never redeployed, or that they landed in the
        # Preview environment / under a mistyped name and the new container cannot
        # see them at all.
        "boot_at": _BOOT_AT,
        "uptime_seconds": int(time.time() - _BOOT_TS),
        "smtp_configured": config is not None,
        # Gmail-draft mode readiness
        "google_client_configured": bool(GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET),
        "google_env_present": {
            "GOOGLE_CLIENT_ID": bool(os.environ.get("GOOGLE_CLIENT_ID")),
            "GOOGLE_CLIENT_SECRET": bool(os.environ.get("GOOGLE_CLIENT_SECRET")),
            "PUBLIC_BASE_URL": bool(os.environ.get("PUBLIC_BASE_URL")),
        },
        "google_authorized": bool(store and store.get("refresh_token")),
        "google_account": store.get("email", "") if store else "",
        "google_oauth_redirect_uri": _google_redirect_uri(request),
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
        "source_file": excel_original_name_cache or "unknown",
        "upload_id": excel_history_id_cache or "unknown",
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
    source_file = job_data.get("source_file", "unknown")
    upload_id = job_data.get("upload_id", "unknown")

    # Detect handle column (so log rows record it for traceability)
    handle_col = None
    for i, h in enumerate(headers):
        if h.lower() in ("handle", "tiktok_handle", "tiktok handle", "username", "user_name") or "handle" in h.lower():
            handle_col = i
            break

    def _make_log(row_no, email, ok, err, handle_val):
        """Build a log entry that includes the source Excel filename + handle."""
        return {
            "row": row_no,
            "email": str(email),
            "handle": str(handle_val) if handle_val else "",
            "ok": bool(ok),
            "error": str(err) if err else "",
            "source_file": source_file,
            "upload_id": upload_id,
        }
    
    total = len(rows)
    sent = 0
    failed = 0
    log = []

    # Reusable SMTP connection (avoid re-auth per email, which Gmail throttles on datacenter IPs)
    smtp_conn = None

    def _get_smtp():
        nonlocal smtp_conn
        if smtp_conn is None:
            smtp_conn = _open_smtp(config["host"], config["port"], timeout=120)
            smtp_conn.login(config["username"], config["password"])
        return smtp_conn

    def _close_smtp():
        nonlocal smtp_conn
        if smtp_conn is not None:
            try:
                smtp_conn.quit()
            except Exception:
                pass
            smtp_conn = None

    _save_state({
        "status": "running",
        "total": total,
        "sent": 0,
        "failed": 0,
        "current_row": 0,
        "log": [],
        "started_at": datetime.now().isoformat(),
    })
    
    try:
        for idx, row in enumerate(rows):
            if _stop_flag.is_set():
                break

            # Build variable map
            variables = {}
            for i, h in enumerate(headers):
                variables[h] = row[i] if i < len(row) else ""

            # Get recipient
            recipient = test_email if test_mode and test_email else (row[email_col] if email_col < len(row) else "")

            if not recipient or "@" not in str(recipient):
                failed += 1
                log.append(_make_log(idx + 1, recipient, False, "No valid email",
                                     row[handle_col] if handle_col is not None and handle_col < len(row) else ""))
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

            # Send email with connection reuse + retry (up to 3 attempts)
            ok = False
            last_err = ""
            for attempt in range(3):
                try:
                    srv = _get_smtp()
                    msg = EmailMessage()
                    msg["From"] = formataddr((from_name or "", config["username"]))
                    msg["To"] = recipient
                    msg["Subject"] = personalized_subject
                    msg.set_content("Please enable HTML to view this email.")
                    msg.add_alternative(personalized_html, subtype="html")
                    srv.send_message(msg)
                    ok = True
                    break
                except Exception as e:
                    last_err = str(e)
                    # Connection may be dead: close and reconnect on next attempt
                    _close_smtp()
                    time.sleep(2)

            if ok:
                sent += 1
                log.append(_make_log(idx + 1, recipient, True, "",
                                     row[handle_col] if handle_col is not None and handle_col < len(row) else ""))
            else:
                failed += 1
                log.append(_make_log(idx + 1, recipient, False, last_err,
                                     row[handle_col] if handle_col is not None and handle_col < len(row) else ""))

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
    finally:
        _close_smtp()

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


@app.get("/api/upload_history")
def upload_history():
    """Persistent history of all Excel files uploaded to this service (survives restarts)."""
    return JSONResponse({"ok": True, "history": _load_upload_history()})


@app.post("/api/search_excel")
async def search_excel(request: Request):
    """
    Search a string across ALL uploaded Excel files (current + historical).
    Returns: which file(s) contain that string, and the matched row data.
    Example payload: {"query": "caryysu", "fields": ["email", "handle"]}
    """
    body = await request.json()
    query = str(body.get("query", "")).strip().lower()
    if not query:
        return JSONResponse({"ok": False, "error": "Query empty"}, status_code=400)

    results = []
    # Search current in-memory cache first (most recent)
    for i, row in enumerate(excel_rows_cache):
        for j, cell in enumerate(row):
            cell_str = str(cell).lower() if cell is not None else ""
            if query in cell_str:
                col_name = excel_columns_cache[j] if j < len(excel_columns_cache) else f"col_{j+1}"
                results.append({
                    "file": excel_original_name_cache or "current.xlsx",
                    "upload_id": excel_history_id_cache or "",
                    "row_no": i + 2,  # +2: 1-based + header row
                    "matched_column": col_name,
                    "matched_value": str(cell),
                    "row_data": [str(c) if c is not None else "" for c in row],
                })
                break  # one match per row (per file)

    # Search historical uploads saved to disk
    history = _load_upload_history()
    uploaded_names = {excel_original_name_cache, "current.xlsx"}
    for record in history:
        fname = record.get("saved_name", "")
        # Avoid duplicating current cache (same file)
        if record.get("upload_id") == excel_history_id_cache:
            continue
        path = UPLOAD_DIR / fname
        if not path.exists() or fname in uploaded_names:
            continue
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb.active
            for i, row in enumerate(ws.iter_rows(values_only=True)):
                if i == 0:
                    continue  # skip header
                for j, cell in enumerate(row):
                    cell_str = str(cell).lower() if cell is not None else ""
                    if query in cell_str:
                        headers_list = record.get("headers", [])
                        col_name = headers_list[j] if j < len(headers_list) else f"col_{j+1}"
                        results.append({
                            "file": record.get("original_name", fname),
                            "upload_id": record.get("upload_id", ""),
                            "uploaded_at": record.get("uploaded_at", ""),
                            "row_no": i + 1,
                            "matched_column": col_name,
                            "matched_value": str(cell),
                            "row_data": [str(c) if c is not None else "" for c in row],
                        })
                        break
        finally:
            wb.close()

    return JSONResponse({"ok": True, "query": query, "matches": results})


@app.post("/api/smtp/diag")
async def smtp_diag():
    """
    Probe SMTP endpoints from THIS server's network (shows which ports are reachable
    from the deployment environment — useful when a provider blocks SMTP egress).
    Connectivity + banner only; does NOT send email.
    """
    results = []
    for host, port in _diagnose_smtp_targets():
        entry = {"host": host, "port": port, "connect": None, "banner": None, "error": None}
        try:
            if port == 465:
                srv = smtplib.SMTP_SSL(host, port, timeout=20)
            else:
                srv = smtplib.SMTP(host, port, timeout=20)
            code, msg = srv.ehlo()
            entry["connect"] = f"OK (EHLO {code})"
            entry["banner"] = (msg or b"").decode(errors="replace")[:120]
            entry["starttls_available"] = srv.has_extn("starttls")
            entry["auth_available"] = srv.has_extn("auth")
            try:
                srv.quit()
            except Exception:
                pass
        except Exception as e:
            entry["connect"] = "FAILED"
            entry["error"] = f"{type(e).__name__}: {str(e)[:100]}"
        results.append(entry)
    https_results = [_probe_https(u) for u in _diagnose_https_targets()]
    return JSONResponse({
        "ok": True,
        "server_time": datetime.now().isoformat(),
        "targets": results,
        "lark_api_reachable": https_results,
    })


@app.post("/api/stop")
def stop_send():
    """Request a graceful stop after the current email attempt finishes."""
    global _stop_flag
    _stop_flag.set()
    return JSONResponse({"ok": True, "message": "Stop requested"})


@app.post("/api/terminate")
def terminate_send():
    """Terminate the current task and invalidate the cached recipient list.

    This is intentionally stronger than /api/stop: after termination, a new
    Excel upload is required before /api/send can start again. The worker checks
    _stop_flag between emails and closes its SMTP connection in finally.
    """
    global _stop_flag, excel_columns_cache, excel_rows_cache, excel_path_cache
    global excel_history_id_cache, excel_original_name_cache
    _stop_flag.set()
    # Invalidate the in-memory list immediately so a stale list cannot be reused.
    excel_columns_cache = []
    excel_rows_cache = []
    excel_path_cache = None
    excel_history_id_cache = None
    excel_original_name_cache = None
    _save_state({
        "status": "terminated",
        "total": 0,
        "sent": 0,
        "failed": 0,
        "current_row": 0,
        "log": [],
        "finished_at": datetime.now().isoformat(),
    })
    return JSONResponse({
        "ok": True,
        "message": "Current task terminated; cached recipient list cleared; upload a new Excel before sending.",
    })


@app.get("/api/status")
def status():
    s = get_state()
    return JSONResponse(s)


# ── Gmail Drafts via Google API (compose-only: this app never calls .send) ──
#
# Strategy change (2026-10-19): instead of SMTP-blasting emails, build one Gmail
# DRAFT per Excel row and let a human hit "Send" in the Gmail UI. The OAuth scope
# requested is gmail.compose only.
GOOGLE_CLIENT_ID = os.environ.get("GOOGLE_CLIENT_ID", "")
GOOGLE_CLIENT_SECRET = os.environ.get("GOOGLE_CLIENT_SECRET", "")
# Local-dev fallback: the downloaded Cloud Console client secret JSON.
# This path is git-ignored (secrets/); on Railway the env vars above are used.
GOOGLE_CLIENT_SECRET_FILE = Path(os.environ.get(
    "GOOGLE_CLIENT_SECRET_FILE", str(BASE_DIR / "secrets" / "google_client_secret.json")))
if not GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET_FILE.exists():
    try:
        with open(GOOGLE_CLIENT_SECRET_FILE) as _f:
            _gconf = json.load(_f)
        _gweb = _gconf.get("web") or _gconf.get("installed") or {}
        GOOGLE_CLIENT_ID = _gweb.get("client_id", "")
        GOOGLE_CLIENT_SECRET = GOOGLE_CLIENT_SECRET or _gweb.get("client_secret", "")
    except (json.JSONDecodeError, IOError):
        pass
GOOGLE_SCOPES = os.environ.get(
    "GOOGLE_SCOPES", "https://www.googleapis.com/auth/gmail.compose")
GOOGLE_TOKEN_FILE = Path(os.environ.get(
    "GOOGLE_TOKEN_FILE", str(BASE_DIR / "data" / "google_tokens.json")))
GOOGLE_DRAFT_LEDGER = Path(os.environ.get(
    "GOOGLE_DRAFT_LEDGER", str(BASE_DIR / "data" / "draft_ledger.json")))
GOOGLE_AUTH_BASE = "https://accounts.google.com/o/oauth2/v2/auth"
GOOGLE_TOKEN_URI = "https://oauth2.googleapis.com/token"
GOOGLE_API_BASE = "https://gmail.googleapis.com/gmail/v1/users/me"
PUBLIC_BASE_URL = os.environ.get("PUBLIC_BASE_URL", "").rstrip("/")


def _google_redirect_uri(request: Request):
    """Callback URL sent to Google; must byte-match the Google Console setting.

    Railway terminates TLS at its edge and forwards http:// to the container, so
    request.base_url alone would produce an http:// redirect_uri that Google
    rejects. Honour the forwarded protocol/host, or an explicit override.
    """
    if PUBLIC_BASE_URL:
        return f"{PUBLIC_BASE_URL}/api/google/oauth/callback"
    base = str(request.base_url).rstrip("/")
    fwd_proto = request.headers.get("x-forwarded-proto", "").split(",")[0].strip()
    fwd_host = request.headers.get("x-forwarded-host", "").split(",")[0].strip()
    if fwd_proto in ("http", "https"):
        from urllib.parse import urlparse

        netloc = fwd_host or urlparse(base).netloc
        base = f"{fwd_proto}://{netloc}"
    return f"{base}/api/google/oauth/callback"


def _google_load_tokens():
    if not GOOGLE_TOKEN_FILE.exists():
        return None
    try:
        with open(GOOGLE_TOKEN_FILE) as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        return None


def _google_save_tokens(store):
    GOOGLE_TOKEN_FILE.parent.mkdir(parents=True, exist_ok=True)
    tmp = GOOGLE_TOKEN_FILE.with_suffix(".json.tmp")
    with open(tmp, "w") as f:
        json.dump(store, f, indent=2)
    os.chmod(tmp, 0o600)
    os.replace(tmp, GOOGLE_TOKEN_FILE)


def _google_http_json(url, method="GET", payload=None, token=None, form=None, timeout=40):
    """Minimal JSON/form HTTP client on stdlib urllib (no extra dependency)."""
    import urllib.request
    import urllib.error
    import urllib.parse

    headers = {"Content-Type": "application/json; charset=utf-8"}
    if token:
        headers["Authorization"] = f"Bearer {token}"
    if form is not None:
        data = urllib.parse.urlencode(form).encode()
        headers["Content-Type"] = "application/x-www-form-urlencoded"
    else:
        data = json.dumps(payload).encode() if payload is not None else None
    req = urllib.request.Request(url, data=data, method=method, headers=headers)
    try:
        with urllib.request.urlopen(req, timeout=timeout) as resp:
            return json.loads(resp.read().decode()), None
    except urllib.error.HTTPError as e:
        try:
            body = json.loads(e.read().decode())
        except Exception:
            body = {"error": {"message": str(e)}}
        return body, f"HTTP {e.code}: {body.get('error', {}).get('message', str(e))}"
    except Exception as e:
        return {}, f"{type(e).__name__}: {e}"


def _google_access_token():
    """Return (access_token, error). Refreshes the stored token when needed."""
    store = _google_load_tokens()
    if not store or not store.get("refresh_token"):
        return None, "尚未完成 Google 授权（请在页面上点击「连接 Gmail 授权」）"
    now = time.time()
    if store.get("access_token") and store.get("expires_at", 0) - 60 > now:
        return store["access_token"], None
    if not GOOGLE_CLIENT_ID or not GOOGLE_CLIENT_SECRET:
        return None, "GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET 未配置"
    resp, err = _google_http_json(
        GOOGLE_TOKEN_URI, method="POST",
        form={
            "client_id": GOOGLE_CLIENT_ID,
            "client_secret": GOOGLE_CLIENT_SECRET,
            "grant_type": "refresh_token",
            "refresh_token": store["refresh_token"],
        })
    if err:
        return None, f"token 刷新失败: {err}"
    store["access_token"] = resp.get("access_token")
    store["expires_at"] = now + int(resp.get("expires_in", 3600))
    _google_save_tokens(store)
    return store["access_token"], None


def _google_account_email():
    store = _google_load_tokens() or {}
    return store.get("email") or ""


def _load_draft_ledger():
    """upload_id -> set of row numbers that already have a draft, for dedupe."""
    if not GOOGLE_DRAFT_LEDGER.exists():
        return {}
    try:
        with open(GOOGLE_DRAFT_LEDGER) as f:
            data = json.load(f)
        return {k: set(v) for k, v in data.items()}
    except (json.JSONDecodeError, IOError):
        return {}


def _save_draft_ledger(ledger):
    GOOGLE_DRAFT_LEDGER.parent.mkdir(parents=True, exist_ok=True)
    try:
        tmp = GOOGLE_DRAFT_LEDGER.with_suffix(".json.tmp")
        with open(tmp, "w") as f:
            json.dump({k: sorted(v) for k, v in ledger.items()}, f, indent=2)
        os.replace(tmp, GOOGLE_DRAFT_LEDGER)
    except (IOError, OSError):
        pass


def _render_email_row(subject, html_body, row, headers, round_params, recipient):
    """Apply round-level fixed params, then per-row Excel variables."""
    import re

    out_subject = subject
    out_html = html_body
    for rp_name, rp_val in (round_params or {}).items():
        rp_val = "" if rp_val is None else str(rp_val)
        out_subject = out_subject.replace("{" + rp_name + "}", rp_val)
        out_html = out_html.replace("{" + rp_name + "}", rp_val)
    variables = {h: (row[i] if i < len(row) else "") for i, h in enumerate(headers)}
    for var_name, var_val in variables.items():
        out_subject = out_subject.replace("{" + var_name + "}", str(var_val))
        out_html = out_html.replace("{" + var_name + "}", str(var_val))
    if "{name}" in out_html and "name" not in headers:
        out_html = out_html.replace("{name}", str(recipient).split("@")[0])
    if "{name}" in out_subject and "name" not in headers:
        out_subject = out_subject.replace("{name}", str(recipient).split("@")[0])
    return out_subject, out_html


def _gmail_create_draft(token, to_addr, subject, html, from_name, cc_addr=None):
    """Create a Gmail draft (users.drafts.create). Returns (draft_id, error)."""
    import base64

    msg = EmailMessage()
    if from_name:
        msg["From"] = formataddr((from_name, _google_account_email() or "me"))
    else:
        msg["From"] = _google_account_email() or "me"
    msg["To"] = to_addr
    if cc_addr:
        msg["Cc"] = cc_addr
    msg["Subject"] = subject
    msg.set_content("Please enable HTML to view this email.")
    msg.add_alternative(html, subtype="html")
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    resp, err = _google_http_json(
        f"{GOOGLE_API_BASE}/drafts", method="POST", payload={"raw": raw}, token=token)
    if err:
        return None, err
    return resp.get("id"), None


_draft_thread = None
_draft_stop = threading.Event()
_draft_state = {"status": "idle", "total": 0, "created": 0, "skipped": 0,
                "failed": 0, "current_row": 0, "log": []}


def _draft_alive():
    return _draft_thread is not None and _draft_thread.is_alive()


def _save_draft_state(state):
    state["updated_at"] = datetime.now().isoformat()
    global _draft_state
    _draft_state = state


def _get_draft_state():
    if _draft_alive():
        return _draft_state
    if _draft_state.get("status") == "running":
        _draft_state["status"] = "interrupted"
    return _draft_state


@app.get("/api/google/status")
def google_status(request: Request):
    store = _google_load_tokens()
    return JSONResponse({
        "ok": True,
        "client_configured": bool(GOOGLE_CLIENT_ID and GOOGLE_CLIENT_SECRET),
        "authorized": bool(store and store.get("refresh_token")),
        "email": _google_account_email(),
        "scope": (store or {}).get("scope", ""),
        "redirect_uri": _google_redirect_uri(request),
        "token_file": str(GOOGLE_TOKEN_FILE),
        # Only this scope is requested; the app never calls drafts.send.
        "compose_only": "gmail.compose" in (store or {}).get("scope", ""),
    })


@app.get("/api/google/connect")
def google_connect(request: Request):
    """Start Google OAuth. Requires the sender account to consent once."""
    if not GOOGLE_CLIENT_ID or not GOOGLE_CLIENT_SECRET:
        return JSONResponse({"ok": False, "error": "GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET 未配置"}, status_code=400)
    from urllib.parse import urlencode

    params = {
        "client_id": GOOGLE_CLIENT_ID,
        "redirect_uri": _google_redirect_uri(request),
        "scope": GOOGLE_SCOPES,
        "response_type": "code",
        "access_type": "offline",       # needed for a refresh token
        "prompt": "consent",            # force refresh token even if previously granted
        "include_granted_scopes": "true",
    }
    return RedirectResponse(f"{GOOGLE_AUTH_BASE}?{urlencode(params)}")


@app.get("/api/google/oauth/callback")
async def google_oauth_callback(request: Request, code: str = "", state: str = "", error: str = ""):
    if error:
        return HTMLResponse(f"<h3>❌ Google 授权被拒绝</h3><p>{error}</p>")
    if not code:
        return HTMLResponse("<h3>❌ 缺少 code 参数</h3>", status_code=400)
    resp, err = _google_http_json(
        GOOGLE_TOKEN_URI, method="POST",
        form={
            "client_id": GOOGLE_CLIENT_ID,
            "client_secret": GOOGLE_CLIENT_SECRET,
            "code": code,
            "grant_type": "authorization_code",
            "redirect_uri": _google_redirect_uri(request),
        })
    if err:
        return HTMLResponse(
            f"<h3>❌ 授权失败</h3><p>{err}</p>"
            f"<p style='color:#888'>常见原因：redirect_uri 与 Google Console 登记值不一致"
            f"（当前使用 {_google_redirect_uri(request)}）</p>", status_code=400)
    # Decode the email/scope from the id_token without signature verification
    # (display-only; the API calls themselves are authenticated by the token).
    email, scope = "", resp.get("scope", "")
    try:
        import base64 as _b64

        idt = resp.get("id_token", "")
        if idt:
            payload = idt.split(".")[1]
            payload += "=" * (-len(payload) % 4)
            email = json.loads(_b64.urlsafe_b64decode(payload)).get("email", "")
    except Exception:
        pass
    _google_save_tokens({
        "access_token": resp.get("access_token"),
        "refresh_token": resp.get("refresh_token"),
        "expires_at": time.time() + int(resp.get("expires_in", 3600)),
        "scope": scope,
        "email": email,
        "authorized_at": datetime.now().isoformat(),
    })
    return HTMLResponse(
        "<h3>✅ Google 授权成功，已保存 token</h3>"
        f"<p>账号：{email or '（未返回邮箱，不影响使用）'}</p>"
        f"<p>权限：{scope}</p>"
        "<p><a href='/'>← 返回 Mail Blaster</a></p>")


@app.post("/api/create_drafts")
async def start_create_drafts(request: Request):
    """Queue a draft-creation job. Creates Gmail drafts only -- sends nothing."""
    global _draft_thread, _draft_stop

    body = await request.json()
    subject = body.get("subject", "")
    html_body = body.get("html", "")
    email_col = int(body.get("email_column", 0))
    from_name = body.get("from_name", "")
    round_params = body.get("round_params", {}) or {}
    limit = int(body.get("limit", 0) or 0)          # 0 = all remaining
    delay_sec = float(body.get("delay_sec", 1.0))   # polite pacing for the API
    skip_sent = bool(body.get("skip_existing", True))

    if not excel_rows_cache:
        return JSONResponse({"ok": False, "error": "未上传 Excel 名单"}, status_code=400)
    if _draft_alive():
        return JSONResponse({"ok": False, "error": "已有草稿任务在运行"}, status_code=409)
    token, terr = _google_access_token()
    if not token:
        return JSONResponse({"ok": False, "error": terr}, status_code=400)
    if not subject.strip():
        return JSONResponse({"ok": False, "error": "邮件主题为空"}, status_code=400)

    upload_id = excel_history_id_cache or "unknown"
    ledger = _load_draft_ledger()
    done_rows = ledger.get(upload_id, set()) if skip_existing else set()

    job_data = {
        "subject": subject,
        "html": html_body,
        "email_col": email_col,
        "from_name": from_name,
        "round_params": round_params,
        "rows": [[str(c) if c is not None else "" for c in row] for row in excel_rows_cache],
        "headers": excel_columns_cache,
        "source_file": excel_original_name_cache or "unknown",
        "upload_id": upload_id,
        "limit": limit,
        "delay_sec": delay_sec,
        "done_rows": done_rows,
    }

    _draft_stop.clear()
    _draft_thread = threading.Thread(target=_run_draft_job, args=(job_data,), daemon=True)
    _draft_thread.start()
    return JSONResponse({"ok": True, "total": len(job_data["rows"]),
                         "skipped_existing": len(done_rows),
                         "source_file": job_data["source_file"]})


def _run_draft_job(job_data):
    subject = job_data["subject"]
    html_body = job_data["html"]
    email_col = job_data["email_col"]
    from_name = job_data["from_name"]
    round_params = job_data["round_params"]
    rows = job_data["rows"]
    headers = job_data["headers"]
    source_file = job_data["source_file"]
    upload_id = job_data["upload_id"]
    limit = job_data["limit"]
    delay_sec = job_data["delay_sec"]

    ledger = _load_draft_ledger()
    created_ids = ledger.setdefault(f"{upload_id}:ids", [])
    done = ledger.get(upload_id, set())

    handle_col = None
    for i, h in enumerate(headers):
        hl = h.lower()
        if "handle" in hl or hl in ("username", "user_name"):
            handle_col = i
            break

    total = len(rows)
    created = skipped = failed = 0
    log = []
    _save_draft_state({"status": "running", "total": total, "created": 0,
                       "skipped": 0, "failed": 0, "current_row": 0, "log": [],
                       "source_file": source_file,
                       "started_at": datetime.now().isoformat()})
    try:
        for idx, row in enumerate(rows):
            if _draft_stop.is_set():
                break
            if limit and created >= limit:
                break
            recipient = row[email_col] if email_col < len(row) else ""
            row_no = idx + 1
            handle_val = row[handle_col] if handle_col is not None and handle_col < len(row) else ""

            entry = {"row": row_no, "email": str(recipient),
                     "handle": str(handle_val or ""),
                     "source_file": source_file, "upload_id": upload_id}
            if not recipient or "@" not in str(recipient):
                failed += 1
                log.append({**entry, "ok": False, "error": "No valid email"})
            elif row_no in done:
                skipped += 1
                log.append({**entry, "ok": True, "error": "", "note": "已存在草稿，跳过"})
            else:
                subj, body_html = _render_email_row(
                    subject, html_body, row, headers, round_params, recipient)
                token, terr = _google_access_token()
                if not token:
                    failed += 1
                    log.append({**entry, "ok": False, "error": f"token: {terr}"})
                else:
                    draft_id, derr = _gmail_create_draft(
                        token, str(recipient), subj, body_html, from_name)
                    if derr:
                        failed += 1
                        log.append({**entry, "ok": False, "error": derr})
                    else:
                        created += 1
                        done.add(row_no)
                        ledger[upload_id] = done
                        if draft_id:
                            created_ids.append(draft_id)
                        _save_draft_ledger(ledger)
                        log.append({**entry, "ok": True, "error": "", "draft_id": draft_id})
            _save_draft_state({"status": "running", "total": total, "created": created,
                               "skipped": skipped, "failed": failed,
                               "current_row": row_no, "log": log[-100:],
                               "source_file": source_file})
            if delay_sec > 0 and not _draft_stop.is_set():
                time.sleep(delay_sec)
    finally:
        _save_draft_ledger(ledger)

    final = "completed" if not _draft_stop.is_set() else "stopped"
    _save_draft_state({"status": final, "total": total, "created": created,
                       "skipped": skipped, "failed": failed, "current_row": total,
                       "log": log[-100:], "source_file": source_file,
                       "finished_at": datetime.now().isoformat()})


@app.get("/api/draft_status")
def draft_status():
    return JSONResponse(_get_draft_state())


@app.post("/api/draft_stop")
def draft_stop():
    _draft_stop.set()
    return JSONResponse({"ok": True, "message": "Stop requested"})


@app.get("/api/draft_ledger")
def draft_ledger():
    ledger = {k: (sorted(v) if isinstance(v, set) else v) for k, v in _load_draft_ledger().items()}
    return JSONResponse({"ok": True, "ledger": ledger})


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

/* Image delete button (top-right corner of thumbnail) */
.img-del-btn{position:absolute;top:2px;right:2px;width:22px;height:22px;border:none;border-radius:50%;background:rgba(0,0,0,0.55);color:#fff;font-size:12px;line-height:22px;text-align:center;cursor:pointer;padding:0;opacity:0;transition:opacity 0.15s}
.image-thumb:hover .img-del-btn{opacity:1}
.img-del-btn:hover{background:#dc3545}

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

<!-- Step 3: Round Params (fixed for this batch) -->
<div class="card" id="round-params-card">
  <h2>3️⃣ 本轮统一参数 <span style="font-weight:normal;font-size:12px;color:#888;">(本批所有邮件一致，按需填写)</span></h2>
  <p class="hint">💡 有需要就填，不需要的留空即可（留空的变量会被替换成空字符串，不影响其他内容）。</p>
  <div class="grid2">
    <div>
      <label>产品分类 <span class="snippet">{category_type}</span></label>
      <input type="text" id="rp-category_type" placeholder="例如: Christmas Light Strip">
    </div>
    <div>
      <label>品牌名称 <span class="snippet">{brand_name}</span></label>
      <input type="text" id="rp-brand_name" placeholder="例如: Eufy">
    </div>
  </div>
  <div class="grid2">
    <div>
      <label>产品链接 <span class="snippet">{product_url}</span></label>
      <input type="text" id="rp-product_url" placeholder="Taplink / TikTok Shop 产品页 / 官网链接">
    </div>
    <div>
      <label>产品价格 <span class="snippet">{product_price}</span></label>
      <input type="text" id="rp-product_price" placeholder="例如: $199">
    </div>
  </div>
  <div>
    <label>产品图片 URL <span class="snippet">{picture_url}</span></label>
    <input type="text" id="rp-picture_url" placeholder="https://你的CDN图片地址...">
  </div>
</div>

<!-- Step 4: Email Content -->
<div class="card">
  <h2>4️⃣ 邮件内容 (HTML)</h2>
  <button class="btn btn-secondary" style="float:right;padding:4px 10px;font-size:12px;" onclick="clearDraft()">清空本地草稿</button>
  <label>邮件主题</label>
  <input type="text" id="subject" placeholder="Hi {handle}, TikTok Shop product opportunities from Liveology US" value="Hi {handle}, TikTok Shop product opportunities from Liveology US">
  <p class="hint">
    💡 变量用 <span class="snippet">{"{"}列名{"}"}</span> 替换（自动识别 Excel 列名）：
    <span class="snippet">{"{"}handle{"}"}</span> / <span class="snippet">{"{"}name{"}"}</span> / <span class="snippet">{"{"}email{"}"}</span> 等
  </p>

  <!-- Image gallery (add images by URL, click-to-insert) -->
  <div style="margin-bottom:8px;">
    <label style="display:inline;">图片库 (点缩略图插入)</label>
    <br>
    <input type="text" id="image-url-input" placeholder="粘贴图片 CDN URL —— https://..." style="display:inline-block;margin-top:6px;width:380px;margin-bottom:6px;">
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

<!-- Step 5: Send Settings -->
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

<!-- Step 5.5: Gmail Draft mode (human clicks Send in Gmail) -->
<div class="card" id="draft-card" style="border-left:4px solid #059669;">
  <h2>📤 Gmail 草稿模式（推荐：程序只做草稿，你人工点发送）</h2>
  <div id="google-status" style="margin-bottom:10px;font-size:13px;"></div>
  <div style="margin-bottom:12px;">
    <button class="btn btn-secondary" onclick="window.open('/api/google/connect','_blank')">连接 Gmail 授权（一次性）</button>
    <a class="btn btn-secondary" href="https://mail.google.com/mail/u/0/#drafts" target="_blank">打开 Gmail 草稿箱</a>
  </div>
  <div class="grid2">
    <div><label>本次创建数量（0 = 全部剩余）</label><input type="number" id="draft-limit" value="20"></div>
    <div><label>每封间隔 (秒)</label><input type="number" id="draft-delay" value="1" step="0.1"></div>
  </div>
  <label><input type="checkbox" id="draft-skip-existing" checked style="width:auto;margin-right:6px;">跳过已创建过草稿的行（防止重复草稿）</label>
  <div style="margin-top:12px;">
    <button class="btn btn-primary" id="btn-drafts" onclick="startDrafts()">📝 创建 Gmail 草稿</button>
    <button class="btn btn-danger" id="btn-draft-stop" onclick="stopDrafts()" disabled>停止</button>
    <span id="draft-note" style="font-size:12px;color:#888;margin-left:8px;">此操作只是把邮件写入草稿箱，不会发出任何邮件</span>
  </div>
  <div class="stats" style="margin-top:12px;">
    <div class="stat"><div class="num" id="dstat-created" style="color:#059669">0</div><div class="label">已建草稿</div></div>
    <div class="stat"><div class="num" id="dstat-skipped">0</div><div class="label">跳过</div></div>
    <div class="stat"><div class="num" id="dstat-failed" style="color:#dc2626">0</div><div class="label">失败</div></div>
    <div class="stat"><div class="num" id="dstat-row">0</div><div class="label">当前行</div></div>
  </div>
  <div class="progress-bar"><div class="progress-fill" id="draft-progress" style="width:0%"></div></div>
  <div class="log-box" id="draft-log"><div class="log-line" style="color:#999;">等待创建草稿...</div></div>
</div>

<!-- Step 6: Source tracing (who came from which file) -->
<div class="card">
  <h2>6️⃣ 名单溯源 <span style="font-weight:normal;font-size:12px;color:#888;">(查询某个邮箱/handle 在哪个 Excel 文件里出现过)</span></h2>
  <div style="display:flex;gap:8px;align-items:flex-end;flex-wrap:wrap;">
    <div style="flex:1;min-width:200px;">
      <label>查询关键词（邮箱或 handle 片段）</label>
      <input type="text" id="trace-query" placeholder="例如: caryysu / someone@gmail.com">
    </div>
    <button class="btn btn-secondary" onclick="searchTrace()">🔍 查询所有上传记录</button>
  </div>
  <div id="trace-results" style="margin-top:8px;"></div>
</div>

<!-- Step 7: Send -->
<div class="card">
  <h2>7️⃣ 发送</h2>
  <div style="display:flex;gap:10px;margin-bottom:12px;">
    <button class="btn btn-success" id="btn-send" onclick="startSend()">开始发送</button>
    <button class="btn btn-danger" id="btn-stop" onclick="stopSend()">🛑 停止</button>
    <button class="btn btn-danger" id="btn-terminate" onclick="terminateSend()">⛔ 终止并清空名单</button>
  </div>
  
  <div id="job-context" class="hint" style="margin-bottom:6px;">当前名单：未上传　|　上次任务：无</div>

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
// Persist editable campaign fields locally so a page refresh does not erase work.
// Excel files are intentionally NOT persisted in localStorage; re-upload the list after refresh.
const DRAFT_STORAGE_KEY = 'mail_blaster_draft_v1';
const DRAFT_FIELD_IDS = [
  'subject', 'rp-category_type', 'rp-brand_name', 'rp-product_url',
  'rp-picture_url', 'rp-product_price', 'test-email', 'min-delay',
  'max-delay', 'smtp-name', 'draft-limit', 'draft-delay'
];
function saveDraft() {
  try {
    const draft = { html: editor.getValue(), saved_at: new Date().toISOString() };
    DRAFT_FIELD_IDS.forEach(id => {
      const el = document.getElementById(id);
      if (el) draft[id] = el.type === 'checkbox' ? el.checked : el.value;
    });
    const testMode = document.getElementById('test-mode');
    if (testMode) draft['test-mode'] = testMode.checked;
    localStorage.setItem(DRAFT_STORAGE_KEY, JSON.stringify(draft));
  } catch (e) { console.warn('Draft save failed:', e); }
}
function restoreDraft() {
  try {
    const raw = localStorage.getItem(DRAFT_STORAGE_KEY);
    if (!raw) { editor.setValue(DEFAULT_HTML); return; }
    const draft = JSON.parse(raw);
    editor.setValue(draft.html || DEFAULT_HTML);
    DRAFT_FIELD_IDS.forEach(id => {
      const el = document.getElementById(id);
      if (!el || draft[id] === undefined) return;
      if (el.type === 'checkbox') el.checked = !!draft[id];
      else el.value = draft[id];
    });
    const testMode = document.getElementById('test-mode');
    if (testMode && draft['test-mode'] !== undefined) testMode.checked = !!draft['test-mode'];
    if (testMode) document.getElementById('test-email-box').style.display = testMode.checked ? 'block' : 'none';
  } catch (e) {
    console.warn('Draft restore failed:', e);
    editor.setValue(DEFAULT_HTML);
  }
}
function clearDraft() {
  localStorage.removeItem(DRAFT_STORAGE_KEY);
  editor.setValue(DEFAULT_HTML);
  DRAFT_FIELD_IDS.forEach(id => { const el = document.getElementById(id); if (el) el.value = ''; });
  document.getElementById('subject').value = 'Hi {handle}, TikTok Shop product opportunities from Liveology US';
  document.getElementById('min-delay').value = 2;
  document.getElementById('max-delay').value = 10;
  document.getElementById('test-mode').checked = true;
  document.getElementById('test-email-box').style.display = 'block';
}
restoreDraft();
editor.on('change', saveDraft);
DRAFT_FIELD_IDS.forEach(id => {
  const el = document.getElementById(id);
  if (el) el.addEventListener('input', saveDraft);
});
const testModeDraft = document.getElementById('test-mode');
if (testModeDraft) testModeDraft.addEventListener('change', saveDraft);

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

// ── Image library (URL only; base64 removed because Gmail does not render data-URI images) ──
function addImageToGallery(name, url) {
  const gallery = document.getElementById('image-gallery');
  const div = document.createElement('div');
  div.className = 'image-thumb';
  div.style.position = 'relative';

  // Build DOM safely (no innerHTML to avoid XSS from name/url)
  const img = document.createElement('img');
  img.src = url;
  img.alt = name;
  const nameDiv = document.createElement('div');
  nameDiv.className = 'name';
  nameDiv.textContent = name;
  const delBtn = document.createElement('button');
  delBtn.className = 'img-del-btn';
  delBtn.textContent = '✕';
  delBtn.title = '从图片库删除';
  delBtn.onclick = (ev) => {
    ev.stopPropagation();  // don't trigger insert
    if (confirm(`确认从图片库删除「${name}」吗？\n（已插入编辑器的内容不受影响）`)) {
      div.remove();
    }
  };
  div.appendChild(img);
  div.appendChild(nameDiv);
  div.appendChild(delBtn);

  // Click thumbnail → insert into editor
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
    currentSourceFile = d.source_file || f.name;
    info.innerHTML = '✅ 文件 <strong>' + currentSourceFile + '</strong> · 共 ' + d.row_count + ' 行 · 列: ' + d.headers.join(', ');
    // A fresh list is now authoritative: re-enable send and refresh the context line
    if (pollTimer) { clearInterval(pollTimer); pollTimer = null; }
    document.getElementById('btn-send').disabled = false;
    document.getElementById('btn-terminate').disabled = false;
    updateJobContext({status: 'idle'});
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
let currentSourceFile = null;  // populated on upload

async function searchTrace() {
  const q = document.getElementById('trace-query').value.trim().toLowerCase();
  const box = document.getElementById('trace-results');
  if (!q) { alert('请输入查询关键词'); return; }
  box.textContent = '查询中...';
  const d = await api('/api/search_excel', {
    method: 'POST', headers: {'Content-Type': 'application/json'},
    body: JSON.stringify({query: q})
  });
  if (!d.ok) { box.textContent = '❌ ' + (d.error || '查询失败'); return; }
  const m = d.matches || [];
  if (m.length === 0) {
    box.textContent = `⚠️ 在所有已上传的 Excel 中没有找到 "${q}"`;
    return;
  }
  box.textContent = '';
  const header = document.createElement('div');
  header.style.cssText = 'padding:6px 0;color:#6b7280;font-size:12px;border-bottom:1px solid #eee;';
  header.textContent = `✅ 在 ${m.length} 行找到 "${q}"，来自 ${new Set(m.map(x=>x.file)).size} 个文件：`;
  box.appendChild(header);
  m.forEach((x) => {
    const row = document.createElement('div');
    row.style.cssText = 'padding:8px 0;border-bottom:1px solid #f0f0f0;font-size:12px;';
    const parts = [];
    parts.push('📄 ' + x.file);
    parts.push('第 ' + x.row_no + ' 行');
    parts.push('列「' + x.matched_column + '」= ' + x.matched_value);
    if (x.uploaded_at) parts.push('上传于 ' + x.uploaded_at.slice(0, 16).replace('T',' '));
    parts.forEach((p, i) => {
      if (i > 0) row.appendChild(document.createTextNode(' · '));
      row.appendChild(document.createTextNode(p));
    });
    box.appendChild(row);
  });
}

function collectRoundParams() {
  const params = {};
  ['category_type', 'brand_name', 'product_url', 'picture_url', 'product_price'].forEach(key => {
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

  // Confirm source file + row count before sending (prevents "wrong list" accidents)
  const fileName = currentSourceFile || '当前内存中的名单';
  const modeLabel = testMode ? `🧪 TEST MODE（全部发往 ${testEmail}）` : '🚨 正式发送（真实收件人会收到邮件！）';
  const ok = confirm(
    `请仔细核对发送信息：\n\n` +
    `📄 名单文件: ${fileName}\n` +
    `👥 名单行数: ${rowCount} 封\n` +
    `📮 模式: ${modeLabel}\n\n` +
    `确认无误后点击「确定」开始发送。`
  );
  if (!ok) return;
  
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
    document.getElementById('btn-terminate').disabled = false;
    // btn-stop stays enabled always (user can stop anytime)
    startPolling();
  } else {
    alert('❌ ' + (d.error||'发送失败'));
  }
}

async function terminateSend() {
  const ok = confirm(
    '确认终止当前任务并清空服务器中的名单吗？\n\n' +
    '当前发送会停止，旧名单会立即失效。\n' +
    '之后必须重新上传新 Excel，才能再次发送。'
  );
  if (!ok) return;
  const btn = document.getElementById('btn-terminate');
  btn.disabled = true;
  btn.textContent = '终止中...';
  try {
    const d = await api('/api/terminate', {method:'POST'});
    if (!d.ok) throw new Error(d.error || 'terminate failed');
    document.getElementById('btn-send').disabled = true;
    currentSourceFile = null;
    rowCount = 0;
    updateJobContext({status: 'terminated'});
    document.getElementById('excel-info').textContent = '⛔ 当前任务已终止，旧名单已清空，请重新上传 Excel';
    btn.textContent = '已终止，请重新上传名单';
    if (pollTimer) clearInterval(pollTimer);
    updateProgress();
  } catch (e) {
    btn.disabled = false;
    btn.textContent = '⛔ 终止并清空名单';
    alert('❌ 终止失败，请重试');
  }
}

async function stopSend() {
  // Always clickable: even if frontend thinks job is idle, backend may still be running
  const btn = document.getElementById('btn-stop');
  btn.disabled = true;
  btn.textContent = '停止中...';
  try {
    await api('/api/stop', {method:'POST'});
    // Immediate feedback, no blocking alert (so UI can keep polling)
    btn.textContent = '停止请求已发出';
    setTimeout(() => { btn.textContent = '停止'; btn.disabled = false; }, 2000);
  } catch (e) {
    btn.textContent = '停止';
    btn.disabled = false;
    alert('❌ 停止请求失败，请重试');
  }
}

function startPolling() {
  if (pollTimer) clearInterval(pollTimer);
  pollTimer = setInterval(updateProgress, 2000);
  updateProgress();
}

// Distinguish "list currently loaded" from "counters of the last job snapshot"
function updateJobContext(d) {
  const box = document.getElementById('job-context');
  if (!box) return;
  const listPart = currentSourceFile
    ? ('当前名单：' + currentSourceFile + ' · ' + (rowCount || 0) + ' 行')
    : '当前名单：未上传';
  let jobPart = '上次任务：无';
  if (d && d.status && d.status !== 'idle') {
    const src = (d.log && d.log[0] && d.log[0].source_file) ? d.log[0].source_file : '';
    const when = d.updated_at ? d.updated_at.replace('T', ' ').slice(5, 16) : '';
    jobPart = '上次任务：' + d.status + ' · 总数 ' + (d.total || 0) +
              (src ? (' · ' + src) : '') + (when ? (' · ' + when) : '');
  }
  box.textContent = listPart + '　|　' + jobPart;
}

async function updateProgress() {
  const d = await api('/api/status');
  // Counters below belong to the LAST job; show zeros when no job has run here
  const stale = !d.status || d.status === 'idle' || d.status === 'terminated';
  document.getElementById('stat-total').textContent = stale ? 0 : (d.total || 0);
  document.getElementById('stat-sent').textContent = stale ? 0 : (d.sent || 0);
  document.getElementById('stat-failed').textContent = stale ? 0 : (d.failed || 0);
  document.getElementById('stat-remain').textContent = stale ? 0 : ((d.total||0) - (d.sent||0) - (d.failed||0));
  updateJobContext(d);

  const pct = (!stale && d.total > 0) ? ((d.sent + d.failed) / d.total * 100) : 0;
  document.getElementById('progress-fill').style.width = pct + '%';
  
  // Log (use textContent via DOM construction to avoid XSS from email/error fields)
  const lb = document.getElementById('log-box');
  lb.textContent = '';
  const log = d.log || [];
  log.slice(-20).forEach(e => {
    const div = document.createElement('div');
    div.className = 'log-line ' + (e.ok ? 'log-ok' : 'log-fail');
    const parts = ['Row ' + e.row];
    if (e.source_file) parts.push('📄 ' + e.source_file);
    if (e.handle) parts.push('@' + e.handle);
    parts.push(e.email);
    parts.push(e.ok ? '✅' : '❌ ' + (e.error || ''));
    parts.forEach((p, i) => {
      if (i > 0) div.appendChild(document.createTextNode(' · '));
      div.appendChild(document.createTextNode(String(p)));
    });
    lb.appendChild(div);
  });
  lb.scrollTop = lb.scrollHeight;
  
  // Status check
  if (d.status && d.status !== 'running' && d.status !== 'idle') {
    document.getElementById('btn-send').disabled = false;
    if (d.status === 'terminated') {
      document.getElementById('btn-send').disabled = true;
      document.getElementById('btn-terminate').disabled = true;
    }
    if (d.status === 'completed') {
      if (pollTimer) clearInterval(pollTimer);
    }
  }
}

// ── Test mode toggle ──
document.getElementById('test-mode').addEventListener('change', function() {
  document.getElementById('test-email-box').style.display = this.checked ? 'block' : 'none';
});

// ── Warn before closing page while sending ──
window.addEventListener('beforeunload', function(e) {
  // Check current UI state: send button disabled = job in progress
  if (document.getElementById('btn-send').disabled) {
    e.preventDefault();
    e.returnValue = '⚠️ 邮件正在发送中，关闭页面将中断任务。确定要离开吗？';
    return e.returnValue;
  }
});

// ── Gmail Draft mode ──
let draftPollTimer = null;

async function checkGoogle() {
  const el = document.getElementById('google-status');
  try {
    const d = await api('/api/google/status');
    if (!d.client_configured) {
      el.innerHTML = '<span class="badge badge-no">❌ Google 凭据未配置</span> 请在 Railway 设置 GOOGLE_CLIENT_ID / GOOGLE_CLIENT_SECRET';
      return;
    }
    if (d.authorized) {
      el.innerHTML = '<span class="badge badge-ok">✅ 已授权 Gmail（只读草稿权限）</span> 账号: ' +
        (d.email || '(未返回)') + (d.compose_only ? ' · 权限: gmail.compose（程序无发信权限调用）' : '');
    } else {
      el.innerHTML = '<span class="badge badge-no">⚠️ 尚未授权</span> 点上方「连接 Gmail 授权」，用发件账号（liveology.ny@gmail.com）登录并同意，一次即可';
    }
  } catch (e) {
    el.innerHTML = '<span class="badge badge-no">❌ 无法获取 Google 授权状态</span> ' + e.message;
  }
}

async function startDrafts() {
  const subject = document.getElementById('subject').value;
  const html = getEditorHtml();
  if (!subject) { alert('请填写邮件主题'); return; }
  if (!html || !html.trim()) { alert('请填写邮件正文'); return; }
  if (emailColIdx === null) { alert('请先上传 Excel 并选择邮箱列'); return; }
  const limit = parseInt(document.getElementById('draft-limit').value || '0', 10);
  const ok = confirm(
    `将创建 Gmail 草稿（不会发送任何邮件）：\n\n` +
    `文件：${currentSourceFile || '未知'}\n` +
    `名单：${rowCount} 行\n` +
    `本次数量：${limit > 0 ? limit : '全部剩余'}\n` +
    `跳过已创建：${document.getElementById('draft-skip-existing').checked ? '是' : '否'}\n\n` +
    `确认继续？`);
  if (!ok) return;
  const d = await api('/api/create_drafts', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({
      subject: subject,
      html: html,
      email_column: emailColIdx,
      from_name: document.getElementById('smtp-name').value,
      round_params: collectRoundParams(),
      limit: limit,
      delay_sec: parseFloat(document.getElementById('draft-delay').value || '1'),
      skip_existing: document.getElementById('draft-skip-existing').checked,
    }),
  });
  if (!d.ok) { alert('❌ ' + (d.error || '启动失败')); return; }
  document.getElementById('btn-drafts').disabled = true;
  document.getElementById('btn-draft-stop').disabled = false;
  pollDraftStatus(true);
}

async function stopDrafts() {
  await api('/api/draft_stop', { method: 'POST' });
  document.getElementById('draft-log').innerHTML +=
    '<div class="log-line" style="color:#f59e0b">⏹ 已请求停止...</div>';
}

function renderDraftStatus(d) {
  document.getElementById('dstat-created').textContent = d.created || 0;
  document.getElementById('dstat-skipped').textContent = d.skipped || 0;
  document.getElementById('dstat-failed').textContent = d.failed || 0;
  document.getElementById('dstat-row').textContent = d.current_row || 0;
  const total = d.total || 0;
  const pct = total ? Math.min(100, Math.round(((d.current_row || 0) / total) * 100)) : 0;
  document.getElementById('draft-progress').style.width = pct + '%';
  const box = document.getElementById('draft-log');
  box.innerHTML = (d.log || []).slice(-40).map(l => {
    if (l.ok) {
      return `<div class="log-line" style="color:#059669">✅ 第${l.row}行 ${l.email} ${l.note || ''} 草稿ID:${l.draft_id || '-'}</div>`;
    }
    return `<div class="log-line" style="color:#dc2626">❌ 第${l.row}行 ${l.email || '?'} · ${l.error || ''}</div>`;
  }).join('') || '<div class="log-line" style="color:#999">等待创建草稿...</div>';
  box.scrollTop = box.scrollHeight;
}

function pollDraftStatus(start) {
  if (draftPollTimer) clearInterval(draftPollTimer);
  const tick = async () => {
    try {
      const d = await api('/api/draft_status');
      renderDraftStatus(d);
      if (d.status === 'running') {
        document.getElementById('btn-drafts').disabled = true;
        document.getElementById('btn-draft-stop').disabled = false;
        // Survive a page refresh mid-job: re-arm polling when a job is live.
        if (!draftPollTimer) draftPollTimer = setInterval(tick, 1500);
      } else {
        if (draftPollTimer) { clearInterval(draftPollTimer); draftPollTimer = null; }
        document.getElementById('btn-drafts').disabled = false;
        document.getElementById('btn-draft-stop').disabled = true;
        if (d.created || d.failed) {
          document.getElementById('draft-log').innerHTML +=
            `<div class="log-line" style="color:#059669;font-weight:600">■ 任务${d.status === 'stopped' ? '已停止' : '完成'}：成功 ${d.created || 0} 封草稿，跳过 ${d.skipped || 0}，失败 ${d.failed || 0}。去 Gmail 草稿箱人工核对后点击发送。</div>`;
        }
      }
    } catch (e) { /* transient */ }
  };
  tick();
  if (start) draftPollTimer = setInterval(tick, 1500);
}

// ── Init ──
checkSMTP();
checkGoogle();
updateProgress();
pollDraftStatus(false);
</script>
</body>
</html>
"""


# ── Start server ──
if __name__ == "__main__":
    import uvicorn
    port = int(os.environ.get("PORT", 8003))
    uvicorn.run(app, host="0.0.0.0", port=port)

