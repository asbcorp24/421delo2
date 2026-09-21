# -*- coding: utf-8 -*-
import json
import os
import io
import csv
import hmac
import secrets
import sqlite3
import datetime as dt
from pathlib import Path
from typing import Optional
from functools import wraps
from models import *
from models import ActivityTemplate, MemoTemplate
from flask import (
    Flask, request, jsonify, render_template, redirect,
    url_for, flash, send_from_directory, send_file, abort
)
from sqlalchemy.orm import joinedload
from flask_sqlalchemy import SQLAlchemy
from flask_login import (
    LoginManager, UserMixin, login_user, logout_user,
    current_user, login_required
)
from sqlalchemy import text
from werkzeug.utils import secure_filename
from datetime import datetime
from sqlalchemy.orm import joinedload
from sqlalchemy import and_, or_, func
import uuid
from werkzeug.utils import secure_filename
import pytesseract
from PIL import Image
import pdfplumber
import json
from docx import Document
import docx
import re
from flask import current_app
import time
import glob
from config import SUPERADMIN_LOGIN, SUPERADMIN_PASSWORD,SECRET_PASSPHRASE
from dotenv import load_dotenv
import os
from urllib.parse import quote
from backup_database import BACKUP_DIR, create_backup, list_backups
load_dotenv()
from sqlalchemy import text as sql_text
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
# =========================
# ????????????
# =========================
import sys
from pathlib import Path
if getattr(sys, 'frozen', False):
    BASE_DIR = Path(sys.executable).resolve().parent
else:
    BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)
# === Crypto Utils ===
import base64
import os
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM


def get_department_prefix(department_id, field_name: str, fallback: str) -> str:
    if department_id:
        department = db.session.get(Department, department_id)
        if department:
            department_prefix = (getattr(department, field_name, None) or "").strip()
            if department_prefix:
                return department_prefix
    return fallback.strip()


def get_department_start_number(department_id, field_name: str, fallback: int) -> int:
    if department_id:
        department = db.session.get(Department, department_id)
        if department:
            value = getattr(department, field_name, None)
            if isinstance(value, int) and value > 0:
                return value
    return fallback


def generate_next_letter_reg_number(department_id=None) -> str:
    prefix = get_department_prefix(
        department_id,
        "letter_prefix",
        get_app_setting("letter_prefix", ""),
    )
    start_number = get_department_start_number(
        department_id,
        "letter_start_number",
        get_app_setting_int("letter_start_number", 1),
    )
    memo_source_id = ensure_memo_source()
    max_number = max(start_number - 1, 0)
    department = db.session.get(Department, department_id) if department_id else None
    series = (department.letter_number_series or "").strip() if department else ""
    letter_query = db.session.query(Letter.reg_number).filter(or_(Letter.source_id.is_(None), Letter.source_id != memo_source_id))
    if series:
        department_ids = [d.id for d in Department.query.filter_by(letter_number_series=series).all()]
        letter_query = letter_query.join(User, Letter.author_id == User.id).filter(User.department_id.in_(department_ids))
    letter_numbers = letter_query.all()
    for value in letter_numbers:
        reg_number = (value[0] or "").strip()
        if series:
            match = re.search(r"(?:-|^)(\d+)$", reg_number)
        elif prefix:
            match = re.search(rf"^{re.escape(prefix)}-(\d+)$", reg_number)
        else:
            match = re.search(r"^(\d+)$", reg_number)
        if match:
            max_number = max(max_number, int(match.group(1)))
    next_number = max_number + 1
    return f"{prefix}-{next_number}" if prefix else str(next_number)


def get_app_setting(key: str, default: str = "") -> str:
    row = db.session.execute(
        sql_text("SELECT value FROM app_settings WHERE key = :key"),
        {"key": key},
    ).fetchone()
    if not row or row[0] is None:
        return default
    return str(row[0])


def set_app_setting(key: str, value: str) -> None:
    db.session.execute(
        sql_text("""
            INSERT INTO app_settings (key, value)
            VALUES (:key, :value)
            ON CONFLICT(key) DO UPDATE SET value = excluded.value
        """),
        {"key": key, "value": value},
    )
    db.session.commit()


def get_app_setting_int(key: str, default: int = 1) -> int:
    raw = (get_app_setting(key, str(default)) or "").strip()
    try:
        value = int(raw)
    except (TypeError, ValueError):
        return default
    return value if value > 0 else default


def ensure_memo_source() -> int:
    source = LetterSource.query.filter_by(name="Служебная записка").first()
    if source:
        return source.id
    source = LetterSource(name="Служебная записка")
    db.session.add(source)
    db.session.commit()
    return source.id


def generate_next_memo_reg_number(department_id=None) -> str:
    prefix = get_department_prefix(
        department_id,
        "memo_prefix",
        get_app_setting("memo_prefix", "421"),
    ) or "421"
    start_number = get_department_start_number(
        department_id,
        "memo_start_number",
        get_app_setting_int("memo_start_number", 1),
    )
    memo_source_id = ensure_memo_source()
    max_number = max(start_number - 1, 0)
    department = db.session.get(Department, department_id) if department_id else None
    series = (department.memo_number_series or "").strip() if department else ""
    memo_query = db.session.query(Letter.reg_number).filter(Letter.source_id == memo_source_id)
    if series:
        department_ids = [d.id for d in Department.query.filter_by(memo_number_series=series).all()]
        memo_query = memo_query.join(User, Letter.author_id == User.id).filter(User.department_id.in_(department_ids))
    memo_numbers = memo_query.all()
    for value in memo_numbers:
        reg_number = (value[0] or "").strip()
        match = re.search(r"(?:-|^)(\d+)$", reg_number) if series else re.search(rf"^{re.escape(prefix)}-(\d+)$", reg_number)
        if match:
            max_number = max(max_number, int(match.group(1)))
    return f"{prefix}-{max_number + 1}"


def user_can_manage_reg_numbers() -> bool:
    return current_user.is_authenticated and getattr(current_user, "role", None) in ("admin", "superadmin")


def validate_reg_number_uniqueness(reg_number: str, *, is_memo: bool, exclude_letter_id: Optional[int] = None) -> bool:
    query = Letter.query.filter(func.lower(Letter.reg_number) == reg_number.lower())
    if exclude_letter_id is not None:
        query = query.filter(Letter.id != exclude_letter_id)
    memo_source_id = ensure_memo_source()
    if is_memo:
        query = query.filter(Letter.source_id == memo_source_id)
    else:
        query = query.filter(or_(Letter.source_id.is_(None), Letter.source_id != memo_source_id))
    return query.first() is None


def get_chat_sender_name(message: ChatMessage) -> str:
    if getattr(message, "sender", None) and getattr(message.sender, "full_name", None):
        return message.sender.full_name
    if getattr(message, "sender_id", None) == -1:
        return "Суперадмин"
    return "Пользователь"


def serialize_chat_message(message: ChatMessage) -> dict:
    return {
        "id": message.id,
        "sender_id": message.sender_id,
        "sender": get_chat_sender_name(message),
        "receiver_id": message.receiver_id,
        "kind": "general" if message.receiver_id is None else "direct",
        "text": message.text or "",
        "file": url_for("static", filename=message.file_path) if message.file_path else None,
        "time": message.created_at.strftime("%H:%M"),
    }


def _derive_key(passphrase: str, salt: bytes) -> bytes:
    kdf = PBKDF2HMAC(
        algorithm=hashes.SHA256(),
        length=32,
        salt=salt,
        iterations=200_000,
    )
    return kdf.derive(passphrase.encode("utf-8"))

def encrypt_str(plaintext: str, passphrase: str) -> str:
    salt = os.urandom(16)
    key = _derive_key(passphrase, salt)
    aes = AESGCM(key)
    nonce = os.urandom(12)
    ct = aes.encrypt(nonce, plaintext.encode("utf-8"), None)
    return base64.b64encode(salt + nonce + ct).decode("utf-8")

def decrypt_str(token_b64: str, passphrase: str) -> str:
    data = base64.b64decode(token_b64)
    salt, nonce, ct = data[:16], data[16:28], data[28:]
    key = _derive_key(passphrase, salt)
    aes = AESGCM(key)
    pt = aes.decrypt(nonce, ct, None)
    return pt.decode("utf-8")
app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))
app.config.update(
    SECRET_KEY=os.getenv("SECRET_KEY", "dev-secret"),
    SQLALCHEMY_DATABASE_URI="sqlite:///app.db",
    SQLALCHEMY_TRACK_MODIFICATIONS=False,
    UPLOAD_FOLDER=str(UPLOAD_DIR),
)
app.config["SQLALCHEMY_ECHO"] = True
import unicodedata
def set_password(self, raw: str):
    self.password_enc = encrypt_str(raw, SECRET_PASSPHRASE)

def check_password(self, raw: str) -> bool:
    return raw == decrypt_str(self.password_enc, SECRET_PASSPHRASE)
def cleanup_old_chat_files(days=3):
    limit_time = dt.datetime.utcnow() - dt.timedelta(days=days)
    folder = os.path.join(app.root_path, "static", "chat_uploads")
    if os.path.exists(folder):
        now = time.time()
        for f in glob.glob(os.path.join(folder, "*")):
            try:
                if now - os.path.getmtime(f) > days * 86400:
                    os.remove(f)
            except Exception as e:
                print(f"cleanup error {f}: {e}")

with app.app_context():
    cleanup_old_chat_files()

from flask_login import UserMixin

class SuperAdmin(UserMixin):
    def __init__(self):
        self.id = -1
        self.username = SUPERADMIN_LOGIN
        self.full_name = "??????????????????"
        self.role = "superadmin"
        self.is_approved = True

    def get_id(self):
        return str(self.id)


def secure_filename_rus(filename):
    filename = filename.strip()
    filename = filename.replace(" ", "_")
    filename = re.sub(r"[^?-??-?A-Za-z0-9._-]", "", filename)
    return filename[:120]

def extract_text_from_file(path):
    ext = os.path.splitext(path)[-1].lower().strip(".")
    text = ""
    try:
        if ext in ("jpg", "jpeg", "png"):
            text = pytesseract.image_to_string(Image.open(path), lang="rus+eng")
        elif ext == "pdf":
            text_parts = []
            with pdfplumber.open(path) as pdf:
                for page in pdf.pages:
                    page_text = page.extract_text()
                    if page_text and page_text.strip():
                        text_parts.append(page_text)
                    else:
                        im = page.to_image(resolution=300).original
                        ocr_text = pytesseract.image_to_string(im, lang="rus+eng")
                        text_parts.append(ocr_text)
            text = "\n".join(text_parts)
        elif ext == "txt":
            with open(path, "r", encoding="utf-8", errors="ignore") as f:
                text = f.read()
        elif ext == "docx":
            doc = docx.Document(path)
            paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            text = "\n".join(paragraphs)
    except Exception as e:
        print(f"text extraction error from {path}: {e}")
        text = ""
    return text.strip()

def safe_from_json(value):
    if not value:
        return []
    if isinstance(value, list):
        return value
    if isinstance(value, str):
        try:
            return json.loads(value)
        except Exception:
            return []
    return []

app.jinja_env.filters['from_json'] = safe_from_json

DB_PATH = os.path.join(BASE_DIR, "app.db")
app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{DB_PATH}"
db.init_app(app)

ACTIVITY_STATUS_META = {
    "in_progress": {"label": "В работе", "class": "bg-warning text-dark", "style": ""},
    "done": {"label": "Вып.", "class": "bg-success", "style": ""},
    "not_done": {"label": "Не вып.", "class": "bg-danger", "style": ""},
    "postponed": {"label": "Перенос", "class": "text-dark", "style": "background-color:#fd7e14;"},
    "cancelled": {"label": "Отмена", "class": "", "style": "background-color:#6f4e37; color:#fff;"},
    "approved": {"label": "Вып.", "class": "bg-success", "style": ""},
    "rejected": {"label": "Отмена", "class": "", "style": "background-color:#6f4e37; color:#fff;"},
    "open": {"label": "В работе", "class": "bg-warning text-dark", "style": ""},
}

ACTIVITY_STATUS_FILTERS = (
    ("active_and_not_done", "В работе и не выполнено"),
    ("in_progress", "В работе"),
    ("not_done", "Не выполнено"),
    ("postponed", "Перенос"),
    ("done", "Выполнено"),
    ("cancelled", "Отмена"),
    ("all", "Все статусы"),
)

ACTIVITY_PRIORITY_META = {
    1: {"label": "1 - низкий", "class": "bg-secondary"},
    2: {"label": "2 - обычный", "class": "bg-info text-dark"},
    3: {"label": "3 - средний", "class": "bg-primary"},
    4: {"label": "4 - высокий", "class": "bg-warning text-dark"},
    5: {"label": "5 - критический", "class": "bg-danger"},
}


def get_activity_priority_meta(priority):
    try:
        priority = int(priority)
    except (TypeError, ValueError):
        priority = 3
    return ACTIVITY_PRIORITY_META.get(priority, ACTIVITY_PRIORITY_META[3])


def parse_activity_priority(raw_priority):
    try:
        priority = int(raw_priority)
    except (TypeError, ValueError):
        return None
    return priority if priority in ACTIVITY_PRIORITY_META else None


def apply_activity_status_filter(query, selected_status):
    """Apply UI status groups while keeping legacy status values visible."""
    if selected_status == "active_and_not_done":
        return query.filter(Activity.status.in_(("open", "in_progress", "not_done")))
    if selected_status == "in_progress":
        return query.filter(Activity.status.in_(("open", "in_progress")))
    if selected_status == "done":
        return query.filter(Activity.status.in_(("done", "approved")))
    if selected_status == "cancelled":
        return query.filter(Activity.status.in_(("cancelled", "rejected")))
    if selected_status in ("not_done", "postponed"):
        return query.filter(Activity.status == selected_status)
    return query


def activity_work_is_locked(activity):
    return activity.status in ("done", "not_done")


def can_modify_activity_work(activity):
    return (
        activity.owner_id == current_user.id
        or current_user.role in ("manager", "admin", "superadmin", "deputy")
        or (
            activity.assigned_department_id is not None
            and activity.assigned_department_id == getattr(current_user, "department_id", None)
        )
    )


def add_activity_history(activity, action, details):
    db.session.add(ActivityHistory(
        activity_id=activity.id,
        author_id=current_user.id if getattr(current_user, "id", -1) != -1 else None,
        action=action,
        details=details,
    ))

def get_activity_status_meta(status: str) -> dict:
    return ACTIVITY_STATUS_META.get(status, {"label": status or "-", "class": "bg-secondary", "style": ""})


@app.context_processor
def inject_template_helpers():
    deps = []
    try:
        deps = Department.query.order_by(Department.name).all()
    except Exception:
        deps = []
    return {
        "get_status_meta": get_activity_status_meta,
        "get_priority_meta": get_activity_priority_meta,
        "deps": deps,
    }


app.jinja_env.globals["get_status_meta"] = get_activity_status_meta
app.jinja_env.globals["get_priority_meta"] = get_activity_priority_meta


@app.get("/manifest.webmanifest")
def pwa_manifest():
    return send_from_directory(app.static_folder, "manifest.webmanifest", mimetype="application/manifest+json")


@app.get("/sw.js")
def service_worker():
    response = send_from_directory(app.static_folder, "sw.js", mimetype="application/javascript")
    response.headers["Service-Worker-Allowed"] = "/"
    response.headers["Cache-Control"] = "no-cache"
    return response

def normalize_activity_statuses():
    # Date fields are stored at midnight, but a deadline remains valid until
    # the end of that calendar day.
    today_start = dt.datetime.combine(dt.date.today(), dt.time.min)
    changed = (
        Activity.query
        .filter(Activity.end_date.isnot(None))
        .filter(Activity.end_date < today_start)
        .filter(Activity.status.in_(["open", "in_progress", "postponed"]))
        .update({"status": "not_done"}, synchronize_session=False)
    )
    if changed:
        db.session.commit()

def ensure_runtime_schema():
    with db.engine.begin() as conn:
        conn.execute(sql_text("""
            CREATE TABLE IF NOT EXISTS user_notifications (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER NOT NULL REFERENCES users(id),
                sender_id INTEGER REFERENCES users(id),
                title VARCHAR(180) NOT NULL,
                message TEXT NOT NULL,
                created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
                read_at DATETIME
            )
        """))
        conn.execute(sql_text("""
            CREATE INDEX IF NOT EXISTS ix_user_notifications_user_read
            ON user_notifications (user_id, read_at, created_at)
        """))
        conn.execute(sql_text("""
            CREATE TABLE IF NOT EXISTS app_settings (
                key VARCHAR(100) PRIMARY KEY,
                value TEXT
            )
        """))
        conn.execute(
            sql_text("INSERT OR IGNORE INTO app_settings (key, value) VALUES (:key, :value)"),
            {"key": "memo_prefix", "value": "421"},
        )
        conn.execute(
            sql_text("INSERT OR IGNORE INTO app_settings (key, value) VALUES (:key, :value)"),
            {"key": "letter_prefix", "value": ""},
        )
        conn.execute(
            sql_text("INSERT OR IGNORE INTO app_settings (key, value) VALUES (:key, :value)"),
            {"key": "letter_start_number", "value": "1"},
        )
        conn.execute(
            sql_text("INSERT OR IGNORE INTO app_settings (key, value) VALUES (:key, :value)"),
            {"key": "memo_start_number", "value": "1"},
        )
        conn.execute(sql_text("""
            CREATE TABLE IF NOT EXISTS letter_sources (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name VARCHAR(200) NOT NULL UNIQUE,
                bin_code VARCHAR(50),
                address VARCHAR(255),
                phone VARCHAR(100),
                email VARCHAR(150),
                contact_person VARCHAR(150),
                notes TEXT
            )
        """))
        for source_name in ("Входящее", "Исходящее", "Электронная почта"):
            conn.execute(
                sql_text("INSERT OR IGNORE INTO letter_sources (name) VALUES (:name)"),
                {"name": source_name},
            )
        letter_source_cols = {row[1] for row in conn.execute(sql_text("PRAGMA table_info(letter_sources)")).fetchall()}
        if "bin_code" not in letter_source_cols:
            conn.execute(sql_text("ALTER TABLE letter_sources ADD COLUMN bin_code VARCHAR(50)"))
        if "address" not in letter_source_cols:
            conn.execute(sql_text("ALTER TABLE letter_sources ADD COLUMN address VARCHAR(255)"))
        if "phone" not in letter_source_cols:
            conn.execute(sql_text("ALTER TABLE letter_sources ADD COLUMN phone VARCHAR(100)"))
        if "email" not in letter_source_cols:
            conn.execute(sql_text("ALTER TABLE letter_sources ADD COLUMN email VARCHAR(150)"))
        if "contact_person" not in letter_source_cols:
            conn.execute(sql_text("ALTER TABLE letter_sources ADD COLUMN contact_person VARCHAR(150)"))
        if "notes" not in letter_source_cols:
            conn.execute(sql_text("ALTER TABLE letter_sources ADD COLUMN notes TEXT"))

        department_cols = {row[1] for row in conn.execute(sql_text("PRAGMA table_info(departments)")).fetchall()}
        if department_cols:
            if "letter_prefix" not in department_cols:
                conn.execute(sql_text("ALTER TABLE departments ADD COLUMN letter_prefix VARCHAR(100)"))
            if "memo_prefix" not in department_cols:
                conn.execute(sql_text("ALTER TABLE departments ADD COLUMN memo_prefix VARCHAR(100)"))
            if "letter_start_number" not in department_cols:
                conn.execute(sql_text("ALTER TABLE departments ADD COLUMN letter_start_number INTEGER"))
            if "memo_start_number" not in department_cols:
                conn.execute(sql_text("ALTER TABLE departments ADD COLUMN memo_start_number INTEGER"))
            if "letter_number_series" not in department_cols:
                conn.execute(sql_text("ALTER TABLE departments ADD COLUMN letter_number_series VARCHAR(100)"))
            if "memo_number_series" not in department_cols:
                conn.execute(sql_text("ALTER TABLE departments ADD COLUMN memo_number_series VARCHAR(100)"))

        cols = {row[1] for row in conn.execute(sql_text("PRAGMA table_info(activities)")).fetchall()}
        if "outgoing_document" not in cols:
            conn.execute(sql_text("ALTER TABLE activities ADD COLUMN outgoing_document VARCHAR(255)"))
        if "assigned_department_id" not in cols:
            conn.execute(sql_text("ALTER TABLE activities ADD COLUMN assigned_department_id INTEGER"))
        if "status_comment" not in cols:
            conn.execute(sql_text("ALTER TABLE activities ADD COLUMN status_comment TEXT"))
        if "postponed_to" not in cols:
            conn.execute(sql_text("ALTER TABLE activities ADD COLUMN postponed_to DATETIME"))
        if "archived_at" not in cols:
            conn.execute(sql_text("ALTER TABLE activities ADD COLUMN archived_at DATETIME"))
        if "archived_by_id" not in cols:
            conn.execute(sql_text("ALTER TABLE activities ADD COLUMN archived_by_id INTEGER"))
        if "completion_memo_id" not in cols:
            conn.execute(sql_text("ALTER TABLE activities ADD COLUMN completion_memo_id INTEGER"))
        if "created_by_id" not in cols:
            conn.execute(sql_text("ALTER TABLE activities ADD COLUMN created_by_id INTEGER"))
        document_cols = {row[1] for row in conn.execute(sql_text("PRAGMA table_info(activity_documents)")).fetchall()}
        if "uploaded_by_id" not in document_cols:
            conn.execute(sql_text("ALTER TABLE activity_documents ADD COLUMN uploaded_by_id INTEGER"))
        conn.execute(sql_text("""CREATE TABLE IF NOT EXISTS activity_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT, activity_id INTEGER NOT NULL REFERENCES activities(id),
            author_id INTEGER REFERENCES users(id), action VARCHAR(100) NOT NULL, details TEXT NOT NULL,
            created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP)"""))
        conn.execute(sql_text("""
            UPDATE activities
            SET created_by_id = (
                SELECT h.author_id FROM activity_history h
                WHERE h.activity_id = activities.id
                  AND h.action = 'Создание задачи'
                  AND h.author_id IS NOT NULL
                ORDER BY h.created_at ASC
                LIMIT 1
            )
            WHERE created_by_id IS NULL
        """))
        conn.execute(sql_text("""
            CREATE TABLE IF NOT EXISTS activity_templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name VARCHAR(200) NOT NULL UNIQUE,
                title VARCHAR(255) NOT NULL,
                description TEXT,
                type_id INTEGER REFERENCES activity_types(id),
                direction VARCHAR(255),
                outgoing_document VARCHAR(255),
                duration_days INTEGER,
                created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
        """))
        conn.execute(sql_text("""
            CREATE TABLE IF NOT EXISTS memo_templates (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name VARCHAR(200) NOT NULL UNIQUE,
                subject VARCHAR(255) NOT NULL,
                body TEXT,
                department_id INTEGER REFERENCES departments(id),
                executor_id INTEGER REFERENCES users(id),
                created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP
            )
        """))
        # Correct titles made before memo-aware task creation was introduced.
        conn.execute(sql_text("""
            UPDATE activities
            SET title = 'Ответ на служебную записку №' || COALESCE(
                (SELECT reg_number FROM letters WHERE letters.id = activities.source_letter_id),
                activities.source_letter_id
            )
            WHERE title LIKE 'Ответ на письмо №%'
              AND source_letter_id IN (
                  SELECT l.id FROM letters l
                  JOIN letter_sources s ON s.id = l.source_id
                  WHERE s.name = 'Служебная записка'
              )
        """))

        letter_cols = {row[1] for row in conn.execute(sql_text("PRAGMA table_info(letters)")).fetchall()}
        if "reg_number" not in letter_cols:
            conn.execute(sql_text("ALTER TABLE letters ADD COLUMN reg_number VARCHAR(100)"))
        if "source_id" not in letter_cols:
            conn.execute(sql_text("ALTER TABLE letters ADD COLUMN source_id INTEGER"))
        if "executor_id" not in letter_cols:
            conn.execute(sql_text("ALTER TABLE letters ADD COLUMN executor_id INTEGER"))
        if "archived_at" not in letter_cols:
            conn.execute(sql_text("ALTER TABLE letters ADD COLUMN archived_at DATETIME"))
        if "archived_by_id" not in letter_cols:
            conn.execute(sql_text("ALTER TABLE letters ADD COLUMN archived_by_id INTEGER"))
        if "approval_status" not in letter_cols:
            conn.execute(sql_text("ALTER TABLE letters ADD COLUMN approval_status VARCHAR(20) NOT NULL DEFAULT 'pending'"))
        if "approved_by_id" not in letter_cols:
            conn.execute(sql_text("ALTER TABLE letters ADD COLUMN approved_by_id INTEGER"))
        if "approved_at" not in letter_cols:
            conn.execute(sql_text("ALTER TABLE letters ADD COLUMN approved_at DATETIME"))

        shared_link_cols = {row[1] for row in conn.execute(sql_text("PRAGMA table_info(shared_links)")).fetchall()}
        if shared_link_cols and "payload" not in shared_link_cols:
            conn.execute(sql_text("ALTER TABLE shared_links ADD COLUMN payload JSON"))

        plan_cols = {row[1] for row in conn.execute(sql_text("PRAGMA table_info(plans)")).fetchall()}
        if "approval_status" not in plan_cols:
            conn.execute(sql_text("ALTER TABLE plans ADD COLUMN approval_status VARCHAR(32) NOT NULL DEFAULT 'approved'"))
        if "approved_by_id" not in plan_cols:
            conn.execute(sql_text("ALTER TABLE plans ADD COLUMN approved_by_id INTEGER"))
        if "approved_at" not in plan_cols:
            conn.execute(sql_text("ALTER TABLE plans ADD COLUMN approved_at DATETIME"))
        if "revision_comment" not in plan_cols:
            conn.execute(sql_text("ALTER TABLE plans ADD COLUMN revision_comment TEXT"))

        # Plan rows are stored separately from the legacy free-text plans so old
        # plans remain available while new monthly plans get measurable results.
        conn.execute(sql_text("""
            CREATE TABLE IF NOT EXISTS plan_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                plan_id INTEGER NOT NULL REFERENCES plans(id) ON DELETE CASCADE,
                executor_id INTEGER NOT NULL REFERENCES users(id),
                position INTEGER NOT NULL DEFAULT 1,
                task_text TEXT NOT NULL,
                deadline_kind VARCHAR(32) NOT NULL DEFAULT 'month',
                status VARCHAR(32) NOT NULL DEFAULT 'planned'
                    CHECK (status IN ('planned', 'done', 'not_done', 'cancelled')),
                comment TEXT,
                completed_at DATETIME,
                created_at DATETIME DEFAULT CURRENT_TIMESTAMP,
                updated_at DATETIME DEFAULT CURRENT_TIMESTAMP
            )
        """))
        conn.execute(sql_text("""
            CREATE INDEX IF NOT EXISTS ix_plan_items_plan_executor
            ON plan_items (plan_id, executor_id)
        """))
        plan_item_cols = {row[1] for row in conn.execute(sql_text("PRAGMA table_info(plan_items)")).fetchall()}
        if "execution_approved" not in plan_item_cols:
            conn.execute(sql_text("ALTER TABLE plan_items ADD COLUMN execution_approved BOOLEAN NOT NULL DEFAULT 0"))
        if "execution_approved_by_id" not in plan_item_cols:
            conn.execute(sql_text("ALTER TABLE plan_items ADD COLUMN execution_approved_by_id INTEGER"))
        if "execution_approved_at" not in plan_item_cols:
            conn.execute(sql_text("ALTER TABLE plan_items ADD COLUMN execution_approved_at DATETIME"))
        if "deadline_date" not in plan_item_cols:
            conn.execute(sql_text("ALTER TABLE plan_items ADD COLUMN deadline_date DATE"))
        if "linked_activity_id" not in plan_item_cols:
            conn.execute(sql_text("ALTER TABLE plan_items ADD COLUMN linked_activity_id INTEGER REFERENCES activities(id)"))

        activity_cols = {row[1] for row in conn.execute(sql_text("PRAGMA table_info(activities)")).fetchall()}
        if "priority" not in activity_cols:
            conn.execute(sql_text("ALTER TABLE activities ADD COLUMN priority INTEGER NOT NULL DEFAULT 3"))

        create_sql_row = conn.execute(
            sql_text("SELECT sql FROM sqlite_master WHERE type='table' AND name='activities'")
        ).fetchone()
        create_sql = create_sql_row[0] if create_sql_row else ""
        if create_sql and "not_done" not in create_sql:
            conn.execute(sql_text("PRAGMA foreign_keys=OFF"))
            conn.execute(sql_text("ALTER TABLE activities RENAME TO activities__old_status_migration"))
            conn.execute(sql_text("""
                CREATE TABLE activities (
                  id INTEGER PRIMARY KEY AUTOINCREMENT,
                  type_id INTEGER NOT NULL,
                  owner_id INTEGER NOT NULL,
                  created_by_id INTEGER,
                  title TEXT NOT NULL,
                  description TEXT,
                  start_date TEXT NOT NULL DEFAULT (datetime('now')),
                  end_date TEXT,
                  status TEXT NOT NULL DEFAULT 'open'
                    CHECK (status IN ('open','in_progress','done','approved','rejected','not_done','postponed','cancelled')),
                  priority INTEGER NOT NULL DEFAULT 3,
                  approved_by_id INTEGER,
                  approved_at TEXT,
                  approve_comment TEXT,
                  source_letter_id INTEGER,
                  completion_memo_id INTEGER REFERENCES letters(id),
                  direction TEXT,
                  decision TEXT,
                  status_comment TEXT,
                  postponed_to TEXT,
                  extra_values TEXT,
                  plan_id INTEGER REFERENCES plans(id),
                  outgoing_document VARCHAR(255),
                  assigned_department_id INTEGER REFERENCES departments(id),
                  FOREIGN KEY (type_id) REFERENCES activity_types(id) ON UPDATE CASCADE ON DELETE RESTRICT,
                  FOREIGN KEY (owner_id) REFERENCES users(id) ON UPDATE CASCADE ON DELETE RESTRICT,
                  FOREIGN KEY (approved_by_id) REFERENCES users(id) ON UPDATE CASCADE ON DELETE SET NULL
                )
            """))
            conn.execute(sql_text("""
                INSERT INTO activities (
                  id, type_id, owner_id, created_by_id, title, description, start_date, end_date, status, priority,
                  approved_by_id, approved_at, approve_comment, source_letter_id, completion_memo_id, direction,
                   decision, status_comment, postponed_to, extra_values, plan_id, outgoing_document, assigned_department_id
                )
                SELECT
                  id, type_id, owner_id, created_by_id, title, description, start_date, end_date, status, COALESCE(priority, 3),
                  approved_by_id, approved_at, approve_comment, source_letter_id, completion_memo_id, direction,
                   decision, status_comment, postponed_to, extra_values, plan_id, outgoing_document, assigned_department_id
                FROM activities__old_status_migration
            """))
            conn.execute(sql_text("DROP TABLE activities__old_status_migration"))
            conn.execute(sql_text("CREATE INDEX IF NOT EXISTS idx_activities_type ON activities(type_id)"))
            conn.execute(sql_text("CREATE INDEX IF NOT EXISTS idx_activities_owner ON activities(owner_id)"))
            conn.execute(sql_text("CREATE INDEX IF NOT EXISTS idx_activities_status ON activities(status)"))
            conn.execute(sql_text("CREATE INDEX IF NOT EXISTS idx_activities_start_date ON activities(start_date)"))
            conn.execute(sql_text("PRAGMA foreign_keys=ON"))

def ensure_required_activity_types():
    needed = ["Задача", "Ответ на письмо", "написать письмо", "написать с/з", "ответ на с/з"]
    for name in needed:
        exists = ActivityType.query.filter(ActivityType.name == name).first()
        if not exists:
            db.session.add(ActivityType(name=name))
    db.session.commit()


def cleanup_broken_activity_types():
    canonical_pairs = [
        ("Задача", [1001, 1006]),
        ("Ответ на письмо", [1002, 1007]),
        ("написать письмо", [1003]),
        ("написать с/з", [1004]),
        ("ответ на с/з", [1005]),
    ]
    for canonical_name, broken_ids in canonical_pairs:
        canonical = ActivityType.query.filter(ActivityType.name == canonical_name).first()
        if not canonical:
            continue
        for broken_id in broken_ids:
            broken = ActivityType.query.get(broken_id)
            if broken and canonical.id != broken.id:
                Activity.query.filter(Activity.type_id == broken.id).update(
                    {"type_id": canonical.id},
                    synchronize_session=False,
                )
                db.session.delete(broken)
    db.session.commit()

with app.app_context():
    ensure_runtime_schema()
    ensure_required_activity_types()
    cleanup_broken_activity_types()

login_manager = LoginManager(app)
login_manager.login_view = "login_page"

@login_manager.user_loader
def load_user(user_id):
    if str(user_id) == "-1":
        return SuperAdmin()
    return User.query.get(int(user_id))


REPORT_SHARE_ENDPOINTS = {
    "activities_analytics": "activities",
    "plans_statistics": "plans_statistics",
    "plans_monthly_report": "plans_monthly_report",
}
REPORT_SHARE_OPTIONS = {
    "activities": {"endpoint": "activities_analytics", "label": "Аналитика задач"},
    "plans_statistics": {"endpoint": "plans_statistics", "label": "Аналитика выполнения планов"},
    "plans_monthly_report": {"endpoint": "plans_monthly_report", "label": "Сводный отчет по планам"},
}


def get_active_report_share(endpoint_name):
    report_key = REPORT_SHARE_ENDPOINTS.get(endpoint_name)
    token = (request.args.get("share") or "").strip()
    if not report_key or not token:
        return None
    link = SharedLink.query.filter_by(
        token=token,
        doc_type="report:{}".format(report_key),
        doc_id=0,
    ).first()
    if not link or (link.expires_at and link.expires_at < dt.datetime.utcnow()):
        return None
    return link


def get_report_share_payload(endpoint_name):
    link = get_active_report_share(endpoint_name)
    if not link:
        return None
    payload = link.payload
    if isinstance(payload, str):
        try:
            payload = json.loads(payload)
        except (TypeError, ValueError):
            payload = {}
    return payload if isinstance(payload, dict) else {}


def role_required(*roles):
    def _decor(fn):
        @wraps(fn)
        def wrapper(*args, **kwargs):
            if get_active_report_share(fn.__name__):
                return fn(*args, **kwargs)
            if not current_user.is_authenticated:
                flash("Требуется вход в систему.", "warning")
                return redirect(url_for("login_page"))
            if current_user.role not in roles:
                abort(403)
            return fn(*args, **kwargs)
        return wrapper
    return _decor


def crm_api_required(fn):
    """Authenticate server-to-server CRM requests with the administrator token."""
    @wraps(fn)
    def wrapper(*args, **kwargs):
        expected_token = get_app_setting("integration_api_token", "")
        authorization = (request.headers.get("Authorization") or "").strip()
        supplied_token = request.headers.get("X-API-Key", "").strip()
        if authorization.lower().startswith("bearer "):
            supplied_token = authorization[7:].strip()
        if not expected_token or not supplied_token or not hmac.compare_digest(expected_token, supplied_token):
            return jsonify({"error": "unauthorized", "message": "Передайте действующий API-токен."}), 401
        return fn(*args, **kwargs)
    return wrapper


def crm_api_pagination():
    try:
        limit = int(request.args.get("limit", 100))
        offset = int(request.args.get("offset", 0))
    except ValueError:
        return None, None
    return min(max(limit, 1), 500), max(offset, 0)


def crm_api_date(value):
    if not value:
        return None
    try:
        return dt.datetime.strptime(value, "%Y-%m-%d").date()
    except ValueError:
        return False


def crm_api_user(user):
    if not user:
        return None
    return {
        "id": user.id,
        "full_name": user.full_name,
        "department": user.department.name if user.department else None,
    }


@app.get("/api/crm/v1/health")
@crm_api_required
def crm_api_health():
    return jsonify({"status": "ok", "api_version": "v1", "server_time": dt.datetime.utcnow().isoformat() + "Z"})


@app.get("/api/crm/v1/tasks")
@crm_api_required
def crm_api_tasks():
    limit, offset = crm_api_pagination()
    if limit is None:
        return jsonify({"error": "invalid_pagination"}), 400
    date_from = crm_api_date(request.args.get("date_from"))
    date_to = crm_api_date(request.args.get("date_to"))
    if date_from is False or date_to is False:
        return jsonify({"error": "invalid_date", "message": "Используйте формат YYYY-MM-DD."}), 400
    query = Activity.query.options(
        joinedload(Activity.owner).joinedload(User.department), joinedload(Activity.type),
        joinedload(Activity.assigned_department), joinedload(Activity.source_letter), joinedload(Activity.completion_memo),
    ).filter(Activity.archived_at.is_(None))
    if date_from:
        query = query.filter(or_(Activity.end_date >= dt.datetime.combine(date_from, dt.time.min), Activity.start_date >= dt.datetime.combine(date_from, dt.time.min)))
    if date_to:
        query = query.filter(or_(Activity.start_date <= dt.datetime.combine(date_to, dt.time.max), Activity.end_date <= dt.datetime.combine(date_to, dt.time.max)))
    status = (request.args.get("status") or "").strip()
    if status:
        query = query.filter(Activity.status == status)
    total = query.count()
    rows = query.order_by(Activity.id.desc()).offset(offset).limit(limit).all()
    data = []
    for activity in rows:
        data.append({
            "id": activity.id, "title": activity.title, "description": activity.description,
            "type": {"id": activity.type_id, "name": activity.type.name if activity.type else None},
            "status": activity.status, "status_label": get_activity_status_meta(activity.status)["label"],
            "priority": activity.priority or 3,
            "start_date": iso_or_none(activity.start_date), "end_date": iso_or_none(activity.end_date),
            "postponed_to": iso_or_none(activity.postponed_to), "status_comment": activity.status_comment,
            "direction": activity.direction, "outgoing_document": activity.outgoing_document,
            "owner": crm_api_user(activity.owner),
            "assigned_department": {"id": activity.assigned_department.id, "name": activity.assigned_department.name} if activity.assigned_department else None,
            "plan_id": activity.plan_id,
            "source_document": {"id": activity.source_letter.id, "number": activity.source_letter.reg_number, "subject": activity.source_letter.subject} if activity.source_letter else None,
            "completion_memo": {"id": activity.completion_memo.id, "number": activity.completion_memo.reg_number, "subject": activity.completion_memo.subject} if activity.completion_memo else None,
        })
    return jsonify({"data": data, "pagination": {"total": total, "limit": limit, "offset": offset}})


@app.get("/api/crm/v1/plans")
@crm_api_required
def crm_api_plans():
    limit, offset = crm_api_pagination()
    if limit is None:
        return jsonify({"error": "invalid_pagination"}), 400
    date_from = crm_api_date(request.args.get("date_from"))
    date_to = crm_api_date(request.args.get("date_to"))
    if date_from is False or date_to is False:
        return jsonify({"error": "invalid_date", "message": "Используйте формат YYYY-MM-DD."}), 400
    query = Plan.query.options(
        joinedload(Plan.creator),
        joinedload(Plan.items).joinedload(PlanItem.executor).joinedload(User.department),
        joinedload(Plan.items).joinedload(PlanItem.linked_activity),
    )
    if date_from:
        query = query.filter(Plan.end_date >= date_from)
    if date_to:
        query = query.filter(Plan.start_date <= date_to)
    total = query.count()
    rows = query.order_by(Plan.start_date.desc(), Plan.id.desc()).offset(offset).limit(limit).all()
    data = []
    for plan in rows:
        data.append({
            "id": plan.id, "start_date": iso_or_none(plan.start_date), "end_date": iso_or_none(plan.end_date),
            "text": plan.plan_text, "approval_status": plan.approval_status,
            "created_at": iso_or_none(plan.created_at), "creator": crm_api_user(plan.creator),
            "items": [{
                "id": item.id, "position": item.position, "text": item.task_text,
                "deadline_kind": item.deadline_kind, "deadline_date": iso_or_none(item.deadline_date),
                "status": item.status, "comment": item.comment,
                "completed_at": iso_or_none(item.completed_at), "executor": crm_api_user(item.executor),
                "linked_task": {"id": item.linked_activity.id, "title": item.linked_activity.title} if item.linked_activity else None,
            } for item in plan.items],
        })
    return jsonify({"data": data, "pagination": {"total": total, "limit": limit, "offset": offset}})


@app.get("/api/crm/v1/memos")
@crm_api_required
def crm_api_memos():
    limit, offset = crm_api_pagination()
    if limit is None:
        return jsonify({"error": "invalid_pagination"}), 400
    date_from = crm_api_date(request.args.get("date_from"))
    date_to = crm_api_date(request.args.get("date_to"))
    if date_from is False or date_to is False:
        return jsonify({"error": "invalid_date", "message": "Используйте формат YYYY-MM-DD."}), 400
    query = Letter.query.options(
        joinedload(Letter.author).joinedload(User.department), joinedload(Letter.executor).joinedload(User.department),
        joinedload(Letter.departments),
    ).filter(Letter.source_id == ensure_memo_source(), Letter.archived_at.is_(None))
    if date_from:
        query = query.filter(or_(Letter.letter_date >= date_from, Letter.created_at >= dt.datetime.combine(date_from, dt.time.min)))
    if date_to:
        query = query.filter(or_(Letter.letter_date <= date_to, Letter.created_at <= dt.datetime.combine(date_to, dt.time.max)))
    approval_status = (request.args.get("approval_status") or "").strip()
    if approval_status:
        query = query.filter(Letter.approval_status == approval_status)
    total = query.count()
    rows = query.order_by(Letter.letter_date.desc().nullslast(), Letter.id.desc()).offset(offset).limit(limit).all()
    data = []
    for memo in rows:
        data.append({
            "id": memo.id, "number": memo.reg_number, "subject": memo.subject, "body": memo.body,
            "date": iso_or_none(memo.letter_date), "created_at": iso_or_none(memo.created_at),
            "approval_status": memo.approval_status, "approved_at": iso_or_none(memo.approved_at),
            "author": crm_api_user(memo.author), "executor": crm_api_user(memo.executor),
            "departments": [{"id": department.id, "name": department.name} for department in memo.departments],
        })
    return jsonify({"data": data, "pagination": {"total": total, "limit": limit, "offset": offset}})


@app.route("/reports/external-links", methods=["GET", "POST"])
@role_required("admin", "superadmin")
def report_external_links():
    created_url = None
    if request.method == "POST":
        report_key = (request.form.get("report_key") or "").strip()
        expires_raw = (request.form.get("expires_at") or "").strip()
        if report_key not in REPORT_SHARE_OPTIONS:
            flash("Выберите отчет.", "warning")
            return redirect(request.url)
        try:
            expires_at = dt.datetime.strptime(expires_raw, "%Y-%m-%dT%H:%M")
        except ValueError:
            flash("Укажите дату и время окончания доступа.", "warning")
            return redirect(request.url)
        if expires_at <= dt.datetime.utcnow():
            flash("Дата окончания доступа должна быть в будущем.", "warning")
            return redirect(request.url)
        link = SharedLink(
            doc_type="report:{}".format(report_key), doc_id=0, token=uuid.uuid4().hex,
            created_by=current_user.id, expires_at=expires_at, payload={},
        )
        db.session.add(link)
        db.session.commit()
        endpoint = REPORT_SHARE_OPTIONS[report_key]["endpoint"]
        created_url = url_for(endpoint, share=link.token, _external=True)
        flash("Внешняя ссылка создана.", "success")

    links = []
    for link in SharedLink.query.filter(SharedLink.doc_type.like("report:%")).order_by(SharedLink.created_at.desc()).all():
        report_key = link.doc_type.split(":", 1)[1]
        option = REPORT_SHARE_OPTIONS.get(report_key)
        if not option:
            continue
        links.append({
            "id": link.id, "label": option["label"], "expires_at": link.expires_at,
            "created_at": link.created_at, "url": url_for(option["endpoint"], share=link.token, _external=True),
        })
    return render_template(
        "report_external_links.html", title="Внешние ссылки на отчеты", report_options=REPORT_SHARE_OPTIONS,
        links=links, created_url=created_url, default_expires=(dt.datetime.now() + dt.timedelta(days=7)).strftime("%Y-%m-%dT%H:%M"),
    )


@app.post("/reports/external-links/<int:link_id>/delete")
@role_required("admin", "superadmin")
def delete_report_external_link(link_id):
    link = SharedLink.query.filter(SharedLink.id == link_id, SharedLink.doc_type.like("report:%")).first_or_404()
    db.session.delete(link)
    db.session.commit()
    flash("Внешняя ссылка отозвана.", "success")
    return redirect(url_for("report_external_links"))


def iso_or_none(dtobj):
    return dtobj.isoformat() if dtobj else None

@app.route("/")
def index():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    return redirect(url_for("login_page"))

@app.get("/login")
def login_page():
    login_users = (
        User.query
        .filter_by(is_approved=True)
        .order_by(User.full_name.asc())
        .all()
    )
    return render_template(
        "login.html",
        title="Вход",
        login_users=login_users,
        superadmin_login=SUPERADMIN_LOGIN,
    )
    return render_template("login.html", title="Вход")


@app.get("/api/chat/tray/login-users")
def chat_tray_login_users():
    login_users = (
        User.query
        .filter_by(is_approved=True)
        .order_by(User.full_name.asc())
        .all()
    )

    lines = []
    if SUPERADMIN_LOGIN:
        lines.append(
            "\t".join([
                "USER",
                quote(SUPERADMIN_LOGIN, safe=""),
                quote("Суперадминистратор", safe=""),
            ])
        )

    for user in login_users:
        lines.append(
            "\t".join([
                "USER",
                quote(user.username or "", safe=""),
                quote(user.full_name or user.username or "", safe=""),
            ])
        )

    payload = "\n".join(lines)
    return current_app.response_class(payload, mimetype="text/plain; charset=utf-8")

@app.get("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login_page"))

@app.get("/dashboard")
@login_required
def dashboard():
    from datetime import date, datetime, timedelta
    normalize_activity_statuses()
    my_acts = Activity.query.filter_by(owner_id=current_user.id).count()
    my_letters = Letter.query.filter_by(author_id=current_user.id).count()
    pending_approval = 0
    if current_user.role in ("manager", "admin", "superadmin", "deputy"):
        pending_approval = Activity.query.filter_by(status="done").count()

    today = date.today()
    quarter_start_month = ((today.month - 1) // 3) * 3 + 1
    first_day = date(today.year, quarter_start_month, 1)
    if quarter_start_month == 10:
        next_quarter = date(today.year + 1, 1, 1)
    else:
        next_quarter = date(today.year, quarter_start_month + 3, 1)
    last_day = next_quarter - timedelta(days=1)
    start_str = request.args.get("date_from", first_day.strftime("%Y-%m-%d"))
    end_str = request.args.get("date_to", last_day.strftime("%Y-%m-%d"))
    my_only = request.args.get("my_only") == "1"
    owner_id = request.args.get("owner_id", "").strip()
    selected_status = request.args.get("status", "active_and_not_done").strip() or "active_and_not_done"

    try:
        date_from = datetime.strptime(start_str, "%Y-%m-%d")
        date_to = datetime.combine(datetime.strptime(end_str, "%Y-%m-%d").date(), datetime.max.time())
    except ValueError:
        date_from = datetime.combine(first_day, datetime.min.time())
        date_to = datetime.combine(last_day, datetime.max.time())

    tasks_query = Activity.query.options(
        joinedload(Activity.type),
        joinedload(Activity.owner),
    ).filter(
        Activity.start_date >= date_from,
        Activity.start_date <= date_to,
    ).order_by(
        Activity.start_date.desc(),
        Activity.id.desc(),
    )
    if selected_status not in {value for value, _label in ACTIVITY_STATUS_FILTERS}:
        selected_status = "active_and_not_done"
    tasks_query = apply_activity_status_filter(tasks_query, selected_status)
    privileged_roles = ("admin", "manager", "superadmin", "deputy")
    filter_users = []
    if current_user.role in privileged_roles:
        if my_only:
            tasks_query = tasks_query.filter(Activity.owner_id == current_user.id)
        elif owner_id:
            try:
                tasks_query = tasks_query.filter(Activity.owner_id == int(owner_id))
            except ValueError:
                pass
        filter_users = User.query.filter_by(is_approved=True).order_by(User.full_name).all()
        dashboard_tasks = tasks_query.limit(10).all()
    else:
        dashboard_tasks = tasks_query.filter(Activity.owner_id == current_user.id).limit(10).all()

    return render_template(
        "dashboard.html",
        title="Главная",
        my_acts=my_acts,
        my_letters=my_letters,
        pending_approval=pending_approval,
        dashboard_tasks=dashboard_tasks,
        date_from=start_str,
        date_to=end_str,
        my_only=my_only,
        owner_id=owner_id,
        selected_status=selected_status,
        status_filters=ACTIVITY_STATUS_FILTERS,
        filter_users=filter_users,
        get_status_meta=get_activity_status_meta,
    )


@app.get("/help")
@login_required
def help_page():
    return render_template("help.html", title="Справка")


@app.route("/admin/document-templates", methods=["GET", "POST"])
@role_required("admin", "superadmin")
def document_templates_page():
    if request.method == "POST":
        template_kind = (request.form.get("template_kind") or "").strip()
        name = (request.form.get("name") or "").strip()
        if not name:
            flash("Укажите название шаблона.", "warning")
            return redirect(url_for("document_templates_page"))

        if template_kind == "activity":
            title = (request.form.get("title") or "").strip()
            type_id = (request.form.get("type_id") or "").strip()
            if not title or not type_id.isdigit() or not db.session.get(ActivityType, int(type_id)):
                flash("Для шаблона задачи укажите название задачи и её тип.", "warning")
                return redirect(url_for("document_templates_page"))
            duration_raw = (request.form.get("duration_days") or "").strip()
            try:
                duration_days = int(duration_raw) if duration_raw else None
            except ValueError:
                duration_days = None
            if duration_days is not None and duration_days < 0:
                flash("Срок шаблона не может быть отрицательным.", "warning")
                return redirect(url_for("document_templates_page"))
            template = ActivityTemplate(
                name=name, title=title, description=(request.form.get("description") or "").strip(),
                type_id=int(type_id), direction=(request.form.get("direction") or "").strip(),
                outgoing_document=(request.form.get("outgoing_document") or "").strip(),
                duration_days=duration_days,
            )
        elif template_kind == "memo":
            subject = (request.form.get("subject") or "").strip()
            department_id = (request.form.get("department_id") or "").strip()
            executor_id = (request.form.get("executor_id") or "").strip()
            if not subject:
                flash("Для шаблона служебной записки укажите краткое содержание.", "warning")
                return redirect(url_for("document_templates_page"))
            if department_id and (not department_id.isdigit() or not db.session.get(Department, int(department_id))):
                flash("Выбран некорректный отдел.", "warning")
                return redirect(url_for("document_templates_page"))
            if executor_id and (not executor_id.isdigit() or not db.session.get(User, int(executor_id))):
                flash("Выбран некорректный исполнитель.", "warning")
                return redirect(url_for("document_templates_page"))
            template = MemoTemplate(
                name=name, subject=subject, body=(request.form.get("body") or "").strip(),
                department_id=int(department_id) if department_id else None,
                executor_id=int(executor_id) if executor_id else None,
            )
        else:
            abort(400)

        try:
            db.session.add(template)
            db.session.commit()
        except Exception:
            db.session.rollback()
            flash("Шаблон с таким названием уже существует.", "warning")
            return redirect(url_for("document_templates_page"))
        flash("Шаблон сохранён.", "success")
        return redirect(url_for("document_templates_page"))

    return render_template(
        "document_templates.html", title="Шаблоны документов",
        activity_templates=ActivityTemplate.query.options(joinedload(ActivityTemplate.type)).order_by(ActivityTemplate.name).all(),
        memo_templates=MemoTemplate.query.options(joinedload(MemoTemplate.department), joinedload(MemoTemplate.executor)).order_by(MemoTemplate.name).all(),
        activity_types=ActivityType.query.order_by(ActivityType.name).all(),
        departments=Department.query.order_by(Department.name).all(),
        users=User.query.filter_by(is_approved=True).order_by(User.full_name).all(),
    )


@app.post("/admin/document-templates/<string:template_kind>/<int:template_id>/delete")
@role_required("admin", "superadmin")
def delete_document_template(template_kind, template_id):
    model = ActivityTemplate if template_kind == "activity" else MemoTemplate if template_kind == "memo" else None
    if model is None:
        abort(404)
    template = db.session.get(model, template_id)
    if not template:
        abort(404)
    db.session.delete(template)
    db.session.commit()
    flash("Шаблон удалён.", "success")
    return redirect(url_for("document_templates_page"))


@app.route("/activities/create", methods=["GET", "POST"])
@login_required
def create_activity():
    from datetime import datetime
    types = ActivityType.query.order_by(ActivityType.name).all()
    templates = ActivityTemplate.query.options(joinedload(ActivityTemplate.type)).order_by(ActivityTemplate.name).all()
    template_id = (request.args.get("template_id") or "").strip()
    selected_template = db.session.get(ActivityTemplate, int(template_id)) if template_id.isdigit() else None

    can_assign_executors = current_user.role in ("admin", "superadmin", "deputy")
    if current_user.role in ("admin", "superadmin"):
        users = User.query.filter_by(is_approved=True).order_by(User.full_name).all()
    elif current_user.role == "deputy":
        users = []
        if current_user.department_id:
            users = (
                User.query.filter_by(
                    is_approved=True,
                    department_id=current_user.department_id,
                )
                .order_by(User.full_name)
                .all()
            )
    else:
        users = [current_user]

    existing_dirs = [d[0] for d in db.session.query(Activity.direction).filter(Activity.direction.isnot(None)).distinct().all()]

    if request.method == "POST":
        title = request.form.get("title", "").strip()
        description = request.form.get("description", "").strip()
        type_id = request.form.get("type_id")
        start_date = request.form.get("start_date")
        end_date = request.form.get("end_date")
        priority = parse_activity_priority(request.form.get("priority", "3"))
        direction = request.form.get("direction", "").strip()
        outgoing_document = request.form.get("outgoing_document", "").strip()
        add_to_plan = request.form.get("add_to_plan") == "1"

        executor_ids = (
            request.form.getlist("executors")
            if can_assign_executors
            else [str(current_user.id)]
        )

        if not title:
            flash("Введите название задачи.")
            return redirect(request.url)
        selected_type = None
        if type_id and type_id.isdigit():
            selected_type = db.session.get(ActivityType, int(type_id))
        if not selected_type:
            flash("Выберите тип задачи.", "warning")
            return redirect(request.url)
        if priority is None:
            flash("Выберите приоритет от 1 до 5.", "warning")
            return redirect(request.url)
        if can_assign_executors and not executor_ids:
            message = "Заместителю должен быть назначен отдел." if current_user.role == "deputy" and not current_user.department_id else "Выберите хотя бы одного ответственного."
            flash(message, "warning")
            return redirect(request.url)
        try:
            executor_ids = [int(user_id) for user_id in executor_ids]
        except ValueError:
            flash("Некорректный исполнитель.", "warning")
            return redirect(request.url)
        allowed_executor_ids = {user.id for user in users}
        if not executor_ids or not set(executor_ids).issubset(allowed_executor_ids):
            flash("Можно назначать задачи только сотрудникам доступного вам отдела.", "warning")
            return redirect(request.url)

        task_start = datetime.strptime(start_date, "%Y-%m-%d") if start_date else datetime.now()
        task_end = datetime.strptime(end_date, "%Y-%m-%d") if end_date else task_start
        created = []
        linked_to_plan = 0
        for executor_id in executor_ids:
            plan_id = None
            matched_plan = None
            if add_to_plan:
                # A task may be added only to the selected executor's own plan,
                # never to a plan belonging to another employee.
                matched_plan = (
                    Plan.query
                    .filter(
                        Plan.start_date <= task_start.date(),
                        Plan.end_date >= task_start.date(),
                        or_(
                            Plan.created_by == executor_id,
                            Plan.items.any(PlanItem.executor_id == executor_id),
                        ),
                    )
                    .order_by(
                        (Plan.items.any(PlanItem.executor_id == executor_id)).desc(),
                        Plan.start_date.desc(),
                        Plan.id.desc(),
                    )
                    .first()
                )
                if matched_plan:
                    plan_id = matched_plan.id
                    linked_to_plan += 1
            a = Activity(
                title=title,
                description=description,
                owner_id=executor_id,
                created_by_id=current_user.id if current_user.id != -1 else None,
                type_id=selected_type.id,
                start_date=task_start,
                end_date=datetime.strptime(end_date, "%Y-%m-%d") if end_date else None,
                status="in_progress",
                priority=priority,
                direction=direction,
                outgoing_document=outgoing_document,
                plan_id=plan_id,
            )
            db.session.add(a)
            db.session.flush()
            if matched_plan:
                last_position = (
                    db.session.query(func.max(PlanItem.position))
                    .filter(PlanItem.plan_id == matched_plan.id)
                    .scalar()
                    or 0
                )
                db.session.add(PlanItem(
                    plan_id=matched_plan.id,
                    executor_id=executor_id,
                    position=last_position + 1,
                    task_text=title,
                    deadline_kind="date" if end_date else "month",
                    deadline_date=task_end.date() if end_date else None,
                    linked_activity_id=a.id,
                ))
            created.append(a)
        db.session.commit()
        for activity in created:
            add_activity_history(
                activity,
                "Создание задачи",
                "Задача создана и назначена исполнителю: {}.".format(
                    activity.owner.full_name if activity.owner else "-"
                ),
            )
        db.session.commit()

        message = f"Создано {len(created)} задач."
        if add_to_plan:
            if linked_to_plan:
                message += f" В планы добавлено: {linked_to_plan}."
            else:
                message += " У выбранного исполнителя нет плана на дату начала задачи."
        flash(message)
        return redirect(url_for("activities_page"))

    return render_template(
        "activity_create.html",
        title="Создание задачи",
        types=types,
        users=users,
        existing_directions=existing_dirs,
        templates=templates,
        selected_template=selected_template,
        can_assign_executors=can_assign_executors,
        executor_scope_label=("Сотрудники вашего отдела" if current_user.role == "deputy" else "Ответственный"),
        initial_start_date=datetime.now().strftime("%Y-%m-%d"),
        initial_end_date=(datetime.now().date() + dt.timedelta(days=selected_template.duration_days)).strftime("%Y-%m-%d") if selected_template and selected_template.duration_days is not None else "",
    )

@app.get("/activities")
@login_required
def activities_page():
    from datetime import datetime, date, timedelta
    normalize_activity_statuses()

    today = date.today()
    first_day = date(today.year, today.month, 1)
    next_month = first_day.replace(day=28) + timedelta(days=4)
    last_day = next_month - timedelta(days=next_month.day)

    start_str = request.args.get("date_from", first_day.strftime("%Y-%m-%d"))
    end_str = request.args.get("date_to", last_day.strftime("%Y-%m-%d"))
    my_only = request.args.get("my_only") == "1"
    owner_id = request.args.get("owner_id", "").strip()
    search = request.args.get("search", "").strip()
    selected_status = request.args.get("status", "active_and_not_done").strip() or "active_and_not_done"
    sort = request.args.get("sort", "date_desc").strip()
    page = request.args.get("page", 1, type=int) or 1

    try:
        date_from = datetime.strptime(start_str, "%Y-%m-%d")
        date_to = datetime.strptime(end_str, "%Y-%m-%d")
    except ValueError:
        date_from = datetime.combine(first_day, datetime.min.time())
        date_to = datetime.combine(last_day, datetime.max.time())

    q = Activity.query.options(
        joinedload(Activity.owner), joinedload(Activity.created_by), joinedload(Activity.type)
    ).filter(Activity.archived_at.is_(None))

    privileged_roles = ("admin", "manager", "superadmin", "deputy")
    if current_user.role not in privileged_roles:
        department_id = getattr(current_user, "department_id", None)
        q = q.filter(
            or_(
                Activity.owner_id == current_user.id,
                Activity.assigned_department_id == department_id if department_id else False,
            )
        )
    else:
        if my_only:
            q = q.filter(Activity.owner_id == current_user.id)
        elif owner_id:
            try:
                q = q.filter(Activity.owner_id == int(owner_id))
            except ValueError:
                pass

    q = q.filter(Activity.start_date >= date_from, Activity.start_date <= date_to)
    if selected_status not in {value for value, _label in ACTIVITY_STATUS_FILTERS}:
        selected_status = "active_and_not_done"
    q = apply_activity_status_filter(q, selected_status)
    if search:
        like_pattern = f"%{search}%"
        q = q.filter(
            or_(
                Activity.title.ilike(like_pattern),
                Activity.description.ilike(like_pattern),
                Activity.direction.ilike(like_pattern),
                Activity.outgoing_document.ilike(like_pattern),
            )
        )
    sort_orders = {
        "date_desc": (Activity.start_date.desc(), Activity.id.desc()),
        "date_asc": (Activity.start_date.asc(), Activity.id.asc()),
        "priority_desc": (Activity.priority.desc(), Activity.start_date.desc(), Activity.id.desc()),
        "priority_asc": (Activity.priority.asc(), Activity.start_date.desc(), Activity.id.desc()),
    }
    if sort not in sort_orders:
        sort = "date_desc"
    pagination = q.order_by(*sort_orders[sort]).paginate(page=page, per_page=50, error_out=False)
    activities = pagination.items
    filter_users = []
    if current_user.role in privileged_roles:
        filter_users = User.query.filter_by(is_approved=True).order_by(User.full_name).all()

    return render_template(
        "activities.html",
        title="Список задач",
        activities=activities,
        date_from=start_str,
        date_to=end_str,
        my_only=my_only,
        owner_id=owner_id,
        selected_status=selected_status,
        status_filters=ACTIVITY_STATUS_FILTERS,
        search=search,
        sort=sort,
        filter_users=filter_users,
        get_status_meta=get_activity_status_meta,
        get_priority_meta=get_activity_priority_meta,
        pagination=pagination,
    )


@app.get("/activities/analytics")
@role_required("manager", "admin", "superadmin")
def activities_analytics():
    normalize_activity_statuses()
    shared_payload = get_report_share_payload("activities_analytics")
    today = dt.date.today()
    quarter_start_month = ((today.month - 1) // 3) * 3 + 1
    default_from = dt.date(today.year, quarter_start_month, 1)
    if quarter_start_month == 10:
        default_to = dt.date(today.year + 1, 1, 1) - dt.timedelta(days=1)
    else:
        default_to = dt.date(today.year, quarter_start_month + 3, 1) - dt.timedelta(days=1)

    date_from_value = str((shared_payload or {}).get("date_from") or request.args.get("date_from") or default_from.isoformat()).strip()
    date_to_value = str((shared_payload or {}).get("date_to") or request.args.get("date_to") or default_to.isoformat()).strip()
    selected_department_id = str((shared_payload or {}).get("department_id") or request.args.get("department_id") or "").strip()
    selected_owner_id = str((shared_payload or {}).get("owner_id") or request.args.get("owner_id") or "").strip()
    try:
        date_from = datetime.strptime(date_from_value, "%Y-%m-%d")
        date_to = datetime.combine(datetime.strptime(date_to_value, "%Y-%m-%d").date(), datetime.max.time())
    except ValueError:
        date_from = datetime.combine(default_from, datetime.min.time())
        date_to = datetime.combine(default_to, datetime.max.time())
        date_from_value = default_from.isoformat()
        date_to_value = default_to.isoformat()

    query = Activity.query.options(joinedload(Activity.owner).joinedload(User.department)).filter(
        Activity.start_date >= date_from,
        Activity.start_date <= date_to,
    )
    is_external_share = shared_payload is not None
    is_manager = getattr(current_user, "role", None) == "manager" and not is_external_share
    current_department_id = getattr(current_user, "department_id", None)
    if is_manager:
        # A manager can assess only their own department; without a department,
        # the safe fallback is their personal tasks.
        if current_department_id:
            query = query.join(Activity.owner).filter(User.department_id == current_department_id)
            selected_department_id = str(current_department_id)
        else:
            query = query.filter(Activity.owner_id == current_user.id)
            selected_department_id = ""
    elif selected_department_id:
        try:
            query = query.join(Activity.owner).filter(User.department_id == int(selected_department_id))
        except ValueError:
            selected_department_id = ""
    if selected_owner_id:
        try:
            query = query.filter(Activity.owner_id == int(selected_owner_id))
        except ValueError:
            selected_owner_id = ""

    activities = query.order_by(Activity.start_date.desc(), Activity.id.desc()).all()
    departments_query = Department.query.order_by(Department.name)
    users_query = User.query.filter_by(is_approved=True).order_by(User.full_name)
    if is_manager:
        if current_department_id:
            departments_query = departments_query.filter(Department.id == current_department_id)
            users_query = users_query.filter(User.department_id == current_department_id)
        else:
            departments_query = departments_query.filter(Department.id == -1)
            users_query = users_query.filter(User.id == current_user.id)
    departments = departments_query.all()
    users = users_query.all()

    buckets = {
        "done": {"label": "Выполнено", "color": "#198754"},
        "not_done": {"label": "Не выполнено", "color": "#dc3545"},
        "in_progress": {"label": "В работе", "color": "#ffc107"},
        "postponed": {"label": "Перенос", "color": "#fd7e14"},
        "cancelled": {"label": "Отменено", "color": "#6f4e37"},
    }

    def status_bucket(status):
        if status in ("done", "approved"):
            return "done"
        if status in ("not_done",):
            return "not_done"
        if status in ("cancelled", "rejected"):
            return "cancelled"
        if status == "postponed":
            return "postponed"
        return "in_progress"

    def empty_row(name, department_name=""):
        return {"name": name, "department": department_name, "total": 0, **{key: 0 for key in buckets}}

    department_rows_by_id = {}
    user_rows_by_id = {}
    summary = empty_row("Все задачи")
    for activity in activities:
        bucket = status_bucket(activity.status)
        owner = activity.owner
        department = owner.department if owner else None
        department_id = department.id if department else 0
        department_name = department.name if department else "Без отдела"
        if department_id not in department_rows_by_id:
            department_rows_by_id[department_id] = empty_row(department_name)
        owner_id = owner.id if owner else 0
        owner_name = owner.full_name if owner else "Удаленный пользователь"
        if owner_id not in user_rows_by_id:
            user_rows_by_id[owner_id] = empty_row(owner_name, department_name)
        for row in (summary, department_rows_by_id[department_id], user_rows_by_id[owner_id]):
            row["total"] += 1
            row[bucket] += 1

    def add_completion(row):
        applicable = row["total"] - row["cancelled"]
        row["completion"] = round(row["done"] * 100 / applicable) if applicable else 0
        return row

    summary = add_completion(summary)
    department_rows = sorted(
        (add_completion(row) for row in department_rows_by_id.values()),
        key=lambda row: (-row["total"], row["name"]),
    )
    user_rows = sorted(
        (add_completion(row) for row in user_rows_by_id.values()),
        key=lambda row: (-row["total"], row["name"]),
    )
    chart_data = {
        "status_labels": [value["label"] for value in buckets.values()],
        "status_values": [summary[key] for key in buckets],
        "status_colors": [value["color"] for value in buckets.values()],
        "department_labels": [row["name"] for row in department_rows],
        "department_done": [row["done"] for row in department_rows],
        "department_not_done": [row["not_done"] for row in department_rows],
        "department_work": [row["in_progress"] + row["postponed"] for row in department_rows],
        "user_labels": [row["name"] for row in user_rows],
        "user_done": [row["done"] for row in user_rows],
        "user_not_done": [row["not_done"] for row in user_rows],
        "user_work": [row["in_progress"] + row["postponed"] for row in user_rows],
    }

    extra_rows_by_key = {}
    for activity in activities:
        raw_values = activity.extra_values
        if isinstance(raw_values, str):
            try:
                raw_values = json.loads(raw_values)
            except (TypeError, ValueError):
                raw_values = []
        if not isinstance(raw_values, list):
            continue
        for value in raw_values:
            if not isinstance(value, dict) or value.get("value") in (None, ""):
                continue
            key = str(value.get("id") or value.get("name") or value.get("rusname") or "unknown")
            row = extra_rows_by_key.setdefault(
                key,
                {
                    "name": value.get("rusname") or value.get("name") or "Дополнительное значение",
                    "type": value.get("type") or "string",
                    "filled": 0,
                    "task_ids": set(),
                    "numeric_count": 0,
                    "numeric_sum": 0.0,
                    "true_count": 0,
                    "false_count": 0,
                },
            )
            row["filled"] += 1
            row["task_ids"].add(activity.id)
            if row["type"] in ("int", "float"):
                try:
                    row["numeric_sum"] += float(str(value["value"]).replace(",", "."))
                    row["numeric_count"] += 1
                except (TypeError, ValueError):
                    pass
            elif row["type"] == "bool":
                if str(value["value"]).lower() in ("true", "1", "yes", "да"):
                    row["true_count"] += 1
                else:
                    row["false_count"] += 1

    extra_rows = []
    for row in extra_rows_by_key.values():
        row["tasks_count"] = len(row.pop("task_ids"))
        row["numeric_average"] = (
            round(row["numeric_sum"] / row["numeric_count"], 2)
            if row["numeric_count"]
            else None
        )
        extra_rows.append(row)
    extra_rows.sort(key=lambda row: (-row["tasks_count"], row["name"]))
    chart_data["extra_labels"] = [row["name"] for row in extra_rows[:10]]
    chart_data["extra_tasks_count"] = [row["tasks_count"] for row in extra_rows[:10]]
    return render_template(
        "activities_analytics.html",
        title="Аналитика задач",
        summary=summary,
        department_rows=department_rows,
        user_rows=user_rows,
        chart_data=chart_data,
        departments=departments,
        users=users,
        date_from=date_from_value,
        date_to=date_to_value,
        selected_department_id=selected_department_id,
        selected_owner_id=selected_owner_id,
        is_manager=is_manager,
        extra_rows=extra_rows,
    )

@app.post("/activities")
@login_required
def  add_activity():
    title = request.form.get("title", "Активность")
    type_id = request.form.get("type_id", type=int)
    description = request.form.get("description")

    if not type_id:
        flash("Выберите тип активности")
        return redirect(url_for("activities_page"))

    a = Activity(
        type_id=type_id,
        owner_id=current_user.id,
        created_by_id=current_user.id if current_user.id != -1 else None,
        title=title,
        description=description,
        start_date=dt.datetime.utcnow(),
        status="in_progress"
    )
    db.session.add(a)
    db.session.commit()
    flash(f"Активность #{a.id} создана")
    return redirect(url_for("activities_page"))


@app.get("/activities/<int:activity_id>")
@login_required
def activity_view(activity_id: int):
    normalize_activity_statuses()
    a = (
        Activity.query
        .options(
            joinedload(Activity.logs).joinedload(ActivityLog.children),
            joinedload(Activity.letters),
            joinedload(Activity.assigned_department),
            joinedload(Activity.completion_memo).joinedload(Letter.source),
        )
        .get_or_404(activity_id)
    )
    # --- проверка прав доступа ---
    is_owner = current_user.id == a.owner_id
    is_admin = current_user.role in ("manager", "admin", "superadmin", "deputy")
    is_department_assignee = (
        a.assigned_department_id is not None
        and a.assigned_department_id == getattr(current_user, "department_id", None)
    )

    if not (is_owner or is_admin or is_department_assignee):
        flash("❌ У вас нет доступа к этой активности", "danger")
        return redirect(url_for("index"))
    # --- проверка прав ---
    can_edit = (
        current_user.id == a.owner_id
        or current_user.role in ("manager", "admin", "superadmin", "deputy")
        or is_department_assignee
    )
    status_locked = activity_work_is_locked(a)
    can_extend_deadline = bool(
        can_edit and a.status == "not_done" and a.end_date
        and a.end_date.date() < dt.date.today()
    )
    can_modify_work = can_edit and not status_locked
    can_change_status = can_edit and (
        not status_locked
        or current_user.role in ("admin", "superadmin", "deputy")
        or can_extend_deadline
    )
    can_delegate = current_user.role in ("deputy", "admin", "superadmin")
    closure_memos_query = Letter.query.filter(
        Letter.source_id == ensure_memo_source(),
        Letter.archived_at.is_(None),
    )
    if current_user.role not in ("manager", "admin", "superadmin"):
        closure_memos_query = (
            closure_memos_query
            .outerjoin(Letter.recipients)
            .outerjoin(Letter.departments)
            .filter(or_(
                Letter.author_id == current_user.id,
                Letter.executor_id == current_user.id,
                LetterRecipient.user_id == current_user.id,
                Department.id == getattr(current_user, "department_id", None),
            ))
            .distinct()
        )
    closure_memos = closure_memos_query.order_by(
        Letter.letter_date.desc().nullslast(), Letter.id.desc()
    ).all()

    return render_template(
        "activity_view.html",
        title=f"Активность #{a.id}",
        a=a,
        can_edit=can_edit,   # 👈 передаём флаг в шаблон
        can_modify_work=can_modify_work,
        can_change_status=can_change_status,
        can_extend_deadline=can_extend_deadline,
        status_locked=status_locked,
        datetime=datetime,
        status_meta=get_activity_status_meta(a.status),
        can_delegate=can_delegate,
        closure_memos=closure_memos,
        delegation_users=User.query.filter_by(is_approved=True).order_by(User.full_name).all() if can_delegate else [],
        delegation_departments=Department.query.order_by(Department.name).all() if can_delegate else [],
        history=ActivityHistory.query.filter_by(activity_id=a.id).order_by(ActivityHistory.created_at.desc()).all(),
    )


@app.get("/activities/archive")
@role_required("admin", "superadmin", "deputy")
def activities_archive_page():
    search = (request.args.get("search") or "").strip()
    query = Activity.query.options(joinedload(Activity.owner), joinedload(Activity.type)).filter(
        Activity.archived_at.is_not(None)
    )
    if search:
        like_pattern = "%{}%".format(search)
        query = query.filter(or_(
            Activity.title.ilike(like_pattern),
            Activity.description.ilike(like_pattern),
            Activity.direction.ilike(like_pattern),
        ))
    activities = query.order_by(Activity.archived_at.desc(), Activity.id.desc()).all()
    return render_template(
        "activities_archive.html", title="Архив задач", activities=activities, search=search,
        get_status_meta=get_activity_status_meta, get_priority_meta=get_activity_priority_meta,
    )


@app.post("/activities/<int:activity_id>/archive")
@role_required("admin", "superadmin", "deputy")
def archive_activity(activity_id):
    activity = Activity.query.get_or_404(activity_id)
    activity.archived_at = dt.datetime.utcnow()
    activity.archived_by_id = current_user.id if current_user.id != -1 else None
    add_activity_history(activity, "Архивирование", "Задача перемещена в архив.")
    db.session.commit()
    flash("Задача архивирована.", "success")
    return redirect(request.referrer or url_for("activities_page"))


@app.post("/activities/<int:activity_id>/restore")
@role_required("admin", "superadmin", "deputy")
def restore_activity(activity_id):
    activity = Activity.query.get_or_404(activity_id)
    activity.archived_at = None
    activity.archived_by_id = None
    add_activity_history(activity, "Восстановление", "Задача восстановлена из архива.")
    db.session.commit()
    flash("Задача восстановлена из архива.", "success")
    return redirect(request.referrer or url_for("activity_view", activity_id=activity.id))


@app.get("/activities/audit")
@role_required("manager", "admin", "superadmin", "deputy")
def activities_audit():
    today = dt.date.today()
    default_from = today.replace(day=1)
    date_from_raw = (request.args.get("date_from") or default_from.isoformat()).strip()
    date_to_raw = (request.args.get("date_to") or today.isoformat()).strip()
    activity_id = (request.args.get("activity_id") or "").strip()
    author_id = (request.args.get("author_id") or "").strip()
    action = (request.args.get("action") or "").strip()
    search = (request.args.get("search") or "").strip()
    try:
        date_from = dt.datetime.strptime(date_from_raw, "%Y-%m-%d")
        date_to = dt.datetime.strptime(date_to_raw, "%Y-%m-%d") + dt.timedelta(days=1)
    except ValueError:
        date_from = dt.datetime.combine(default_from, dt.time.min)
        date_to = dt.datetime.combine(today + dt.timedelta(days=1), dt.time.min)
        date_from_raw, date_to_raw = default_from.isoformat(), today.isoformat()

    query = (
        ActivityHistory.query
        .join(Activity, Activity.id == ActivityHistory.activity_id)
        .outerjoin(User, User.id == ActivityHistory.author_id)
        .options(joinedload(ActivityHistory.author))
        .filter(ActivityHistory.created_at >= date_from, ActivityHistory.created_at < date_to)
    )
    if current_user.role == "manager":
        if current_user.department_id:
            query = query.filter(Activity.owner.has(User.department_id == current_user.department_id))
        else:
            query = query.filter(Activity.owner_id == current_user.id)
    if activity_id.isdigit():
        query = query.filter(ActivityHistory.activity_id == int(activity_id))
    if author_id.isdigit():
        query = query.filter(ActivityHistory.author_id == int(author_id))
    if action:
        query = query.filter(ActivityHistory.action == action)
    if search:
        query = query.filter(or_(Activity.title.ilike("%{}%".format(search)), ActivityHistory.details.ilike("%{}%".format(search))))

    rows = query.order_by(ActivityHistory.created_at.desc()).limit(1000).all()
    action_options = [row[0] for row in db.session.query(ActivityHistory.action).distinct().order_by(ActivityHistory.action).all()]
    users_query = User.query.filter_by(is_approved=True).order_by(User.full_name)
    if current_user.role == "manager" and current_user.department_id:
        users_query = users_query.filter(User.department_id == current_user.department_id)
    return render_template(
        "activities_audit.html", title="Аудит задач", rows=rows,
        date_from=date_from_raw, date_to=date_to_raw, activity_id=activity_id,
        author_id=author_id, action=action, search=search,
        action_options=action_options, users=users_query.all(),
    )


@app.post("/activities/<int:activity_id>/redirect")
@role_required("deputy", "admin", "superadmin")
def redirect_activity(activity_id):
    activity = Activity.query.options(joinedload(Activity.owner), joinedload(Activity.assigned_department)).get_or_404(activity_id)
    user_id = (request.form.get("target_user_id") or "").strip()
    department_id = (request.form.get("target_department_id") or "").strip()
    if bool(user_id) == bool(department_id):
        flash("Выберите одного сотрудника или один отдел для перенаправления.", "warning")
        return redirect(url_for("activity_view", activity_id=activity.id))

    previous_target = activity.owner.full_name if activity.owner else "не назначен"
    if activity.assigned_department:
        previous_target += f" (отдел {activity.assigned_department.name})"
    if user_id:
        try:
            target_user = db.session.get(User, int(user_id))
        except ValueError:
            target_user = None
        if not target_user or not target_user.is_approved:
            flash("Выбранный сотрудник недоступен.", "warning")
            return redirect(url_for("activity_view", activity_id=activity.id))
        activity.owner_id = target_user.id
        activity.assigned_department_id = target_user.department_id
        new_target = target_user.full_name
        if target_user.department:
            new_target += f" (отдел {target_user.department.name})"
    else:
        try:
            target_department = db.session.get(Department, int(department_id))
        except ValueError:
            target_department = None
        if not target_department:
            flash("Выбранный отдел не найден.", "warning")
            return redirect(url_for("activity_view", activity_id=activity.id))
        activity.assigned_department_id = target_department.id
        new_target = f"отдел {target_department.name}"

    db.session.add(ActivityLog(
        activity_id=activity.id,
        text=f"Задача перенаправлена: {previous_target} -> {new_target}. Выполнил: {current_user.full_name}.",
        entry_date=dt.datetime.utcnow(),
    ))
    add_activity_history(
        activity,
        "Перенаправление",
        "Задача перенаправлена: {} -> {}.".format(previous_target, new_target),
    )
    db.session.commit()
    flash("Задача перенаправлена. Запись добавлена в журнал.", "success")
    return redirect(url_for("activity_view", activity_id=activity.id))


@app.route("/activities/<int:activity_id>/edit", methods=["GET", "POST"])
@role_required("admin", "superadmin")
def edit_activity(activity_id):
    activity = Activity.query.get_or_404(activity_id)
    users = User.query.filter_by(is_approved=True).order_by(User.full_name).all()
    types = ActivityType.query.order_by(ActivityType.name).all()
    if request.method == "POST":
        previous_values = {
            "title": activity.title or "",
            "description": activity.description or "",
            "owner": activity.owner.full_name if activity.owner else "-",
            "type": activity.type.name if activity.type else "-",
            "start_date": activity.start_date.strftime("%d.%m.%Y") if activity.start_date else "-",
            "end_date": activity.end_date.strftime("%d.%m.%Y") if activity.end_date else "-",
            "status": get_activity_status_meta(activity.status)["label"],
            "priority": activity.priority or 3,
            "direction": activity.direction or "",
            "outgoing_document": activity.outgoing_document or "",
        }
        title = (request.form.get("title") or "").strip()
        if not title:
            flash("Введите название задачи.", "warning")
            return redirect(request.url)
        try:
            owner_id = int(request.form.get("owner_id") or "")
        except ValueError:
            flash("Выберите исполнителя.", "warning")
            return redirect(request.url)
        owner = db.session.get(User, owner_id)
        if not owner:
            flash("Исполнитель не найден.", "warning")
            return redirect(request.url)
        try:
            start_date_raw = (request.form.get("start_date") or "").strip()
            end_date_raw = (request.form.get("end_date") or "").strip()
            start_date = datetime.strptime(start_date_raw, "%Y-%m-%d") if start_date_raw else activity.start_date
            end_date = datetime.strptime(end_date_raw, "%Y-%m-%d") if end_date_raw else None
        except ValueError:
            flash("Некорректный формат даты.", "warning")
            return redirect(request.url)
        status = (request.form.get("status") or "in_progress").strip()
        if status not in ACTIVITY_STATUS_META:
            flash("Некорректный статус задачи.", "warning")
            return redirect(request.url)
        priority = parse_activity_priority(request.form.get("priority", activity.priority))
        if priority is None:
            flash("Выберите приоритет от 1 до 5.", "warning")
            return redirect(request.url)
        type_id = (request.form.get("type_id") or "").strip()
        selected_type = db.session.get(ActivityType, int(type_id)) if type_id.isdigit() else None
        activity.title = title
        activity.description = (request.form.get("description") or "").strip()
        activity.owner_id = owner.id
        activity.type_id = selected_type.id if selected_type else None
        activity.start_date = start_date
        activity.end_date = end_date
        activity.status = status
        activity.priority = priority
        activity.direction = (request.form.get("direction") or "").strip()
        activity.outgoing_document = (request.form.get("outgoing_document") or "").strip()
        updated_values = {
            "Название": (previous_values["title"], activity.title or ""),
            "Описание": (previous_values["description"], activity.description or ""),
            "Исполнитель": (previous_values["owner"], owner.full_name),
            "Тип": (previous_values["type"], selected_type.name if selected_type else "-"),
            "Дата начала": (previous_values["start_date"], activity.start_date.strftime("%d.%m.%Y") if activity.start_date else "-"),
            "Дата окончания": (previous_values["end_date"], activity.end_date.strftime("%d.%m.%Y") if activity.end_date else "-"),
            "Статус": (previous_values["status"], get_activity_status_meta(activity.status)["label"]),
            "Приоритет": (get_activity_priority_meta(previous_values["priority"])["label"], get_activity_priority_meta(activity.priority)["label"]),
            "Тема": (previous_values["direction"], activity.direction or ""),
            "Исходный документ": (previous_values["outgoing_document"], activity.outgoing_document or ""),
        }
        changes = ["{}: «{}» -> «{}»".format(label, before or "-", after or "-") for label, (before, after) in updated_values.items() if before != after]
        if changes:
            add_activity_history(activity, "Редактирование задачи", "; ".join(changes))
        db.session.commit()
        flash("Задача обновлена.", "success")
        return redirect(url_for("activity_view", activity_id=activity.id))
    return render_template(
        "activity_edit.html",
        title=f"Редактирование задачи #{activity.id}",
        activity=activity,
        users=users,
        types=types,
        status_meta=ACTIVITY_STATUS_META,
    )


@app.post("/activities/<int:activity_id>/logs")
@login_required
def add_activity_log(activity_id: int):
    a = Activity.query.get_or_404(activity_id)
    if a.owner_id != current_user.id and current_user.role not in ("manager", "admin", "superadmin", "deputy"):
        abort(403)
    if activity_work_is_locked(a):
        abort(403)

    text = request.form.get("text", "").strip()
    entry_date_str = request.form.get("entry_date")  # <-- берём из формы

    if not text:
        return {"success": False, "error": "Текст пустой"}, 400

    # Парсим дату/время, если передана
    from datetime import datetime as _dt
    entry_dt = None
    if entry_date_str:
        try:
            entry_dt = _dt.strptime(entry_date_str, "%Y-%m-%dT%H:%M")
        except ValueError:
            entry_dt = _dt.utcnow()

    log = ActivityLog(
        activity_id=a.id,
        text=text,
        entry_date=entry_dt or _dt.utcnow()
    )
    db.session.add(log)
    add_activity_history(a, "Журнал действий", "Добавлено действие: {}".format(text[:300]))
    db.session.commit()

    return redirect(url_for("activity_view", activity_id=a.id))

import uuid
from werkzeug.utils import secure_filename
import mimetypes



@app.post("/activities/<int:activity_id>/upload")
@login_required
def upload_activity_document(activity_id):
    a = Activity.query.get_or_404(activity_id)
    if activity_work_is_locked(a) or not can_modify_activity_work(a):
        abort(403)
    file = request.files.get("file")

    if not file or not file.filename:
        flash("⚠️ Файл не выбран", "warning")
        return redirect(url_for("activity_view", activity_id=a.id))

    # Проверяем расширение
    allowed_ext = {"png", "jpg", "jpeg", "pdf", "txt", "docx"}
    ext = file.filename.rsplit(".", 1)[-1].lower()
    if ext not in allowed_ext:
        flash(f"🚫 Недопустимый формат файла: .{ext}", "danger")
        return redirect(url_for("activity_view", activity_id=a.id))

    # Проверяем MIME-тип
    mime, _ = mimetypes.guess_type(file.filename)
    if mime and not mime.startswith(("image/", "application/", "text/")):
        flash("🚫 Неверный тип файла", "danger")
        return redirect(url_for("activity_view", activity_id=a.id))

    # Создаём папку если нет
    upload_dir = os.path.join(app.root_path, "static", "uploads")
    os.makedirs(upload_dir, exist_ok=True)

    # Генерируем безопасное имя
    filename = secure_filename_rus(file.filename)
    random_name = f"{uuid.uuid4().hex}_{filename}"
    save_path = os.path.join(upload_dir, random_name)
    file.save(save_path)

    # Ограничиваем размер до 10 МБ
    max_size = 10 * 1024 * 1024
    if os.path.getsize(save_path) > max_size:
        os.remove(save_path)
        flash("🚫 Файл слишком большой (макс. 10 МБ)", "danger")
        return redirect(url_for("activity_view", activity_id=a.id))

    # Распознаём текст
    recognized_text = extract_text_from_file(save_path)

    # Сохраняем запись в БД
    doc = ActivityDocument(
        activity_id=a.id,
        filename=filename,
        filepath=f"uploads/{random_name}",  # 🔗 путь внутри static/
        doc_rec=recognized_text,
        uploaded_by_id=current_user.id if current_user.id != -1 else None,
    )
    db.session.add(doc)
    add_activity_history(a, "Документ", "Добавлен документ: {}.".format(filename))
    db.session.commit()

    flash("📄 Документ успешно загружен и распознан", "success")
    return redirect(url_for("activity_view", activity_id=a.id))
def has_document_access(user_id: int, doc_type: str, doc_id: int) -> bool:
    """Проверка прав на документ активности или письма"""
    if current_user.role in ("admin", "manager"):
        return True

    if doc_type == "activity":
        doc = ActivityDocument.query.get(doc_id)
        if not doc:
            return False
        return (doc.activity.owner_id == user_id) or (
            DocumentAccess.query.filter_by(doc_type="activity", doc_id=doc_id, user_id=user_id).count() > 0
        )

    if doc_type == "letter":
        doc = LetterDocument.query.get(doc_id)
        if not doc:
            return False
        return (doc.letter.author_id == user_id) or (
            DocumentAccess.query.filter_by(doc_type="letter", doc_id=doc_id, user_id=user_id).count() > 0
        )

    return False

@app.get("/documents/<int:doc_id>")
@login_required
def download_document(doc_id: int):
    d = ActivityDocument.query.get_or_404(doc_id)
    return send_from_directory(app.config["UPLOAD_FOLDER"], os.path.basename(d.filepath), as_attachment=True)

@app.post("/activities/<int:activity_id>/done")
@login_required
def mark_activity_done(activity_id: int):
    a = Activity.query.get_or_404(activity_id)
    if a.owner_id != current_user.id and current_user.role not in ("manager", "admin"):
        abort(403)
    a.status = "done"
    if not a.end_date:
        a.end_date = dt.datetime.utcnow()
    db.session.commit()
    flash("Отмечено как выполнено")
    return redirect(url_for("activity_view", activity_id=activity_id))

@app.post("/activities/<int:activity_id>/status")
@login_required
def update_activity_status(activity_id: int):
    a = Activity.query.get_or_404(activity_id)
    if not can_modify_activity_work(a):
        abort(403)

    new_status = request.form.get("status", "").strip()
    comment = request.form.get("comment", "").strip()
    completion_memo_id = (request.form.get("completion_memo_id") or "").strip()
    if a.status in ("done", "not_done") and current_user.role not in ("admin", "superadmin", "deputy"):
        if new_status != "extend":
            abort(403)
    allowed_statuses = {"in_progress", "done", "not_done", "postponed", "cancelled", "extend"}
    if new_status not in allowed_statuses:
        flash("Недопустимый статус", "warning")
        return redirect(url_for("activity_view", activity_id=activity_id))

    postponed_to = None
    if new_status == "not_done" and not comment:
        flash("Укажите причину невыполнения", "warning")
        return redirect(url_for("activity_view", activity_id=activity_id))
    if new_status in ("postponed", "extend"):
        postponed_to_raw = request.form.get("postponed_to", "").strip()
        if not postponed_to_raw or not comment:
            flash("Для продления укажите новую дату и причину", "warning")
            return redirect(url_for("activity_view", activity_id=activity_id))
        try:
            postponed_to = dt.datetime.strptime(postponed_to_raw, "%Y-%m-%d")
        except ValueError:
            flash("Некорректная дата переноса", "warning")
            return redirect(url_for("activity_view", activity_id=activity_id))
        if new_status == "extend":
            deadline = a.end_date.date() if a.end_date else None
            if a.status != "not_done" or not deadline or deadline >= dt.date.today():
                flash("Продлить можно только просроченную задачу.", "warning")
                return redirect(url_for("activity_view", activity_id=activity_id))
            if postponed_to.date() < dt.date.today():
                flash("Новый срок не может быть в прошлом.", "warning")
                return redirect(url_for("activity_view", activity_id=activity_id))

    completion_memo = None
    if new_status == "done" and completion_memo_id:
        if not completion_memo_id.isdigit():
            flash("Выберите служебную записку из списка.", "warning")
            return redirect(url_for("activity_view", activity_id=activity_id))
        completion_memo = Letter.query.filter(
            Letter.id == int(completion_memo_id),
            Letter.source_id == ensure_memo_source(),
            Letter.archived_at.is_(None),
        ).first()
        if not completion_memo:
            flash("Выбранная служебная записка не найдена или находится в архиве.", "warning")
            return redirect(url_for("activity_view", activity_id=activity_id))
        memo_access = (
            current_user.role in ("manager", "admin", "superadmin")
            or completion_memo.author_id == current_user.id
            or completion_memo.executor_id == current_user.id
            or any(recipient.user_id == current_user.id for recipient in completion_memo.recipients)
            or any(department.id == getattr(current_user, "department_id", None) for department in completion_memo.departments)
        )
        if not memo_access:
            abort(403)

    previous_status = a.status
    stored_status = "postponed" if new_status == "extend" else new_status
    a.status = stored_status
    a.completion_memo_id = completion_memo.id if new_status == "done" and completion_memo else None
    if new_status == "done" and not a.end_date:
        a.end_date = dt.datetime.utcnow()
    if stored_status == "postponed":
        a.postponed_to = postponed_to
        a.end_date = postponed_to
    elif new_status == "in_progress":
        a.postponed_to = None
    a.status_comment = comment or None
    status_text = "Продление срока" if new_status == "extend" else get_activity_status_meta(stored_status)["label"]
    log_text = "Статус изменён: {} -> {}.".format(
        get_activity_status_meta(previous_status)["label"],
        status_text,
    )
    if stored_status == "postponed":
        log_text += " Новый срок: {}.".format(postponed_to.strftime("%d.%m.%Y"))
    if comment:
        log_text += " Комментарий: {}".format(comment)
    if completion_memo:
        log_text += " Задача закрыта служебной запиской №{}.".format(
            completion_memo.reg_number or completion_memo.id
        )
    db.session.add(ActivityLog(
        activity_id=a.id,
        text=log_text,
        entry_date=dt.datetime.utcnow(),
    ))
    add_activity_history(a, "Смена статуса", log_text)
    db.session.commit()
    flash("Статус задачи обновлён", "success")
    return redirect(url_for("activity_view", activity_id=activity_id))

@app.post("/activities/<int:activity_id>/approve")
@role_required("manager", "admin")
def approve_activity(activity_id: int):
    a = Activity.query.get_or_404(activity_id)
    approve = request.form.get("approve", "true").lower() != "false"
    comment = (request.form.get("comment") or "").strip()
    a.status = "approved" if approve else "in_progress"
    a.approved_by_id = current_user.id
    a.approved_at = dt.datetime.utcnow()
    a.approve_comment = comment
    if not approve:
        log_text = "Утверждение отклонено: задача возвращена в работу."
        if comment:
            log_text += " Комментарий руководителя: {}".format(comment)
        db.session.add(ActivityLog(
            activity_id=a.id,
            text=log_text,
            entry_date=dt.datetime.utcnow(),
        ))
    db.session.commit()
    flash("Задача утверждена" if approve else "Задача возвращена в работу")
    return redirect(url_for("activity_view", activity_id=activity_id))

# =========================
# Письма: HTML
# =========================
from sqlalchemy import or_

@app.get("/letters")
@login_required
def letters_page():
    from datetime import datetime, date, timedelta
    memo_source_id = ensure_memo_source()
    filter_type = request.args.get("filter", "all")
    source_id = (request.args.get("source_id") or "").strip()
    executor_id = (request.args.get("executor_id") or "").strip()
    search = request.args.get("search", "").strip()
    page = request.args.get("page", 1, type=int) or 1

    # даты фильтра
    today = date.today()
    first_day = date(today.year, today.month, 1)
    last_day = (first_day.replace(month=today.month % 12 + 1, day=1) - timedelta(days=1))

    start_str = request.args.get("date_from", first_day.strftime("%Y-%m-%d"))
    end_str = request.args.get("date_to", last_day.strftime("%Y-%m-%d"))

    try:
        date_from = datetime.strptime(start_str, "%Y-%m-%d").date()
        date_to = datetime.strptime(end_str, "%Y-%m-%d").date()
    except ValueError:
        date_from = first_day
        date_to = last_day

    load_options = [
        joinedload(Letter.author),
        joinedload(Letter.recipients).joinedload(LetterRecipient.user),
        joinedload(Letter.departments),
    ]
    if hasattr(Letter, "executor"):
        load_options.append(joinedload(Letter.executor))
    if hasattr(Letter, "source"):
        load_options.append(joinedload(Letter.source))

    q = (
        Letter.query
        .options(*load_options)
        .outerjoin(Letter.recipients)
        .outerjoin(Letter.departments)
        .filter(
            Letter.archived_at.is_(None),
            or_(Letter.source_id.is_(None), Letter.source_id != memo_source_id),
        )
    )

    # === фильтрация по пользователю ===
    if current_user.role in ("admin", "manager", "superadmin"):
        if filter_type == "my":
            q = q.filter(LetterRecipient.user_id == current_user.id)
        elif filter_type == "dept":
            q = q.filter(Department.id == current_user.department_id)
        elif filter_type == "sent":
            q = q.filter(Letter.author_id == current_user.id)
    else:
        if filter_type == "my":
            q = q.filter(LetterRecipient.user_id == current_user.id)
        elif filter_type == "dept":
            q = q.filter(Department.id == current_user.department_id)
        elif filter_type == "sent":
            q = q.filter(Letter.author_id == current_user.id)
        else:
            q = q.filter(
                or_(
                    LetterRecipient.user_id == current_user.id,
                    Department.id == current_user.department_id,
                    Letter.author_id == current_user.id,
                )
            )

    # === фильтрация по датам ===
    q = q.filter(
        or_(
            Letter.letter_date.between(date_from, date_to),
            Letter.created_at.between(date_from, date_to)
        )
    )

    if search:
        like_pattern = f"%{search}%"
        letter_filters = [
            Letter.subject.ilike(like_pattern),
            Letter.body.ilike(like_pattern),
        ]
        if hasattr(Letter, "reg_number"):
            letter_filters.append(Letter.reg_number.ilike(like_pattern))
        q = q.filter(or_(*letter_filters))

    if source_id and hasattr(Letter, "source_id"):
        try:
            q = q.filter(Letter.source_id == int(source_id))
        except ValueError:
            source_id = ""

    if executor_id:
        try:
            executor_int = int(executor_id)
            executor_filters = [
                and_(
                    LetterRecipient.user_id == executor_int,
                    LetterRecipient.role == "executor",
                )
            ]
            if hasattr(Letter, "executor_id"):
                executor_filters.insert(0, Letter.executor_id == executor_int)
            q = q.filter(
                or_(
                    *executor_filters,
                )
            )
        except ValueError:
            executor_id = ""

    pagination = q.distinct().order_by(
        Letter.letter_date.desc().nullslast(),
        Letter.created_at.desc(),
        Letter.id.desc(),
    ).paginate(page=page, per_page=50, error_out=False)
    letters = pagination.items

    sources = LetterSource.query.order_by(LetterSource.name).all()
    source_lookup = {source.id: source for source in sources}
    executor_ids = sorted({
        le.executor_id for le in letters
        if getattr(le, "executor_id", None)
    })
    executor_lookup = {}
    if executor_ids:
        executor_lookup = {
            user.id: user
            for user in User.query.filter(User.id.in_(executor_ids)).all()
        }
    filter_users = []
    if current_user.role in ("admin", "manager", "superadmin"):
        filter_users = User.query.filter_by(is_approved=True).order_by(User.full_name).all()

    return render_template(
        "letters.html",
        title="Письма",
        letters=letters,
        filter_type=filter_type,
        date_from=start_str,
        date_to=end_str,
        sources=sources,
        source_lookup=source_lookup,
        filter_users=filter_users,
        source_id=source_id,
        executor_id=executor_id,
        search=search,
        pagination=pagination,
        executor_lookup=executor_lookup,
    )

@app.get("/letters/new")
@login_required
def new_letter_page():
    deps = Department.query.order_by(Department.name).all()
    users = User.query.filter_by(is_approved=True).order_by(User.full_name).all()
    sources = LetterSource.query.order_by(LetterSource.name).all()
    import datetime
    initial_reg_number = generate_next_letter_reg_number(getattr(current_user, "department_id", None))

    return render_template(
        "letter_create.html",
        title="Создание письма",
        deps=deps,
        users=users,
        sources=sources,
        letter_prefix=get_app_setting("letter_prefix", ""),
        initial_reg_number=initial_reg_number,
        number_department_id=getattr(current_user, "department_id", None),
        can_manage_reg_numbers=user_can_manage_reg_numbers(),
        now=datetime.datetime.now
    )


@app.get("/api/registration-number")
@login_required
def registration_number_preview():
    document_type = (request.args.get("document_type") or "").strip()
    department_id = (request.args.get("department_id") or "").strip()
    try:
        department_id = int(department_id) if department_id else getattr(current_user, "department_id", None)
    except ValueError:
        return {"error": "Некорректный отдел."}, 400

    if document_type == "letter":
        number = generate_next_letter_reg_number(department_id)
    elif document_type == "memo":
        number = generate_next_memo_reg_number(department_id)
    else:
        return {"error": "Неизвестный тип документа."}, 400
    return {"number": number}


@app.post("/letters")
@login_required
def add_letter():
    subject = request.form.get("subject", "").strip()
    body = request.form.get("body", "").strip()
    source_id = (request.form.get("source_id") or "").strip()
    executor_id = (request.form.get("executor_id") or request.form.get("executors") or "").strip()
    requested_reg_number = (request.form.get("reg_number") or "").strip()

    # --- получаем строку даты из формы ---
    letter_date_str = request.form.get("letter_date", "").strip()

    # --- преобразуем строку в объект datetime.date ---
    letter_date = None
    if letter_date_str:
        try:
            letter_date = datetime.strptime(letter_date_str, "%Y-%m-%d").date()
        except ValueError:
            flash("Некорректный формат даты письма.")
            return redirect(url_for("letters_page"))

    executor = None
    if executor_id:
        try:
            executor = db.session.get(User, int(executor_id))
        except ValueError:
            executor = None
        if not executor:
            flash("Некорректный исполнитель.", "warning")
            return redirect(url_for("new_letter_page"))

    number_department_id = executor.department_id if executor else getattr(current_user, "department_id", None)

    # --- создаём объект письма ---
    le = Letter(
        subject=subject,
        body=body,
        author_id=current_user.id,
        letter_date=letter_date
    )
    le.reg_number = (
        requested_reg_number
        if user_can_manage_reg_numbers() and requested_reg_number
        else generate_next_letter_reg_number(number_department_id)
    )
    if not validate_reg_number_uniqueness(le.reg_number, is_memo=False):
        flash("Письмо с таким номером уже существует.", "warning")
        return redirect(url_for("new_letter_page"))

    if source_id:
        try:
            le.source_id = int(source_id)
        except ValueError:
            flash("Некорректное предприятие.", "warning")
            return redirect(url_for("new_letter_page"))

    if executor:
        le.executor_id = executor.id

    # --- добавляем выбранные отделы ---
    for dep_id in request.form.getlist("departments"):
        dep = Department.query.get(dep_id)
        if dep:
            le.departments.append(dep)

    db.session.add(le)
    db.session.commit()

    # --- добавляем исполнителей и ознакомление ---
    # --- добавляем исполнителей и ознакомление ---
    for key, role in [("executors", "executor"), ("observers", "cc")]:
        ids_str = request.form.get(key, "")
        if ids_str:
            for uid in [x.strip() for x in ids_str.split(",") if x.strip()]:
                db.session.add(LetterRecipient(letter_id=le.id, user_id=int(uid), role=role))

    db.session.commit()
    flash("Письмо создано.")
    return redirect(url_for("letters_page"))

# === Основной маршрут ===
@app.post("/letters/<int:letter_id>/upload")
@login_required
def upload_letter_document(letter_id: int):
    letter = Letter.query.get_or_404(letter_id)
    f = request.files.get("file")

    if not f or not f.filename:
        flash("⚠️ Файл не выбран", "warning")
        return redirect(url_for("letter_view", letter_id=letter_id))

    # Разрешённые типы
    allowed_ext = {"png", "jpg", "jpeg", "pdf", "txt", "docx"}
    ext = f.filename.rsplit(".", 1)[-1].lower()
    if ext not in allowed_ext:
        flash(f"🚫 Недопустимый формат файла: .{ext}", "danger")
        return redirect(url_for("letter_view", letter_id=letter_id))

    # Создаём папку
    upload_dir = os.path.join(current_app.root_path, "static", "uploads")
    os.makedirs(upload_dir, exist_ok=True)

    # Безопасное оригинальное имя
    original_name = secure_filename_rus(f.filename)

    # Генерируем случайное имя
    random_name = f"{uuid.uuid4().hex}_{original_name}"
    save_path = os.path.join(upload_dir, random_name)
    f.save(save_path)

    # OCR / Распознавание текста
    recognized_text = extract_text_from_file(save_path)

    # Сохраняем документ в БД
    doc = LetterDocument(
        letter_id=letter.id,
        filename=original_name,
        filepath=f"uploads/{random_name}",
        doc_rec=recognized_text
    )

    db.session.add(doc)
    db.session.commit()

    flash("📄 Документ успешно загружен и распознан", "success")
    return redirect(url_for("letter_view", letter_id=letter.id))
@app.get("/letters/<int:letter_id>")
@login_required
def letter_view(letter_id):
    le = Letter.query.get_or_404(letter_id)
    deps = Department.query.all()
    executor_lookup = {}
    source_lookup = {}
    if getattr(le, "executor_id", None):
        executor_user = User.query.get(le.executor_id)
        if executor_user:
            executor_lookup[executor_user.id] = executor_user
    if getattr(le, "source_id", None):
        source_obj = LetterSource.query.get(le.source_id)
        if source_obj:
            source_lookup[source_obj.id] = source_obj

    recipient_ids = [r.user_id for r in le.recipients]
    is_author = le.author_id == current_user.id
    is_recipient = current_user.id in recipient_ids
    is_admin = current_user.role in ("admin", "superadmin")

    # ✅ Проверяем, выдан ли доступ через DocumentAccess
    has_custom_access = (
        DocumentAccess.query.filter_by(
            doc_type="letter",
            doc_id=le.id,
            user_id=current_user.id
        ).count() > 0
    )

    # 🔒 Финальная проверка всех возможных прав
    has_access = is_author or is_recipient or is_admin or has_custom_access

    if not has_access:
        author = User.query.get(le.author_id)
        return render_template(
            "letter_request_access.html",
            title="Доступ к письму запрещён",
            le=le,
            author=author,
        ), 403

    can_edit = is_author or is_admin

    return render_template(
        "letter_view.html",
        le=le,
        deps=deps,
        executor_lookup=executor_lookup,
        source_lookup=source_lookup,
        can_edit=can_edit,
        title=request.form.get("title", "").strip() or f"????? ?? ?????? ?{le.id}"
    )


@app.post("/letters/<int:letter_id>/archive")
@role_required("admin", "superadmin")
def archive_letter(letter_id):
    letter = Letter.query.get_or_404(letter_id)
    if letter.source_id == ensure_memo_source():
        abort(404)
    letter.archived_at = dt.datetime.utcnow()
    letter.archived_by_id = current_user.id if current_user.id != -1 else None
    db.session.commit()
    flash("Письмо перемещено в архив.", "success")
    return redirect(url_for("letters_page"))


@app.post("/letters/<int:letter_id>/restore")
@role_required("admin", "superadmin")
def restore_letter(letter_id):
    letter = Letter.query.get_or_404(letter_id)
    if letter.source_id == ensure_memo_source():
        abort(404)
    letter.archived_at = None
    letter.archived_by_id = None
    db.session.commit()
    flash("Письмо восстановлено из архива.", "success")
    return redirect(url_for("letter_view", letter_id=letter.id))


@app.get("/letters/archive")
@role_required("admin", "superadmin")
def letters_archive_page():
    memo_source_id = ensure_memo_source()
    letters = (
        Letter.query
        .options(
            joinedload(Letter.author),
            joinedload(Letter.source),
            joinedload(Letter.executor),
            joinedload(Letter.recipients).joinedload(LetterRecipient.user),
        )
        .filter(
            Letter.archived_at.isnot(None),
            or_(Letter.source_id.is_(None), Letter.source_id != memo_source_id),
        )
        .order_by(Letter.archived_at.desc(), Letter.id.desc())
        .all()
    )
    return render_template("letters_archive.html", title="Архив писем", letters=letters)


@app.get("/memos/<int:memo_id>")
@login_required
def memo_view(memo_id):
    memo = Letter.query.get_or_404(memo_id)
    if memo.source_id != ensure_memo_source():
        abort(404)

    recipient_ids = [r.user_id for r in memo.recipients]
    is_author = memo.author_id == current_user.id
    is_recipient = current_user.id in recipient_ids
    is_privileged = current_user.role in ("admin", "superadmin", "manager")
    has_custom_access = (
        DocumentAccess.query.filter_by(
            doc_type="letter",
            doc_id=memo.id,
            user_id=current_user.id,
        ).count() > 0
    )

    has_access = is_author or is_recipient or is_privileged or has_custom_access
    if not has_access:
        author = User.query.get(memo.author_id)
        return render_template(
            "letter_request_access.html",
            title="Доступ к служебной записке запрещён",
            le=memo,
            author=author,
        ), 403

    executor_lookup = {}
    if getattr(memo, "executor_id", None):
        executor_user = User.query.get(memo.executor_id)
        if executor_user:
            executor_lookup[executor_user.id] = executor_user

    can_edit = is_author or is_privileged
    linked_activities = Activity.query.filter(
        or_(Activity.source_letter_id == memo.id, Activity.completion_memo_id == memo.id)
    ).order_by(Activity.id.desc()).all()
    return render_template(
        "memo_view.html",
        memo=memo,
        executor_lookup=executor_lookup,
        can_edit=can_edit,
        linked_activities=linked_activities,
        title=f"Служебная записка {memo.reg_number or memo.id}",
    )


@app.post("/memos/<int:memo_id>/archive")
@role_required("admin", "superadmin")
def archive_memo(memo_id):
    memo = Letter.query.get_or_404(memo_id)
    if memo.source_id != ensure_memo_source():
        abort(404)
    memo.archived_at = dt.datetime.utcnow()
    memo.archived_by_id = current_user.id if current_user.id != -1 else None
    db.session.commit()
    flash("Служебная записка перемещена в архив.", "success")
    return redirect(url_for("memos_page"))


@app.post("/memos/<int:memo_id>/approve")
@role_required("manager", "admin", "superadmin")
def approve_memo(memo_id):
    memo = Letter.query.get_or_404(memo_id)
    if memo.source_id != ensure_memo_source():
        abort(404)
    memo.approval_status = "approved"
    memo.approved_by_id = current_user.id if current_user.id != -1 else None
    memo.approved_at = dt.datetime.utcnow()
    db.session.commit()
    flash("Служебная записка утверждена.", "success")
    return redirect(url_for("memo_view", memo_id=memo.id))


@app.post("/memos/<int:memo_id>/restore")
@role_required("admin", "superadmin")
def restore_memo(memo_id):
    memo = Letter.query.get_or_404(memo_id)
    if memo.source_id != ensure_memo_source():
        abort(404)
    memo.archived_at = None
    memo.archived_by_id = None
    db.session.commit()
    flash("Служебная записка восстановлена из архива.", "success")
    return redirect(url_for("memo_view", memo_id=memo.id))


@app.get("/memos/archive")
@role_required("admin", "superadmin")
def memos_archive_page():
    memo_source_id = ensure_memo_source()
    memos = (
        Letter.query
        .options(
            joinedload(Letter.author),
            joinedload(Letter.recipients).joinedload(LetterRecipient.user),
            joinedload(Letter.departments),
        )
        .filter(Letter.source_id == memo_source_id, Letter.archived_at.isnot(None))
        .order_by(Letter.archived_at.desc())
        .all()
    )
    executor_lookup = {
        user.id: user
        for user in User.query.filter_by(is_approved=True).all()
    }
    return render_template(
        "memos_archive.html",
        title="Архив служебных записок",
        memos=memos,
        executor_lookup=executor_lookup,
    )


@app.route("/memos/<int:memo_id>/edit", methods=["GET", "POST"])
@role_required("admin", "superadmin")
def edit_memo(memo_id):
    memo = Letter.query.get_or_404(memo_id)
    if memo.source_id != ensure_memo_source():
        abort(404)
    departments = Department.query.order_by(Department.name).all()
    users = User.query.filter_by(is_approved=True).order_by(User.full_name).all()
    if request.method == "POST":
        subject = (request.form.get("subject") or "").strip()
        reg_number = (request.form.get("reg_number") or "").strip()
        department_id = (request.form.get("department_id") or "").strip()
        executor_id = (request.form.get("executor_id") or "").strip()
        memo_date_raw = (request.form.get("memo_date") or "").strip()
        if not subject or not reg_number:
            flash("Заполните номер и краткое содержание.", "warning")
            return redirect(request.url)
        try:
            department = db.session.get(Department, int(department_id))
        except ValueError:
            department = None
        if not department:
            flash("Выберите отдел.", "warning")
            return redirect(request.url)
        try:
            memo_date = datetime.strptime(memo_date_raw, "%Y-%m-%d").date() if memo_date_raw else None
        except ValueError:
            flash("Некорректный формат даты.", "warning")
            return redirect(request.url)
        executor = None
        if executor_id:
            try:
                executor = db.session.get(User, int(executor_id))
            except ValueError:
                executor = None
            if not executor:
                flash("Исполнитель не найден.", "warning")
                return redirect(request.url)
        if not validate_reg_number_uniqueness(reg_number, is_memo=True, exclude_letter_id=memo.id):
            flash("Служебная записка с таким номером уже существует.", "warning")
            return redirect(request.url)
        memo.reg_number = reg_number
        memo.subject = subject
        memo.body = (request.form.get("body") or "").strip()
        memo.letter_date = memo_date
        memo.executor_id = executor.id if executor else None
        memo.departments[:] = [department]
        LetterRecipient.query.filter_by(letter_id=memo.id, role="executor").delete(synchronize_session=False)
        if executor:
            db.session.add(LetterRecipient(letter_id=memo.id, user_id=executor.id, role="executor"))
        db.session.commit()
        flash("Служебная записка обновлена.", "success")
        return redirect(url_for("memo_view", memo_id=memo.id))
    return render_template(
        "memo_edit.html",
        title=f"Редактирование служебной записки {memo.reg_number or memo.id}",
        memo=memo,
        departments=departments,
        users=users,
    )

# @app.get("/letters/<int:letter_id>")
# @login_required
# def letter_view(letter_id):
#     le = Letter.query.get_or_404(letter_id)
#     deps = Department.query.all()
#
#     recipient_ids = [r.user_id for r in le.recipients]
#     is_author = le.author_id == current_user.id
#     is_recipient = current_user.id in recipient_ids
#     is_admin = current_user.role == "admin"
#
#     # 🔒 Проверяем доступ
#     has_access = is_author or is_recipient or is_admin
#
#     if not has_access:
#         # Не автор и не получатель — показываем страницу с просьбой дать доступ
#         author = User.query.get(le.author_id)
#         return render_template(
#             "letter_request_access.html",
#             title="Доступ к письму запрещён",
#             le=le,
#             author=author,
#         ), 403
#
#     can_edit = is_author or is_admin
#
#     return render_template(
#         "letter_view.html",
#         le=le,
#         deps=deps,
#         can_edit=can_edit,
#         title=f"Письмо #{le.id}"
#     )
#

@app.post("/letters/<int:letter_id>/read")
@login_required
def mark_letter_read(letter_id: int):
    le = Letter.query.get_or_404(letter_id)
    updated = 0
    for r in le.recipients:
        if (r.user_id == current_user.id) or (r.department_id and r.department_id == current_user.department_id):
            if not r.is_read:
                r.is_read = True
                r.read_at = dt.datetime.utcnow()
                updated += 1
    if updated:
        db.session.commit()
        flash("Письмо отмечено прочитанным")
    return redirect(url_for("letter_view", letter_id=letter_id))

@app.post("/letters/<int:letter_id>/sign")
@login_required
def sign_letter(letter_id: int):
    le = Letter.query.get_or_404(letter_id)
    signed_result = request.form.get("signed_result", "")
    le.signed_result = signed_result
    db.session.commit()
    flash("Подпись-результат сохранена")
    return redirect(url_for("letter_view", letter_id=letter_id))

# Связи «из письма → активность»
@app.route("/letters/<int:letter_id>/create-activity", methods=["GET", "POST"])
@login_required
def create_activity_from_letter(letter_id):
    le = Letter.query.get_or_404(letter_id)

    if request.method == "POST":
        from datetime import datetime

        type_id = request.form.get("type_id")
        start_date = request.form.get("start_date")
        end_date = request.form.get("end_date")
        priority = parse_activity_priority(request.form.get("priority", "3"))
        selected_type = db.session.get(ActivityType, int(type_id)) if type_id and type_id.isdigit() else None
        if not selected_type:
            flash("Выберите тип задачи.", "warning")
            return redirect(request.url)
        if priority is None:
            flash("Выберите приоритет от 1 до 5.", "warning")
            return redirect(request.url)

        memo_source = LetterSource.query.filter_by(name="Служебная записка").first()
        source_label = "служебную записку" if memo_source and le.source_id == memo_source.id else "письмо"
        is_memo = source_label == "служебную записку"


        # --- создаём активность "Ответ на письмо №..." ---
        a = Activity(
            title=request.form.get("title", "").strip() or "Ответ на {} №{}".format(source_label, le.reg_number or le.id),
            description=request.form.get("description", "").strip() or (le.body[:300] + "..."),
            owner_id=current_user.id,
            created_by_id=current_user.id if current_user.id != -1 else None,
            type_id=selected_type.id,
            priority=priority,
            source_letter_id=le.id,
            status="in_progress",
            start_date=datetime.now(),
            direction=request.form.get("direction", "").strip(),
            outgoing_document=request.form.get("outgoing_document", "").strip()
        )

        if start_date:
            a.start_date = datetime.strptime(start_date, "%Y-%m-%d")
        if end_date:
            a.end_date = datetime.strptime(end_date, "%Y-%m-%d")

        db.session.add(a)
        db.session.commit()

        flash("Создана активность: ответ на {} №{}.".format(source_label, le.reg_number or le.id))
        return redirect(url_for("memo_view", memo_id=le.id) if is_memo else url_for("letter_view", letter_id=le.id))

    types = ActivityType.query.order_by(ActivityType.name).all()
    existing_dirs = [d[0] for d in db.session.query(Activity.direction).filter(Activity.direction.isnot(None)).distinct().all()]
    return render_template("activity_create_from_letter.html", le=le, types=types, existing_directions=existing_dirs)

# Связи «из активности → письмо»
@app.get("/memos")
@login_required
def memos_page():
    from datetime import datetime, date, timedelta

    memo_source_id = ensure_memo_source()
    today = date.today()
    first_day = date(today.year, today.month, 1)
    last_day = (first_day.replace(month=today.month % 12 + 1, day=1) - timedelta(days=1))

    start_str = request.args.get("date_from", first_day.strftime("%Y-%m-%d"))
    end_str = request.args.get("date_to", last_day.strftime("%Y-%m-%d"))
    department_id = (request.args.get("department_id") or "").strip()
    executor_id = (request.args.get("executor_id") or "").strip()
    search = request.args.get("search", "").strip()
    page = request.args.get("page", 1, type=int) or 1

    try:
        date_from = datetime.strptime(start_str, "%Y-%m-%d").date()
        date_to = datetime.strptime(end_str, "%Y-%m-%d").date()
    except ValueError:
        date_from = first_day
        date_to = last_day

    q = (
        Letter.query
        .options(
            joinedload(Letter.author),
            joinedload(Letter.recipients).joinedload(LetterRecipient.user),
            joinedload(Letter.departments),
        )
        .outerjoin(Letter.recipients)
        .outerjoin(Letter.departments)
        .filter(Letter.source_id == memo_source_id, Letter.archived_at.is_(None))
    )

    privileged_roles = ("admin", "manager", "superadmin")
    if current_user.role not in privileged_roles:
        q = q.filter(
            or_(
                Letter.author_id == current_user.id,
                Department.id == current_user.department_id,
                LetterRecipient.user_id == current_user.id,
            )
        )

    q = q.filter(
        or_(
            Letter.letter_date.between(date_from, date_to),
            Letter.created_at.between(date_from, date_to),
        )
    )

    if department_id:
        try:
            q = q.filter(Department.id == int(department_id))
        except ValueError:
            department_id = ""

    if executor_id:
        try:
            executor_int = int(executor_id)
            q = q.filter(
                or_(
                    Letter.executor_id == executor_int,
                    and_(LetterRecipient.user_id == executor_int, LetterRecipient.role == "executor"),
                )
            )
        except ValueError:
            executor_id = ""

    if search:
        like_pattern = f"%{search}%"
        q = q.filter(
            or_(
                Letter.reg_number.ilike(like_pattern),
                Letter.subject.ilike(like_pattern),
                Letter.body.ilike(like_pattern),
            )
        )

    pagination = q.distinct().order_by(
        Letter.letter_date.desc().nullslast(),
        Letter.created_at.desc(),
        Letter.id.desc(),
    ).paginate(page=page, per_page=50, error_out=False)
    memos = pagination.items

    departments = Department.query.order_by(Department.name).all()
    department_lookup = {dep.id: dep for dep in departments}
    executor_ids = sorted({memo.executor_id for memo in memos if getattr(memo, "executor_id", None)})
    executor_lookup = {}
    if executor_ids:
        executor_lookup = {user.id: user for user in User.query.filter(User.id.in_(executor_ids)).all()}
    filter_users = []
    if current_user.role in privileged_roles:
        filter_users = User.query.filter_by(is_approved=True).order_by(User.full_name).all()

    return render_template(
        "memos.html",
        title="Служебные записки",
        memos=memos,
        date_from=start_str,
        date_to=end_str,
        departments=departments,
        department_lookup=department_lookup,
        department_id=department_id,
        executor_id=executor_id,
        executor_lookup=executor_lookup,
        filter_users=filter_users,
        search=search,
        pagination=pagination,
    )


@app.get("/memos/analytics")
@role_required("manager", "admin", "superadmin")
def memos_analytics():
    today = dt.date.today()
    default_from = today.replace(day=1)
    default_to = (default_from.replace(day=28) + dt.timedelta(days=4)).replace(day=1) - dt.timedelta(days=1)
    date_from_value = (request.args.get("date_from") or default_from.isoformat()).strip()
    date_to_value = (request.args.get("date_to") or default_to.isoformat()).strip()
    department_id = (request.args.get("department_id") or "").strip()
    executor_id = (request.args.get("executor_id") or "").strip()
    try:
        date_from = dt.datetime.strptime(date_from_value, "%Y-%m-%d").date()
        date_to = dt.datetime.strptime(date_to_value, "%Y-%m-%d").date()
    except ValueError:
        date_from, date_to = default_from, default_to
        date_from_value, date_to_value = default_from.isoformat(), default_to.isoformat()

    query = (
        Letter.query.options(joinedload(Letter.author).joinedload(User.department), joinedload(Letter.executor), joinedload(Letter.departments))
        .outerjoin(Letter.departments)
        .filter(
            Letter.source_id == ensure_memo_source(), Letter.archived_at.is_(None),
            or_(Letter.letter_date.between(date_from, date_to), Letter.created_at.between(date_from, date_to)),
        )
    )
    is_manager = current_user.role == "manager"
    if is_manager and current_user.department_id:
        query, department_id = query.filter(Department.id == current_user.department_id), str(current_user.department_id)
    elif department_id.isdigit():
        query = query.filter(Department.id == int(department_id))
    if executor_id.isdigit():
        query = query.filter(Letter.executor_id == int(executor_id))
    memos = query.distinct().all()
    summary = {"total": len(memos), "approved": 0, "pending": 0}
    department_rows, executor_rows = {}, {}
    for memo in memos:
        status = "approved" if memo.approval_status == "approved" else "pending"
        summary[status] += 1
        department = memo.departments[0] if memo.departments else (memo.author.department if memo.author else None)
        department_name = department.name if department else "Без отдела"
        executor_name = memo.executor.full_name if memo.executor else "Не назначен"
        for bucket, name in ((department_rows, department_name), (executor_rows, executor_name)):
            row = bucket.setdefault(name, {"name": name, "total": 0, "approved": 0, "pending": 0})
            row["total"] += 1
            row[status] += 1
    for row in list(department_rows.values()) + list(executor_rows.values()):
        row["approval_percent"] = round(100 * row["approved"] / row["total"]) if row["total"] else 0
    department_rows = sorted(department_rows.values(), key=lambda row: (-row["total"], row["name"]))
    executor_rows = sorted(executor_rows.values(), key=lambda row: (-row["total"], row["name"]))
    summary["approval_percent"] = round(100 * summary["approved"] / summary["total"]) if summary["total"] else 0
    chart_data = {
        "status": [summary["approved"], summary["pending"]],
        "departments": [row["name"] for row in department_rows],
        "department_approved": [row["approved"] for row in department_rows],
        "department_pending": [row["pending"] for row in department_rows],
    }
    users = User.query.filter_by(is_approved=True).order_by(User.full_name)
    if is_manager and current_user.department_id:
        users = users.filter(User.department_id == current_user.department_id)
    return render_template(
        "memos_analytics.html", title="Аналитика служебных записок", summary=summary,
        department_rows=department_rows, executor_rows=executor_rows, chart_data=chart_data,
        departments=Department.query.order_by(Department.name).all(), users=users.all(),
        date_from=date_from_value, date_to=date_to_value, department_id=department_id,
        executor_id=executor_id, is_manager=is_manager,
    )


@app.get("/memos/new")
@login_required
def new_memo_page():
    departments = Department.query.order_by(Department.name).all()
    users = User.query.filter_by(is_approved=True).order_by(User.full_name).all()
    import datetime
    templates = MemoTemplate.query.options(joinedload(MemoTemplate.department), joinedload(MemoTemplate.executor)).order_by(MemoTemplate.name).all()
    template_id = (request.args.get("template_id") or "").strip()
    selected_template = db.session.get(MemoTemplate, int(template_id)) if template_id.isdigit() else None
    number_department_id = (
        selected_template.department_id if selected_template and selected_template.department_id
        else getattr(current_user, "department_id", None)
    )
    initial_reg_number = generate_next_memo_reg_number(number_department_id)

    return render_template(
        "memo_create.html",
        title="Создание служебной записки",
        departments=departments,
        users=users,
        letter_prefix=get_app_setting("letter_prefix", ""),
        memo_prefix=get_app_setting("memo_prefix", "421"),
        initial_reg_number=initial_reg_number,
        number_department_id=number_department_id,
        can_manage_reg_numbers=user_can_manage_reg_numbers(),
        now=datetime.datetime.now,
        templates=templates,
        selected_template=selected_template,
    )


@app.post("/memos")
@login_required
def add_memo():
    subject = request.form.get("subject", "").strip()
    body = request.form.get("body", "").strip()
    department_id = (request.form.get("department_id") or "").strip()
    executor_id = (request.form.get("executor_id") or "").strip()
    memo_date_str = request.form.get("memo_date", "").strip()
    requested_reg_number = (request.form.get("reg_number") or "").strip()

    memo_date = None
    if memo_date_str:
        try:
            memo_date = datetime.strptime(memo_date_str, "%Y-%m-%d").date()
        except ValueError:
            flash("Некорректный формат даты служебной записки.", "warning")
            return redirect(url_for("new_memo_page"))

    department = None
    if department_id:
        try:
            department = db.session.get(Department, int(department_id))
        except ValueError:
            department = None
    if not department:
        flash("Выберите отдел служебной записки.", "warning")
        return redirect(url_for("new_memo_page"))

    memo = Letter(
        reg_number=(
            requested_reg_number
            if user_can_manage_reg_numbers() and requested_reg_number
            else generate_next_memo_reg_number(department.id)
        ),
        subject=subject,
        body=body,
        author_id=current_user.id,
        source_id=ensure_memo_source(),
        letter_date=memo_date,
    )
    if not validate_reg_number_uniqueness(memo.reg_number, is_memo=True):
        flash("Служебная записка с таким номером уже существует.", "warning")
        return redirect(url_for("new_memo_page"))

    if executor_id:
        try:
            memo.executor_id = int(executor_id)
        except ValueError:
            flash("Некорректный исполнитель.", "warning")
            return redirect(url_for("new_memo_page"))

    memo.departments.append(department)

    db.session.add(memo)
    db.session.commit()

    if memo.executor_id:
        db.session.add(LetterRecipient(letter_id=memo.id, user_id=memo.executor_id, role="executor"))
        db.session.commit()

    flash(f"Служебная записка {memo.reg_number} создана.", "success")
    return redirect(url_for("memos_page"))


@app.get("/admin/settings")
@role_required("admin", "superadmin")
def admin_settings():
    return render_template(
        "admin_settings.html",
        title="Настройки",
        letter_prefix=get_app_setting("letter_prefix", ""),
        memo_prefix=get_app_setting("memo_prefix", "421"),
        letter_start_number=get_app_setting_int("letter_start_number", 1),
        memo_start_number=get_app_setting_int("memo_start_number", 1),
        integration_api_token=get_app_setting("integration_api_token", ""),
        departments=Department.query.order_by(Department.name).all(),
    )


@app.post("/admin/settings")
@role_required("admin", "superadmin")
def save_admin_settings():
    letter_prefix = (request.form.get("letter_prefix") or "").strip()
    memo_prefix = (request.form.get("memo_prefix") or "").strip()
    letter_start_number = (request.form.get("letter_start_number") or "1").strip()
    memo_start_number = (request.form.get("memo_start_number") or "1").strip()
    if not memo_prefix:
        flash("Префикс номера служебной записки не может быть пустым.", "warning")
        return redirect(url_for("admin_settings"))
    try:
        letter_start_number_int = int(letter_start_number)
        memo_start_number_int = int(memo_start_number)
    except ValueError:
        flash("Стартовые номера должны быть целыми числами.", "warning")
        return redirect(url_for("admin_settings"))
    if letter_start_number_int < 1 or memo_start_number_int < 1:
        flash("Стартовые номера должны быть больше нуля.", "warning")
        return redirect(url_for("admin_settings"))

    departments = Department.query.order_by(Department.name).all()
    for department in departments:
        department_letter_prefix = (request.form.get(f"department_letter_prefix_{department.id}") or "").strip()
        department_memo_prefix = (request.form.get(f"department_memo_prefix_{department.id}") or "").strip()
        department_letter_start = (request.form.get(f"department_letter_start_number_{department.id}") or "").strip()
        department_memo_start = (request.form.get(f"department_memo_start_number_{department.id}") or "").strip()
        department_letter_series = (request.form.get(f"department_letter_number_series_{department.id}") or "").strip()
        department_memo_series = (request.form.get(f"department_memo_number_series_{department.id}") or "").strip()
        try:
            department_letter_start_int = int(department_letter_start) if department_letter_start else None
            department_memo_start_int = int(department_memo_start) if department_memo_start else None
        except ValueError:
            flash(f"Стартовые номера отдела «{department.name}» должны быть целыми числами.", "warning")
            return redirect(url_for("admin_settings"))
        if (
            (department_letter_start_int is not None and department_letter_start_int < 1)
            or (department_memo_start_int is not None and department_memo_start_int < 1)
        ):
            flash(f"Стартовые номера отдела «{department.name}» должны быть больше нуля.", "warning")
            return redirect(url_for("admin_settings"))
        department.letter_prefix = department_letter_prefix or None
        department.memo_prefix = department_memo_prefix or None
        department.letter_start_number = department_letter_start_int
        department.memo_start_number = department_memo_start_int
        department.letter_number_series = department_letter_series or None
        department.memo_number_series = department_memo_series or None

    set_app_setting("letter_prefix", letter_prefix)
    set_app_setting("memo_prefix", memo_prefix)
    set_app_setting("letter_start_number", str(letter_start_number_int))
    set_app_setting("memo_start_number", str(memo_start_number_int))
    db.session.commit()
    flash("Настройки сохранены.", "success")
    return redirect(url_for("admin_settings"))


@app.post("/admin/settings/integration-token")
@role_required("admin", "superadmin")
def regenerate_integration_api_token():
    set_app_setting("integration_api_token", secrets.token_urlsafe(32))
    flash("Токен интеграции CRM создан. Передайте его только администратору внешней CRM.", "success")
    return redirect(url_for("admin_settings"))


@app.get("/admin/backups")
@role_required("admin", "superadmin")
def admin_backups():
    backups = []
    for path in list_backups():
        stat = path.stat()
        backups.append({
            "name": path.name,
            "size_kb": round(stat.st_size / 1024, 1),
            "created_at": dt.datetime.fromtimestamp(stat.st_mtime),
        })
    return render_template("admin_backups.html", title="Резервные копии БД", backups=backups)


@app.post("/admin/backups/create")
@role_required("admin", "superadmin")
def create_database_backup():
    try:
        archive = create_backup()
    except (OSError, sqlite3.Error) as exc:
        flash(f"Не удалось создать резервную копию: {exc}", "danger")
    else:
        flash(f"Архив базы данных создан: {archive.name}", "success")
    return redirect(url_for("admin_backups"))


@app.get("/admin/backups/<filename>")
@role_required("admin", "superadmin")
def download_database_backup(filename):
    if Path(filename).name != filename or not filename.endswith(".zip"):
        abort(404)
    archive = BACKUP_DIR / filename
    if not archive.is_file():
        abort(404)
    return send_from_directory(BACKUP_DIR, filename, as_attachment=True, download_name=filename)


@app.get("/admin/letter-sources")
@role_required("admin", "superadmin")
def admin_letter_sources():
    sources = LetterSource.query.order_by(LetterSource.name).all()
    return render_template("admin_letter_sources.html", title="Предприятия", sources=sources)

@app.post("/admin/letter-sources")
@role_required("admin", "superadmin")
def add_letter_source():
    name = request.form.get("name", "").strip()
    bin_code = request.form.get("bin_code", "").strip()
    address = request.form.get("address", "").strip()
    phone = request.form.get("phone", "").strip()
    email = request.form.get("email", "").strip()
    contact_person = request.form.get("contact_person", "").strip()
    notes = request.form.get("notes", "").strip()
    if not name:
        flash("Название предприятия не может быть пустым.")
        return redirect(url_for("admin_letter_sources"))
    if LetterSource.query.filter(func.lower(LetterSource.name) == name.lower()).first():
        flash("Такое предприятие уже существует.")
        return redirect(url_for("admin_letter_sources"))

    db.session.add(
        LetterSource(
            name=name,
            bin_code=bin_code or None,
            address=address or None,
            phone=phone or None,
            email=email or None,
            contact_person=contact_person or None,
            notes=notes or None,
        )
    )
    db.session.commit()
    flash(f"Предприятие «{name}» добавлено.")
    return redirect(url_for("admin_letter_sources"))


@app.post("/letter-sources/quick-add")
@login_required
def quick_add_letter_source():
    name = (request.form.get("name") or "").strip()
    if not name:
        return jsonify(success=False, error="Введите название предприятия."), 400

    existing = LetterSource.query.filter(func.lower(LetterSource.name) == name.lower()).first()
    if existing:
        return jsonify(success=True, id=existing.id, name=existing.name, existed=True)

    source = LetterSource(name=name)
    db.session.add(source)
    db.session.commit()
    return jsonify(success=True, id=source.id, name=source.name, existed=False)


@app.post("/letters/<int:letter_id>/update-reg-number")
@role_required("admin", "superadmin")
def update_letter_reg_number(letter_id):
    letter = Letter.query.get_or_404(letter_id)
    reg_number = (request.form.get("reg_number") or "").strip()
    if not reg_number:
        flash("Номер не может быть пустым.", "warning")
        return redirect(request.referrer or url_for("letters_page"))

    is_memo = letter.source_id == ensure_memo_source()
    if not validate_reg_number_uniqueness(reg_number, is_memo=is_memo, exclude_letter_id=letter.id):
        flash("Документ с таким номером уже существует.", "warning")
        return redirect(request.referrer or url_for("letters_page"))

    letter.reg_number = reg_number
    db.session.commit()
    flash("Номер документа обновлён.", "success")

    if is_memo:
        return redirect(url_for("memo_view", memo_id=letter.id))
    return redirect(url_for("letter_view", letter_id=letter.id))

@app.post("/admin/letter-sources/<int:source_id>/edit")
@role_required("admin", "superadmin")
def edit_letter_source(source_id):
    source = LetterSource.query.get_or_404(source_id)
    name = request.form.get("name", "").strip()
    bin_code = request.form.get("bin_code", "").strip()
    address = request.form.get("address", "").strip()
    phone = request.form.get("phone", "").strip()
    email = request.form.get("email", "").strip()
    contact_person = request.form.get("contact_person", "").strip()
    notes = request.form.get("notes", "").strip()
    if not name:
        flash("Название предприятия не может быть пустым.")
        return redirect(url_for("admin_letter_sources"))

    duplicate = LetterSource.query.filter(
        func.lower(LetterSource.name) == name.lower(),
        LetterSource.id != source.id,
    ).first()
    if duplicate:
        flash("Такое предприятие уже существует.")
        return redirect(url_for("admin_letter_sources"))

    source.name = name
    source.bin_code = bin_code or None
    source.address = address or None
    source.phone = phone or None
    source.email = email or None
    source.contact_person = contact_person or None
    source.notes = notes or None
    db.session.commit()
    flash("Изменения сохранены.")
    return redirect(url_for("admin_letter_sources"))

@app.post("/admin/letter-sources/<int:source_id>/delete")
@role_required("admin", "superadmin")
def delete_letter_source(source_id):
    source = LetterSource.query.get_or_404(source_id)
    Letter.query.filter_by(source_id=source.id).update({"source_id": None})
    db.session.delete(source)
    db.session.commit()
    flash(f"Предприятие «{source.name}» удалено.")
    return redirect(url_for("admin_letter_sources"))

@app.get("/activities/<int:activity_id>/create-letter")
@login_required
def create_letter_from_activity(activity_id: int):
    a = Activity.query.get_or_404(activity_id)

    # создаём письмо
    subject = request.form.get("subject", f"Результаты активности №{a.id}: {a.title}")
    body = request.form.get("body", a.description or "")

    # дата письма берётся из формы или текущая
    letter_date_str = request.form.get("letter_date", "")
    if letter_date_str:
        try:
            letter_date = datetime.strptime(letter_date_str, "%Y-%m-%d").date()
        except ValueError:
            letter_date = datetime.now().date()
    else:
        letter_date = datetime.now().date()

    le = Letter(
        author_id=current_user.id,
        subject=subject,
        body=body,
        letter_date=letter_date
    )

    # ✅ сохраняем сам объект письма
    db.session.add(le)
    db.session.flush()
    print("📨 DEBUG: добавляем письмо к активности" )
    # ✅ устанавливаем связь письмо ↔ активность
    if hasattr(le, "activities"):
        le.activities.append(a)
    elif hasattr(a, "letters"):
        a.letters.append(le)
    print("📨 DEBUG: добавляем письмо к активности", a.id)
    print("📨 DEBUG: до добавления:", [x.id for x in a.letters])
    le.activities.append(a)
    print("📨 DEBUG: после добавления:", [x.id for x in le.activities])
    db.session.commit()
    print("📨 DEBUG: commit успешен")

    db.session.commit()

    flash(f"Создано письмо №{le.id} по активности №{a.id}")
    return redirect(url_for("activity_view", activity_id=a.id))


# =========================
# Тех. хелпер
# =========================
@app.errorhandler(403)
def _403(e): return render_template("error.html", title="Доступ запрещён", code=403), 403

@app.errorhandler(404)
def _404(e): return render_template("error.html", title="Не найдено", code=404), 404

# =========================
# Запуск
# =========================
# --- Регистрация ---
@app.get("/register")
def register_page():
    deps = Department.query.order_by(Department.name).all()
    return render_template("register.html", title="Регистрация", deps=deps)

@app.post("/register")
def register_form():
    data = request.form
    full_name = data.get("full_name")
    username = data.get("username")
    password = data.get("password")
    dep_id = data.get("department_id")

    if User.query.filter_by(username=username).first():
        flash("Пользователь с таким логином уже существует!")
        return redirect(url_for("register_page"))

    u = User(full_name=full_name, username=username,
             department_id=dep_id, role="worker", is_approved=False)
    u.set_password(password)
    db.session.add(u)
    db.session.commit()
    flash("Регистрация отправлена на подтверждение администратору.")
    return redirect(url_for("login_page"))

# --- Вход ---
from flask import request, redirect, url_for, flash
from flask_login import login_user, UserMixin
from config import SUPERADMIN_LOGIN, SUPERADMIN_PASSWORD
from sqlalchemy import text, func

@app.post("/login")
def login_form():
    username = (request.form.get("username") or "").strip()
    password = request.form.get("password") or ""

    print("=== DEBUG LOGIN ===")
    print("INPUT username:", repr(username))
    print("INPUT password:", repr(password))

    if username == SUPERADMIN_LOGIN and password == SUPERADMIN_PASSWORD:
        super_user =  SuperAdmin()
        login_user(super_user, remember=True)
        flash("🔐 Вход выполнен как суперадминистратор", "success")
        return redirect(url_for("admin_users"))
    # --- обычная проверка пользователей ---
    try:
        user = User.query.filter(func.trim(User.username) == username).first()
        print("filtered(trim) user:", user)
    except Exception as e:
        print("DB query error:", e)
        user = None

    if not user:
        flash("Пользователь не найден.", "warning")
        print("⚠️ user not found in DB")
        return redirect(url_for("login_page"))

    if not getattr(user, "is_approved", True):
        flash("Ваш аккаунт ожидает подтверждения администратора.", "info")
        return redirect(url_for("login_page"))

    # --- Проверка пароля (шифрованного) ---
    try:
        ok = user.check_password(password)
    except Exception as e:
        print("check_password exception:", e)
        ok = False
    print("check_password returned:", ok)

    if ok:
        login_user(user)
        flash("Вход успешен", "success")
        return redirect(url_for("dashboard"))
    else:
        flash("Неверный пароль.", "danger")
        return redirect(url_for("login_page"))

# --- Админка: список пользователей ---
@app.get("/admin/users")
@role_required("admin","superadmin")
def admin_users():
    users = User.query.order_by(User.created_at.desc()).all()
    deps = Department.query.order_by(Department.name).all()
    return render_template("admin_users.html", title="Пользователи", users=users, deps=deps)

@app.post("/admin/users")
@role_required("admin", "superadmin")
def admin_add_user():
    full_name = (request.form.get("full_name") or "").strip()
    username = (request.form.get("username") or "").strip()
    password = request.form.get("password") or ""
    dep_id = (request.form.get("department_id") or "").strip()
    role = (request.form.get("role") or "worker").strip()

    if not full_name or not username or not password:
        flash("Заполните ФИО, логин и пароль.", "warning")
        return redirect(url_for("admin_users"))

    if User.query.filter(func.trim(User.username) == username).first():
        flash("Пользователь с таким логином уже существует.", "warning")
        return redirect(url_for("admin_users"))

    allowed_roles = {"worker", "deputy", "manager"}
    if current_user.role == "superadmin":
        allowed_roles.add("admin")
    if role not in allowed_roles:
        role = "worker"

    department_id = None
    if dep_id:
        try:
            department_id = int(dep_id)
        except ValueError:
            flash("Некорректный отдел.", "warning")
            return redirect(url_for("admin_users"))

    u = User(
        full_name=full_name,
        username=username,
        department_id=department_id,
        role=role,
        is_approved=True,
    )
    u.set_password(password)
    db.session.add(u)
    db.session.commit()
    flash(f"Пользователь {u.full_name} добавлен.", "success")
    return redirect(url_for("admin_users"))


@app.route("/admin/users/<int:user_id>/edit", methods=["GET", "POST"])
@role_required("admin", "superadmin")
def edit_user(user_id):
    """Изменение учетной записи без необходимости пересоздавать пользователя."""
    user = User.query.get_or_404(user_id)
    if user.role in ("admin", "superadmin") and current_user.role != "superadmin":
        abort(403)

    departments = Department.query.order_by(Department.name).all()
    allowed_roles = {"worker", "deputy", "manager"}
    if current_user.role == "superadmin":
        allowed_roles.add("admin")

    if request.method == "POST":
        full_name = (request.form.get("full_name") or "").strip()
        username = (request.form.get("username") or "").strip()
        password = request.form.get("password") or ""
        role = (request.form.get("role") or "worker").strip()
        department_raw = (request.form.get("department_id") or "").strip()

        if not full_name or not username:
            flash("Заполните ФИО и логин.", "warning")
            return redirect(url_for("edit_user", user_id=user.id))
        if role not in allowed_roles:
            flash("Недопустимая роль.", "warning")
            return redirect(url_for("edit_user", user_id=user.id))
        duplicate = User.query.filter(
            func.trim(User.username) == username,
            User.id != user.id,
        ).first()
        if duplicate:
            flash("Пользователь с таким логином уже существует.", "warning")
            return redirect(url_for("edit_user", user_id=user.id))

        department_id = None
        if department_raw:
            try:
                department_id = int(department_raw)
            except ValueError:
                flash("Некорректный отдел.", "warning")
                return redirect(url_for("edit_user", user_id=user.id))
            if not db.session.get(Department, department_id):
                flash("Выбранный отдел не найден.", "warning")
                return redirect(url_for("edit_user", user_id=user.id))

        user.full_name = full_name
        user.username = username
        user.department_id = department_id
        user.role = role
        user.is_approved = request.form.get("is_approved") == "1"
        if password:
            user.set_password(password)
        db.session.commit()
        flash(f"Данные пользователя {user.full_name} сохранены.", "success")
        return redirect(url_for("admin_users"))

    return render_template(
        "admin_user_edit.html",
        title="Редактирование пользователя",
        user=user,
        departments=departments,
        allowed_roles=allowed_roles,
    )

@app.post("/admin/users/<int:user_id>/approve")
@role_required("admin", "superadmin")
def approve_user(user_id):
    u = User.query.get_or_404(user_id)
    u.is_approved = True
    db.session.commit()
    flash(f"Пользователь {u.full_name} подтверждён.")
    return redirect(url_for("admin_users"))

@app.post("/admin/users/<int:user_id>/delete")
@role_required("admin", "superadmin")
def delete_user(user_id):
    u = User.query.get_or_404(user_id)
    if getattr(current_user, "id", None) == u.id:
        flash("Нельзя удалить самого себя.", "warning")
        return redirect(url_for("admin_users"))

    if u.role in ("admin", "superadmin") and current_user.role != "superadmin":
        flash("Удалять администраторов может только суперадминистратор.", "warning")
        return redirect(url_for("admin_users"))

    replacement_user = (
        User.query
        .filter(User.id != u.id)
        .order_by(User.is_approved.desc(), User.created_at.asc())
        .first()
    )
    replacement_user_id = replacement_user.id if replacement_user else None

    owned_records_count = (
        Activity.query.filter_by(owner_id=u.id).count()
        + Letter.query.filter_by(author_id=u.id).count()
        + Plan.query.filter_by(created_by=u.id).count()
        + Task.query.filter_by(created_by=u.id).count()
        + GeneralDocument.query.filter_by(uploaded_by=u.id).count()
        + SharedLink.query.filter_by(created_by=u.id).count()
    )
    if owned_records_count and replacement_user_id is None:
        flash("Нельзя удалить пользователя: некому передать связанные записи.", "danger")
        return redirect(url_for("admin_users"))

    try:
        Activity.query.filter_by(owner_id=u.id).update({"owner_id": replacement_user_id}, synchronize_session=False)
        Activity.query.filter_by(approved_by_id=u.id).update({"approved_by_id": None}, synchronize_session=False)
        Letter.query.filter_by(author_id=u.id).update({"author_id": replacement_user_id}, synchronize_session=False)
        Plan.query.filter_by(created_by=u.id).update({"created_by": replacement_user_id}, synchronize_session=False)
        Task.query.filter_by(created_by=u.id).update({"created_by": replacement_user_id}, synchronize_session=False)
        GeneralDocument.query.filter_by(uploaded_by=u.id).update({"uploaded_by": replacement_user_id}, synchronize_session=False)
        SharedLink.query.filter_by(created_by=u.id).update({"created_by": replacement_user_id}, synchronize_session=False)

        LetterRecipient.query.filter_by(user_id=u.id).delete(synchronize_session=False)
        PlanExecutor.query.filter_by(user_id=u.id).delete(synchronize_session=False)
        TaskAssignee.query.filter_by(user_id=u.id).delete(synchronize_session=False)
        DownloadLog.query.filter_by(user_id=u.id).delete(synchronize_session=False)
        DocumentAccess.query.filter(
            (DocumentAccess.user_id == u.id) | (DocumentAccess.granted_by == u.id)
        ).delete(synchronize_session=False)
        ChatMessage.query.filter(
            (ChatMessage.sender_id == u.id) | (ChatMessage.receiver_id == u.id)
        ).delete(synchronize_session=False)

        db.session.delete(u)
        db.session.commit()
    except Exception as exc:
        db.session.rollback()
        flash(f"Не удалось удалить пользователя: {exc}", "danger")
        return redirect(url_for("admin_users"))
    flash("Пользователь удалён.")
    return redirect(url_for("admin_users"))
# --- Админка: отделы ---
@app.get("/admin/departments")
@role_required("admin")
def admin_departments():
    deps = Department.query.order_by(Department.name).all()
    return render_template("admin_departments.html", title="Отделы", deps=deps)

@app.post("/admin/departments")
@role_required("admin")
def add_department():
    name = request.form.get("name", "").strip()
    if not name:
        flash("Название отдела не может быть пустым.")
        return redirect(url_for("admin_departments"))
    if Department.query.filter_by(name=name).first():
        flash("Такой отдел уже существует.")
        return redirect(url_for("admin_departments"))

    d = Department(name=name)
    db.session.add(d)
    db.session.commit()
    flash(f"Отдел «{name}» добавлен.")
    return redirect(url_for("admin_departments"))

@app.post("/admin/departments/<int:dep_id>/edit")
@role_required("admin")
def edit_department(dep_id):
    d = Department.query.get_or_404(dep_id)
    name = request.form.get("name", "").strip()
    if not name:
        flash("Название не может быть пустым.")
        return redirect(url_for("admin_departments"))
    d.name = name
    db.session.commit()
    flash("Изменения сохранены.")
    return redirect(url_for("admin_departments"))

@app.post("/admin/departments/<int:dep_id>/delete")
@role_required("admin")
def delete_department(dep_id):
    d = Department.query.get_or_404(dep_id)
    db.session.delete(d)
    db.session.commit()
    flash(f"Отдел «{d.name}» удалён.")
    return redirect(url_for("admin_departments"))
# --- Админка: типы активностей ---
@app.get("/admin/activity-types")
@role_required("admin")
def admin_activity_types():
    types = ActivityType.query.order_by(ActivityType.name).all()
    return render_template("admin_activity_types.html", title="Типы активностей", types=types)

@app.post("/admin/activity-types")
@role_required("admin")
def add_activity_type():
    name = request.form.get("name", "").strip()
    if not name:
        flash("Название типа активности не может быть пустым.")
        return redirect(url_for("admin_activity_types"))
    if ActivityType.query.filter_by(name=name).first():
        flash("Такой тип уже существует.")
        return redirect(url_for("admin_activity_types"))

    t = ActivityType(name=name)
    db.session.add(t)
    db.session.commit()
    flash(f"Тип активности «{name}» добавлен.")
    return redirect(url_for("admin_activity_types"))

@app.post("/admin/activity-types/<int:type_id>/edit")
@role_required("admin")
def edit_activity_type(type_id):
    t = ActivityType.query.get_or_404(type_id)
    name = request.form.get("name", "").strip()
    if not name:
        flash("Название не может быть пустым.")
        return redirect(url_for("admin_activity_types"))
    t.name = name
    db.session.commit()
    flash("Изменения сохранены.")
    return redirect(url_for("admin_activity_types"))

@app.post("/admin/activity-types/<int:type_id>/delete")
@role_required("admin")
def delete_activity_type(type_id):
    t = ActivityType.query.get_or_404(type_id)
    db.session.delete(t)
    db.session.commit()
    flash(f"Тип активности «{t.name}» удалён.")
    return redirect(url_for("admin_activity_types"))
@app.get("/activities/<int:activity_id>/create-letter")
@login_required

def new_letter_from_activity(activity_id):
    activity = Activity.query.get_or_404(activity_id)
    deps = Department.query.order_by(Department.name).all()
    users = User.query.filter_by(is_approved=True).order_by(User.full_name).all()
    sources = LetterSource.query.order_by(LetterSource.name).all()

    # создаём черновик письма с автозаполнением
    subject = f"Письмо по активности №{activity.id} — {activity.title}"
    body = f"Описание активности:\n{activity.description or ''}\n\nСтатус: {activity.status}"

    return render_template(
        "letter_create.html",
        title="Письмо на основе активности",
        deps=deps,
        users=users,
        sources=sources,
        subject=subject,
        body=body,
        from_activity=activity,
        now=datetime.now  # ← вот это добавляем
    )
# ✅ AJAX: переключение статуса "выполнено"
@app.post("/logs/<int:log_id>/toggle_done")
@login_required
def toggle_log_done(log_id: int):
    log = ActivityLog.query.get_or_404(log_id)
    if activity_work_is_locked(log.activity) or not can_modify_activity_work(log.activity):
        abort(403)
    log.is_done = not log.is_done
    add_activity_history(
        log.activity,
        "Журнал действий",
        "{} выполнение действия: {}".format("Отмечено" if log.is_done else "Снята отметка", log.text[:300]),
    )
    db.session.commit()
    return {"success": True, "is_done": log.is_done}


# ✅ AJAX: добавление поддействия
@app.post("/logs/<int:log_id>/add_child")
@login_required
def add_child_log(log_id: int):
    parent = ActivityLog.query.get_or_404(log_id)
    if activity_work_is_locked(parent.activity) or not can_modify_activity_work(parent.activity):
        abort(403)
    text = request.form.get("text", "").strip()
    if not text:
        return {"success": False, "error": "Пустой текст"}, 400

    child = ActivityLog(activity_id=parent.activity_id, parent_id=log_id, text=text)
    db.session.add(child)
    add_activity_history(parent.activity, "Журнал действий", "Добавлено поддействие: {}".format(text[:300]))
    db.session.commit()

    return {
        "success": True,
        "child": {
            "id": child.id,
            "text": child.text,
            "is_done": child.is_done,
            "entry_date": child.entry_date.strftime("%d.%m.%Y %H:%M")
        }
    }
# ✅ AJAX: удаление действия (и всех поддействий)
@app.post("/logs/<int:log_id>/delete")
@login_required
def delete_log(log_id: int):
    log = ActivityLog.query.get_or_404(log_id)
    if activity_work_is_locked(log.activity) or not can_modify_activity_work(log.activity):
        abort(403)

    # Рекурсивное удаление всех поддействий
    def delete_children(node):
        for child in node.children:
            delete_children(child)
            db.session.delete(child)

    deleted_text = log.text
    delete_children(log)
    db.session.delete(log)
    add_activity_history(log.activity, "Журнал действий", "Удалено действие: {}".format(deleted_text[:300]))
    db.session.commit()

    return jsonify(success=True)
@app.post("/logs/<int:log_id>/toggle_done")
@login_required
def toggle_done(log_id: int):
    log = ActivityLog.query.get_or_404(log_id)
    if activity_work_is_locked(log.activity) or not can_modify_activity_work(log.activity):
        abort(403)
    log.is_done = not log.is_done
    add_activity_history(
        log.activity,
        "Журнал действий",
        "{} выполнение действия: {}".format("Отмечено" if log.is_done else "Снята отметка", log.text[:300]),
    )
    db.session.commit()
    return jsonify(success=True, is_done=log.is_done)
@app.post("/logs/<int:log_id>/edit")
@login_required
def edit_log(log_id: int):
    log = ActivityLog.query.get_or_404(log_id)
    if activity_work_is_locked(log.activity) or not can_modify_activity_work(log.activity):
        abort(403)
    text = request.form.get("text", "").strip()
    if not text:
        return jsonify(success=False, error="Пустой текст"), 400
    previous_text = log.text
    log.text = text
    add_activity_history(
        log.activity,
        "Журнал действий",
        "Изменено действие: «{}» -> «{}»".format(previous_text[:150], text[:150]),
    )
    db.session.commit()
    return jsonify(success=True)


@app.route('/reports', methods=['GET', 'POST'])
@role_required("admin")
def manage_reports():
    if request.method == 'POST':
        rid = request.form.get('id')
        name = request.form.get('name')
        sql_text = request.form.get('sql')
        json_data = request.form.get('json')
        js_code = request.form.get('js')

        if rid:
            db.session.execute(
                text("UPDATE otchets SET name=:n, sql=:s, json=:j, js=:k WHERE id=:id"),
                {"n": name, "s": sql_text, "j": json_data, "k": js_code, "id": rid},
            )
        else:
            db.session.execute(
                text("INSERT INTO otchets (name, sql, json, js) VALUES (:n, :s, :j, :k)"),
                {"n": name, "s": sql_text, "j": json_data, "k": js_code},
            )
        db.session.commit()
        flash('Отчёт сохранён', 'success')
        return redirect(url_for('manage_reports'))

    rows = db.session.execute(text("SELECT * FROM otchets ORDER BY id DESC")).fetchall()
    return render_template('otchet.html', rows=rows)

@app.context_processor
def inject_now():
    from datetime import datetime
    return {'now': datetime.now}


@app.context_processor
def inject_template_helpers():
    return {
        "has_endpoint": lambda endpoint: endpoint in app.view_functions,
    }
@app.get("/admin/values")
@login_required
def admin_values_list():
    if current_user.role != "admin":
        abort(403)
    values = ValueTemplate.query.order_by(ValueTemplate.rusname).all()
    return render_template("admin_values.html", title="Шаблоны значений", values=values)

@app.post("/admin/values")
@login_required
def admin_values_create():
    if current_user.role != "admin":
        abort(403)
    name = request.form.get("name","").strip()
    rusname = request.form.get("rusname","").strip()
    vtype = request.form.get("type","string").strip()
    if not name or not rusname:
        flash("Имя и заголовок обязательны","warning")
        return redirect(url_for("admin_values_list"))
    vt = ValueTemplate(name=name, rusname=rusname, type=vtype, is_active=True)
    db.session.add(vt)
    db.session.commit()
    flash("Шаблон добавлен","success")
    return redirect(url_for("admin_values_list"))

@app.post("/admin/values/<int:vid>/toggle")
@login_required
def admin_values_toggle(vid):
    if current_user.role != "admin":
        abort(403)
    vt = ValueTemplate.query.get_or_404(vid)
    vt.is_active = not vt.is_active
    db.session.commit()
    return redirect(url_for("admin_values_list"))

@app.post("/admin/values/<int:vid>/delete")
@login_required
def admin_values_delete(vid):
    if current_user.role != "admin":
        abort(403)
    vt = ValueTemplate.query.get_or_404(vid)
    db.session.delete(vt)
    db.session.commit()
    flash("Шаблон удалён","info")
    return redirect(url_for("admin_values_list"))

# ====== API: получить активные шаблоны в JSON (для JS на форме активности) ======



# ====== ДОРАБОТКА СОЗДАНИЯ АКТИВНОСТИ: читаем direction/decision/extra_values ======

@app.get("/api/value-templates")
@login_required
def api_value_templates():
    templates = ValueTemplate.query.filter_by(is_active=True).all()
    data = [
        {"id": v.id, "name": v.name, "rusname": v.rusname, "type": v.type}
        for v in templates
    ]
    return jsonify(data)
# -------------------------------
#  Управление шаблонами значений
# -------------------------------
@app.get("/admin/value-templates")
@login_required
def value_templates_page():
    if current_user.role != "admin":
        abort(403)
    values = ValueTemplate.query.order_by(ValueTemplate.rusname).all()
    return render_template("value_templates.html", values=values, title="Шаблоны значений")


@app.post("/admin/value-templates/add")
@login_required
def add_value_template():
    if current_user.role != "admin":
        abort(403)
    name = request.form.get("name")
    rusname = request.form.get("rusname")
    type_ = request.form.get("type")
    if not name or not rusname:
        flash("Имя и описание обязательны", "warning")
        return redirect(url_for("value_templates_page"))
    vt = ValueTemplate(name=name.strip(), rusname=rusname.strip(), type=type_)
    db.session.add(vt)
    db.session.commit()
    flash(f"Добавлен шаблон: {rusname}", "success")
    return redirect(url_for("value_templates_page"))


@app.post("/admin/value-templates/<int:vt_id>/toggle")
@login_required
def toggle_value_template(vt_id):
    if current_user.role != "admin":
        abort(403)
    vt = ValueTemplate.query.get_or_404(vt_id)
    vt.is_active = not vt.is_active
    db.session.commit()
    flash(f"Шаблон {'включён' if vt.is_active else 'отключён'}", "info")
    return redirect(url_for("value_templates_page"))


@app.post("/admin/value-templates/<int:vt_id>/delete")
@login_required
def delete_value_template(vt_id):
    if current_user.role != "admin":
        abort(403)
    vt = ValueTemplate.query.get_or_404(vt_id)
    db.session.delete(vt)
    db.session.commit()
    flash("Шаблон удалён", "danger")
    return redirect(url_for("value_templates_page"))
@app.post("/activities/<int:activity_id>/add_value")
@login_required
def add_activity_value(activity_id):
    a = Activity.query.get_or_404(activity_id)
    if activity_work_is_locked(a) or not can_modify_activity_work(a):
        abort(403)
    tpl_id = request.form.get("template_id")
    val = request.form.get("value")

    tpl = ValueTemplate.query.get_or_404(tpl_id)

    # --- если extra_values хранится как строка, преобразуем ---
    if not a.extra_values:
        values = []
    else:
        try:
            if isinstance(a.extra_values, str):
                values = json.loads(a.extra_values)
            else:
                values = a.extra_values
        except Exception:
            values = []

    # --- добавляем новое значение ---
    values.append({
        "id": tpl.id,
        "name": tpl.name,
        "rusname": tpl.rusname,
        "type": tpl.type,
        "value": val
    })

    # --- сохраняем список обратно как JSON ---
    a.extra_values = json.dumps(values, ensure_ascii=False)
    add_activity_history(a, "Дополнительное значение", "Добавлено значение: {}.".format(tpl.rusname))
    db.session.commit()

    return jsonify(success=True)

@app.post("/activities/<int:activity_id>/delete_value/<int:value_id>")
@login_required
def delete_activity_value(activity_id, value_id):
    a = Activity.query.get_or_404(activity_id)
    if activity_work_is_locked(a) or not can_modify_activity_work(a):
        abort(403)
    if isinstance(a.extra_values, str):
        a.extra_values = json.loads(a.extra_values)
    deleted = next((value for value in a.extra_values if int(value.get("id", 0)) == value_id), None)
    a.extra_values = [v for v in a.extra_values if int(v.get("id", 0)) != value_id]
    if deleted:
        add_activity_history(
            a,
            "Дополнительное значение",
            "Удалено значение: {}.".format(deleted.get("rusname") or deleted.get("name") or value_id),
        )
    db.session.commit()
    return jsonify(success=True)
@app.post("/letters/<int:letter_id>/save_extras")
@login_required
def save_letter_extras(letter_id):
    import json
    le = Letter.query.get_or_404(letter_id)
    data = request.get_json()
    new_extras = data.get("extras", [])

    # Загружаем старые значения, если есть
    existing = []
    if le.extra_values:
        try:
            existing = json.loads(le.extra_values)
        except Exception as e:
            print("⚠️ Ошибка парсинга JSON:", e)

    # Добавляем новые к старым
    existing.extend(new_extras)

    # Сохраняем обратно
    le.extra_values = existing
    db.session.commit()

    print(f"✅ Обновлённые extras для письма #{le.id}:", le.extra_values)
    return jsonify({"success": True, "count": len(existing)})
# =========================
# Планы (создание, исполнители, активности)
# =========================
PLAN_ITEM_STATUSES = {
    "planned": {"label": "Запланировано", "class": "bg-secondary"},
    "done": {"label": "Выполнено", "class": "bg-success"},
    "not_done": {"label": "Невыполнено", "class": "bg-danger"},
    "cancelled": {"label": "Отменено", "class": "bg-dark"},
}
PLAN_DEADLINES = {
    "month": "В течение месяца",
    "q1": "1 квартал",
    "q2": "2 квартал",
    "q3": "3 квартал",
    "q4": "4 квартал",
    "date": "До указанной даты",
}


def get_plan_period(month_value):
    from calendar import monthrange
    from datetime import date, datetime

    if month_value:
        try:
            period_start = datetime.strptime(month_value, "%Y-%m").date().replace(day=1)
        except ValueError:
            return None
    else:
        period_start = date.today().replace(day=1)
    period_end = period_start.replace(day=monthrange(period_start.year, period_start.month)[1])
    return period_start, period_end


def plan_item_status_meta(status):
    return PLAN_ITEM_STATUSES.get(status, {"label": status or "-", "class": "bg-secondary"})


def parse_plan_deadline_date(deadline_kind, raw_date):
    if deadline_kind != "date":
        return None
    try:
        return dt.datetime.strptime((raw_date or "").strip(), "%Y-%m-%d").date()
    except ValueError:
        return None


def build_plan_department_rows(items):
    """Return one consistent aggregation for department and employee reports."""
    departments = {}
    employees = {}
    for item in items:
        user = item.executor
        department = user.department if user else None
        department_id = department.id if department else 0
        department_name = department.name if department else "Без отдела"
        department_row = departments.setdefault(department_id, {
            "name": department_name,
            "total": 0, "planned": 0, "done": 0, "not_done": 0, "cancelled": 0,
            "executor_map": {},
        })
        employee_row = department_row["executor_map"].setdefault(item.executor_id, {
            "user": user,
            "items": [],
            "total": 0, "planned": 0, "done": 0, "not_done": 0, "cancelled": 0,
        })
        for row in (department_row, employee_row):
            row["total"] += 1
            row[item.status] = row.get(item.status, 0) + 1
        employee_row["items"].append(item)
        employees[item.executor_id] = employee_row

    department_rows = list(departments.values())
    for row in department_rows:
        row["completion"] = round((row["done"] / row["total"]) * 100, 1) if row["total"] else 0
        row["executors"] = list(row.pop("executor_map").values())
        for employee in row["executors"]:
            employee["completion"] = round((employee["done"] / employee["total"]) * 100, 1) if employee["total"] else 0
        row["executors"].sort(key=lambda employee: (employee["user"].full_name or "").lower())
    department_rows.sort(key=lambda row: (row["name"] or "").lower())
    executor_rows = sorted(employees.values(), key=lambda row: (row["user"].full_name or "").lower())
    return department_rows, executor_rows


@app.route("/plans", methods=["GET", "POST"])
@login_required
def plans_list():
    from datetime import datetime, date

    # --- фильтрация по дате ---
    start_str = request.args.get("date_from")
    end_str = request.args.get("date_to")

    q = Plan.query

    if start_str:
        try:
            date_from = datetime.strptime(start_str, "%Y-%m-%d").date()
            q = q.filter(Plan.start_date >= date_from)
        except ValueError:
            pass

    if end_str:
        try:
            date_to = datetime.strptime(end_str, "%Y-%m-%d").date()
            q = q.filter(Plan.end_date <= date_to)
        except ValueError:
            pass

    is_manager = current_user.role in ("admin", "superadmin")
    if not is_manager:
        q = q.filter(Plan.items.any(PlanItem.executor_id == current_user.id))

    plans = q.order_by(Plan.start_date.desc(), Plan.id.desc()).all()
    plan_summary = {}
    for plan in plans:
        item_counts = {status: 0 for status in PLAN_ITEM_STATUSES}
        for item in plan.items:
            item_counts[item.status] = item_counts.get(item.status, 0) + 1
        plan_summary[plan.id] = item_counts
    return render_template(
        "plans.html", title="Планы", plans=plans, date_from=start_str, date_to=end_str,
        plan_summary=plan_summary, is_manager=is_manager,
    )


@app.route("/plans/create", methods=["GET", "POST"])
@login_required
def plan_create():
    if request.method == "POST":
        period = get_plan_period((request.form.get("month") or "").strip())
        if not period:
            flash("Выберите месяц и год плана.", "warning")
            return redirect(url_for("plan_create"))
        start_date, end_date = period
        plan_text = (request.form.get("plan_text") or "").strip() or f"План на {start_date.strftime('%m.%Y')}"
        existing_plan = Plan.query.filter_by(created_by=current_user.id, start_date=start_date).first()
        if existing_plan:
            flash("На выбранный месяц у вас уже создан план. Откройте его для редактирования.", "warning")
            return redirect(url_for("plan_detail", plan_id=existing_plan.id))
        executor_ids = request.form.getlist("executor_id")
        tasks = request.form.getlist("task_text")
        deadlines = request.form.getlist("deadline_kind")
        deadline_dates = request.form.getlist("deadline_date")

        rows = []
        for index, task in enumerate(tasks):
            task = (task or "").strip()
            executor_id = executor_ids[index] if index < len(executor_ids) else ""
            deadline = deadlines[index] if index < len(deadlines) else "month"
            deadline_date = parse_plan_deadline_date(
                deadline, deadline_dates[index] if index < len(deadline_dates) else ""
            )
            if not task and not executor_id:
                continue
            if not task or not executor_id or deadline not in PLAN_DEADLINES or (deadline == "date" and not deadline_date):
                flash("В каждой строке укажите исполнителя, мероприятие и срок.", "warning")
                return redirect(url_for("plan_create"))
            rows.append((int(executor_id), task, deadline, deadline_date))

        if not rows:
            flash("Добавьте хотя бы один пункт плана.", "warning")
            return redirect(url_for("plan_create"))

        is_manager = current_user.role in ("admin", "superadmin")
        if not is_manager and any(executor_id != current_user.id for executor_id, _, _, _ in rows):
            abort(403)
        p = Plan(
            start_date=start_date,
            end_date=end_date,
            plan_text=plan_text,
            created_by=current_user.id,
            approval_status="approved" if is_manager else "submitted",
            approved_by_id=current_user.id if is_manager else None,
            approved_at=dt.datetime.utcnow() if is_manager else None,
        )
        db.session.add(p)
        db.session.flush()
        for position, (executor_id, task, deadline, deadline_date) in enumerate(rows, start=1):
            db.session.add(PlanItem(
                plan_id=p.id, executor_id=executor_id, position=position,
                task_text=task, deadline_kind=deadline, deadline_date=deadline_date, status="planned",
            ))
        db.session.commit()
        flash("План и его пункты созданы", "success")
        return redirect(url_for("plan_detail", plan_id=p.id))

    is_manager = current_user.role in ("admin", "superadmin")
    users = User.query.filter_by(is_approved=True).order_by(User.full_name).all() if is_manager else [current_user]
    return render_template(
        "plan_create.html", title="Создание плана", users=users,
        deadline_options=PLAN_DEADLINES, is_manager=is_manager,
    )


@app.route("/plans/<int:plan_id>")
@login_required
def plan_detail(plan_id):
    plan = Plan.query.get_or_404(plan_id)
    is_manager = current_user.role in ("admin", "superadmin")
    own_items = [item for item in plan.items if item.executor_id == current_user.id]
    is_creator = plan.created_by == current_user.id
    if not is_manager and not own_items and not is_creator:
        abort(403)
    items = plan.items if is_manager else own_items
    users = User.query.filter_by(is_approved=True).order_by(User.full_name).all() if is_manager else []
    task_owner_ids = {item.executor_id for item in items}
    available_tasks = (
        Activity.query.options(joinedload(Activity.owner), joinedload(Activity.type))
        .filter(Activity.archived_at.is_(None), Activity.owner_id.in_(task_owner_ids or {-999999}))
        .order_by(Activity.owner_id, Activity.start_date.desc(), Activity.id.desc())
        .all()
    )
    return render_template(
        "plan_detail.html", title="План", plan=plan, items=items, users=users,
        is_manager=is_manager, is_creator=is_creator,
        can_edit_definition=is_manager or (is_creator and plan.approval_status != "approved"),
        deadline_options=PLAN_DEADLINES,
        status_options=PLAN_ITEM_STATUSES, get_plan_status_meta=plan_item_status_meta,
        available_tasks=available_tasks,
    )


@app.post("/plans/<int:plan_id>/items")
@login_required
def plan_item_create(plan_id):
    plan = Plan.query.get_or_404(plan_id)
    is_manager = current_user.role in ("admin", "superadmin")
    if not is_manager and (plan.created_by != current_user.id or plan.approval_status == "approved"):
        abort(403)
    task_text = (request.form.get("task_text") or "").strip()
    executor_id = request.form.get("executor_id", type=int)
    deadline_kind = request.form.get("deadline_kind", "month")
    deadline_date = parse_plan_deadline_date(deadline_kind, request.form.get("deadline_date"))
    if not task_text or not executor_id or deadline_kind not in PLAN_DEADLINES or (deadline_kind == "date" and not deadline_date):
        flash("Заполните исполнителя, мероприятие и срок.", "warning")
        return redirect(url_for("plan_detail", plan_id=plan.id))
    if not is_manager and executor_id != current_user.id:
        abort(403)
    position = (max((item.position for item in plan.items), default=0) + 1)
    db.session.add(PlanItem(
        plan_id=plan.id, executor_id=executor_id, position=position,
        task_text=task_text, deadline_kind=deadline_kind, deadline_date=deadline_date,
    ))
    db.session.commit()
    flash("Пункт плана добавлен", "success")
    return redirect(url_for("plan_detail", plan_id=plan.id))


@app.post("/plan-items/<int:item_id>/update")
@login_required
def plan_item_update(item_id):
    item = PlanItem.query.get_or_404(item_id)
    plan = item.plan
    is_manager = current_user.role in ("admin", "superadmin")
    if not is_manager and item.executor_id != current_user.id:
        abort(403)

    if item.execution_approved and not is_manager:
        flash("Выполнение уже принято администратором: пункт доступен только для просмотра.", "warning")
        return redirect(url_for("plan_detail", plan_id=item.plan_id))

    can_update_result = is_manager or plan.approval_status == "approved"
    if can_update_result:
        status = request.form.get("status", item.status)
        if status not in PLAN_ITEM_STATUSES:
            abort(400)
        item.status = status
        item.comment = (request.form.get("comment") or "").strip() or None
        linked_activity_id = request.form.get("linked_activity_id", type=int)
        if linked_activity_id:
            linked_activity = db.session.get(Activity, linked_activity_id)
            if not linked_activity or linked_activity.archived_at or linked_activity.owner_id != item.executor_id:
                flash("Можно привязать только действующую задачу исполнителя пункта плана.", "warning")
                return redirect(url_for("plan_detail", plan_id=item.plan_id))
            item.linked_activity_id = linked_activity.id
        elif request.form.get("linked_activity_id") == "":
            item.linked_activity_id = None
        if status == "done" and not item.completed_at:
            item.completed_at = dt.datetime.utcnow()
    if is_manager or (plan.created_by == current_user.id and plan.approval_status != "approved"):
        item.task_text = (request.form.get("task_text") or item.task_text).strip()
        deadline_kind = request.form.get("deadline_kind", item.deadline_kind)
        if deadline_kind in PLAN_DEADLINES:
            deadline_date = parse_plan_deadline_date(deadline_kind, request.form.get("deadline_date"))
            if deadline_kind == "date" and not deadline_date:
                flash("Для срока «До указанной даты» выберите дату.", "warning")
                return redirect(url_for("plan_detail", plan_id=item.plan_id))
            item.deadline_kind = deadline_kind
            item.deadline_date = deadline_date
        executor_id = request.form.get("executor_id", type=int)
        if executor_id and is_manager:
            item.executor_id = executor_id
    db.session.commit()
    flash("Результат по пункту плана сохранен", "success")
    return redirect(url_for("plan_detail", plan_id=item.plan_id))


@app.post("/plan-items/<int:item_id>/delete")
@login_required
def plan_item_delete(item_id):
    item = PlanItem.query.get_or_404(item_id)
    plan = item.plan
    is_manager = current_user.role in ("admin", "superadmin")
    if not is_manager and (plan.created_by != current_user.id or plan.approval_status == "approved" or item.executor_id != current_user.id):
        abort(403)
    plan_id = item.plan_id
    db.session.delete(item)
    db.session.commit()
    flash("Пункт плана удален", "success")
    return redirect(url_for("plan_detail", plan_id=plan_id))


@app.post("/plans/<int:plan_id>/delete")
@role_required("admin", "superadmin")
def plan_delete(plan_id):
    plan = Plan.query.get_or_404(plan_id)
    # Legacy activities can remain in the task register after their plan is removed.
    Activity.query.filter_by(plan_id=plan.id).update({"plan_id": None}, synchronize_session=False)
    db.session.delete(plan)
    db.session.commit()
    flash("План удален.", "success")
    return redirect(url_for("plans_list"))


@app.post("/plans/<int:plan_id>/approve")
@role_required("admin", "superadmin")
def plan_approve(plan_id):
    plan = Plan.query.get_or_404(plan_id)
    plan.approval_status = "approved"
    plan.approved_by_id = current_user.id if current_user.id != -1 else None
    plan.approved_at = dt.datetime.utcnow()
    plan.revision_comment = None
    db.session.commit()
    flash("План принят. Исполнитель теперь может указывать только результат выполнения.", "success")
    return redirect(url_for("plan_detail", plan_id=plan.id))


@app.post("/plans/<int:plan_id>/return-for-revision")
@role_required("admin", "superadmin")
def plan_return_for_revision(plan_id):
    plan = Plan.query.get_or_404(plan_id)
    comment = (request.form.get("revision_comment") or "").strip()
    if not comment:
        flash("Укажите, что требуется доработать.", "warning")
        return redirect(url_for("plan_detail", plan_id=plan.id))
    plan.approval_status = "revision"
    plan.approved_by_id = None
    plan.approved_at = None
    plan.revision_comment = comment
    db.session.commit()
    flash("План возвращен на доработку.", "warning")
    return redirect(url_for("plan_detail", plan_id=plan.id))


@app.post("/plans/<int:plan_id>/resubmit")
@login_required
def plan_resubmit(plan_id):
    plan = Plan.query.get_or_404(plan_id)
    if plan.created_by != current_user.id or plan.approval_status != "revision":
        abort(403)
    plan.approval_status = "submitted"
    db.session.commit()
    flash("План повторно направлен администратору на принятие.", "success")
    return redirect(url_for("plan_detail", plan_id=plan.id))


@app.post("/plan-items/<int:item_id>/approve-execution")
@role_required("admin", "superadmin")
def plan_item_approve_execution(item_id):
    item = PlanItem.query.get_or_404(item_id)
    if item.status == "planned":
        flash("Сначала исполнитель должен указать результат выполнения.", "warning")
        return redirect(url_for("plan_detail", plan_id=item.plan_id))
    item.execution_approved = True
    item.execution_approved_by_id = current_user.id if current_user.id != -1 else None
    item.execution_approved_at = dt.datetime.utcnow()
    db.session.commit()
    flash("Выполнение принято. Пункт переведен в режим только для чтения.", "success")
    return redirect(url_for("plan_detail", plan_id=item.plan_id))


@app.route("/plans/<int:plan_id>/executors", methods=["GET", "POST"])
@role_required("admin")
def plan_executors(plan_id):
    plan = Plan.query.get_or_404(plan_id)
    users = User.query.filter_by(is_approved=True).order_by(User.full_name).all()

    if request.method == "POST":
        user_id = request.form.get("user_id")
        topic = request.form.get("topic")
        start_date = request.form.get("start_date")
        end_date = request.form.get("end_date")

        pe = PlanExecutor(
            plan_id=plan.id,
            user_id=user_id,
            topic=topic,
            start_date=dt.datetime.strptime(start_date, "%Y-%m-%d").date(),
            end_date=dt.datetime.strptime(end_date, "%Y-%m-%d").date(),
        )
        db.session.add(pe)
        db.session.commit()

        # создаём активность для исполнителя
        a = Activity(
            title=f"Задача по плану: {topic}",
            description=plan.plan_text,
            owner_id=user_id,
            created_by_id=current_user.id if current_user.id != -1 else None,
            start_date=pe.start_date,
            end_date=pe.end_date,
            status="open",
            plan_id=plan.id,
            type_id=1001
        )
        db.session.add(a)
        db.session.commit()

        flash("Исполнитель добавлен и активность создана", "success")
        return redirect(url_for("plan_executors", plan_id=plan.id))

    return render_template("plan_executors.html", title="Исполнители", plan=plan, users=users)


@app.get("/plans/<int:plan_id>/activities")
@login_required
def plan_activities(plan_id):
    plan = Plan.query.get_or_404(plan_id)
    acts = Activity.query.filter_by(plan_id=plan.id).order_by(Activity.start_date.desc()).all()
    return render_template("plan_activities.html", title="Активности по плану", plan=plan, activities=acts)
@app.post("/plans/<int:plan_id>/executors/ajax")
@role_required("admin")
def plan_executors_add_ajax(plan_id):
    plan = Plan.query.get_or_404(plan_id)
    data = request.get_json()
    try:
        user_id = int(data.get("user_id"))
        topic = data.get("topic", "").strip()
        start_date = dt.datetime.strptime(data.get("start_date"), "%Y-%m-%d").date()
        end_date = dt.datetime.strptime(data.get("end_date"), "%Y-%m-%d").date()
    except Exception as e:
        return jsonify(success=False, error=f"Неверные данные: {e}")

    if not topic or not user_id:
        return jsonify(success=False, error="Не указаны исполнитель или тема")

    pe = PlanExecutor(
        plan_id=plan.id,
        user_id=user_id,
        topic=topic,
        start_date=start_date,
        end_date=end_date
    )
    db.session.add(pe)
    db.session.commit()

    # 🔹 создаём активность
    act = Activity(
        title=f"Задача по плану: {topic}",
        description=plan.plan_text,
        owner_id=user_id,
        created_by_id=current_user.id if current_user.id != -1 else None,
        start_date=start_date,
        end_date=end_date,
        status="open",
        plan_id=plan.id,    type_id=1001
    )
    db.session.add(act)
    db.session.commit()

    return jsonify(
        success=True,
        executor={
            "id": pe.id,
            "user_name": pe.user.full_name,
            "topic": pe.topic,
            "start_date": pe.start_date.strftime("%Y-%m-%d"),
            "end_date": pe.end_date.strftime("%Y-%m-%d"),
            "status": pe.status,
        }
    )
@app.post("/plans/<int:plan_id>/toggle_complete")
@role_required("admin")
def toggle_plan_complete(plan_id):
    plan = Plan.query.get_or_404(plan_id)
    plan.is_completed = not plan.is_completed
    db.session.commit()
    return jsonify(success=True, is_completed=plan.is_completed)
# === 📅 Единый календарь сроков ===
@app.route("/calendar")
@login_required
def deadlines_calendar():
    return render_template("plans_calendar.html", title="Календарь сроков")


@app.route("/plans/calendar")
@login_required
def plans_calendar():
    """Keep the old plans URL working while the calendar becomes a global section."""
    return redirect(url_for("deadlines_calendar"))

# === API для подгрузки сроков задач и планов ===
@app.route("/api/activities/events")
@login_required
def activities_events():
    selected_kinds = set(filter(None, (request.args.get("kinds") or "task,plan,postponed,overdue").split(",")))
    today = dt.date.today()
    privileged_roles = ("admin", "manager", "superadmin", "deputy")
    is_privileged = current_user.role in privileged_roles

    activity_query = (
        Activity.query.options(joinedload(Activity.owner))
        .filter(Activity.archived_at.is_(None))
    )
    if not is_privileged:
        department_id = getattr(current_user, "department_id", None)
        activity_query = activity_query.filter(or_(
            Activity.owner_id == current_user.id,
            Activity.assigned_department_id == department_id if department_id else False,
        ))
    events = []
    for activity in activity_query.all():
        deadline = activity.end_date.date() if activity.end_date else None
        start_date = activity.start_date.date() if activity.start_date else None
        postponed_to = activity.postponed_to.date() if activity.postponed_to else None
        is_postponed = activity.status == "postponed" and postponed_to and postponed_to >= today
        is_overdue = (
            deadline and deadline < today
            and activity.status in ("open", "in_progress", "postponed")
            and not is_postponed
        )
        if is_postponed:
            kind, event_date, color, prefix = "postponed", postponed_to, "#f28c28", "Перенос"
        elif is_overdue:
            kind, event_date, color, prefix = "overdue", deadline, "#dc3545", "Просрочено"
        else:
            kind, event_date, color, prefix = (
                "task", deadline or start_date, "#0dcaf0",
                "Задача" if deadline else "Начало задачи",
            )
        if not event_date or kind not in selected_kinds:
            continue
        events.append({
            "id": "activity-{}-{}".format(activity.id, kind),
            "title": "{}: {}".format(prefix, activity.title),
            "start": event_date.isoformat(),
            "allDay": True,
            "color": color,
            "kind": kind,
            "status": activity.status,
            "owner": activity.owner.full_name if activity.owner else "-",
            "url": url_for("activity_view", activity_id=activity.id),
        })

    if "plan" in selected_kinds:
        plan_items_query = PlanItem.query.options(joinedload(PlanItem.plan), joinedload(PlanItem.executor))
        if not is_privileged:
            plan_items_query = plan_items_query.filter(PlanItem.executor_id == current_user.id)
        for item in plan_items_query.all():
            plan = item.plan
            if not plan:
                continue
            event_date = item.deadline_date or plan.end_date
            if not event_date:
                continue
            status_meta = plan_item_status_meta(item.status)
            events.append({
                "id": "plan-item-{}".format(item.id),
                "title": "План: {}".format(item.task_text),
                "start": event_date.isoformat(),
                "allDay": True,
                "color": "#6f42c1" if item.status == "planned" else "#198754" if item.status == "done" else "#6c757d",
                "kind": "plan",
                "status": status_meta["label"],
                "owner": item.executor.full_name if item.executor else "-",
                "deadline": PLAN_DEADLINES.get(item.deadline_kind, "Срок плана"),
                "url": url_for("plan_detail", plan_id=plan.id),
            })

    return jsonify(events)
# === 🔔 API: уведомления о задачах со сроком завтра ===
@app.route("/api/notifications")
@login_required
def notifications():
    today = dt.date.today()
    tasks = Activity.query.filter(
        Activity.owner_id == current_user.id,
        Activity.status.in_(["open", "in_progress", "postponed"]),
        Activity.end_date.isnot(None),
    ).all()
    reminders = []
    for task in tasks:
        deadline = task.end_date.date() if isinstance(task.end_date, dt.datetime) else task.end_date
        days_left = (deadline - today).days
        if days_left < 0:
            prefix = "Просрочена на {} дн.".format(abs(days_left))
        elif days_left in (1, 3, 7):
            prefix = "Срок через {} дн.".format(days_left)
        else:
            continue
        reminders.append({"title": "{}: {}".format(prefix, task.title), "days_left": days_left})

    count = len(reminders)
    if count == 0:
        return jsonify({"count": 0})
    return jsonify({
        "count": count,
        "titles": [item["title"] for item in reminders],
        "message": "У вас {} задач(и), требующих внимания по сроку.".format(count),
    })


@app.route("/notifications/send", methods=["GET", "POST"])
@role_required("admin", "superadmin", "deputy")
def send_user_notification():
    users = User.query.filter_by(is_approved=True).order_by(User.full_name).all()
    departments = Department.query.order_by(Department.name).all()

    if request.method == "POST":
        target_type = (request.form.get("target_type") or "").strip()
        target_id = (request.form.get("target_id") or "").strip()
        title = (request.form.get("title") or "").strip()
        message = (request.form.get("message") or "").strip()
        recipients = []

        if not title or not message:
            flash("Заполните заголовок и текст уведомления.", "warning")
            return redirect(request.url)

        if target_type == "all":
            recipients = users
        elif target_type == "user" and target_id.isdigit():
            user = db.session.get(User, int(target_id))
            recipients = [user] if user and user.is_approved else []
        elif target_type == "department" and target_id.isdigit():
            recipients = User.query.filter_by(
                department_id=int(target_id), is_approved=True
            ).order_by(User.full_name).all()

        if not recipients:
            flash("Не удалось найти получателей уведомления.", "warning")
            return redirect(request.url)

        sender_id = current_user.id if current_user.id != -1 else None
        db.session.add_all([
            UserNotification(
                user_id=user.id,
                sender_id=sender_id,
                title=title,
                message=message,
            )
            for user in recipients
        ])
        db.session.commit()
        flash("Уведомление отправлено: получателей {}.".format(len(recipients)), "success")
        return redirect(url_for("send_user_notification"))

    return render_template(
        "notification_send.html",
        title="Отправка уведомления",
        users=users,
        departments=departments,
    )


@app.get("/notifications")
@login_required
def user_notifications_page():
    notifications_list = (
        UserNotification.query
        .filter_by(user_id=current_user.id)
        .order_by(UserNotification.created_at.desc())
        .limit(100)
        .all()
    )
    unread = [item for item in notifications_list if item.read_at is None]
    if unread:
        now = dt.datetime.utcnow()
        for item in unread:
            item.read_at = now
        db.session.commit()
    return render_template(
        "notifications.html",
        title="Уведомления",
        notifications_list=notifications_list,
    )


@app.get("/api/user-notifications")
@login_required
def user_notifications_api():
    notifications_list = (
        UserNotification.query
        .filter_by(user_id=current_user.id)
        .order_by(UserNotification.created_at.desc())
        .limit(20)
        .all()
    )
    unread_count = sum(item.read_at is None for item in notifications_list)
    return jsonify({
        "unread_count": unread_count,
        "items": [
            {
                "id": item.id,
                "title": item.title,
                "message": item.message,
                "created_at": item.created_at.strftime("%d.%m.%Y %H:%M"),
                "is_read": item.read_at is not None,
            }
            for item in notifications_list
        ],
    })


@app.post("/api/user-notifications/<int:notification_id>/read")
@login_required
def mark_user_notification_read(notification_id):
    notification = UserNotification.query.filter_by(
        id=notification_id, user_id=current_user.id
    ).first_or_404()
    if notification.read_at is None:
        notification.read_at = dt.datetime.utcnow()
        db.session.commit()
    return jsonify(success=True)
@app.route("/plans/statistics")
@role_required("admin", "superadmin")
def plans_statistics():
    shared_payload = get_report_share_payload("plans_statistics")
    month_value = str((shared_payload or {}).get("month") or request.args.get("month") or "").strip()
    period = get_plan_period(month_value)
    if not period:
        flash("Неверный месяц.", "warning")
        return redirect(url_for("plans_statistics"))
    period_start, period_end = period
    month_value = period_start.strftime("%Y-%m")

    items = (
        PlanItem.query.options(joinedload(PlanItem.linked_activity)).join(Plan).join(User, PlanItem.executor_id == User.id)
        .filter(Plan.start_date <= period_end, Plan.end_date >= period_start)
        .order_by(PlanItem.executor_id, PlanItem.position)
        .all()
    )
    is_manager = shared_payload is not None or getattr(current_user, "role", None) in ("admin", "superadmin")
    if not is_manager:
        items = [item for item in items if item.executor_id == current_user.id]

    totals = {status: sum(1 for item in items if item.status == status) for status in PLAN_ITEM_STATUSES}
    department_rows, rows = build_plan_department_rows(items)

    total = len(items)
    summary = {
        "total": total,
        "planned": totals["planned"],
        "done": totals["done"],
        "not_done": totals["not_done"],
        "cancelled": totals["cancelled"],
        "completion": round((totals["done"] / total) * 100, 1) if total else 0,
    }
    chart_data = {
        "status_labels": ["Выполнено", "Невыполнено", "Отменено", "Запланировано"],
        "status_values": [summary["done"], summary["not_done"], summary["cancelled"], summary["planned"]],
        "executor_labels": [row["user"].full_name for row in rows],
        "executor_done": [row["done"] for row in rows],
        "executor_not_done": [row["not_done"] for row in rows],
        "executor_planned": [row["planned"] for row in rows],
        "department_labels": [row["name"] for row in department_rows],
        "department_done": [row["done"] for row in department_rows],
        "department_not_done": [row["not_done"] for row in department_rows],
        "department_planned": [row["planned"] for row in department_rows],
    }
    return render_template(
        "plans_statistics.html", title="Аналитика выполнения планов",
        month_value=month_value, period_start=period_start, period_end=period_end,
        summary=summary, department_rows=department_rows, executor_rows=rows, chart_data=chart_data,
    )


@app.route("/plans/monthly-report")
@role_required("admin", "superadmin")
def plans_monthly_report():
    shared_payload = get_report_share_payload("plans_monthly_report")
    month_value = str((shared_payload or {}).get("month") or request.args.get("month") or "").strip()
    period = get_plan_period(month_value)
    if not period:
        flash("Неверный месяц.", "warning")
        return redirect(url_for("plans_monthly_report"))
    period_start, period_end = period
    month_value = period_start.strftime("%Y-%m")
    plans = Plan.query.filter(Plan.start_date <= period_end, Plan.end_date >= period_start).all()
    items = (
        PlanItem.query.join(Plan).join(User, PlanItem.executor_id == User.id)
        .filter(Plan.start_date <= period_end, Plan.end_date >= period_start)
        .order_by(User.full_name, PlanItem.position)
        .all()
    )
    department_rows, executor_rows = build_plan_department_rows(items)
    counts = {status: sum(1 for item in items if item.status == status) for status in PLAN_ITEM_STATUSES}
    summary = {
        "plans_count": len(plans), "departments_count": len(department_rows),
        "executors_count": len(executor_rows), "items_total": len(items),
        "planned_count": counts["planned"], "done_count": counts["done"],
        "not_done_count": counts["not_done"], "cancelled_count": counts["cancelled"],
        "completion_percent": round((counts["done"] / len(items)) * 100, 1) if items else 0,
    }

    return render_template(
        "plans_monthly_report.html",
        title="Сводный отчет по планам",
        month_value=month_value,
        period_start=period_start,
        period_end=period_end,
        department_rows=department_rows,
        summary=summary,
        get_plan_status_meta=plan_item_status_meta,
        deadline_options=PLAN_DEADLINES,
    )


@app.route("/my-plans")
@login_required
def my_plans():
    from datetime import date

    # Получаем все планы, где текущий пользователь — исполнитель
    executors = PlanExecutor.query.filter_by(user_id=current_user.id).all()
    plan_ids = list({ex.plan_id for ex in executors})

    plans = Plan.query.filter(Plan.id.in_(plan_ids)).all()
    my_plans_data = []

    for p in plans:
        # Все активности этого пользователя в рамках плана
        my_activities = Activity.query.filter_by(plan_id=p.id, owner_id=current_user.id).all()

        total = len(my_activities)
        done = sum(1 for a in my_activities if a.status in ("done", "approved"))
        progress = round((done / total) * 100, 1) if total > 0 else 0

        # Дни до конца плана
        days_left = (p.end_date - date.today()).days if p.end_date else None

        my_plans_data.append({
            "id": p.id,
            "start_date": p.start_date,
            "end_date": p.end_date,
            "text": p.plan_text,
            "progress": progress,
            "days_left": days_left,
        })

    return render_template("my_plans.html", title="Мои планы", plans=my_plans_data)
@app.get("/exports/<string:kind>.<string:format_name>")
@login_required
def export_journal(kind, format_name):
    if format_name != "csv" or kind not in ("activities", "letters", "memos", "plans"):
        abort(404)
    today = dt.date.today()
    date_from = request.args.get("date_from", "")
    date_to = request.args.get("date_to", "")
    headers, rows, title = [], [], ""
    if kind == "activities":
        query = Activity.query.options(joinedload(Activity.owner), joinedload(Activity.type))
        if date_from:
            query = query.filter(Activity.start_date >= dt.datetime.strptime(date_from, "%Y-%m-%d"))
        if date_to:
            query = query.filter(Activity.start_date <= dt.datetime.combine(dt.datetime.strptime(date_to, "%Y-%m-%d").date(), dt.datetime.max.time()))
        headers, title = ["№", "Название", "Исполнитель", "Статус", "Срок"], "Журнал задач"
        rows = [[a.id, a.title, a.owner.full_name if a.owner else "-", get_activity_status_meta(a.status)["label"], a.end_date.strftime("%d.%m.%Y") if a.end_date else "-"] for a in query.order_by(Activity.id.desc()).all()]
    elif kind in ("letters", "memos"):
        memo_source = LetterSource.query.filter_by(name="Служебная записка").first()
        query = Letter.query.options(joinedload(Letter.executor), joinedload(Letter.source))
        query = query.filter(Letter.source_id == memo_source.id) if kind == "memos" and memo_source else (query.filter(Letter.source_id != memo_source.id) if kind == "letters" and memo_source else query)
        headers, title = ["Номер", "Дата", "Тема", "Исполнитель"], "Журнал служебных записок" if kind == "memos" else "Журнал писем"
        rows = [[item.reg_number or item.id, item.letter_date.strftime("%d.%m.%Y") if item.letter_date else "-", item.subject, item.executor.full_name if item.executor else "-"] for item in query.order_by(Letter.id.desc()).all()]
    else:
        headers, title = ["№", "Период", "Автор", "Статус"], "Журнал планов"
        rows = [[item.id, "{} - {}".format(item.start_date.strftime("%d.%m.%Y"), item.end_date.strftime("%d.%m.%Y")), item.creator.full_name if item.creator else "-", item.approval_status] for item in Plan.query.options(joinedload(Plan.creator)).order_by(Plan.id.desc()).all()]
    payload = io.StringIO(newline="")
    writer = csv.writer(payload, delimiter=";")
    writer.writerow([title]); writer.writerow(headers); writer.writerows(rows)
    response = app.response_class(payload.getvalue().encode("utf-8-sig"), mimetype="text/csv; charset=utf-8")
    response.headers["Content-Disposition"] = "attachment; filename={}.csv".format(kind)
    return response

@app.route("/search", methods=["GET"])
@login_required
def global_search():
    """Single search across tasks, letters, memos, and attachment text."""
    q = request.args.get("q", "").strip()
    results_acts, results_lets = [], []
    task_results, letter_results, memo_results = [], [], []

    if q:
        like = "%{}%".format(q)
        task_results = Activity.query.filter(or_(
            Activity.title.ilike(like), Activity.description.ilike(like),
            Activity.direction.ilike(like), Activity.outgoing_document.ilike(like),
        )).order_by(Activity.start_date.desc()).limit(50).all()
        memo_source = LetterSource.query.filter_by(name="Служебная записка").first()
        letters_query = Letter.query.filter(or_(
            Letter.subject.ilike(like), Letter.body.ilike(like), Letter.reg_number.ilike(like),
        ))
        if memo_source:
            letter_results = letters_query.filter(Letter.source_id != memo_source.id).order_by(Letter.letter_date.desc()).limit(50).all()
            memo_results = letters_query.filter(Letter.source_id == memo_source.id).order_by(Letter.letter_date.desc()).limit(50).all()
        else:
            letter_results = letters_query.order_by(Letter.letter_date.desc()).limit(50).all()
        # === Поиск по активности ===
        sql_a = text("""
        SELECT ad.id AS doc_id, ad.filename, a.id AS parent_id, a.title AS parent_title,
               'activity' AS source,
               snippet(activity_documents_fts, 0, '<mark>', '</mark>', '...', 20) AS snippet
        FROM activity_documents_fts
        JOIN activity_documents ad ON ad.id = activity_documents_fts.rowid
        LEFT JOIN activities a ON a.id = ad.activity_id
        WHERE activity_documents_fts MATCH :q
        ORDER BY rank;
        """)
        results_acts = db.session.execute(sql_a, {"q": q}).fetchall()

        # === Поиск по письмам ===
        sql_l = text("""
        SELECT ld.id AS doc_id, ld.filename, l.id AS parent_id, l.subject AS parent_title,
               'letter' AS source,
               snippet(letter_documents_fts, 0, '<mark>', '</mark>', '...', 20) AS snippet
        FROM letter_documents_fts
        JOIN letter_documents ld ON ld.id = letter_documents_fts.rowid
        LEFT JOIN letters l ON l.id = ld.letter_id
        WHERE letter_documents_fts MATCH :q
        ORDER BY rank;
        """)
        results_lets = db.session.execute(sql_l, {"q": q}).fetchall()

    return render_template(
        "search_results.html",
        title="Поиск по документам",
        q=q,
        task_results=task_results,
        letter_results=letter_results,
        memo_results=memo_results,
        results_acts=results_acts,
        results_lets=results_lets
    )

def init_fts_indexes():
    """Создание виртуальных таблиц FTS5 и триггеров"""
    with db.engine.begin() as conn:
        # === activity_documents ===
        conn.exec_driver_sql("""
        CREATE VIRTUAL TABLE IF NOT EXISTS activity_documents_fts
        USING fts5(doc_rec, activity_id UNINDEXED, content='activity_documents', content_rowid='id');
        """)
        conn.exec_driver_sql("""
        CREATE TRIGGER IF NOT EXISTS activity_documents_ai AFTER INSERT ON activity_documents
        BEGIN
            INSERT INTO activity_documents_fts(rowid, doc_rec, activity_id)
            VALUES (new.id, new.doc_rec, new.activity_id);
        END;
        """)
        conn.exec_driver_sql("""
        CREATE TRIGGER IF NOT EXISTS activity_documents_ad AFTER DELETE ON activity_documents
        BEGIN
            DELETE FROM activity_documents_fts WHERE rowid = old.id;
        END;
        """)
        conn.exec_driver_sql("""
        CREATE TRIGGER IF NOT EXISTS activity_documents_au AFTER UPDATE ON activity_documents
        BEGIN
            UPDATE activity_documents_fts
            SET doc_rec = new.doc_rec, activity_id = new.activity_id
            WHERE rowid = old.id;
        END;
        """)
        conn.exec_driver_sql("""
        INSERT INTO activity_documents_fts(rowid, doc_rec, activity_id)
        SELECT id, doc_rec, activity_id FROM activity_documents
        WHERE id NOT IN (SELECT rowid FROM activity_documents_fts);
        """)

        # === letter_documents ===
        conn.exec_driver_sql("""
        CREATE VIRTUAL TABLE IF NOT EXISTS letter_documents_fts
        USING fts5(doc_rec, letter_id UNINDEXED, content='letter_documents', content_rowid='id');
        """)
        conn.exec_driver_sql("""
        CREATE TRIGGER IF NOT EXISTS letter_documents_ai AFTER INSERT ON letter_documents
        BEGIN
            INSERT INTO letter_documents_fts(rowid, doc_rec, letter_id)
            VALUES (new.id, new.doc_rec, new.letter_id);
        END;
        """)
        conn.exec_driver_sql("""
        CREATE TRIGGER IF NOT EXISTS letter_documents_ad AFTER DELETE ON letter_documents
        BEGIN
            DELETE FROM letter_documents_fts WHERE rowid = old.id;
        END;
        """)
        conn.exec_driver_sql("""
        CREATE TRIGGER IF NOT EXISTS letter_documents_au AFTER UPDATE ON letter_documents
        BEGIN
            UPDATE letter_documents_fts
            SET doc_rec = new.doc_rec, letter_id = new.letter_id
            WHERE rowid = old.id;
        END;
        """)
        conn.exec_driver_sql("""
        INSERT INTO letter_documents_fts(rowid, doc_rec, letter_id)
        SELECT id, doc_rec, letter_id FROM letter_documents
        WHERE id NOT IN (SELECT rowid FROM letter_documents_fts);
        """)

        print("📚 FTS5 индексы инициализированы.")
from flask import send_file, abort
import os

# === Скачивание документа активности ===
@app.route("/documents/<int:doc_id>/download")
@login_required
def download_activity_document(doc_id):
    doc = ActivityDocument.query.get_or_404(doc_id)
    allowed = False

    if current_user.role in ("admin", "manager"):
        allowed = True
    elif doc.activity.owner_id == current_user.id:
        allowed = True
    else:
        allowed = (
            DocumentAccess.query.filter_by(
                doc_type="activity",
                doc_id=doc.id,
                user_id=current_user.id
            ).count() > 0
        )

    if not allowed:
        # Автор документа
        author = User.query.get(doc.activity.owner_id)
        return render_template(
            "document_request_access.html",
            title="Доступ к документу запрещён",
            doc=doc,
            author=author,
            doc_type="activity"
        ), 403

    file_path = os.path.join("static", doc.filepath)
    if not os.path.exists(file_path):
        abort(404, "Файл отсутствует на сервере")

    # 🧾 Лог загрузки
    db.session.add(DownloadLog(
        user_id=current_user.id,
        doc_type="activity",
        doc_id=doc.id
    ))
    db.session.commit()

    return send_file(file_path, as_attachment=True, download_name=doc.filename)

# === Скачивание документа письма ===
@app.route("/letters/documents/<int:doc_id>/download")
@login_required
def download_letter_document(doc_id):
    doc = LetterDocument.query.get_or_404(doc_id)
    allowed = False

    if current_user.role in ("admin", "manager"):
        allowed = True
    elif doc.letter.author_id == current_user.id:
        allowed = True
    else:
        allowed = (
            DocumentAccess.query.filter_by(
                doc_type="letter",
                doc_id=doc.id,
                user_id=current_user.id
            ).count() > 0
        )

    if not allowed:
        author = User.query.get(doc.letter.author_id)
        return render_template(
            "document_request_access.html",
            title="Доступ к документу запрещён",
            doc=doc,
            author=author,
            doc_type="letter"
        ), 403

    file_path = os.path.join("static", doc.filepath)
    if not os.path.exists(file_path):
        abort(404, "Файл отсутствует на сервере")

    db.session.add(DownloadLog(
        user_id=current_user.id,
        doc_type="letter",
        doc_id=doc.id
    ))
    db.session.commit()

    return send_file(file_path, as_attachment=True, download_name=doc.filename)

@app.get("/document_access")
@login_required
def document_access_manager():
    """Управление доступом к документам (активности + письма)"""
    if current_user.role in ("admin", "manager"):
        activity_docs = ActivityDocument.query.all()
        letter_docs = LetterDocument.query.all()
    else:
        activity_docs = (
            ActivityDocument.query
            .join(Activity)
            .filter(Activity.owner_id == current_user.id)
            .all()
        )
        letter_docs = (
            LetterDocument.query
            .join(Letter)
            .filter(Letter.author_id == current_user.id)
            .all()
        )

    # соберём все доступы сразу
    access_map = {}
    for acc in DocumentAccess.query.all():
        access_map.setdefault((acc.doc_type, acc.doc_id), []).append(acc)

    users = User.query.order_by(User.full_name).all()

    return render_template(
        "document_access_manager.html",
        activity_docs=activity_docs,
        letter_docs=letter_docs,
        users=users,
        access_map=access_map,
        is_admin=current_user.role in ("admin", "manager")
    )

@app.post("/document_access/grant")
@login_required
def grant_doc_access():
    doc_type = request.form["doc_type"]
    doc_id = int(request.form["doc_id"])
    user_id = int(request.form["user_id"])

    # Автор или админ может выдавать
    if not has_document_access(current_user.id, doc_type, doc_id):
        return jsonify({"error": "Нет прав выдавать доступ"}), 403

    exists = DocumentAccess.query.filter_by(
        doc_type=doc_type, doc_id=doc_id, user_id=user_id
    ).first()
    if exists:
        return jsonify({"error": "Уже есть доступ"}), 400

    acc = DocumentAccess(
        doc_type=doc_type,
        doc_id=doc_id,
        user_id=user_id,
        granted_by=current_user.id
    )
    db.session.add(acc)
    db.session.commit()
    return jsonify({"success": True})


@app.post("/document_access/revoke")
@login_required
def revoke_doc_access():
    doc_type = request.form["doc_type"]
    doc_id = int(request.form["doc_id"])
    user_id = int(request.form["user_id"])

    # Автор или админ может отзывать
    if not has_document_access(current_user.id, doc_type, doc_id):
        return jsonify({"error": "Нет прав отзывать доступ"}), 403

    DocumentAccess.query.filter_by(
        doc_type=doc_type, doc_id=doc_id, user_id=user_id
    ).delete()
    db.session.commit()
    return jsonify({"success": True})
@app.route("/chat")
@login_required
def chat_page():
    users = User.query.filter(User.id != current_user.id, User.is_approved == True).all()
    return render_template("chat.html", title="Чат", users=users)


@app.post("/api/chat/send")
@login_required
def chat_send():
    text = request.form.get("text", "").strip()
    receiver_id = request.form.get("receiver_id")
    file = request.files.get("file")

    file_path = None
    if file and file.filename:
        filename = secure_filename_rus(file.filename)
        upload_dir = os.path.join(app.root_path, "static", "chat_uploads")
        os.makedirs(upload_dir, exist_ok=True)
        unique_name = f"{uuid.uuid4().hex}_{filename}"
        file.save(os.path.join(upload_dir, unique_name))
        file_path = f"chat_uploads/{unique_name}"

    msg = ChatMessage(
        sender_id=current_user.id,
        receiver_id=int(receiver_id) if receiver_id and receiver_id != "all" else None,
        text=text,
        file_path=file_path
    )
    db.session.add(msg)
    db.session.commit()
    return jsonify(success=True, message=serialize_chat_message(msg))

@app.get("/api/chat/messages")
@login_required
def chat_messages():
    """Сообщения либо общего чата, либо между пользователями"""
    user_id = request.args.get("user_id")

    q = ChatMessage.query

    if user_id and user_id != "all":
        uid = int(user_id)
        q = q.filter(
            or_(
                and_(ChatMessage.sender_id == current_user.id, ChatMessage.receiver_id == uid),
                and_(ChatMessage.sender_id == uid, ChatMessage.receiver_id == current_user.id)
            )
        )
    else:
        q = q.filter(ChatMessage.receiver_id.is_(None))  # общий чат

    q = q.order_by(ChatMessage.created_at.desc()).limit(50)

    msgs = reversed(q.all())
    data = [
        serialize_chat_message(m)
        for m in msgs
    ]
    return jsonify(data)
@app.get("/api/chat/users")
@login_required
def chat_users():
    """Возвращает список всех активных пользователей для чата"""
    users = User.query.filter(User.id != current_user.id, User.is_approved == True).all()
    data = [
        {"id": u.id, "name": u.full_name, "role": u.role}
        for u in users
    ]
    return jsonify(data)


@app.get("/api/chat/incoming")
@login_required
def chat_incoming():
    direct_after_raw = (request.args.get("direct_after_id") or "0").strip()
    general_after_raw = (request.args.get("general_after_id") or "0").strip()
    try:
        direct_after_id = int(direct_after_raw)
    except ValueError:
        direct_after_id = 0
    try:
        general_after_id = int(general_after_raw)
    except ValueError:
        general_after_id = 0

    direct_messages = (
        ChatMessage.query
        .filter(ChatMessage.receiver_id == current_user.id)
        .filter(ChatMessage.sender_id != current_user.id)
        .filter(ChatMessage.id > direct_after_id)
        .order_by(ChatMessage.id.asc())
        .limit(20)
        .all()
    )

    general_messages = (
        ChatMessage.query
        .filter(ChatMessage.receiver_id.is_(None))
        .filter(ChatMessage.sender_id != current_user.id)
        .filter(ChatMessage.id > general_after_id)
        .order_by(ChatMessage.id.asc())
        .limit(20)
        .all()
    )

    direct_data = [
        {
            "id": m.id,
            "sender": get_chat_sender_name(m),
            "sender_id": m.sender_id,
            "text": m.text or "",
            "time": m.created_at.strftime("%H:%M"),
            "kind": "direct",
        }
        for m in direct_messages
    ]

    general_data = [
        {
            "id": m.id,
            "sender": get_chat_sender_name(m),
            "sender_id": m.sender_id,
            "text": m.text or "",
            "time": m.created_at.strftime("%H:%M"),
            "kind": "general",
        }
        for m in general_messages
    ]

    direct_last_id = direct_data[-1]["id"] if direct_data else direct_after_id
    general_last_id = general_data[-1]["id"] if general_data else general_after_id

    return jsonify({
        "direct_messages": direct_data,
        "general_messages": general_data,
        "direct_last_id": direct_last_id,
        "general_last_id": general_last_id,
    })


@app.get("/api/chat/tray/incoming")
@login_required
def chat_tray_incoming():
    direct_after_raw = (request.args.get("direct_after_id") or "0").strip()
    general_after_raw = (request.args.get("general_after_id") or "0").strip()
    try:
        direct_after_id = int(direct_after_raw)
    except ValueError:
        direct_after_id = 0
    try:
        general_after_id = int(general_after_raw)
    except ValueError:
        general_after_id = 0

    direct_messages = (
        ChatMessage.query
        .filter(ChatMessage.receiver_id == current_user.id)
        .filter(ChatMessage.sender_id != current_user.id)
        .filter(ChatMessage.id > direct_after_id)
        .order_by(ChatMessage.id.asc())
        .limit(20)
        .all()
    )

    general_messages = (
        ChatMessage.query
        .filter(ChatMessage.receiver_id.is_(None))
        .filter(ChatMessage.sender_id != current_user.id)
        .filter(ChatMessage.id > general_after_id)
        .order_by(ChatMessage.id.asc())
        .limit(20)
        .all()
    )

    direct_last_id = direct_messages[-1].id if direct_messages else direct_after_id
    general_last_id = general_messages[-1].id if general_messages else general_after_id

    all_messages = [
        {
            "id": m.id,
            "kind": "direct",
            "sender_id": m.sender_id or 0,
            "sender": get_chat_sender_name(m),
            "text": m.text or "",
            "time": m.created_at.strftime("%H:%M"),
        }
        for m in direct_messages
    ] + [
        {
            "id": m.id,
            "kind": "general",
            "sender_id": m.sender_id or 0,
            "sender": get_chat_sender_name(m),
            "text": m.text or "",
            "time": m.created_at.strftime("%H:%M"),
        }
        for m in general_messages
    ]
    all_messages.sort(key=lambda item: item["id"])

    lines = [
        f"DIRECT_LAST\t{direct_last_id}",
        f"GENERAL_LAST\t{general_last_id}",
    ]
    for item in all_messages:
        lines.append(
            "\t".join([
                "MSG",
                str(item["id"]),
                item["kind"],
                str(item["sender_id"]),
                quote(item["sender"], safe=""),
                quote(item["text"], safe=""),
                quote(item["time"], safe=""),
            ])
        )

    payload = "\n".join(lines)
    return current_app.response_class(payload, mimetype="text/plain; charset=utf-8")


@app.get("/api/chat/tray/users")
@login_required
def chat_tray_users():
    users = (
        User.query
        .filter(User.id != current_user.id, User.is_approved == True)
        .order_by(User.full_name.asc())
        .all()
    )

    lines = []
    for user in users:
        lines.append(
            "\t".join([
                "USER",
                str(user.id),
                quote(user.full_name or "", safe=""),
                quote(user.role or "", safe=""),
            ])
        )

    payload = "\n".join(lines)
    return current_app.response_class(payload, mimetype="text/plain; charset=utf-8")
@app.post("/letters/<int:letter_id>/request-access")
@login_required
def request_letter_access(letter_id):
    le = Letter.query.get_or_404(letter_id)
    author = User.query.get(le.author_id)
    message = request.form.get("message", "").strip()

    # 📎 Формируем ссылку на письмо
    letter_url = url_for("letter_view", letter_id=le.id, _external=True)

    # 📨 Текст уведомления автору
    msg_text = (
        f"📬 Пользователь <b>{current_user.full_name}</b> просит открыть доступ "
        f"к письму №{le.id} — <a href='{letter_url}'>перейти к письму</a>.<br><br>"
        f"Сообщение: {message or '(без текста)'}"
    )

    # 💬 Создаём сообщение автору (внутренний чат)
    chat_msg = ChatMessage(
        sender_id=current_user.id,
        receiver_id=author.id,
        text=msg_text
    )
    db.session.add(chat_msg)
    db.session.commit()

    flash("Запрос автору письма отправлен.", "success")
    return redirect(url_for("letters_page"))
@app.post("/letters/<int:letter_id>/grant-access")
@login_required
def grant_letter_access(letter_id):
    """Автор письма выдаёт доступ другому пользователю"""
    le = Letter.query.get_or_404(letter_id)

    # Только автор или админ может выдавать
    if current_user.id != le.author_id and current_user.role != "admin":
        return jsonify({"error": "Нет прав выдавать доступ"}), 403

    user_id = int(request.form.get("user_id", 0))
    if not user_id:
        return jsonify({"error": "Не указан пользователь"}), 400

    exists = DocumentAccess.query.filter_by(
        doc_type="letter", doc_id=le.id, user_id=user_id
    ).first()
    if exists:
        return jsonify({"error": "Уже есть доступ"}), 400

    acc = DocumentAccess(
        doc_type="letter",
        doc_id=le.id,
        user_id=user_id,
        granted_by=current_user.id
    )
    db.session.add(acc)
    db.session.commit()

    return jsonify({"success": True, "user_id": user_id})
@app.get("/letters/<int:letter_id>/access-list")
@login_required
def letter_access_list(letter_id):
    """Возвращает список пользователей, имеющих доступ к письму"""
    le = Letter.query.get_or_404(letter_id)

    # Только автор, админ или менеджер может смотреть
    if current_user.id != le.author_id and current_user.role not in ("admin", "manager"):
        return jsonify({"error": "Нет прав"}), 403

    accesses = (
        DocumentAccess.query
        .join(User, DocumentAccess.user_id == User.id)
        .filter(DocumentAccess.doc_type == "letter", DocumentAccess.doc_id == le.id)
        .with_entities(DocumentAccess.id, User.full_name.label("user_name"), User.id.label("user_id"))
        .all()
    )

    data = [{"id": a.id, "user_id": a.user_id, "user_name": a.user_name} for a in accesses]
    return jsonify(data)


@app.post("/letters/<int:letter_id>/revoke-access")
@login_required
def revoke_letter_access(letter_id):
    """Отзыв доступа к письму"""
    le = Letter.query.get_or_404(letter_id)

    if current_user.id != le.author_id and current_user.role != "admin":
        return jsonify({"error": "Нет прав отзывать доступ"}), 403

    user_id = int(request.form.get("user_id", 0))
    if not user_id:
        return jsonify({"error": "Не указан пользователь"}), 400

    DocumentAccess.query.filter_by(doc_type="letter", doc_id=le.id, user_id=user_id).delete()
    db.session.commit()

    return jsonify({"success": True})
@app.post("/documents/<doc_type>/<int:doc_id>/request-access")
@login_required
def request_document_access(doc_type, doc_id):
    """Пользователь запрашивает доступ к документу у автора"""
    if doc_type == "letter":
        doc = LetterDocument.query.get_or_404(doc_id)
        author = User.query.get(doc.letter.author_id)
        letter_url = url_for("letter_view", letter_id=doc.letter.id, _external=True)
        msg_text = (
            f"📎 Пользователь <b>{current_user.full_name}</b> просит доступ к документу "
            f"<i>{doc.filename}</i> из письма №{doc.letter.id}.<br>"
            f"<a href='{letter_url}'>Открыть письмо</a><br>"
            f"Сообщение: {request.form.get('message') or '(без текста)'}"
        )

    elif doc_type == "activity":
        doc = ActivityDocument.query.get_or_404(doc_id)
        author = User.query.get(doc.activity.owner_id)
        activity_url = url_for("activity_view", activity_id=doc.activity.id, _external=True)
        msg_text = (
            f"📎 Пользователь <b>{current_user.full_name}</b> просит доступ к документу "
            f"<i>{doc.filename}</i> из активности №{doc.activity.id}.<br>"
            f"<a href='{activity_url}'>Открыть активность</a><br>"
            f"Сообщение: {request.form.get('message') or '(без текста)'}"
        )

    else:
        abort(400, "Неизвестный тип документа")

    # 💬 Отправляем сообщение автору
    chat_msg = ChatMessage(
        sender_id=current_user.id,
        receiver_id=author.id,
        text=msg_text
    )
    db.session.add(chat_msg)
    db.session.commit()

    flash("Запрос автору отправлен.", "success")
    return redirect(url_for("dashboard"))
@app.get("/admin/download-logs")
@role_required("admin", "manager")
def admin_download_logs():
    logs = (
        DownloadLog.query
        .order_by(DownloadLog.timestamp.desc())
        .limit(100)
        .all()
    )
    return render_template("admin_download_logs.html", logs=logs, title="Логи загрузок документов")
@app.route("/files/upload", methods=["GET", "POST"])
@login_required
def upload_general_document():
    if request.method == "POST":
        file = request.files.get("file")
        desc = request.form.get("description", "").strip()

        if not file or not file.filename:
            flash("⚠️ Файл не выбран", "warning")
            return redirect(request.url)

        allowed_ext = {"png", "jpg", "jpeg", "pdf", "txt", "docx", "xlsx", "zip", "rar"}
        ext = file.filename.rsplit(".", 1)[-1].lower()
        if ext not in allowed_ext:
            flash(f"🚫 Недопустимый формат файла: .{ext}", "danger")
            return redirect(request.url)

        upload_dir = os.path.join(app.root_path, "static", "general_uploads")
        os.makedirs(upload_dir, exist_ok=True)

        filename = secure_filename_rus(file.filename)
        random_name = f"{uuid.uuid4().hex}_{filename}"
        save_path = os.path.join(upload_dir, random_name)
        file.save(save_path)

        doc = GeneralDocument(
            filename=filename,
            filepath=f"general_uploads/{random_name}",
            uploaded_by=current_user.id,
            description=desc
        )
        db.session.add(doc)
        db.session.commit()
        flash("📁 Документ успешно загружен.", "success")
        return redirect(url_for("files_list"))

    return render_template("file_upload.html", title="Загрузка документа")
@app.get("/files")
@login_required
def files_list():
    if current_user.role in ("admin", "manager"):
        docs = GeneralDocument.query.order_by(GeneralDocument.uploaded_at.desc()).all()
    else:
        docs = (
            GeneralDocument.query.outerjoin(DocumentAccess,
                (DocumentAccess.doc_type == "file") &
                (DocumentAccess.doc_id == GeneralDocument.id)
            )
            .filter(
                (GeneralDocument.uploaded_by == current_user.id) |
                (DocumentAccess.user_id == current_user.id)
            )
            .distinct()
            .order_by(GeneralDocument.uploaded_at.desc())
            .all()
        )

    # ✅ добавляем активные ссылки
    shared_links = {
        link.doc_id: link
        for link in SharedLink.query.filter_by(doc_type="file").all()
    }

    return render_template("files_list.html", docs=docs, shared_links=shared_links, title="Общие документы")

@app.route("/files/<int:file_id>")
@login_required
def view_file(file_id):
    doc = GeneralDocument.query.get_or_404(file_id)

    # Проверка доступа
    allowed = (
        current_user.role in ("admin", "manager") or
        doc.uploaded_by == current_user.id or
        DocumentAccess.query.filter_by(doc_type="file", doc_id=doc.id, user_id=current_user.id).count() > 0
    )

    if not allowed:
        author = User.query.get(doc.uploaded_by)
        return render_template(
            "document_request_access.html",
            title="Доступ к документу запрещён",
            doc=doc,
            author=author,
            doc_type="file"
        ), 403

    file_path = os.path.join("static", doc.filepath)
    if not os.path.exists(file_path):
        abort(404, "Файл отсутствует на сервере")

    # Лог загрузки
    db.session.add(DownloadLog(
        user_id=current_user.id,
        doc_type="file",
        doc_id=doc.id
    ))
    db.session.commit()

    return send_file(file_path, as_attachment=True, download_name=doc.filename)
@app.post("/files/<int:file_id>/share")
@login_required
def share_file(file_id):
    doc = GeneralDocument.query.get_or_404(file_id)
    if doc.uploaded_by != current_user.id and current_user.role != "admin":
        abort(403)

    # Удаляем старую ссылку, если есть
    SharedLink.query.filter_by(doc_type="file", doc_id=doc.id).delete()

    token = uuid.uuid4().hex
    days = int(request.form.get("days", 0))  # 0 = бессрочно
    expires_at = None
    if days > 0:
        expires_at = dt.datetime.utcnow() + dt.timedelta(days=days)

    link = SharedLink(
        doc_type="file",
        doc_id=doc.id,
        token=token,
        created_by=current_user.id,
        expires_at=expires_at
    )
    db.session.add(link)
    db.session.commit()

    flash("✅ Ссылка создана", "success")
    return redirect(url_for("files_list"))

@app.route("/share/<token>")
@login_required
def open_shared_link(token):
    """Открытие документа по внутренней ссылке"""
    link = SharedLink.query.filter_by(token=token).first_or_404()

    # Проверяем срок действия
    if link.expires_at and link.expires_at < dt.datetime.utcnow():
        abort(403, "Ссылка просрочена")

    # Проверяем тип документа
    if link.doc_type == "file":
        doc = GeneralDocument.query.get_or_404(link.doc_id)
        file_path = os.path.join("static", doc.filepath)
        if not os.path.exists(file_path):
            abort(404, "Файл отсутствует на сервере")

        # Проверка доступа (только авторизованные пользователи)
        allowed = (
            current_user.role in ("admin", "manager") or
            doc.uploaded_by == current_user.id or
            DocumentAccess.query.filter_by(
                doc_type="file", doc_id=doc.id, user_id=current_user.id
            ).count() > 0
        )
        if not allowed:
            author = User.query.get(doc.uploaded_by)
            return render_template(
                "document_request_access.html",
                title="Доступ к документу запрещён",
                doc=doc,
                author=author,
                doc_type="file"
            ), 403

        # Лог загрузки
        db.session.add(DownloadLog(
            user_id=current_user.id,
            doc_type="file",
            doc_id=doc.id
        ))
        db.session.commit()

        return send_file(file_path, as_attachment=True, download_name=doc.filename)

    abort(400, "Неизвестный тип документа")
@app.get("/files/<int:file_id>/access")
@login_required
def file_access(file_id):
    """Управление доступом к файлу (только автор или админ)"""
    doc = GeneralDocument.query.get_or_404(file_id)
    if doc.uploaded_by != current_user.id and current_user.role != "admin":
        abort(403)

    users = User.query.filter(User.id != current_user.id, User.is_approved == True).order_by(User.full_name).all()
    access_list = (
        DocumentAccess.query
        .join(User, DocumentAccess.user_id == User.id)
        .filter(DocumentAccess.doc_type == "file", DocumentAccess.doc_id == doc.id)
        .with_entities(DocumentAccess.user_id, User.full_name)
        .all()
    )

    return render_template("file_access.html", doc=doc, users=users, access_list=access_list, title="Доступ к файлу")


@app.post("/files/<int:file_id>/grant-access")
@login_required
def grant_file_access(file_id):
    doc = GeneralDocument.query.get_or_404(file_id)
    if doc.uploaded_by != current_user.id and current_user.role != "admin":
        abort(403)

    user_id = int(request.form.get("user_id", 0))
    if not user_id:
        return jsonify({"error": "Не выбран пользователь"}), 400

    exists = DocumentAccess.query.filter_by(doc_type="file", doc_id=doc.id, user_id=user_id).first()
    if exists:
        return jsonify({"error": "Уже есть доступ"}), 400

    db.session.add(DocumentAccess(
        doc_type="file",
        doc_id=doc.id,
        user_id=user_id,
        granted_by=current_user.id
    ))
    db.session.commit()
    return jsonify({"success": True})


@app.post("/files/<int:file_id>/revoke-access")
@login_required
def revoke_file_access(file_id):
    doc = GeneralDocument.query.get_or_404(file_id)
    if doc.uploaded_by != current_user.id and current_user.role != "admin":
        abort(403)

    user_id = int(request.form.get("user_id", 0))
    DocumentAccess.query.filter_by(doc_type="file", doc_id=doc.id, user_id=user_id).delete()
    db.session.commit()
    return jsonify({"success": True})
@app.get("/admin/users/<int:user_id>/show-password")
@role_required("admin")
def show_password(user_id):
    if current_user.role != "admin":
        abort(403)
    u = User.query.get_or_404(user_id)
    return render_template("admin_show_password.html",
                           user=u,
                           password=u.get_password())
@app.post("/admin/users/<int:user_id>/toggle-admin")
@role_required("superadmin")
def toggle_admin(user_id):
    """Суперадмин назначает/снимает администратора"""
    u = User.query.get_or_404(user_id)
    if u.role == "admin":
        u.role = "worker"
        flash(f"❎ {u.full_name} больше не администратор", "info")
    else:
        u.role = "admin"
        flash(f"✅ {u.full_name} назначен администратором", "success")
    db.session.commit()
    return redirect(url_for("admin_users"))


@app.post("/admin/users/<int:user_id>/role")
@role_required("admin", "superadmin")
def update_user_role(user_id):
    user = User.query.get_or_404(user_id)
    role = (request.form.get("role") or "worker").strip()
    allowed_roles = {"worker", "deputy", "manager"}
    if current_user.role == "superadmin":
        allowed_roles.add("admin")
    if role not in allowed_roles:
        abort(400)
    user.role = role
    db.session.commit()
    flash(f"Роль пользователя {user.full_name} обновлена.", "success")
    return redirect(url_for("admin_users"))
@app.get("/admin/users/<int:user_id>/password")
@role_required("admin", "superadmin")
def admin_show_password(user_id):
    """Просмотр расшифрованного пароля"""
    u = User.query.get_or_404(user_id)
    try:
        decrypted = decrypt_str(u.password_enc, SECRET_PASSPHRASE)
    except Exception:
        decrypted = "(ошибка расшифровки)"
    return jsonify({"password": decrypted})
# === Личный кабинет ===
@app.get("/profile")
@login_required
def profile_page():
    """Страница профиля"""
    return render_template("profile.html", title="Личный кабинет", user=current_user)


@app.post("/profile/change-password")
@login_required
def change_password():
    """Смена пароля пользователем"""
    old_pass = request.form.get("old_password", "")
    new_pass = request.form.get("new_password", "")
    confirm = request.form.get("confirm_password", "")

    # Проверяем ввод
    if not old_pass or not new_pass or not confirm:
        flash("Заполните все поля.", "warning")
        return redirect(url_for("profile_page"))

    if new_pass != confirm:
        flash("Новый пароль и подтверждение не совпадают.", "danger")
        return redirect(url_for("profile_page"))

    # Проверка старого пароля
    user = User.query.get(current_user.id)
    if not user.check_password(old_pass):
        flash("Неверный текущий пароль.", "danger")
        return redirect(url_for("profile_page"))

    # Устанавливаем новый
    user.set_password(new_pass)
    db.session.commit()
    flash("✅ Пароль успешно изменён.", "success")
    return redirect(url_for("profile_page"))
# === Доска заданий ===
from datetime import datetime, date

@app.get("/board")
@login_required
def board_page():
    """Просмотр доски заданий с вычисленным статусом"""
    if current_user.role in ("admin", "superadmin"):
        tasks = Task.query.order_by(Task.end_date.desc()).all()
    else:
        tasks = (
            Task.query.join(TaskAssignee)
            .filter(TaskAssignee.user_id == current_user.id)
            .order_by(Task.end_date.desc())
            .all()
        )

    # ⚙️ Добавим вычисленные поля для отображения
    today = date.today()
    for t in tasks:
        t.all_done = all(a.is_done for a in t.assignees)
        t.expired = t.end_date < today
        t.in_progress = not t.all_done and not t.expired

    users = User.query.all()
    return render_template("board.html", tasks=tasks, users=users)


@app.post("/board/add")
@role_required("admin", "superadmin")
def add_task():
    title = request.form.get("title")
    desc = request.form.get("description")
    assigned_to_list = request.form.getlist("assigned_to")  # ✅ несколько пользователей
    start = request.form.get("start_date")
    end = request.form.get("end_date")

    if not title or not assigned_to_list:
        flash("Введите название и выберите хотя бы одного исполнителя.", "warning")
        return redirect(url_for("board_page"))

    try:
        start_date = datetime.strptime(start, "%Y-%m-%d").date()
        end_date = datetime.strptime(end, "%Y-%m-%d").date()
    except Exception:
        flash("Неверный формат даты.", "danger")
        return redirect(url_for("board_page"))

    task = Task(
        title=title,
        description=desc,
        created_by=current_user.id,
        start_date=start_date,
        end_date=end_date,
    )
    db.session.add(task)
    db.session.flush()  # получаем id задания

    # Добавляем всех исполнителей
    for uid in assigned_to_list:
        db.session.add(TaskAssignee(task_id=task.id, user_id=int(uid)))

    db.session.commit()
    flash("✅ Задание создано и назначено пользователям.", "success")
    return redirect(url_for("board_page"))

@app.post("/board/<int:task_id>/toggle")
@login_required
def toggle_task(task_id):
    link = TaskAssignee.query.filter_by(task_id=task_id, user_id=current_user.id).first_or_404()
    link.is_done = not link.is_done
    link.done_at = dt.datetime.utcnow() if link.is_done else None
    db.session.commit()
    return redirect(url_for("board_page"))

if __name__ == "__main__":
    with app.app_context():
        db.create_all()
        init_fts_indexes()  # <=== 👈 вызываем вручную при запуске
        print("✔ Проверка таблиц выполнена:", [t.name for t in db.metadata.sorted_tables])
    app.run(
        host="0.0.0.0",
        port=5001,
        debug=True,
        use_reloader=False,
    )

