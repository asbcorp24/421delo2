# models.py
# -*- coding: utf-8 -*-
import datetime as dt
from flask_login import UserMixin
from werkzeug.security import generate_password_hash, check_password_hash
from flask_sqlalchemy import SQLAlchemy
# ВАЖНО: app.py должен создать db = SQLAlchemy(app) до этого импорта
from datetime import datetime
db = SQLAlchemy()  # создаём экземпляр, но не привязываем к app
# --- Справочники ---
from config import SECRET_PASSPHRASE
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.ciphers.aead import AESGCM
import base64, os

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


class User(UserMixin, db.Model):
    __tablename__ = "users"
    id = db.Column(db.Integer, primary_key=True)
    full_name = db.Column(db.String(255), nullable=False)
    username = db.Column(db.String(120), unique=True, nullable=False)
    password_enc = db.Column(db.Text, nullable=False)
    role = db.Column(db.String(20), nullable=False, default="worker")
    department_id = db.Column(db.Integer, db.ForeignKey("departments.id"))
    is_approved = db.Column(db.Boolean, default=False, nullable=False)
    is_enabled = db.Column(db.Boolean, default=True, nullable=False)
    created_at = db.Column(db.DateTime, default=dt.datetime.utcnow)

    department = db.relationship("Department", backref="users")

    def set_password(self, raw_password: str):
        self.password_enc = encrypt_str(raw_password, SECRET_PASSPHRASE)

    def get_password(self) -> str:
        return decrypt_str(self.password_enc, SECRET_PASSPHRASE)

    def check_password(self, raw_password: str) -> bool:
        try:
            return raw_password == self.get_password()
        except Exception as e:
            print("❌ decrypt error:", e)
            return False

    @property
    def is_active(self):
        """Flask-Login uses this property to reject disabled accounts."""
        return bool(self.is_enabled)
class Workshop(db.Model):
    __tablename__ = "workshops"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), unique=True, nullable=False)
    created_at = db.Column(db.DateTime, default=dt.datetime.utcnow, nullable=False)


class Department(db.Model):
    __tablename__ = "departments"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), unique=True, nullable=False)
    workshop_id = db.Column(db.Integer, db.ForeignKey("workshops.id"))
    workshop = db.relationship("Workshop", backref="departments")
    letter_prefix = db.Column(db.String(100))
    memo_prefix = db.Column(db.String(100))
    letter_start_number = db.Column(db.Integer)
    memo_start_number = db.Column(db.Integer)
    letter_number_series = db.Column(db.String(100))
    memo_number_series = db.Column(db.String(100))


# class User(UserMixin, db.Model):
#     __tablename__ = "users"
#     id = db.Column(db.Integer, primary_key=True)
#     full_name = db.Column(db.String(255), nullable=False)
#     username = db.Column(db.String(120), unique=True, nullable=False)
#     password_hash = db.Column(db.String(255), nullable=False)
#     role = db.Column(db.String(20), nullable=False, default="worker")
#     department_id = db.Column(db.Integer, db.ForeignKey("departments.id"))
#
#     # 👇 добавь эти две строки:
#     is_approved = db.Column(db.Boolean, default=False, nullable=False)
#     created_at = db.Column(db.DateTime, default=dt.datetime.utcnow)
#
#     department = db.relationship("Department", backref="users")
#
#     def set_password(self, raw_password: str):
#         from app import SECRET_PASSPHRASE, encrypt_str
#         self.password_enc = encrypt_str(raw_password, SECRET_PASSPHRASE)
#
#     def get_password(self) -> str:
#         from app import SECRET_PASSPHRASE, decrypt_str
#         return decrypt_str(self.password_enc, SECRET_PASSPHRASE)
#
#     def check_password(self, raw_password: str) -> bool:
#         try:
#             return raw_password == self.get_password()
#         except Exception:
#             return False
#     def set_password(self, pw: str):
#         from werkzeug.security import generate_password_hash
#         self.password_hash = generate_password_hash(pw)
#
#     def check_password(self, pw: str) -> bool:
#         from werkzeug.security import check_password_hash
#         return check_password_hash(self.password_hash, pw)

class ActivityType(db.Model):
    __tablename__ = "activity_types"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), unique=True, nullable=False)


# --- Связи многие-ко-многим ---

LetterActivityLink = db.Table(
    "letter_activity_link",
    db.Column("letter_id", db.Integer, db.ForeignKey("letters.id"), primary_key=True),
    db.Column("activity_id", db.Integer, db.ForeignKey("activities.id"), primary_key=True),
)

LetterLetterLink = db.Table(
    "letter_letter_link",
    db.Column("src_letter_id", db.Integer, db.ForeignKey("letters.id"), primary_key=True),
    db.Column("dst_letter_id", db.Integer, db.ForeignKey("letters.id"), primary_key=True),
)


# --- Активности ---

class Activity(db.Model):
    __tablename__ = "activities"
    id = db.Column(db.Integer, primary_key=True)

    type_id = db.Column(db.Integer, db.ForeignKey("activity_types.id"), nullable=True)
    type = db.relationship("ActivityType")

    owner_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False)
    owner = db.relationship("User", foreign_keys=[owner_id], backref="activities")
    created_by_id = db.Column(db.Integer, db.ForeignKey("users.id"))
    created_by = db.relationship("User", foreign_keys=[created_by_id])

    title = db.Column(db.String(255), nullable=False)
    description = db.Column(db.Text)
    start_date = db.Column(db.DateTime, default=dt.datetime.utcnow)
    end_date = db.Column(db.DateTime)
    status = db.Column(db.String(32), default="open")
    priority = db.Column(db.Integer, nullable=False, default=3)
    complexity_level = db.Column(db.Integer, nullable=False, default=3)
    plan_id = db.Column(db.Integer, db.ForeignKey("plans.id"))
    # The task was created with "add to plan" selected, but its monthly plan
    # did not exist yet. It is consumed when a suitable plan is created.
    plan_requested = db.Column(db.Boolean, nullable=False, default=False)
    approved_by_id = db.Column(db.Integer, db.ForeignKey("users.id"))
    approved_by = db.relationship("User", foreign_keys=[approved_by_id])
    approved_at = db.Column(db.DateTime)
    approve_comment = db.Column(db.Text)
    direction = db.Column(db.String(255))  # Направление (можно выбрать из datalist)
    outgoing_document = db.Column(db.String(255))
    assigned_department_id = db.Column(db.Integer, db.ForeignKey("departments.id"))
    assigned_department = db.relationship("Department", foreign_keys=[assigned_department_id])
    decision = db.Column(db.Text)  # Решение по активности
    status_comment = db.Column(db.Text)
    postponed_to = db.Column(db.DateTime)
    archived_at = db.Column(db.DateTime)
    archived_by_id = db.Column(db.Integer, db.ForeignKey("users.id"))
    extra_values = db.Column(db.JSON)  # JSON с добавочными значениями
    # связь с письмом (старое поле можно оставить для обратной совместимости)
    source_letter_id = db.Column(db.Integer, db.ForeignKey("letters.id"))
    source_letter = db.relationship("Letter", foreign_keys=[source_letter_id])
    completion_memo_id = db.Column(db.Integer, db.ForeignKey("letters.id"))
    completion_memo = db.relationship("Letter", foreign_keys=[completion_memo_id])

    # 🔗 связь многие-ко-многим между письмами и активностями
    letters = db.relationship(
        "Letter",
        secondary=LetterActivityLink,
        back_populates="activities"
    )
    logs = db.relationship(
        "ActivityLog",
        backref="activity",
        cascade="all, delete-orphan",
        lazy="joined"
    )



class ActivityDocument(db.Model):
    __tablename__ = "activity_documents"
    id = db.Column(db.Integer, primary_key=True)
    activity_id = db.Column(db.Integer, db.ForeignKey("activities.id"), nullable=False)
    filename = db.Column(db.String(255), nullable=False)
    uploaded_by_id = db.Column(db.Integer, db.ForeignKey("users.id"))
    filepath = db.Column(db.String(500), nullable=False)
    uploaded_at = db.Column(db.DateTime, default=dt.datetime.utcnow)

    activity = db.relationship("Activity", backref="documents")
    doc_rec = db.Column(db.Text)  # ✅ сюда сохраняем распознанный текст
    ocr_status = db.Column(db.String(20), nullable=False, default="completed")
    ocr_error = db.Column(db.Text)
    ocr_started_at = db.Column(db.DateTime)
    ocr_completed_at = db.Column(db.DateTime)

class ActivityHistory(db.Model):
    __tablename__ = "activity_history"
    id = db.Column(db.Integer, primary_key=True)
    activity_id = db.Column(db.Integer, db.ForeignKey("activities.id"), nullable=False, index=True)
    author_id = db.Column(db.Integer, db.ForeignKey("users.id"))
    action = db.Column(db.String(100), nullable=False)
    details = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime, default=dt.datetime.utcnow, nullable=False)
    author = db.relationship("User", foreign_keys=[author_id])

# --- Письма ---
LetterDepartments = db.Table(
    "letter_departments",
    db.Column("letter_id", db.Integer, db.ForeignKey("letters.id"), primary_key=True),
    db.Column("department_id", db.Integer, db.ForeignKey("departments.id"), primary_key=True)
)

# --- Письма ---

# --- Письма ---
class Letter(db.Model):
    __tablename__ = "letters"
    id = db.Column(db.Integer, primary_key=True)
    reg_number = db.Column(db.String(100))
    subject = db.Column(db.String(255), nullable=False)
    body = db.Column(db.Text)
    author_id = db.Column(db.Integer, db.ForeignKey("users.id"))
    source_id = db.Column(db.Integer, db.ForeignKey("letter_sources.id"))
    executor_id = db.Column(db.Integer, db.ForeignKey("users.id"))
    archived_at = db.Column(db.DateTime)
    archived_by_id = db.Column(db.Integer, db.ForeignKey("users.id"))
    approval_status = db.Column(db.String(20), nullable=False, default="pending")
    approved_by_id = db.Column(db.Integer, db.ForeignKey("users.id"))
    approved_at = db.Column(db.DateTime)

    departments = db.relationship(
        "Department",
        secondary="letter_departments",
        backref="letters"
    )

    letter_date = db.Column(db.Date)
    created_at = db.Column(db.DateTime, default=dt.datetime.utcnow)
    author = db.relationship("User", foreign_keys=[author_id], backref="letters")
    source = db.relationship("LetterSource", backref="letters")
    executor = db.relationship("User", foreign_keys=[executor_id], backref="assigned_letters")
    approved_by = db.relationship("User", foreign_keys=[approved_by_id])
    extra_values = db.Column(db.JSON, default=list)
    recipients = db.relationship("LetterRecipient", backref="letter", cascade="all, delete")
    documents = db.relationship("LetterDocument", backref="letter", cascade="all, delete")

    # 🔗 связь многие-ко-многим с активностями
    activities = db.relationship(
        "Activity",
        secondary=LetterActivityLink,
        back_populates="letters"
    )

class LetterRecipient(db.Model):
    __tablename__ = "letter_recipients"
    id = db.Column(db.Integer, primary_key=True)
    letter_id = db.Column(db.Integer, db.ForeignKey("letters.id"), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False)
    role = db.Column(db.String(20), nullable=False)  # executor / observer
    is_read = db.Column(db.Boolean, default=False)
    read_at = db.Column(db.DateTime)

    user = db.relationship("User")


class LetterDocument(db.Model):
    __tablename__ = "letter_documents"
    id = db.Column(db.Integer, primary_key=True)
    letter_id = db.Column(db.Integer, db.ForeignKey("letters.id"), nullable=False)
    filename = db.Column(db.String(255), nullable=False)
    filepath = db.Column(db.String(255), nullable=False)
    doc_rec = db.Column(db.Text)  # ✅ сюда сохраняем распознанный текст
    ocr_status = db.Column(db.String(20), nullable=False, default="completed")
    ocr_error = db.Column(db.Text)
    ocr_started_at = db.Column(db.DateTime)
    ocr_completed_at = db.Column(db.DateTime)

class ActivityLog(db.Model):
    __tablename__ = "activity_logs"
    id = db.Column(db.Integer, primary_key=True)
    activity_id = db.Column(db.Integer, db.ForeignKey("activities.id"), nullable=False)

    text = db.Column(db.Text, nullable=False)
    entry_date = db.Column(db.DateTime, default=datetime.utcnow)

    parent_id = db.Column(db.Integer, db.ForeignKey("activity_logs.id"))
    is_done = db.Column(db.Boolean, default=False)

    # 🔁 рекурсивная связь (родитель → дети)
    children = db.relationship(
        "ActivityLog",
        backref=db.backref("parent", remote_side=[id]),
        cascade="all, delete-orphan",
        lazy="joined"
    )
class ValueTemplate(db.Model):
    __tablename__ = "value_templates"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(64), unique=True, nullable=False)
    rusname = db.Column(db.String(255), nullable=False)
    type = db.Column(db.String(32), nullable=False)  # string, int, float, bool, date, text
    is_active = db.Column(db.Boolean, default=True)

class Plan(db.Model):
    __tablename__ = "plans"
    id = db.Column(db.Integer, primary_key=True)
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=False)
    plan_text = db.Column(db.Text, nullable=False)
    created_by = db.Column(db.Integer, db.ForeignKey("users.id"))
    created_at = db.Column(db.DateTime, default=dt.datetime.utcnow)
    is_completed = db.Column(db.Boolean, default=False)
    approval_status = db.Column(db.String(32), nullable=False, default="submitted")
    approved_by_id = db.Column(db.Integer, db.ForeignKey("users.id"))
    approved_at = db.Column(db.DateTime)
    revision_comment = db.Column(db.Text)
    creator = db.relationship("User", foreign_keys=[created_by], backref="created_plans")
    approved_by = db.relationship("User", foreign_keys=[approved_by_id])
    executors = db.relationship("PlanExecutor", backref="plan", cascade="all, delete-orphan")
    items = db.relationship("PlanItem", backref="plan", cascade="all, delete-orphan", order_by="PlanItem.position")


class PlanExecutor(db.Model):
    __tablename__ = "plan_executors"
    id = db.Column(db.Integer, primary_key=True)
    plan_id = db.Column(db.Integer, db.ForeignKey("plans.id"), nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False)
    topic = db.Column(db.String(255), nullable=False)
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=False)
    status = db.Column(db.String(32), default="pending")

    user = db.relationship("User", backref="plan_assignments")


class PlanItem(db.Model):
    """A measurable monthly-plan line used in reports and performance analytics."""
    __tablename__ = "plan_items"

    id = db.Column(db.Integer, primary_key=True)
    plan_id = db.Column(db.Integer, db.ForeignKey("plans.id"), nullable=False)
    executor_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False)
    position = db.Column(db.Integer, nullable=False, default=1)
    task_text = db.Column(db.Text, nullable=False)
    deadline_kind = db.Column(db.String(32), nullable=False, default="month")
    deadline_date = db.Column(db.Date)
    planned_hours = db.Column(db.Float, nullable=False, default=0)
    status = db.Column(db.String(32), nullable=False, default="planned")
    comment = db.Column(db.Text)
    completed_at = db.Column(db.DateTime)
    execution_approved = db.Column(db.Boolean, nullable=False, default=False)
    execution_approved_by_id = db.Column(db.Integer, db.ForeignKey("users.id"))
    execution_approved_at = db.Column(db.DateTime)
    linked_activity_id = db.Column(db.Integer, db.ForeignKey("activities.id"))
    created_at = db.Column(db.DateTime, default=dt.datetime.utcnow)
    updated_at = db.Column(db.DateTime, default=dt.datetime.utcnow, onupdate=dt.datetime.utcnow)

    executor = db.relationship("User", foreign_keys=[executor_id], backref="plan_items")
    execution_approved_by = db.relationship("User", foreign_keys=[execution_approved_by_id])
    linked_activity = db.relationship("Activity", foreign_keys=[linked_activity_id])


class Protocol(db.Model):
    """Administrative document grouping existing tasks into a protocol."""
    __tablename__ = "protocols"

    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(255), nullable=False)
    protocol_date = db.Column(db.Date, nullable=False, default=dt.date.today)
    description = db.Column(db.Text)
    created_by_id = db.Column(db.Integer, db.ForeignKey("users.id"))
    created_at = db.Column(db.DateTime, nullable=False, default=dt.datetime.utcnow)

    creator = db.relationship("User", foreign_keys=[created_by_id], backref="created_protocols")
    items = db.relationship(
        "ProtocolItem",
        backref="protocol",
        cascade="all, delete-orphan",
        order_by="ProtocolItem.position",
    )


class ProtocolItem(db.Model):
    __tablename__ = "protocol_items"
    __table_args__ = (db.UniqueConstraint("protocol_id", "activity_id", name="uq_protocol_item_activity"),)

    id = db.Column(db.Integer, primary_key=True)
    protocol_id = db.Column(db.Integer, db.ForeignKey("protocols.id"), nullable=False)
    activity_id = db.Column(db.Integer, db.ForeignKey("activities.id"), nullable=False)
    position = db.Column(db.Integer, nullable=False, default=1)
    due_date = db.Column(db.Date)
    comment = db.Column(db.Text)
    created_at = db.Column(db.DateTime, nullable=False, default=dt.datetime.utcnow)

    activity = db.relationship("Activity", foreign_keys=[activity_id])


class DocumentAccess(db.Model):
    __tablename__ = "document_access"
    id = db.Column(db.Integer, primary_key=True)
    doc_type = db.Column(db.String(20), nullable=False)  # 'activity' или 'letter'
    doc_id = db.Column(db.Integer, nullable=False)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False)
    granted_by = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False)
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    user = db.relationship("User", foreign_keys=[user_id])
    granted = db.relationship("User", foreign_keys=[granted_by])
class ChatMessage(db.Model):
    __tablename__ = "chat_messages"
    id = db.Column(db.Integer, primary_key=True)
    sender_id = db.Column(db.Integer, db.ForeignKey("users.id"))
    receiver_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=True)
    text = db.Column(db.Text)
    file_path = db.Column(db.String(255))
    created_at = db.Column(db.DateTime, default=datetime.utcnow)

    sender = db.relationship("User", foreign_keys=[sender_id])
    receiver = db.relationship("User", foreign_keys=[receiver_id])
class DownloadLog(db.Model):
    __tablename__ = "download_logs"

    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False)
    doc_type = db.Column(db.String(50), nullable=False)   # 'letter' или 'activity'
    doc_id = db.Column(db.Integer, nullable=False)
    timestamp = db.Column(db.DateTime, default=dt.datetime.utcnow)
    user = db.relationship("User")

    def __repr__(self):
        return f"<DownloadLog {self.doc_type}#{self.doc_id} by {self.user_id}>"
class GeneralDocument(db.Model):
    __tablename__ = "general_documents"

    id = db.Column(db.Integer, primary_key=True)
    filename = db.Column(db.String(255), nullable=False)
    filepath = db.Column(db.String(255), nullable=False)
    uploaded_by = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False)
    uploaded_at = db.Column(db.DateTime, default=dt.datetime.utcnow)
    description = db.Column(db.Text)
    user = db.relationship("User")

    def __repr__(self):
        return f"<GeneralDocument {self.filename}>"
class SharedLink(db.Model):
    __tablename__ = "shared_links"

    id = db.Column(db.Integer, primary_key=True)
    doc_type = db.Column(db.String(50), nullable=False)
    doc_id = db.Column(db.Integer, nullable=False)
    token = db.Column(db.String(64), unique=True, nullable=False)
    created_by = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False)
    created_at = db.Column(db.DateTime, default=dt.datetime.utcnow)
    expires_at = db.Column(db.DateTime, nullable=True)
    payload = db.Column(db.JSON, default=dict)
    user = db.relationship("User", foreign_keys=[created_by])


class UserNotification(db.Model):
    __tablename__ = "user_notifications"
    id = db.Column(db.Integer, primary_key=True)
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"), nullable=False, index=True)
    sender_id = db.Column(db.Integer, db.ForeignKey("users.id"))
    title = db.Column(db.String(180), nullable=False)
    message = db.Column(db.Text, nullable=False)
    created_at = db.Column(db.DateTime, default=dt.datetime.utcnow, nullable=False)
    read_at = db.Column(db.DateTime)

    recipient = db.relationship("User", foreign_keys=[user_id], backref="notifications")
    sender = db.relationship("User", foreign_keys=[sender_id])
class Task(db.Model):
    __tablename__ = "tasks"
    id = db.Column(db.Integer, primary_key=True)
    title = db.Column(db.String(255), nullable=False)
    description = db.Column(db.Text)
    created_by = db.Column(db.Integer, db.ForeignKey("users.id"))
    start_date = db.Column(db.Date, nullable=False)
    end_date = db.Column(db.Date, nullable=False)
    created_at = db.Column(db.DateTime, default=dt.datetime.utcnow)

    assignees = db.relationship("TaskAssignee", back_populates="task", cascade="all, delete")

    creator = db.relationship("User", foreign_keys=[created_by])

class TaskAssignee(db.Model):
    __tablename__ = "task_assignees"
    id = db.Column(db.Integer, primary_key=True)
    task_id = db.Column(db.Integer, db.ForeignKey("tasks.id"))
    user_id = db.Column(db.Integer, db.ForeignKey("users.id"))
    is_done = db.Column(db.Boolean, default=False)
    done_at = db.Column(db.DateTime)

    task = db.relationship("Task", back_populates="assignees")
    user = db.relationship("User", backref="task_links")


class LetterSource(db.Model):
    __tablename__ = "letter_sources"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), unique=True, nullable=False)
    bin_code = db.Column(db.String(50))
    address = db.Column(db.String(255))
    phone = db.Column(db.String(100))
    email = db.Column(db.String(150))
    contact_person = db.Column(db.String(150))
    notes = db.Column(db.Text)


class ActivityTemplate(db.Model):
    __tablename__ = "activity_templates"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), unique=True, nullable=False)
    title = db.Column(db.String(255), nullable=False)
    description = db.Column(db.Text)
    type_id = db.Column(db.Integer, db.ForeignKey("activity_types.id"))
    direction = db.Column(db.String(255))
    outgoing_document = db.Column(db.String(255))
    duration_days = db.Column(db.Integer)
    created_at = db.Column(db.DateTime, default=dt.datetime.utcnow, nullable=False)

    type = db.relationship("ActivityType")


class MemoTemplate(db.Model):
    __tablename__ = "memo_templates"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), unique=True, nullable=False)
    subject = db.Column(db.String(255), nullable=False)
    body = db.Column(db.Text)
    department_id = db.Column(db.Integer, db.ForeignKey("departments.id"))
    executor_id = db.Column(db.Integer, db.ForeignKey("users.id"))
    created_at = db.Column(db.DateTime, default=dt.datetime.utcnow, nullable=False)

    department = db.relationship("Department", foreign_keys=[department_id])
    executor = db.relationship("User", foreign_keys=[executor_id])

__all__ = [
    "db",
    "Workshop",
    "Department",
    "User",
    "ActivityType",
    "LetterSource",
    "Activity",
    "ActivityLog",
    "ActivityDocument",
    "Letter",
    "LetterRecipient",
    "LetterActivityLink",
    "LetterLetterLink","LetterDocument","ValueTemplate","PlanExecutor","Plan","PlanItem","Protocol","ProtocolItem","DocumentAccess","ChatMessage",
    "DownloadLog","GeneralDocument","SharedLink","Task","TaskAssignee","UserNotification","ActivityHistory",
    "ActivityTemplate", "MemoTemplate"
]
