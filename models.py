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

class Department(db.Model):
    __tablename__ = "departments"
    id = db.Column(db.Integer, primary_key=True)
    name = db.Column(db.String(200), unique=True, nullable=False)


class User(UserMixin, db.Model):
    __tablename__ = "users"
    id = db.Column(db.Integer, primary_key=True)
    full_name = db.Column(db.String(255), nullable=False)
    username = db.Column(db.String(120), unique=True, nullable=False)
    password_hash = db.Column(db.String(255), nullable=False)
    role = db.Column(db.String(20), nullable=False, default="worker")
    department_id = db.Column(db.Integer, db.ForeignKey("departments.id"))

    # 👇 добавь эти две строки:
    is_approved = db.Column(db.Boolean, default=False, nullable=False)
    created_at = db.Column(db.DateTime, default=dt.datetime.utcnow)

    department = db.relationship("Department", backref="users")

    def set_password(self, pw: str):
        from werkzeug.security import generate_password_hash
        self.password_hash = generate_password_hash(pw)

    def check_password(self, pw: str) -> bool:
        from werkzeug.security import check_password_hash
        return check_password_hash(self.password_hash, pw)

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

    title = db.Column(db.String(255), nullable=False)
    description = db.Column(db.Text)
    start_date = db.Column(db.DateTime, default=dt.datetime.utcnow)
    end_date = db.Column(db.DateTime)
    status = db.Column(db.String(32), default="open")

    approved_by_id = db.Column(db.Integer, db.ForeignKey("users.id"))
    approved_by = db.relationship("User", foreign_keys=[approved_by_id])
    approved_at = db.Column(db.DateTime)
    approve_comment = db.Column(db.Text)

    # связь с письмом (старое поле можно оставить для обратной совместимости)
    source_letter_id = db.Column(db.Integer, db.ForeignKey("letters.id"))

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
    filepath = db.Column(db.String(500), nullable=False)
    uploaded_at = db.Column(db.DateTime, default=dt.datetime.utcnow)

    activity = db.relationship("Activity", backref="documents")


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
    subject = db.Column(db.String(255), nullable=False)
    body = db.Column(db.Text)
    author_id = db.Column(db.Integer, db.ForeignKey("users.id"))

    departments = db.relationship(
        "Department",
        secondary="letter_departments",
        backref="letters"
    )

    letter_date = db.Column(db.Date)
    created_at = db.Column(db.DateTime, default=dt.datetime.utcnow)
    author = db.relationship("User", backref="letters")

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
__all__ = [
    "db",
    "Department",
    "User",
    "ActivityType",
    "Activity",
    "ActivityLog",
    "ActivityDocument",
    "Letter",
    "LetterRecipient",
    "LetterActivityLink",
    "LetterLetterLink","LetterDocument",
]
