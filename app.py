# -*- coding: utf-8 -*-
import os
import datetime as dt
from pathlib import Path
from functools import wraps
from models import *
from flask import (
    Flask, request, jsonify, render_template, redirect,
    url_for, flash, send_from_directory, abort
)
from flask_sqlalchemy import SQLAlchemy
from flask_login import (
    LoginManager, UserMixin, login_user, logout_user,
    current_user, login_required
)
from werkzeug.utils import secure_filename
from datetime import datetime
# =========================
# Конфигурация
# =========================
BASE_DIR = Path(__file__).resolve().parent
UPLOAD_DIR = BASE_DIR / "uploads"
UPLOAD_DIR.mkdir(exist_ok=True)

app = Flask(__name__, template_folder=str(BASE_DIR / "templates"))
app.config.update(
    SECRET_KEY=os.getenv("SECRET_KEY", "dev-secret"),
    SQLALCHEMY_DATABASE_URI="sqlite:///app.db",
    SQLALCHEMY_TRACK_MODIFICATIONS=False,
    UPLOAD_FOLDER=str(UPLOAD_DIR),
)

DB_PATH = os.path.join(BASE_DIR, "app.db")

app.config["SQLALCHEMY_DATABASE_URI"] = f"sqlite:///{DB_PATH}"

db.init_app(app)  # <--- вот это ключевая строка


login_manager = LoginManager(app)
login_manager.login_view = "login_page"

# =========================
# Модели (ORM)
# =========================
# Если у тебя есть отдельный models.py — замени этот блок на:

# Я оставляю модели здесь для самодостаточности файла.

from werkzeug.security import generate_password_hash, check_password_hash

# =========================
# Login manager
# =========================
@login_manager.user_loader
def load_user(user_id: str):
    return User.query.get(int(user_id))

# =========================
# Утилиты
# =========================
def role_required(*roles):
    def _decor(fn):
        @wraps(fn)
        def wrapper(*args, **kwargs):
            if not current_user.is_authenticated:
                return login_manager.unauthorized()
            if current_user.role not in roles:
                abort(403)
            return fn(*args, **kwargs)
        return wrapper
    return _decor

def iso_or_none(dtobj):
    return dtobj.isoformat() if dtobj else None

# =========================
# Маршруты: аутентификация и базовые страницы
# =========================
@app.route("/")
def index():
    if current_user.is_authenticated:
        return redirect(url_for("dashboard"))
    return redirect(url_for("login_page"))

@app.get("/login")
def login_page():
    return render_template("login.html", title="Вход")



@app.get("/logout")
@login_required
def logout():
    logout_user()
    return redirect(url_for("login_page"))

@app.get("/dashboard")
@login_required
def dashboard():
    # простая сводка
    my_acts = Activity.query.filter_by(owner_id=current_user.id).count()
    my_letters = Letter.query.filter_by(author_id=current_user.id).count()
    pending_approval = 0
    if current_user.role in ("manager", "admin"):
        # активности, которые завершены и ждут утверждения (условно)
        pending_approval = Activity.query.filter_by(status="done").count()
    return render_template(
        "dashboard.html",
        title="Главная",
        my_acts=my_acts,
        my_letters=my_letters,
        pending_approval=pending_approval
    )
@app.route("/activities/create", methods=["GET", "POST"])
@login_required
def create_activity():
    from datetime import datetime

    types = ActivityType.query.order_by(ActivityType.name).all()

    if request.method == "POST":
        title = request.form.get("title", "").strip()
        description = request.form.get("description", "").strip()
        type_id = request.form.get("type_id")
        start_date = request.form.get("start_date")
        end_date = request.form.get("end_date")

        # Проверка на обязательные поля
        if not title:
            flash("Введите название активности.")
            return redirect(request.url)

        # Создаём новую активность
        a = Activity(
            title=title,
            description=description,
            owner_id=current_user.id,
            type_id=int(type_id) if type_id else None,
            start_date=datetime.strptime(start_date, "%Y-%m-%d") if start_date else datetime.now(),
            end_date=datetime.strptime(end_date, "%Y-%m-%d") if end_date else None,
            status="open"
        )

        db.session.add(a)
        db.session.commit()

        flash("Активность успешно создана.")
        return redirect(url_for("activities_page"))

    return render_template("activity_create.html", title="Создание активности", types=types)

# =========================
# Активности: HTML + JSON
# =========================
@app.get("/activities")
@login_required
def activities_page():
    # Администраторы и руководители видят все активности
    if current_user.role in ("admin", "manager"):
        activities = Activity.query.order_by(Activity.start_date.desc()).all()
    else:
        # Обычный пользователь видит только свои
        activities = (
            Activity.query.filter_by(owner_id=current_user.id)
            .order_by(Activity.start_date.desc())
            .all()
        )

    return render_template("activities.html", title="Активности", activities=activities)

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
        title=title,
        description=description,
        start_date=dt.datetime.utcnow(),
        status="open"
    )
    db.session.add(a)
    db.session.commit()
    flash(f"Активность #{a.id} создана")
    return redirect(url_for("activities_page"))

@app.get("/activities/<int:activity_id>")
@login_required
def activity_view(activity_id: int):
    a = Activity.query.get_or_404(activity_id)
    return render_template("activity_view.html", title=f"Активность #{a.id}", a=a)

@app.post("/activities/<int:activity_id>/logs")
@login_required
def add_activity_log(activity_id: int):
    a = Activity.query.get_or_404(activity_id)

    # Проверка прав
    if a.owner_id != current_user.id and current_user.role not in ("manager", "admin"):
        abort(403)

    text = request.form.get("text", "").strip()
    entry_date_str = request.form.get("entry_date")  # дата из формы

    if not text:
        flash("Текст лог-записи обязателен")
        return redirect(url_for("activity_view", activity_id=activity_id))

    # Преобразуем дату, если указана
    try:
        if entry_date_str:
            entry_date = datetime.strptime(entry_date_str, "%Y-%m-%d")
        else:
            entry_date = datetime.utcnow()
    except ValueError:
        entry_date = datetime.utcnow()

    # Добавляем лог
    al = ActivityLog(activity_id=a.id, text=text, entry_date=entry_date)
    db.session.add(al)
    db.session.commit()

    flash("Лог добавлен")
    return redirect(url_for("activity_view", activity_id=activity_id))
@app.post("/activities/<int:activity_id>/upload")
@login_required
def upload_activity_document(activity_id: int):
    a = Activity.query.get_or_404(activity_id)
    f = request.files.get("file")
    if not f or not f.filename:
        flash("Файл не выбран")
        return redirect(url_for("activity_view", activity_id=activity_id))
    fname = secure_filename(f.filename)
    path = UPLOAD_DIR / fname
    f.save(path)
    ad = ActivityDocument(activity_id=a.id, filename=fname, filepath=str(path))
    db.session.add(ad)
    db.session.commit()
    flash("Документ загружен")
    return redirect(url_for("activity_view", activity_id=activity_id))

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

@app.post("/activities/<int:activity_id>/approve")
@role_required("manager", "admin")
def approve_activity(activity_id: int):
    a = Activity.query.get_or_404(activity_id)
    approve = request.form.get("approve", "true").lower() != "false"
    comment = request.form.get("comment")
    a.status = "approved" if approve else "rejected"
    a.approved_by_id = current_user.id
    a.approved_at = dt.datetime.utcnow()
    a.approve_comment = comment
    db.session.commit()
    flash("Решение по активности сохранено")
    return redirect(url_for("activity_view", activity_id=activity_id))

# =========================
# Письма: HTML
# =========================
from sqlalchemy import or_

@app.get("/letters")
@login_required
def letters_page():
    filter_type = request.args.get("filter", "all")

    q = Letter.query.outerjoin(Letter.recipients).outerjoin(Letter.departments)

    if current_user.role in ("admin", "manager"):
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

    letters = q.distinct().order_by(Letter.created_at.desc()).all()
    return render_template("letters.html", title="Письма", letters=letters, filter_type=filter_type)

@app.get("/letters/new")
@login_required
def new_letter_page():
    deps = Department.query.order_by(Department.name).all()
    users = User.query.filter_by(is_approved=True).order_by(User.full_name).all()
    import datetime

    return render_template(
        "letter_create.html",
        title="Создание письма",
        deps=deps,
        users=users,
        now=datetime.datetime.now
    )

@app.post("/letters")
@login_required
def add_letter():
    subject = request.form.get("subject", "").strip()
    body = request.form.get("body", "").strip()

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

    # --- создаём объект письма ---
    le = Letter(
        subject=subject,
        body=body,
        author_id=current_user.id,
        letter_date=letter_date
    )

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

@app.post("/letters/<int:letter_id>/upload")
@login_required
def upload_letter_doc(letter_id):
    le = Letter.query.get_or_404(letter_id)
    f = request.files["file"]
    if f:
        filename = secure_filename(f.filename)
        path = os.path.join("uploads", filename)
        f.save(path)
        db.session.add(LetterDocument(letter_id=le.id, filename=filename, filepath=path))
        db.session.commit()
        flash("Документ добавлен.")
    return redirect(url_for("letter_view", letter_id=le.id))

@app.get("/letters/<int:letter_id>")
@login_required
def letter_view(letter_id):
    le = Letter.query.get_or_404(letter_id)
    deps = Department.query.all()
    return render_template("letter_view.html", le=le, deps=deps, title=f"Письмо #{le.id}")

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


        # --- создаём активность "Ответ на письмо №..." ---
        a = Activity(
            title=f"Ответ на письмо №{le.id}",
            description=request.form.get("description", "").strip() or (le.body[:300] + "..."),
            owner_id=current_user.id,
            type_id=int(type_id) if type_id else 1000,
            source_letter_id=le.id,
            status="open",
            start_date=datetime.now()
        )

        if start_date:
            a.start_date = datetime.strptime(start_date, "%Y-%m-%d")
        if end_date:
            a.end_date = datetime.strptime(end_date, "%Y-%m-%d")

        db.session.add(a)
        db.session.commit()

        flash(f"Создана активность: ответ на письмо №{le.id}.")
        # ✅ после создания — сразу переходим на просмотр письма
        return redirect(url_for("letter_view", letter_id=le.id))

    types = ActivityType.query.order_by(ActivityType.name).all()
    return render_template("activity_create_from_letter.html", le=le, types=types)

# Связи «из активности → письмо»
@app.post("/activities/<int:activity_id>/create-letter")
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
@app.post("/login")
def login_form():
    username = (request.form.get("username") or "").strip()
    password = request.form.get("password") or ""
    import os
    from sqlalchemy import text, func

    print("=== DEBUG LOGIN ===")
    print("PWD:", os.getcwd())
    print("DB URI:", app.config.get('SQLALCHEMY_DATABASE_URI'))
    print("INPUT username raw repr:", repr(request.form.get("username")))
    print("STRIPPED username:", repr(username))
    print("INPUT password repr:", repr(password))

    # что видит ORM
    try:
        users = User.query.all()
        print("ORM users count:", len(users))
        for u in users:
            print("  ORM:", u.id, repr(u.username), repr(u.full_name), getattr(u, 'is_approved', None))
    except Exception as e:
        print("ORM error:", e)

    # raw SQL check
    try:
        rows = db.session.execute(text("SELECT id, username, full_name, is_approved FROM users")).fetchall()
        print("RAW SQL rows count:", len(rows))
        for r in rows:
            print("  RAW:", r)
    except Exception as e:
        print("RAW SQL error:", e)

    # попробуем поиск с trim
    try:
        user = User.query.filter(func.trim(User.username) == username).first()
        print("filtered(trim) user:", user)
    except Exception as e:
        print("filter trim error:", e)
        user = None

    if not user:
        flash("Пользователь не найден.")
        print("⚠️ user not found in DB (after trim search)")
        return redirect(url_for("login_page"))

    print("found user:", user.id, repr(user.username), "approved:", getattr(user, 'is_approved', None))
    if not getattr(user, 'is_approved', True):
        flash("Ваш аккаунт ожидает подтверждения администратора.")
        return redirect(url_for("login_page"))

    try:
        ok = user.check_password(password)
    except Exception as e:
        print("check_password exception:", e)
        ok = False
    print("check_password returned:", ok)
    if ok:
        login_user(user)
        flash("Вход успешен")
        return redirect(url_for("dashboard"))
    else:
        flash("Неверный пароль.")
        return redirect(url_for("login_page"))
# --- Админка: список пользователей ---
@app.get("/admin/users")
@role_required("admin")
def admin_users():
    users = User.query.order_by(User.created_at.desc()).all()
    return render_template("admin_users.html", title="Пользователи", users=users)

@app.post("/admin/users/<int:user_id>/approve")
@role_required("admin")
def approve_user(user_id):
    u = User.query.get_or_404(user_id)
    u.is_approved = True
    db.session.commit()
    flash(f"Пользователь {u.full_name} подтверждён.")
    return redirect(url_for("admin_users"))

@app.post("/admin/users/<int:user_id>/delete")
@role_required("admin")
def delete_user(user_id):
    u = User.query.get_or_404(user_id)
    db.session.delete(u)
    db.session.commit()
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

    # создаём черновик письма с автозаполнением
    subject = f"Письмо по активности №{activity.id} — {activity.title}"
    body = f"Описание активности:\n{activity.description or ''}\n\nСтатус: {activity.status}"

    return render_template(
        "letter_create.html",
        title="Письмо на основе активности",
        deps=deps,
        users=users,
        subject=subject,
        body=body,
        from_activity=activity,
        now=datetime.now  # ← вот это добавляем
    )
@app.context_processor
def inject_now():
    from datetime import datetime
    return {'now': datetime.now}
if __name__ == "__main__":
    with app.app_context():
        db.create_all()
        print("✔ Проверка таблиц выполнена:", [t.name for t in db.metadata.sorted_tables])
    app.run(debug=True, port=5001)

