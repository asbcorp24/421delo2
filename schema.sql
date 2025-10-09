-- Включаем внешние ключи
PRAGMA foreign_keys = ON;

-- =========================
-- Справочники
-- =========================
CREATE TABLE IF NOT EXISTS departments (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  name          TEXT NOT NULL UNIQUE
);

CREATE TABLE IF NOT EXISTS users (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  full_name     TEXT NOT NULL,
  username      TEXT NOT NULL UNIQUE,
  password_hash TEXT NOT NULL,
  role          TEXT NOT NULL DEFAULT 'worker' CHECK (role IN ('worker','manager','admin')),
  department_id INTEGER,
  is_approved   INTEGER NOT NULL DEFAULT 0 CHECK (is_approved IN (0,1)),
  created_at    TEXT DEFAULT (datetime('now')),
  FOREIGN KEY (department_id) REFERENCES departments(id) ON UPDATE CASCADE ON DELETE SET NULL
);
CREATE INDEX IF NOT EXISTS idx_users_department ON users(department_id);
CREATE INDEX IF NOT EXISTS idx_users_username   ON users(username);

CREATE TABLE IF NOT EXISTS activity_types (
  id   INTEGER PRIMARY KEY AUTOINCREMENT,
  name TEXT NOT NULL UNIQUE
);

-- =========================
-- Активности
-- =========================
CREATE TABLE IF NOT EXISTS activities (
  id              INTEGER PRIMARY KEY AUTOINCREMENT,
  type_id         INTEGER NOT NULL,
  owner_id        INTEGER NOT NULL,            -- исполнитель (пользователь активности)
  title           TEXT NOT NULL,
  description     TEXT,
  start_date      TEXT NOT NULL DEFAULT (datetime('now')),
  end_date        TEXT,
  status          TEXT NOT NULL DEFAULT 'open' CHECK (status IN ('open','in_progress','done','approved','rejected')),
  approved_by_id  INTEGER,
  approved_at     TEXT,
  approve_comment TEXT,

  FOREIGN KEY (type_id)        REFERENCES activity_types(id) ON UPDATE CASCADE ON DELETE RESTRICT,
  FOREIGN KEY (owner_id)       REFERENCES users(id)          ON UPDATE CASCADE ON DELETE RESTRICT,
  FOREIGN KEY (approved_by_id) REFERENCES users(id)          ON UPDATE CASCADE ON DELETE SET NULL
);

CREATE INDEX IF NOT EXISTS idx_activities_type       ON activities(type_id);
CREATE INDEX IF NOT EXISTS idx_activities_owner      ON activities(owner_id);
CREATE INDEX IF NOT EXISTS idx_activities_status     ON activities(status);
CREATE INDEX IF NOT EXISTS idx_activities_start_date ON activities(start_date);

CREATE TABLE IF NOT EXISTS activity_logs (
  id          INTEGER PRIMARY KEY AUTOINCREMENT,
  activity_id INTEGER NOT NULL,
  entry_date  TEXT NOT NULL DEFAULT (datetime('now')),
  text        TEXT NOT NULL,
  FOREIGN KEY (activity_id) REFERENCES activities(id) ON UPDATE CASCADE ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_activity_logs_activity ON activity_logs(activity_id);
CREATE INDEX IF NOT EXISTS idx_activity_logs_date     ON activity_logs(entry_date);

CREATE TABLE IF NOT EXISTS activity_documents (
  id          INTEGER PRIMARY KEY AUTOINCREMENT,
  activity_id INTEGER NOT NULL,
  filename    TEXT NOT NULL,
  filepath    TEXT NOT NULL,
  uploaded_at TEXT NOT NULL DEFAULT (datetime('now')),
  FOREIGN KEY (activity_id) REFERENCES activities(id) ON UPDATE CASCADE ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_activity_docs_activity ON activity_documents(activity_id);

-- =========================
-- Письма
-- =========================
CREATE TABLE IF NOT EXISTS letters (
  id             INTEGER PRIMARY KEY AUTOINCREMENT,
  author_id      INTEGER NOT NULL,
  subject        TEXT NOT NULL,
  body           TEXT NOT NULL,
  created_at     TEXT NOT NULL DEFAULT (datetime('now')),
  signed_result  TEXT, -- «подпись по письму — результат работы»
  FOREIGN KEY (author_id) REFERENCES users(id) ON UPDATE CASCADE ON DELETE RESTRICT
);

CREATE INDEX IF NOT EXISTS idx_letters_author    ON letters(author_id);
CREATE INDEX IF NOT EXISTS idx_letters_created   ON letters(created_at);

-- Получатели писем: конкретный пользователь ИЛИ отдел (одно из двух полей может быть заполнено)
CREATE TABLE IF NOT EXISTS letter_recipients (
  id            INTEGER PRIMARY KEY AUTOINCREMENT,
  letter_id     INTEGER NOT NULL,
  user_id       INTEGER,
  department_id INTEGER,
  role          TEXT NOT NULL DEFAULT 'cc' CHECK (role IN ('executor','cc')),
  is_read       INTEGER NOT NULL DEFAULT 0 CHECK (is_read IN (0,1)),
  read_at       TEXT,

  FOREIGN KEY (letter_id)     REFERENCES letters(id)     ON UPDATE CASCADE ON DELETE CASCADE,
  FOREIGN KEY (user_id)       REFERENCES users(id)       ON UPDATE CASCADE ON DELETE CASCADE,
  FOREIGN KEY (department_id) REFERENCES departments(id) ON UPDATE CASCADE ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_letterrec_letter     ON letter_recipients(letter_id);
CREATE INDEX IF NOT EXISTS idx_letterrec_user       ON letter_recipients(user_id);
CREATE INDEX IF NOT EXISTS idx_letterrec_department ON letter_recipients(department_id);
CREATE INDEX IF NOT EXISTS idx_letterrec_role       ON letter_recipients(role);

-- Связи письмо ↔ активность (многие-ко-многим)
CREATE TABLE IF NOT EXISTS letter_activity_link (
  letter_id   INTEGER NOT NULL,
  activity_id INTEGER NOT NULL,
  PRIMARY KEY (letter_id, activity_id),
  FOREIGN KEY (letter_id)   REFERENCES letters(id)    ON UPDATE CASCADE ON DELETE CASCADE,
  FOREIGN KEY (activity_id) REFERENCES activities(id) ON UPDATE CASCADE ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_lal_letter   ON letter_activity_link(letter_id);
CREATE INDEX IF NOT EXISTS idx_lal_activity ON letter_activity_link(activity_id);

-- Связи письмо ↔ письмо (многие-ко-многим)
CREATE TABLE IF NOT EXISTS letter_letter_link (
  src_letter_id INTEGER NOT NULL,
  dst_letter_id INTEGER NOT NULL,
  PRIMARY KEY (src_letter_id, dst_letter_id),
  FOREIGN KEY (src_letter_id) REFERENCES letters(id) ON UPDATE CASCADE ON DELETE CASCADE,
  FOREIGN KEY (dst_letter_id) REFERENCES letters(id) ON UPDATE CASCADE ON DELETE CASCADE
);

CREATE INDEX IF NOT EXISTS idx_lll_src ON letter_letter_link(src_letter_id);
CREATE INDEX IF NOT EXISTS idx_lll_dst ON letter_letter_link(dst_letter_id);
