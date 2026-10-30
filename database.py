"""SQLite schema upgrades and consistent, recoverable backups."""
import os
import math
import secrets
import sqlite3
from contextlib import contextmanager
from datetime import datetime, timezone
from pathlib import Path
from werkzeug.security import generate_password_hash, check_password_hash

SCHEMA_VERSION = 2


def connect(path):
    conn = sqlite3.connect(path, timeout=15)
    conn.row_factory = sqlite3.Row
    conn.execute('PRAGMA foreign_keys=ON')
    return conn


@contextmanager
def transaction(path, write=False):
    conn = connect(path)
    try:
        conn.execute('BEGIN IMMEDIATE' if write else 'BEGIN')
        yield conn
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def backup_database(path, directory):
    directory = Path(directory)
    directory.mkdir(parents=True, exist_ok=True, mode=0o700)
    stamp = datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%S%fZ')
    target = directory / f'class-{stamp}-{secrets.token_hex(3)}.db'
    fd = os.open(target, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600)
    os.close(fd)
    source = sqlite3.connect(f'file:{Path(path).resolve()}?mode=ro', uri=True)
    dest = sqlite3.connect(target)
    try:
        source.backup(dest)
        if dest.execute('PRAGMA integrity_check').fetchone()[0] != 'ok':
            raise RuntimeError('备份完整性检查失败')
    finally:
        dest.close()
        source.close()
    return target


def init_database(path, backup_dir, instance_dir, initial_password=None):
    """Additive migration, with backup before changing an existing database."""
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    probe = sqlite3.connect(path)
    exists = bool(probe.execute("SELECT 1 FROM sqlite_master WHERE name='users'").fetchone())
    has_meta = probe.execute("SELECT 1 FROM sqlite_master WHERE name='schema_meta'").fetchone()
    version = probe.execute('SELECT version FROM schema_meta').fetchone()[0] if has_meta else 0
    probe.close()
    if version > SCHEMA_VERSION:
        raise RuntimeError('数据库版本高于当前程序，请勿降级运行')
    if version == SCHEMA_VERSION:
        return
    if exists:
        backup_database(path, backup_dir)
    # Foreign keys are disabled only for the transactional legacy exams table rebuild.
    conn = sqlite3.connect(path, timeout=15)
    conn.row_factory = sqlite3.Row
    try:
        conn.execute('BEGIN IMMEDIATE')
        # A second worker may have finished the migration while this one waited.
        if conn.execute("SELECT 1 FROM sqlite_master WHERE name='schema_meta'").fetchone():
            if conn.execute('SELECT version FROM schema_meta').fetchone()[0] == SCHEMA_VERSION:
                conn.rollback()
                return
        definitions = [
            'CREATE TABLE IF NOT EXISTS users (id INTEGER PRIMARY KEY AUTOINCREMENT, username TEXT UNIQUE NOT NULL, password TEXT NOT NULL, email TEXT)',
            'CREATE TABLE IF NOT EXISTS classes (id INTEGER PRIMARY KEY AUTOINCREMENT, teacher_id INTEGER NOT NULL, name TEXT NOT NULL, grade TEXT, school_year TEXT, active INTEGER NOT NULL DEFAULT 1, created_at REAL NOT NULL DEFAULT 0, UNIQUE(teacher_id,name))',
            'CREATE TABLE IF NOT EXISTS students (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, teacher_id INTEGER, class_id INTEGER)',
            'CREATE TABLE IF NOT EXISTS exams (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, teacher_id INTEGER, class_id INTEGER)',
            'CREATE TABLE IF NOT EXISTS exam_subjects (id INTEGER PRIMARY KEY AUTOINCREMENT, exam_id INTEGER, subject TEXT, max_score REAL, pass_score REAL, FOREIGN KEY(exam_id) REFERENCES exams(id))',
            'CREATE TABLE IF NOT EXISTS scores (id INTEGER PRIMARY KEY AUTOINCREMENT, student_id INTEGER, exam_id INTEGER, subject TEXT, score REAL, exam TEXT, FOREIGN KEY(student_id) REFERENCES students(id))',
            'CREATE TABLE IF NOT EXISTS reset_tokens (id INTEGER PRIMARY KEY AUTOINCREMENT, email TEXT, token TEXT, expires_at REAL)',
        ]
        for statement in definitions:
            conn.execute(statement)
        additions = {
            'users': {'email': 'TEXT', 'active': 'INTEGER NOT NULL DEFAULT 1', 'session_version': 'INTEGER NOT NULL DEFAULT 1', 'must_change_password': 'INTEGER NOT NULL DEFAULT 0'},
            'students': {'teacher_id': 'INTEGER', 'class_id': 'INTEGER'},
            'exams': {'teacher_id': 'INTEGER', 'class_id': 'INTEGER'},
            'scores': {'exam_id': 'INTEGER', 'exam': 'TEXT', 'status': "TEXT NOT NULL DEFAULT 'recorded'"},
            'reset_tokens': {'attempts': 'INTEGER NOT NULL DEFAULT 0', 'created_at': 'REAL NOT NULL DEFAULT 0', 'delivered': 'INTEGER NOT NULL DEFAULT 0'},
        }
        for table, columns in additions.items():
            old = {r['name'] for r in conn.execute(f'PRAGMA table_info({table})')}
            for name, kind in columns.items():
                if name not in old:
                    conn.execute(f'ALTER TABLE {table} ADD COLUMN {name} {kind}')
        # Old deployments enforced global exam-name uniqueness, breaking two teachers
        # creating e.g. "期中考试". Preserve IDs and remove just that legacy restriction.
        global_unique = False
        for index in conn.execute('PRAGMA index_list(exams)').fetchall():
            if index['unique']:
                columns = [r['name'] for r in conn.execute('SELECT * FROM pragma_index_info(?)', (index['name'],))]
                global_unique |= columns == ['name']
        if global_unique:
            sequence = conn.execute("SELECT seq FROM sqlite_sequence WHERE name='exams'").fetchone() if conn.execute("SELECT 1 FROM sqlite_master WHERE name='sqlite_sequence'").fetchone() else None
            conn.execute('CREATE TABLE exams_upgrade (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, teacher_id INTEGER, class_id INTEGER)')
            conn.execute('INSERT INTO exams_upgrade SELECT id,name,teacher_id,class_id FROM exams')
            conn.execute('DROP TABLE exams')
            conn.execute('ALTER TABLE exams_upgrade RENAME TO exams')
            if sequence:
                conn.execute("UPDATE sqlite_sequence SET seq=MAX(seq,?) WHERE name='exams'", (sequence[0],))
        if not conn.execute('SELECT 1 FROM users').fetchone():
            password = initial_password or secrets.token_urlsafe(18)
            if len(password) < 8 or password == 'admin123':
                raise RuntimeError('初始管理员密码至少8位，且不能使用旧默认密码')
            conn.execute('INSERT INTO users(username,password,must_change_password) VALUES(?,?,1)', ('admin', generate_password_hash(password)))
            if not initial_password:
                folder = Path(instance_dir)
                folder.mkdir(parents=True, exist_ok=True, mode=0o700)
                target = folder / 'admin-initial-password'
                with open(os.open(target, os.O_WRONLY | os.O_CREAT | os.O_EXCL, 0o600), 'w') as output:
                    output.write(password + '\n')
        admin = conn.execute("SELECT * FROM users WHERE username='admin'").fetchone()
        if admin:
            if check_password_hash(admin['password'], 'admin123'):
                conn.execute('UPDATE users SET must_change_password=1 WHERE id=?', (admin['id'],))
            for table in ['students', 'exams']:
                conn.execute(f'UPDATE {table} SET teacher_id=? WHERE teacher_id IS NULL', (admin['id'],))
        # Give every existing owner a default class, then attach legacy records.
        owners = conn.execute('SELECT id FROM users').fetchall()
        now = datetime.now(timezone.utc).timestamp()
        for owner in owners:
            conn.execute('INSERT OR IGNORE INTO classes(teacher_id,name,grade,school_year,created_at) VALUES(?,?,?,?,?)',
                         (owner['id'], '默认班级', None, None, now))
            class_row = conn.execute('SELECT id FROM classes WHERE teacher_id=? AND name=?', (owner['id'], '默认班级')).fetchone()
            for table in ['students', 'exams']:
                conn.execute(f'UPDATE {table} SET class_id=? WHERE teacher_id=? AND class_id IS NULL', (class_row['id'], owner['id']))
        # Migrate pre-exam_id databases without overwriting any existing associations.
        legacy = conn.execute('SELECT s.*,st.teacher_id FROM scores s JOIN students st ON st.id=s.student_id WHERE s.exam_id IS NULL').fetchall()
        defaults = dict((s, (mx, ps)) for s, mx, ps in default_subjects())
        for score in legacy:
            name = score['exam'] or '历史考试'
            exam = conn.execute('SELECT id FROM exams WHERE name=? AND teacher_id IS ?', (name, score['teacher_id'])).fetchone()
            eid = exam['id'] if exam else conn.execute('INSERT INTO exams(name,teacher_id) VALUES(?,?)', (name, score['teacher_id'])).lastrowid
            conn.execute('UPDATE scores SET exam_id=? WHERE id=?', (eid, score['id']))
            if not conn.execute('SELECT 1 FROM exam_subjects WHERE exam_id=? AND subject=?', (eid, score['subject'])).fetchone():
                mx, ps = defaults.get(score['subject'], (100, 60))
                conn.execute('INSERT INTO exam_subjects(exam_id,subject,max_score,pass_score) VALUES(?,?,?,?)', (eid, score['subject'], mx, ps))
        # Existing reset links become invalid once, so plaintext legacy tokens cannot survive.
        conn.execute('DELETE FROM reset_tokens')
        conn.execute('CREATE TABLE IF NOT EXISTS rate_limits (bucket TEXT PRIMARY KEY, started_at REAL NOT NULL, count INTEGER NOT NULL)')
        for name, table, cols in [
            ('scores_lookup', 'scores', 'student_id,exam_id,subject'),
            ('subjects_lookup', 'exam_subjects', 'exam_id,subject'),
            ('students_owner', 'students', 'teacher_id'),
            ('exams_owner', 'exams', 'teacher_id'),
            ('classes_owner', 'classes', 'teacher_id'),
            ('students_class', 'students', 'class_id'),
            ('exams_class', 'exams', 'class_id'),
        ]:
            conn.execute(f'CREATE INDEX IF NOT EXISTS {name} ON {table}({cols})')
        # Existing invalid legacy records are kept for repair. New writes must be valid.
        for event in ['INSERT', 'UPDATE']:
            conn.execute(f'''CREATE TRIGGER IF NOT EXISTS scores_valid_{event.lower()} BEFORE {event} ON scores
                WHEN NEW.status NOT IN ('recorded','absent') OR NEW.status IS NULL
                  OR (NEW.status='recorded' AND (NEW.score IS NULL OR typeof(NEW.score) NOT IN ('integer','real') OR NEW.score < 0 OR NEW.score >= 1e308))
                  OR (NEW.status='absent' AND NEW.score IS NOT NULL)
                BEGIN SELECT RAISE(ABORT, 'invalid score'); END''')
        conn.execute('CREATE TABLE IF NOT EXISTS schema_meta (version INTEGER NOT NULL)')
        conn.execute('DELETE FROM schema_meta')
        conn.execute('INSERT INTO schema_meta VALUES(?)', (SCHEMA_VERSION,))
        conn.commit()
    except Exception:
        conn.rollback()
        raise
    finally:
        conn.close()


def default_subjects():
    return [(s, 150, 90) for s in ['语文', '数学', '英语']] + [(s, 100, 60) for s in ['物理', '化学', '生物']]


def finite_number(value):
    try:
        number = float(value)
    except (TypeError, ValueError, OverflowError):
        raise ValueError('请输入有效的数字') from None
    if not math.isfinite(number) or abs(number) >= 1e308:
        raise ValueError('请输入有限的有效数字，不能使用 NaN 或无穷大')
    return number
