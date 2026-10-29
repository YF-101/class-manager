# ===== 班主任学生成绩系统 —— 第2步优化版：模板化 + 现代化样式 =====
# 用 Flask + Jinja 模板(templates/)渲染页面，样式集中在 base.html

import os
import re
import io
import json
import base64
import time
import random
import smtplib
import sqlite3
from functools import wraps
from urllib import request as urlreq
from email.mime.text import MIMEText
from email.utils import formataddr
from flask import (Flask, request, render_template, session,
                   redirect, send_file)
from werkzeug.security import generate_password_hash, check_password_hash
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "student-grade-system-secret-2026")
# 数据库用绝对路径（基于 app.py 所在目录），换任何运行目录都找得到
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB = os.path.join(BASE_DIR, "class.db")


@app.template_filter("num")
def num_filter(x):
    """数字显示：整数不带小数点(150)，小数才保留(88.5)"""
    try:
        x = float(x)
    except (TypeError, ValueError):
        return x
    return int(x) if x.is_integer() else x

def default_subjects():
    """山东高考模板：返回 (科目, 满分, 及格线)，用于一键生成考试科目配置"""
    return [(s, 150, 90) for s in ["语文", "数学", "英语"]] + \
           [(s, 100, 60) for s in ["物理", "化学", "生物"]]


def get_db():
    conn = sqlite3.connect(DB)
    conn.row_factory = sqlite3.Row
    return conn


def send_email(to_addr, subject, body):
    """统一发信入口：配置了 Mailjet 就用 Mailjet（公网），否则用 SMTP（本地）"""
    if os.environ.get("MAILJET_KEY"):
        return send_email_via_mailjet(to_addr, subject, body)
    return send_email_via_smtp(to_addr, subject, body)


def send_email_via_mailjet(to_addr, subject, body):
    """通过 Mailjet HTTP API 发邮件（本地/公网都能用，公网在白名单内）"""
    key = os.environ.get("MAILJET_KEY")
    secret = os.environ.get("MAILJET_SECRET")
    sender = os.environ.get("MAILJET_SENDER")     # 已验证的发件邮箱
    if not (key and secret and sender):
        raise RuntimeError("未配置 Mailjet：需要 MAILJET_KEY / MAILJET_SECRET / MAILJET_SENDER")
    payload = {"Messages": [{
        "From": {"Email": sender, "Name": "班主任成绩系统"},
        "To": [{"Email": to_addr}],
        "Subject": subject, "TextPart": body}]}
    auth = base64.b64encode(f"{key}:{secret}".encode()).decode()
    req = urlreq.Request(
        "https://api.mailjet.com/v3.1/send",
        data=json.dumps(payload).encode("utf-8"),
        headers={"Content-Type": "application/json",
                 "Authorization": "Basic " + auth})
    with urlreq.urlopen(req) as resp:
        return resp.status


def send_email_via_smtp(to_addr, subject, body):
    """通过 SMTP 发邮件（支持 QQ/163/Gmail 等任意邮箱）。
    账号、授权码、服务器地址都从环境变量读取，不写死在代码里
    （防止传到 GitHub 泄露）。"""
    user = os.environ.get("SMTP_USER")      # 发件邮箱
    pwd = os.environ.get("SMTP_PASS")       # 授权码 / 应用专用密码
    host = os.environ.get("SMTP_HOST", "smtp.qq.com")   # SMTP 服务器
    port = int(os.environ.get("SMTP_PORT", "465"))      # 端口
    if not user or not pwd:
        raise RuntimeError("未配置发件邮箱：请先设置环境变量 SMTP_USER 和 SMTP_PASS")
    msg = MIMEText(body, "plain", "utf-8")
    msg["Subject"] = subject
    msg["From"] = formataddr(("班主任成绩系统", user))   # 发件人显示成名称
    msg["To"] = to_addr
    with smtplib.SMTP_SSL(host, port) as s:
        s.login(user, pwd)
        s.sendmail(user, [to_addr], msg.as_string())


def get_student_scores(conn, sid):
    """查某学生全部成绩，并带上考试名、该科满分、及格线（配置来自 exam_subjects）"""
    return conn.execute(
        "SELECT s.id, s.subject, s.score, s.exam_id, e.name exam_name, "
        "es.max_score, es.pass_score FROM scores s "
        "JOIN exams e ON s.exam_id=e.id "
        "LEFT JOIN exam_subjects es ON es.exam_id=s.exam_id AND es.subject=s.subject "
        "WHERE s.student_id=? ORDER BY s.exam_id, s.subject", (sid,)).fetchall()


def login_required(f):
    """装饰器：没登录就先跳转到登录页，保护需要登录才能访问的页面"""
    @wraps(f)
    def wrapper(*args, **kwargs):
        if "user" not in session:
            return redirect("/login")
        return f(*args, **kwargs)
    return wrapper


def admin_required(f):
    """装饰器：只有 admin 管理员能访问（用户管理等）"""
    @wraps(f)
    def wrapper(*args, **kwargs):
        if session.get("user") != "admin":
            return redirect("/")
        return f(*args, **kwargs)
    return wrapper


def get_my_student(conn, sid):
    """查一个学生，但只能查当前老师自己的（admin 不受限）。查不到返回 None。"""
    uid = session.get("uid")
    if session.get("user") == "admin":
        return conn.execute("SELECT * FROM students WHERE id=?", (sid,)).fetchone()
    return conn.execute("SELECT * FROM students WHERE id=? AND teacher_id=?",
                        (sid, uid)).fetchone()


def get_my_exam(conn, exam_id):
    """查一场考试，但只能查当前老师自己的（admin 不受限）。"""
    uid = session.get("uid")
    if session.get("user") == "admin":
        return conn.execute("SELECT * FROM exams WHERE id=?", (exam_id,)).fetchone()
    return conn.execute("SELECT * FROM exams WHERE id=? AND teacher_id=?",
                        (exam_id, uid)).fetchone()


# 启动时建表（若没有）
conn = get_db()
conn.execute("CREATE TABLE IF NOT EXISTS students ("
             "id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, teacher_id INTEGER)")
conn.execute("CREATE TABLE IF NOT EXISTS scores ("
             "id INTEGER PRIMARY KEY AUTOINCREMENT, "
             "student_id INTEGER, exam_id INTEGER, subject TEXT, score REAL, "
             "FOREIGN KEY(student_id) REFERENCES students(id))")
conn.execute("CREATE TABLE IF NOT EXISTS exams ("
             "id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT NOT NULL, teacher_id INTEGER)")
conn.execute("CREATE TABLE IF NOT EXISTS exam_subjects ("
             "id INTEGER PRIMARY KEY AUTOINCREMENT, exam_id INTEGER, "
             "subject TEXT, max_score REAL, pass_score REAL, "
             "FOREIGN KEY(exam_id) REFERENCES exams(id))")
conn.execute("CREATE TABLE IF NOT EXISTS users ("
             "id INTEGER PRIMARY KEY AUTOINCREMENT, "
             "username TEXT UNIQUE NOT NULL, password TEXT NOT NULL, email TEXT)")
conn.execute("CREATE TABLE IF NOT EXISTS reset_tokens ("
             "id INTEGER PRIMARY KEY AUTOINCREMENT, email TEXT, "
             "token TEXT, expires_at REAL)")

# 旧库迁移：给 students / exams 补 teacher_id 列（已存在的列会报错，忽略即可）
for _tbl in ("students", "exams"):
    try:
        conn.execute(f"ALTER TABLE {_tbl} ADD COLUMN teacher_id INTEGER")
    except sqlite3.OperationalError:
        pass  # 列已存在
# 旧数据全部归到 admin（id=1）
conn.execute("UPDATE students SET teacher_id=1 WHERE teacher_id IS NULL")
conn.execute("UPDATE exams SET teacher_id=1 WHERE teacher_id IS NULL")

# 第一次运行没有账号，创建一个默认管理员（密码哈希后存储）
if conn.execute("SELECT COUNT(*) FROM users").fetchone()[0] == 0:
    conn.execute("INSERT INTO users (username, password) VALUES (?,?)",
                 ("admin", generate_password_hash("admin123")))
conn.commit()
conn.close()


# ----- 登录 / 登出 -----
@app.route("/login", methods=["GET", "POST"])
def login():
    error = None
    info = None
    if request.args.get("registered"):
        info = "注册成功，请登录"
    if request.args.get("reset"):
        info = "密码已重置，请用新密码登录"
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        conn = get_db()
        u = conn.execute("SELECT * FROM users WHERE username=?", (username,)).fetchone()
        conn.close()
        if u and check_password_hash(u["password"], password):
            session["user"] = username        # 登录成功，记录到 session
            session["uid"] = u["id"]         # 记录用户 id，用于数据隔离
            return redirect("/")
        error = "用户名或密码错误"
    return render_template("login.html", error=error, info=info)


@app.route("/logout")
def logout():
    session.pop("user", None)
    return redirect("/login")


# ----- 公开注册：其他班主任可自行注册账号 -----
@app.route("/register", methods=["GET", "POST"])
def register():
    error = None
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        confirm = request.form.get("confirm", "")
        email = request.form.get("email", "").strip()
        conn = get_db()
        if not username:
            error = "用户名不能为空"
        elif len(password) < 6:
            error = "密码至少要 6 位"
        elif password != confirm:
            error = "两次输入的密码不一致"
        elif email and not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email):
            error = "邮箱格式不正确"
        elif conn.execute("SELECT 1 FROM users WHERE username=?",
                          (username,)).fetchone():
            error = "用户名已被占用，换一个"
        elif email and conn.execute("SELECT 1 FROM users WHERE email=?",
                                    (email,)).fetchone():
            error = "该邮箱已被注册"
        else:
            conn.execute("INSERT INTO users (username, password, email) VALUES (?,?,?)",
                         (username, generate_password_hash(password),
                          email or None))
            conn.commit()
            conn.close()
            return redirect("/login?registered=1")
        conn.close()
    return render_template("register.html", error=error)


# ----- 忘记密码：输入邮箱 → 发验证码 -----
@app.route("/forgot_password", methods=["GET", "POST"])
def forgot_password():
    error = None
    info = None
    if request.method == "POST":
        email = request.form.get("email", "").strip()
        conn = get_db()
        u = conn.execute("SELECT * FROM users WHERE email=?", (email,)).fetchone()
        if not u:
            error = "该邮箱没有注册过账号"
        else:
            token = f"{random.randint(0, 999999):06d}"      # 6位随机验证码
            expires = time.time() + 600                      # 10分钟有效
            conn.execute("DELETE FROM reset_tokens WHERE email=?", (email,))
            conn.execute("INSERT INTO reset_tokens (email, token, expires_at) "
                         "VALUES (?,?,?)", (email, token, expires))
            conn.commit()
            try:
                send_email(email, "成绩系统 · 密码重置验证码",
                           f"您正在重置密码，验证码：{token}\n"
                           f"验证码 10 分钟内有效。若非本人操作，请忽略此邮件。")
                info = f"验证码已发送到 {email}，请去邮箱查收（注意垃圾箱）"
            except Exception as e:
                error = f"邮件发送失败：{e}"
        conn.close()
    return render_template("forgot_password.html", error=error, info=info)


# ----- 重置密码：输入邮箱 + 验证码 + 新密码 -----
@app.route("/reset_password", methods=["GET", "POST"])
def reset_password():
    error = None
    if request.method == "POST":
        email = request.form.get("email", "").strip()
        token = request.form.get("token", "").strip()
        new = request.form.get("new_password", "")
        confirm = request.form.get("confirm", "")
        conn = get_db()
        r = conn.execute("SELECT * FROM reset_tokens WHERE email=? "
                         "ORDER BY id DESC LIMIT 1", (email,)).fetchone()
        if not r or r["token"] != token:
            error = "验证码不正确"
        elif time.time() > r["expires_at"]:
            error = "验证码已过期，请重新获取"
        elif len(new) < 6:
            error = "新密码至少要 6 位"
        elif new != confirm:
            error = "两次输入的新密码不一致"
        else:
            conn.execute("UPDATE users SET password=? WHERE email=?",
                         (generate_password_hash(new), email))
            conn.execute("DELETE FROM reset_tokens WHERE email=?", (email,))
            conn.commit()
            conn.close()
            return redirect("/login?reset=1")
        conn.close()
    return render_template("reset_password.html", error=error)


# ----- 修改自己的密码 -----
@app.route("/change_password", methods=["GET", "POST"])
@login_required
def change_password():
    error = None
    info = None
    if request.method == "POST":
        old = request.form.get("old_password", "")
        new = request.form.get("new_password", "")
        confirm = request.form.get("confirm_password", "")
        conn = get_db()
        u = conn.execute("SELECT * FROM users WHERE username=?",
                         (session["user"],)).fetchone()
        if not check_password_hash(u["password"], old):
            error = "旧密码不正确"
        elif len(new) < 6:
            error = "新密码至少要 6 位"
        elif new != confirm:
            error = "两次输入的新密码不一致"
        else:
            conn.execute("UPDATE users SET password=? WHERE id=?",
                         (generate_password_hash(new), u["id"]))
            conn.commit()
            info = "密码修改成功！"
        conn.close()
    return render_template("change_password.html", error=error, info=info)


# ----- 用户管理（仅管理员） -----
@app.route("/users")
@login_required
@admin_required
def user_list():
    conn = get_db()
    users = conn.execute("SELECT id, username FROM users ORDER BY id").fetchall()
    conn.close()
    return render_template("users.html", users=users, error=None)


@app.route("/add_user", methods=["POST"])
@login_required
@admin_required
def add_user():
    username = request.form.get("username", "").strip()
    password = request.form.get("password", "")
    conn = get_db()
    error = None
    if not username or len(password) < 6:
        error = "用户名不能为空，密码至少 6 位"
    elif conn.execute("SELECT 1 FROM users WHERE username=?",
                      (username,)).fetchone():
        error = f"用户名 {username} 已存在"
    else:
        conn.execute("INSERT INTO users (username, password) VALUES (?,?)",
                     (username, generate_password_hash(password)))
        conn.commit()
    users = conn.execute("SELECT id, username FROM users ORDER BY id").fetchall()
    conn.close()
    return render_template("users.html", users=users, error=error)


@app.route("/delete_user/<int:uid>", methods=["POST"])
@login_required
@admin_required
def delete_user(uid):
    conn = get_db()
    u = conn.execute("SELECT username FROM users WHERE id=?", (uid,)).fetchone()
    if u and u["username"] != "admin":      # 不能删除管理员 admin
        conn.execute("DELETE FROM users WHERE id=?", (uid,))
        conn.commit()
    conn.close()
    return redirect("/users")


# ----- 首页：学生列表（只看自己班的） -----
@app.route("/")
@login_required
def index():
    conn = get_db()
    uid = session.get("uid")
    if session.get("user") == "admin":
        rows = conn.execute("SELECT * FROM students ORDER BY id").fetchall()
    else:
        rows = conn.execute("SELECT * FROM students WHERE teacher_id=? ORDER BY id",
                            (uid,)).fetchall()
    conn.close()
    return render_template("index.html", students=rows)


# ----- 添加学生 -----
@app.route("/add", methods=["POST"])
@login_required
def add_student():
    name = request.form.get("name", "").strip()
    if name:
        conn = get_db()
        uid = session.get("uid")
        # 重名检查只在自己班范围内
        if session.get("user") == "admin":
            exists = conn.execute("SELECT id FROM students WHERE name=?", (name,)).fetchone()
        else:
            exists = conn.execute("SELECT id FROM students WHERE name=? AND teacher_id=?",
                                  (name, uid)).fetchone()
        if exists:
            conn.close()
            return f'<script>alert("已存在同名学生：{name}"); location.href="/"</script>'
        conn.execute("INSERT INTO students (name, teacher_id) VALUES (?,?)",
                     (name, uid if session.get("user") != "admin" else 1))
        conn.commit()
        conn.close()
    return '<script>location.href="/"</script>'


# ----- 录入成绩页（GET 显示表单，POST 提交保存） -----
@app.route("/score/<int:sid>", methods=["GET", "POST"])
@login_required
def score_page(sid):
    conn = get_db()
    stu = get_my_student(conn, sid)
    if not stu:
        conn.close()
        return "没有这个学生（或不是你班的）"

    if session.get("user") == "admin":
        all_exams = conn.execute("SELECT id, name FROM exams ORDER BY id").fetchall()
    else:
        all_exams = conn.execute("SELECT id, name FROM exams WHERE teacher_id=? ORDER BY id",
                                 (session.get("uid"),)).fetchall()
    # 当前选中的考试（URL 或表单）
    exam_id = request.values.get("exam_id", type=int)
    if not exam_id and all_exams:
        exam_id = all_exams[0]["id"]
    cur_exam = conn.execute("SELECT * FROM exams WHERE id=?",
                            (exam_id,)).fetchone() if exam_id else None
    cur_subs = conn.execute(
        "SELECT subject, max_score, pass_score FROM exam_subjects "
        "WHERE exam_id=? ORDER BY id", (exam_id,)).fetchall() if exam_id else []

    error = None
    if request.method == "POST":
        post_exam_id = request.form.get("exam_id", type=int)
        subject = request.form.get("subject", "").strip()
        score_raw = request.form.get("score", "").strip()
        cfgrow = conn.execute(
            "SELECT max_score, pass_score FROM exam_subjects "
            "WHERE exam_id=? AND subject=?", (post_exam_id, subject)).fetchone()
        try:
            score_val = float(score_raw)
        except ValueError:
            error = "成绩必须是数字"
        else:
            if not cfgrow:
                error = "该考试没有配置这个科目，请先到考试管理里添加"
            elif score_val < 0 or score_val > cfgrow["max_score"]:
                mx = num_filter(cfgrow["max_score"])
                error = f"{subject} 满分 {mx} 分，请输入 0 ~ {mx}"
            else:
                # 同一考试同一科目已录过 → 覆盖；否则新增
                ex = conn.execute(
                    "SELECT id FROM scores WHERE student_id=? AND exam_id=? AND subject=?",
                    (sid, post_exam_id, subject)).fetchone()
                if ex:
                    conn.execute("UPDATE scores SET score=? WHERE id=?",
                                 (score_val, ex["id"]))
                else:
                    conn.execute(
                        "INSERT INTO scores (student_id, exam_id, subject, exam, score) "
                        "VALUES (?,?,?,?,?)",
                        (sid, post_exam_id, subject,
                         conn.execute("SELECT name FROM exams WHERE id=?",
                                      (post_exam_id,)).fetchone()["name"],
                         score_val))
                conn.commit()
        # 表单提交后，页面停在刚操作的考试
        exam_id = post_exam_id
        cur_exam = conn.execute("SELECT * FROM exams WHERE id=?",
                                (exam_id,)).fetchone()
        cur_subs = conn.execute(
            "SELECT subject, max_score, pass_score FROM exam_subjects "
            "WHERE exam_id=? ORDER BY id", (exam_id,)).fetchall()

    rows = get_student_scores(conn, sid)
    conn.close()
    return render_template("score.html", stu=stu, rows=rows,
                           all_exams=all_exams, cur_exam=cur_exam,
                           cur_subs=cur_subs, error=error)


# ----- 学生详情页：显示各科成绩 + 不及格标红 -----
@app.route("/student/<int:sid>")
@login_required
def student_page(sid):
    conn = get_db()
    stu = get_my_student(conn, sid)
    if not stu:
        conn.close()
        return "没有这个学生（或不是你班的）"
    rows = get_student_scores(conn, sid)
    if session.get("user") == "admin":
        all_exams = [dict(e) for e in conn.execute("SELECT id, name FROM exams ORDER BY id").fetchall()]
    else:
        all_exams = [dict(e) for e in conn.execute(
            "SELECT id, name FROM exams WHERE teacher_id=? ORDER BY id",
            (session.get("uid"),)).fetchall()]

    # 折线图：科目 -> {考试id: 分数}
    chart = {}
    for r in rows:
        chart.setdefault(r["subject"], {})[str(r["exam_id"])] = r["score"]

    conn.close()
    return render_template("student.html", stu=stu, rows=rows,
                           all_exams=all_exams, chart=chart)


@app.route("/help")
@login_required
def help_page():
    return render_template("help.html")


# ===== 考试管理：自定义考试名称 + 每场考试的科目/满分/及格线 =====

@app.route("/exams")
@login_required
def exam_list():
    conn = get_db()
    uid = session.get("uid")
    if session.get("user") == "admin":
        exams = conn.execute("SELECT id, name FROM exams ORDER BY id").fetchall()
    else:
        exams = conn.execute("SELECT id, name FROM exams WHERE teacher_id=? ORDER BY id",
                            (uid,)).fetchall()
    # 每场考试录了多少条成绩
    counts = {}
    for r in conn.execute("SELECT exam_id, COUNT(*) c FROM scores GROUP BY exam_id"):
        counts[r["exam_id"]] = r["c"]
    conn.close()
    return render_template("exams.html", exams=exams, counts=counts, error=None)


@app.route("/create_exam", methods=["POST"])
@login_required
def create_exam():
    name = request.form.get("name", "").strip()
    conn = get_db()
    error = None
    uid = session.get("uid")
    if not name:
        error = "考试名称不能为空"
    elif session.get("user") == "admin":
        if conn.execute("SELECT 1 FROM exams WHERE name=?", (name,)).fetchone():
            error = f"已存在考试：{name}"
    else:
        if conn.execute("SELECT 1 FROM exams WHERE name=? AND teacher_id=?",
                        (name, uid)).fetchone():
            error = f"已存在考试：{name}"
    if error is None:
        exam_id = conn.execute("INSERT INTO exams (name, teacher_id) VALUES (?,?)",
                               (name, uid if session.get("user") != "admin" else 1)
                               ).lastrowid
        # 新考试自动按山东高考模板生成科目，省得手动一个个加
        for subj, mx, ps in default_subjects():
            conn.execute("INSERT INTO exam_subjects (exam_id,subject,max_score,pass_score) "
                         "VALUES (?,?,?,?)", (exam_id, subj, mx, ps))
        conn.commit()
        conn.close()
        return redirect(f"/exam/{exam_id}")
    if session.get("user") == "admin":
        exams = conn.execute("SELECT id, name FROM exams ORDER BY id").fetchall()
    else:
        exams = conn.execute("SELECT id, name FROM exams WHERE teacher_id=? ORDER BY id",
                            (uid,)).fetchall()
    counts = {}
    for r in conn.execute("SELECT exam_id, COUNT(*) c FROM scores GROUP BY exam_id"):
        counts[r["exam_id"]] = r["c"]
    conn.close()
    return render_template("exams.html", exams=exams, counts=counts, error=error)


@app.route("/delete_exam/<int:exam_id>", methods=["POST"])
@login_required
def delete_exam(exam_id):
    conn = get_db()
    if get_my_exam(conn, exam_id):
        conn.execute("DELETE FROM exam_subjects WHERE exam_id=?", (exam_id,))
        conn.execute("DELETE FROM scores WHERE exam_id=?", (exam_id,))
        conn.execute("DELETE FROM exams WHERE id=?", (exam_id,))
        conn.commit()
    conn.close()
    return redirect("/exams")


@app.route("/exam/<int:exam_id>")
@login_required
def exam_config(exam_id):
    conn = get_db()
    exam = get_my_exam(conn, exam_id)
    if not exam:
        conn.close()
        return "没有这场考试（或不是你的）"
    subs = conn.execute("SELECT * FROM exam_subjects WHERE exam_id=? ORDER BY id",
                        (exam_id,)).fetchall()
    conn.close()
    return render_template("exam_config.html", exam=exam, subs=subs, error=None)


@app.route("/exam/<int:exam_id>/add_subject", methods=["POST"])
@login_required
def add_subject(exam_id):
    subject = request.form.get("subject", "").strip()
    max_s = request.form.get("max_score", "").strip()
    pass_s = request.form.get("pass_score", "").strip()
    conn = get_db()
    exam = get_my_exam(conn, exam_id)
    if not exam:
        conn.close()
        return "没有这场考试（或不是你的）"
    error = None
    try:
        max_s = float(max_s)
        pass_s = float(pass_s) if pass_s else max_s * 0.6
    except ValueError:
        error = "满分和及格线必须是数字"
    else:
        if not subject:
            error = "科目名不能为空"
        elif max_s <= 0 or pass_s < 0 or pass_s > max_s:
            error = "满分要大于0，及格线要在 0 ~ 满分 之间"
        elif conn.execute("SELECT 1 FROM exam_subjects WHERE exam_id=? AND subject=?",
                          (exam_id, subject)).fetchone():
            error = f"该考试已有科目：{subject}"
        else:
            conn.execute("INSERT INTO exam_subjects (exam_id,subject,max_score,pass_score) "
                         "VALUES (?,?,?,?)", (exam_id, subject, max_s, pass_s))
            conn.commit()
    exam = conn.execute("SELECT * FROM exams WHERE id=?", (exam_id,)).fetchone()
    subs = conn.execute("SELECT * FROM exam_subjects WHERE exam_id=? ORDER BY id",
                        (exam_id,)).fetchall()
    conn.close()
    return render_template("exam_config.html", exam=exam, subs=subs, error=error)


@app.route("/exam/<int:exam_id>/update_subject/<int:cfg_id>", methods=["POST"])
@login_required
def update_subject(exam_id, cfg_id):
    max_s = request.form.get("max_score", "").strip()
    pass_s = request.form.get("pass_score", "").strip()
    conn = get_db()
    exam = get_my_exam(conn, exam_id)
    if not exam:
        conn.close()
        return "没有这场考试（或不是你的）"
    error = None
    try:
        max_s = float(max_s)
        pass_s = float(pass_s)
    except ValueError:
        error = "满分和及格线必须是数字"
    else:
        if max_s <= 0 or pass_s < 0 or pass_s > max_s:
            error = "满分要大于0，及格线要在 0 ~ 满分 之间"
        else:
            conn.execute("UPDATE exam_subjects SET max_score=?, pass_score=? WHERE id=?",
                         (max_s, pass_s, cfg_id))
            conn.commit()
    exam = conn.execute("SELECT * FROM exams WHERE id=?", (exam_id,)).fetchone()
    subs = conn.execute("SELECT * FROM exam_subjects WHERE exam_id=? ORDER BY id",
                        (exam_id,)).fetchall()
    conn.close()
    return render_template("exam_config.html", exam=exam, subs=subs, error=error)


@app.route("/exam/<int:exam_id>/delete_subject/<int:cfg_id>", methods=["POST"])
@login_required
def delete_subject(exam_id, cfg_id):
    subj = request.form.get("subject", "")
    conn = get_db()
    if not get_my_exam(conn, exam_id):
        conn.close()
        return "没有这场考试（或不是你的）"
    conn.execute("DELETE FROM exam_subjects WHERE id=?", (cfg_id,))
    conn.execute("DELETE FROM scores WHERE exam_id=? AND subject=?",
                 (exam_id, subj))
    conn.commit()
    conn.close()
    return redirect(f"/exam/{exam_id}")


@app.route("/exam/<int:exam_id>/apply_template", methods=["POST"])
@login_required
def apply_template(exam_id):
    """一键补齐/重置为山东高考模板：缺的科目补上，已有科目不覆盖"""
    conn = get_db()
    if not get_my_exam(conn, exam_id):
        conn.close()
        return "没有这场考试（或不是你的）"
    for subj, mx, ps in default_subjects():
        if not conn.execute("SELECT 1 FROM exam_subjects WHERE exam_id=? AND subject=?",
                            (exam_id, subj)).fetchone():
            conn.execute("INSERT INTO exam_subjects (exam_id,subject,max_score,pass_score) "
                         "VALUES (?,?,?,?)", (exam_id, subj, mx, ps))
    conn.commit()
    conn.close()
    return redirect(f"/exam/{exam_id}")


# ----- 全班汇总：按某场考试排名 + 不及格名单（只看自己班） -----
@app.route("/summary")
@login_required
def summary():
    conn = get_db()
    uid = session.get("uid")
    if session.get("user") == "admin":
        all_exams = conn.execute("SELECT id, name FROM exams ORDER BY id").fetchall()
    else:
        all_exams = conn.execute("SELECT id, name FROM exams WHERE teacher_id=? ORDER BY id",
                                 (uid,)).fetchall()
    exam_id = request.args.get("exam_id", type=int)
    if not exam_id and all_exams:
        exam_id = all_exams[0]["id"]
    exam = get_my_exam(conn, exam_id) if exam_id else None
    if session.get("user") == "admin":
        rows = conn.execute(
            "SELECT st.id sid, st.name, s.subject, s.score, es.max_score, es.pass_score "
            "FROM scores s JOIN students st ON s.student_id=st.id "
            "LEFT JOIN exam_subjects es ON es.exam_id=s.exam_id AND es.subject=s.subject "
            "WHERE s.exam_id=? ORDER BY st.id", (exam_id,)).fetchall()
    else:
        rows = conn.execute(
            "SELECT st.id sid, st.name, s.subject, s.score, es.max_score, es.pass_score "
            "FROM scores s JOIN students st ON s.student_id=st.id "
            "LEFT JOIN exam_subjects es ON es.exam_id=s.exam_id AND es.subject=s.subject "
            "WHERE s.exam_id=? AND st.teacher_id=? ORDER BY st.id",
            (exam_id, uid)).fetchall()
    conn.close()

    # 按学生分组，算总分、平均分、不及格科数（及格线读 exam_subjects）
    stu_map = {}
    fails = []
    for r in rows:
        d = stu_map.setdefault(r["sid"], {"name": r["name"], "total": 0,
                                          "count": 0, "fail_count": 0})
        d["total"] += r["score"]
        d["count"] += 1
        if r["score"] < r["pass_score"]:
            d["fail_count"] += 1
            fails.append({"name": r["name"], "subject": r["subject"],
                          "score": r["score"], "pass_score": r["pass_score"]})

    ranking = []
    for sid, d in stu_map.items():
        ranking.append({"sid": sid, "name": d["name"], "total": round(d["total"], 1),
                        "avg": round(d["total"] / d["count"], 1),
                        "count": d["count"], "fail_count": d["fail_count"]})
    ranking.sort(key=lambda x: x["avg"], reverse=True)   # 按平均分降序

    return render_template("summary.html", exam=exam, all_exams=all_exams,
                           ranking=ranking, fails=fails)


# ----- 删除学生（连带删除他的成绩） -----
@app.route("/delete/<int:sid>", methods=["POST"])
@login_required
def delete_student(sid):
    conn = get_db()
    if get_my_student(conn, sid):
        conn.execute("DELETE FROM scores WHERE student_id=?", (sid,))   # 先删他的成绩
        conn.execute("DELETE FROM students WHERE id=?", (sid,))          # 再删学生本人
        conn.commit()
    conn.close()
    return '<script>location.href="/"</script>'


# ----- 修改一条成绩 -----
@app.route("/edit_score/<int:score_id>", methods=["GET", "POST"])
@login_required
def edit_score(score_id):
    conn = get_db()
    r = conn.execute(
        "SELECT s.id, s.subject, s.score, s.student_id, s.exam_id, st.name, "
        "st.teacher_id, e.name exam_name, es.max_score, es.pass_score "
        "FROM scores s JOIN students st ON s.student_id=st.id "
        "JOIN exams e ON s.exam_id=e.id "
        "LEFT JOIN exam_subjects es ON es.exam_id=s.exam_id AND es.subject=s.subject "
        "WHERE s.id=?", (score_id,)).fetchone()
    if not r:
        conn.close()
        return "没有这条成绩"
    # 归属检查：非 admin 只能改自己班学生的成绩
    if session.get("user") != "admin" and r["teacher_id"] != session.get("uid"):
        conn.close()
        return "不能修改其他老师的成绩"

    error = None
    if request.method == "POST":
        score_val = request.form.get("score", "").strip()
        try:
            score_val = float(score_val)
        except ValueError:
            error = "成绩必须是数字"
        else:
            if score_val < 0 or score_val > r["max_score"]:
                mx = num_filter(r["max_score"])
                error = f"{r['subject']} 满分 {mx} 分，请输入 0 ~ {mx}"
            else:
                conn.execute("UPDATE scores SET score=? WHERE id=?", (score_val, score_id))
                conn.commit()
                conn.close()
                return (f'<script>location.href="/score/{r["student_id"]}'
                        f'?exam_id={r["exam_id"]}"</script>')
    conn.close()
    return render_template("edit_score.html", rec=r, error=error)


# ----- 导出全班成绩为 Excel -----
@app.route("/export_excel")
@login_required
def export_excel():
    conn = get_db()
    uid = session.get("uid")
    is_admin = session.get("user") == "admin"
    # 只查自己的考试
    if is_admin:
        exams = conn.execute("SELECT id, name FROM exams ORDER BY id").fetchall()
        students = conn.execute("SELECT id, name FROM students ORDER BY id").fetchall()
    else:
        exams = conn.execute("SELECT id, name FROM exams WHERE teacher_id=? ORDER BY id",
                             (uid,)).fetchall()
        students = conn.execute("SELECT id, name FROM students WHERE teacher_id=? ORDER BY id",
                                (uid,)).fetchall()
    # 所有成绩：key=(student_id, exam_id, subject) -> (score, pass_score)
    score_map = {}
    for r in conn.execute(
        "SELECT s.student_id, s.exam_id, s.subject, s.score, es.pass_score "
        "FROM scores s LEFT JOIN exam_subjects es "
        "ON es.exam_id=s.exam_id AND es.subject=s.subject"
    ):
        score_map[(r["student_id"], r["exam_id"], r["subject"])] = (r["score"], r["pass_score"])
    conn.close()

    SUBJECTS = ["语文", "数学", "英语", "物理", "化学", "生物"]

    wb = Workbook()
    wb.remove(wb.active)  # 删除默认空表

    head_fill = PatternFill("solid", fgColor="1E3A8A")
    red_fill = PatternFill("solid", fgColor="FECACA")

    for exam in exams:
        ws = wb.create_sheet(title=exam["name"][:31])
        # 表头
        headers = ["姓名"] + SUBJECTS
        ws.append(headers)
        for c in ws[1]:
            c.fill = head_fill
            c.font = Font(color="FFFFFF", bold=True)
            c.alignment = Alignment(horizontal="center")
        # 每个学生一行
        for stu in students:
            row = [stu["name"]]
            for subj in SUBJECTS:
                info = score_map.get((stu["id"], exam["id"], subj))
                row.append(info[0] if info else "")
            ws.append(row)
            # 不及格标红
            for col_idx, subj in enumerate(SUBJECTS, start=2):
                info = score_map.get((stu["id"], exam["id"], subj))
                if info and info[1] is not None and info[0] < info[1]:
                    ws.cell(row=ws.max_row, column=col_idx).fill = red_fill
                    ws.cell(row=ws.max_row, column=col_idx).font = Font(color="DC2626", bold=True)
        # 列宽
        ws.column_dimensions["A"].width = 12
        for i in range(2, 8):
            ws.column_dimensions[chr(64 + i)].width = 8

    if not exams:
        ws = wb.create_sheet(title="全班成绩")
        ws.append(["暂无考试数据"])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return send_file(
        buf, as_attachment=True, download_name="全班成绩.xlsx",
        mimetype="application/vnd.openxmlformats-officedocument."
                 "spreadsheetml.sheet")


if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8080))   # 云平台会通过 PORT 环境变量指定端口
    app.run(host="0.0.0.0", port=port, debug=False)
