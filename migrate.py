"""
一次性数据库迁移脚本：把旧模型升级为"自定义考试/满分/及格线"
- 建 exams（考试）、exam_subjects（考试科目配置）
- scores 加 exam_id
- 把现有考试、成绩、科目配置迁移过来
幂等：可重复执行，已有数据不会重复、不会覆盖
云端执行：python3 migrate.py
"""
import sqlite3
import os

DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "class.db")
conn = sqlite3.connect(DB)
conn.row_factory = sqlite3.Row

conn.execute("CREATE TABLE IF NOT EXISTS exams ("
             "id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL)")
conn.execute("CREATE TABLE IF NOT EXISTS exam_subjects ("
             "id INTEGER PRIMARY KEY AUTOINCREMENT, exam_id INTEGER, "
             "subject TEXT, max_score REAL, pass_score REAL)")
cols = [r[1] for r in conn.execute("PRAGMA table_info(scores)")]
if "exam_id" not in cols:
    conn.execute("ALTER TABLE scores ADD COLUMN exam_id INTEGER")
# users 表加 email（注册 + 找回密码用）
ucols = [r[1] for r in conn.execute("PRAGMA table_info(users)")]
if "email" not in ucols:
    conn.execute("ALTER TABLE users ADD COLUMN email TEXT")


def template():
    return [(s, 150, 90) for s in ["语文", "数学", "英语"]] + \
           [(s, 100, 60) for s in ["物理", "化学", "生物"]]


for er in conn.execute("SELECT DISTINCT exam FROM scores").fetchall():
    ename = er["exam"]
    if not ename:
        continue
    r = conn.execute("SELECT id FROM exams WHERE name=?", (ename,)).fetchone()
    eid = r["id"] if r else conn.execute(
        "INSERT INTO exams (name) VALUES (?)", (ename,)).lastrowid
    # 标准模板科目
    for subj, mx, ps in template():
        if not conn.execute(
                "SELECT 1 FROM exam_subjects WHERE exam_id=? AND subject=?",
                (eid, subj)).fetchone():
            conn.execute("INSERT INTO exam_subjects (exam_id,subject,max_score,pass_score) "
                         "VALUES (?,?,?,?)", (eid, subj, mx, ps))
    # 该考试实际出现、但模板没有的科目 → 补默认配置，保证成绩能关联上
    for sr in conn.execute(
            "SELECT DISTINCT subject FROM scores WHERE exam=?",
            (ename,)).fetchall():
        subj = sr["subject"]
        if not conn.execute(
                "SELECT 1 FROM exam_subjects WHERE exam_id=? AND subject=?",
                (eid, subj)).fetchone():
            conn.execute("INSERT INTO exam_subjects (exam_id,subject,max_score,pass_score) "
                         "VALUES (?,?,?,?)", (eid, subj, 100, 60))
    conn.execute("UPDATE scores SET exam_id=? WHERE exam=?", (eid, ename))

conn.commit()
print("迁移完成：考试数", conn.execute("SELECT COUNT(*) FROM exams").fetchone()[0],
      "| 科目配置数", conn.execute("SELECT COUNT(*) FROM exam_subjects").fetchone()[0],
      "| 已关联成绩", conn.execute(
          "SELECT COUNT(*) FROM scores WHERE exam_id IS NOT NULL").fetchone()[0])
