"""
重置演示数据：清空旧学生/成绩/考试，生成期中、期末两场考试，
随机生成 12 个学生和 6 门课（语数外 150、物化生 100）的成绩。
用法：python3 reset_demo.py
"""
import sqlite3
import random
import os

DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "class.db")

# 山东高考科目配置：(科目, 满分, 及格线)
SUBJECTS = [
    ("语文", 150, 90),
    ("数学", 150, 90),
    ("英语", 150, 90),
    ("物理", 100, 60),
    ("化学", 100, 60),
    ("生物", 100, 60),
]

# 随机姓名池
SURNAMES = ["王", "李", "张", "刘", "陈", "杨", "赵", "黄", "周", "吴", "徐", "孙",
            "马", "朱", "胡", "郭", "何", "林", "罗", "郑", "梁", "谢", "宋", "唐"]
GIVEN = ["伟", "芳", "娜", "敏", "静", "丽", "强", "磊", "军", "洋",
         "勇", "艳", "杰", "娟", "涛", "明", "超", "秀英", "雪", "晨",
         "子涵", "雨泽", "欣怡", "浩然", "梓萱", "一诺", "思远", "嘉琪",
         "宇轩", "梦琪", "天佑", "若曦", "志远", "书瑶", "文昊", "雅静"]

EXAMS = ["期中考试", "期末考试"]


def gen_names(n):
    random.seed(42)  # 固定种子，每次生成一样的名字，避免重复
    names = set()
    while len(names) < n:
        name = random.choice(SURNAMES) + random.choice(GIVEN)
        names.add(name)
    return list(names)


def score_for(full, pass_line):
    """随机生成分数：大部分及格，少数不及格，偶尔高分。"""
    r = random.random()
    if r < 0.15:  # 15% 不及格
        return round(random.uniform(pass_line * 0.5, pass_line - 3), 1)
    elif r < 0.35:  # 20% 优秀
        return round(random.uniform(full * 0.85, full * 0.98), 1)
    else:  # 65% 中等
        return round(random.uniform(pass_line, full * 0.85), 1)


def main():
    # Initialise schema and snapshot before this intentionally destructive demo reset.
    from app import DB as configured_db, BACKUP_DIR
    from database import backup_database
    backup_database(configured_db, BACKUP_DIR)
    conn = sqlite3.connect(configured_db)
    conn.row_factory = sqlite3.Row
    cur = conn.cursor()
    admin = cur.execute("SELECT id FROM users WHERE username='admin'").fetchone()
    teacher_id = admin['id']
    conn.execute('BEGIN IMMEDIATE')

    # 1. 清空旧数据（保留 users / reset_tokens）
    cur.execute("DELETE FROM scores")
    cur.execute("DELETE FROM exam_subjects")
    cur.execute("DELETE FROM exams")
    cur.execute("DELETE FROM students")
    # 重置自增
    cur.execute("DELETE FROM sqlite_sequence WHERE name IN ('students','scores','exams','exam_subjects')")

    # 2. 创建考试
    exam_ids = {}
    for name in EXAMS:
        cur.execute("INSERT INTO exams (name,teacher_id) VALUES (?,?)", (name,teacher_id))
        exam_ids[name] = cur.lastrowid

    # 3. 配置每科满分/及格线
    for exam_name, eid in exam_ids.items():
        for subj, full, pass_line in SUBJECTS:
            cur.execute(
                "INSERT INTO exam_subjects (exam_id, subject, max_score, pass_score) VALUES (?,?,?,?)",
                (eid, subj, full, pass_line)
            )

    # 4. 生成学生和成绩
    names = gen_names(12)
    for name in names:
        cur.execute("INSERT INTO students (name,teacher_id) VALUES (?,?)", (name,teacher_id))
        sid = cur.lastrowid
        for exam_name, eid in exam_ids.items():
            for subj, full, pass_line in SUBJECTS:
                sc = score_for(full, pass_line)
                cur.execute(
                    "INSERT INTO scores (student_id, subject, score, exam, exam_id) VALUES (?,?,?,?,?)",
                    (sid, subj, sc, exam_name, eid)
                )

    conn.commit()

    # 统计
    n_stu = cur.execute("SELECT COUNT(*) c FROM students").fetchone()["c"]
    n_exam = cur.execute("SELECT COUNT(*) c FROM exams").fetchone()["c"]
    n_score = cur.execute("SELECT COUNT(*) c FROM scores").fetchone()["c"]
    conn.close()

    print(f"✅ 演示数据生成完成")
    print(f"   学生数: {n_stu}")
    print(f"   考试数: {n_exam}（期中、期末）")
    print(f"   成绩条数: {n_score}")
    print(f"   科目: 语数外(150/及格90) + 物化生(100/及格60)")


if __name__ == "__main__":
    main()
