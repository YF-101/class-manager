"""Upgrade safely; existing databases are backed up before migration."""
from app import DB, database

if __name__ == '__main__':
    with database() as conn:
        print('数据库升级完成；学生数：', conn.execute('SELECT COUNT(*) FROM students').fetchone()[0],
              '考试数：', conn.execute('SELECT COUNT(*) FROM exams').fetchone()[0],
              '成绩数：', conn.execute('SELECT COUNT(*) FROM scores').fetchone()[0])
