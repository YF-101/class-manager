"""Consistent SQLite backup; suitable for a PythonAnywhere daily scheduled task."""
import argparse
import os
from pathlib import Path
from database import backup_database


def main():
    parser = argparse.ArgumentParser(description='备份学生成绩数据库并检查完整性')
    parser.add_argument('--database', default=os.environ.get('CLASS_MANAGER_DB', str(Path(__file__).resolve().parent / 'class.db')))
    parser.add_argument('--directory', default=os.environ.get('CLASS_MANAGER_BACKUPS', str(Path(__file__).resolve().parent.parent / '.class_manager_backups')))
    args = parser.parse_args()
    if not Path(args.database).is_file():
        parser.error('数据库文件不存在')
    print('备份已保存：', backup_database(args.database, args.directory))


if __name__ == '__main__':
    main()
