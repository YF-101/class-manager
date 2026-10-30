# 班主任班级成绩管理系统

基于 Flask、Jinja2、SQLite 的多老师成绩管理网站。

支持注册登录、独立老师数据、多班级管理、学生与考试管理、自定义科目/满分/及格线、成绩录入修改与单条删除、缺考登记、完整成绩总分排名（同分并列）、各科统计、得分率趋势图、动态科目 Excel 导出、不及格标红和打印。

## 本地运行

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -r requirements.txt
python app.py
```

打开 http://127.0.0.1:8080 。首次创建数据库时，随机管理员初始密码保存在 `.instance/admin-initial-password`，登录后必须更改。不再提供公开的默认密码。旧数据库升级保留原密码；仍使用旧默认密码的管理员会被要求先改密。

数据默认保存在项目 `class.db`。可用 `CLASS_MANAGER_DB` 指定路径；未配置环境密钥时，本地随机密钥保存在 `.instance/secret-key`。`.instance`、数据库、备份都不能提交或公开。

## 测试、升级与备份

```bash
python -m unittest discover -s tests -v
python migrate.py
python backup.py
```

回归测试只用临时数据库和模拟邮件，不接触真实数据。首次升级会先备份现有数据库；迁移可重复运行。

**生产部署、密钥、数据库恢复和每日备份任务设置见 [DEPLOYMENT.md](DEPLOYMENT.md)。** 在 PythonAnywhere 使用 Web 面板配置的虚拟环境安装依赖，再 Reload。

## 统计规则

- 所有科目都有有效正常成绩的学生，按总分排名；同分并列，后续名次跳位。
- 未录入、明确缺考与异常记录分开显示，暂不参加完整成绩排名。0分是正常成绩，不等于缺考。
- 平均分、及格率、优秀率仅使用有效正常成绩为分母；优秀线为满分85%。
- 不同考试的趋势图使用得分率，便于比较不同满分的考试；缺考和未录入不连线。
- 停用老师立即阻止登录并使原会话失效，保留其学生和考试数据，可恢复账号。
- 管理员可以管理其他老师的数据，但不能把一个老师的学生录入另一个老师的考试。

## 主要文件

- `app.py`：路由、会话、权限、邮件、成绩统计和导出。
- `database.py`：事务、数据库迁移、备份和数值校验。
- `templates/`：界面与打印样式。
- `migrate.py` / `backup.py`：升级与一致性备份命令。
- `tests/test_regressions.py`：权限、安全、业务与旧库迁移回归测试。

下一步适合增加：学号、Excel 批量导入和整班表格录分。学生查询入口需要独立认证，不应仅凭可猜测学号公开成绩。
