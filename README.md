# 班主任班级成绩管理系统（class-manager）

一个基于 Flask 的班级成绩管理 Web 应用，面向班主任/老师，支持多用户注册、自定义考试科目与满分、成绩录入、排名分析、折线图、Excel 导出。

## 在线演示

- **访问地址**：https://1039535644.pythonanywhere.com
- 任何人可自行注册账号使用

## 技术栈

- **后端**：Python 3、Flask、SQLite
- **前端**：HTML、CSS、JavaScript、Chart.js
- **部署**：PythonAnywhere（免费云主机）
- **版本管理**：Git / GitHub

## 主要功能

- **多用户体系**：注册、登录、密码哈希存储、管理员后台管理账号
- **学生管理**：增删改查、防重名
- **自定义考试**：新建考试（如月考、模拟考），自定义每科满分与及格线
- **成绩录入**：按考试+科目录入，同一考试同一科目自动覆盖更新
- **成绩分析**：按考试平均分排名、不及格名单、单科成绩折线图
- **Excel 导出**：不及格成绩自动标红
- **打印成绩单**：适配浏览器打印 / 存 PDF
- **找回密码**：预留邮箱验证码通道（SMTP / Mailjet 双通道）

## 本地运行

```bash
cd class-manager
pip install -r requirements.txt
python3 app.py
# 浏览器打开 http://127.0.0.1:8080
```

## 项目结构

```
class-manager/
├── app.py              # 主程序：路由、数据库、发信逻辑
├── migrate.py          # 数据库迁移脚本（幂等）
├── requirements.txt     # 依赖
├── Procfile            # 部署配置
├── .gitignore
├── templates/          # Jinja2 页面模板
│   ├── login.html / register.html / forgot_password.html
│   ├── index.html / score.html / student.html / summary.html
│   ├── exams.html / exam_config.html / edit_score.html
│   └── users.html / change_password.html / base.html
└── class.db            # SQLite 数据库（不公开）
```

## 作者

YF-101
