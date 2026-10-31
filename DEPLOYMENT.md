# 安全升级与日常备份

## 部署前

1. 本地运行回归测试：`python -m unittest discover -s tests -v`。
2. 将确认的修改提交到自己的 GitHub 仓库。不要提交 class.db、.instance、任何密钥或备份。
3. 在 PythonAnywhere 控制台确认当前仓库状态，备份后再更新代码，不要使用会覆盖本地改动的 `git reset --hard`。

## 服务器更新步骤

先暂时停止录入，确认数据库实际路径，使用 SQLite backup API 备份旧数据库。以下适用于交接文档中的默认目录；若改过数据库位置，请替换路径：

```bash
cd /home/1039535644/class_manager
git status --short
python3 - <<'PY'
import sqlite3, os
from pathlib import Path
from datetime import datetime
folder = Path('/home/1039535644/.class_manager_backups')
folder.mkdir(mode=0o700, exist_ok=True)
path = folder / ('before-update-' + datetime.now().strftime('%Y%m%d-%H%M%S') + '.db')
fd = os.open(path, os.O_CREAT | os.O_EXCL | os.O_WRONLY, 0o600)
os.close(fd)
source = sqlite3.connect('file:class.db?mode=ro', uri=True)
target = sqlite3.connect(path)
source.backup(target)
assert target.execute('PRAGMA integrity_check').fetchone()[0] == 'ok'
target.close()
source.close()
print('Backup:', path)
PY
```

在服务器的 `~/.class_manager_env` 保留已有邮件设置，并新增：

```text
CLASS_MANAGER_ENV=production
SECRET_KEY=独立生成的至少32字符随机值
```

随机值可用 `python3 -c 'import secrets; print(secrets.token_hex(32))'` 生成。不要将密钥发到聊天、写进 Git 或和其他项目共用。配置文件权限设置为600。若进程环境已设置旧 SECRET_KEY，先更新该环境配置，因为进程环境变量优先。

更新仓库与当前 Web 使用的虚拟环境：

```bash
git pull --ff-only origin main
# 先激活 PythonAnywhere Web 面板中配置的虚拟环境，再运行：
python -m pip install -r requirements.txt
python migrate.py
```

确认无误后，在 PythonAnywhere Web 面板 Reload。首次升级会：

- 在数据库目录的上级 `.class_manager_backups` 自动备份数据库（可用 CLASS_MANAGER_BACKUPS 改位置）；备份失败则不执行升级。
- 保留老师、学生、考试和成绩，添加会话版本、账号状态、缺考标识等字段。
- 去除旧库“所有老师考试名必须全局唯一”的限制，允许不同老师创建同名考试。
- 使旧验证码失效；更换会话密钥后，所有人需要重新登录。
- 若 admin 仍用旧默认密码，登录后强制先改密。其他已有正常密码保持有效。

生产环境未配置强密钥时程序会拒绝启动。HTTPS 部署才应设置 production，因为该模式启用 Secure Cookie。本地 HTTP 使用默认开发模式。

## 全新安装

本地运行 `python app.py`。未设置 SECRET_KEY 时在 `.instance/secret-key` 生成持久化随机密钥；该目录不可公开。

全新数据库不再使用公开默认密码。可以预先设置 `ADMIN_INITIAL_PASSWORD`（至少8位且不能是旧默认密码）；否则随机初始密码保存在数据库旁 `.instance/admin-initial-password`，文件权限600。首次管理员登录必须改密，之后删除初始密码文件。已有数据库不会重置密码。

## 更新后验收

- 老师登录、学生列表、考试配置、录分与修改、缺考、统计页面、Excel 导出。
- 用两个测试老师确认互相看不到学生，也不能更改对方科目或成绩。
- 用测试学生和家长账号完成首次改密、退出和重新登录；确认只能访问绑定学生，直接提交录分、删除等请求也被拒绝。
- 同一老师创建两个测试班级，确认学生和家长的趋势图不包含无关班级的考试名称；同一老师内转班后，已录成绩的历史考试仍保留。
- 停用测试学生或家长账号后，已登录会话应失效；恢复账号后，必须重新登录。
- 自定义科目能导出；同名考试、带斜线的考试名能导出。
- 用自己的邮箱验证真实验证码发送、收件、重置；不要使用学生邮箱做联调。
- 如果应用无法启动，查看 Web 的 error log，检查密钥、虚拟环境与数据库路径。不要将完整邮件配置粘贴到公共日志。

## 每日备份与恢复

`python backup.py` 创建一致性快照并运行完整性检查，不导入 Web 应用，不自动迁移。可使用 `--database` 和 `--directory` 指定位置。

在 PythonAnywhere Tasks 配置每日执行（Python 替换为 Web 所用虚拟环境中的绝对路径）：

```bash
/home/1039535644/你的虚拟环境/bin/python /home/1039535644/class_manager/backup.py --directory /home/1039535644/.class_manager_backups
```

首次升级自动备份已在程序实现；每日定时任务需要在主机面板设置。备份不会自动清理，以免误删唯一可恢复版本，需按实际空间定期把备份复制到安全位置。

恢复时暂停 Web 与定时任务，先备份当前库，再用选定备份替换数据库。先在临时目录验证恢复副本，再启动 Web。不要在 Web 正在写入时直接覆盖数据库；应用升级后退回旧代码前，也要确认数据库版本兼容。

## 邮件配置

保持 GApps → Brevo → Mailjet → SMTP 的配置优先选择顺序。一次发送选定一个已配置通道；失败会提示稍后重试，不自动切换到另一通道，避免超时后重复发送。验证码只有在发送调用成功后才可验证；真实送达仍需联调。

GApps 当前使用已有 GET 接口，后续若改为带鉴权的 POST，需同步修改 Google Apps Script。仅凭 Python 仓库无法验证或修复远端脚本是否限制第三方调用。


## 2026-10-30 实际部署记录

- 公网 Web 使用 Python3.10，独立虚拟环境为 `/home/1039535644/.virtualenvs/class-manager`。
- 后续更新依赖、运行迁移和备份时，使用该环境的绝对 Python 路径。
- 生产配置位于 `~/.class_manager_env`，已启用生产模式、独立随机密钥和 Secure Cookie；PythonAnywhere 已启用强制 HTTPS。
- 首次上线在服务器 Python3.10 环境运行40项回归测试通过。测试日志保存在 `/home/1039535644/class-manager-release-tests.log`。
- 部署前一致性备份保存在 `/home/1039535644/.class_manager_backups/before-release-20261030-010529.db`，自动迁移也另外创建了备份。部署前后五张业务表原有列数据与ID逐项一致。
- 本账号的 Tasks 页面明确提示计划任务仅付费账号开放，因此未创建每日定时任务。免费账号可在每次重要录入后手动执行下面的备份命令，无需升级付费计划：

```bash
/home/1039535644/.virtualenvs/class-manager/bin/python /home/1039535644/class_manager/backup.py
```

真实邮件发送与收件未在部署过程中触发，原邮件配置已保留。
