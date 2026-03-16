# 量子推送机器人 v4.0

桌面程序（CustomTkinter），支持长期运行的两类检测任务：
- 邮件检测：按“邮箱检测规则”拉取邮件并筛选附件后推送到量子机器人。
- 文件夹检测：监控本地目录新文件/变更并推送到量子机器人。

## 目录结构

- `gui_app.py`：桌面入口
- `desktop_pages/`：UI 页面模块
- `mail_forwarder/`：检测、筛选、上传、推送核心逻辑
- `settings/`：全部业务配置（可持久化）
- `state/`：仅运行状态（去重与处理进度）
- `QuantumBot.spec`：PyInstaller 打包配置

## 配置文件（v4.0）

程序不再依赖 `.env` 兜底；未配置完整时任务不会启动。

- `settings/app_config.json`：IMAP 基础配置
- `settings/webhook_aliases.json`：机器人别名与 webhook 地址
- `settings/subject_attachment_rules.json`：邮箱检测规则（启用状态、主题关键字、附件类型、附件关键字、目标机器人）
- `settings/folder_monitor_config.json`：文件夹检测配置

运行状态文件：

- `state/mail_state.json`：邮件检测去重状态（如已处理 UID）
- `state/file_sent_state.json`：文件夹文件发送去重状态

## 环境要求

- Python 3.11（推荐）

安装依赖：

```bash
python3.11 -m venv .venv311
source .venv311/bin/activate   # Windows: .venv311\Scripts\activate
pip install -r requirements.txt
pip install -r requirements_gui.txt
```

## 启动

```bash
python gui_app.py
```

Windows 也可使用：

```bat
start_gui.bat
```

## 快速自检

```bash
python quick_check.py
```

## 打包（Windows EXE）

推荐在独立虚拟环境中打包（避免无关依赖导致体积膨胀）：

```powershell
py -3.11 -m venv .venv-pack
.venv-pack\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt -r requirements_gui.txt
pip install pyinstaller
pyinstaller --noconfirm --clean QuantumBot.spec
```

产物：

- `dist/QuantumBot/QuantumBot.exe`

## Git 上传规则（已更新）

默认不提交本机敏感配置与运行状态：

- 忽略：`settings/*.json`
- 忽略：`state/*.json`
- 保留目录：`settings/.gitkeep`、`state/.gitkeep`

建议提交前检查：

```bash
git status
```

确认仅包含代码与文档变更后再推送。
