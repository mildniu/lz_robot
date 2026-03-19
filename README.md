# 量子推送机器人 v5.1

桌面程序当前包含两套桌面入口：

- `gui_app.py`：现有稳定版（CustomTkinter）
- `gui_qt_app.py`：新界面并行版（PySide6）

两套入口共用同一套业务核心与 `settings/*.json` 配置文件，支持长期运行的两类检测任务：

- 邮件检测：按“邮箱检测规则”从指定 IMAP 邮箱拉取最新邮件，筛选附件后直接推送，或交给脚本/EXE 处理后再推送。
- 文件夹检测：监控本地目录中新文件或变化文件，并推送到指定机器人。
- 规则处理程序支持 `.py` 和 `.exe`。
- 邮件规则支持两种检测方式：`周期检测`（每隔 N 分钟）和 `定时检测`（每天固定时刻一次）。
- 邮件规则、文件夹检测均支持独立启动、独立停止、独立日志和运行中热刷新。

## 当前版本重点

- 多 IMAP 邮箱别名，规则可绑定不同邮箱。
- 机器人别名统一管理，邮件规则、文件夹检测、脚本处理共用。
- 邮件规则支持单条保存，顶部摘要只在保存成功后更新。
- 附件保存名带 UID，避免同名覆盖。
- 脚本处理链路支持将图片、文件、文字通过当前选中的机器人推送。
- 处理程序超时可配置，默认 `300` 秒。

## 目录结构

- `gui_app.py`：桌面入口
- `desktop_pages/`：UI 页面模块
- `mail_forwarder/`：邮件检测、筛选、推送、脚本执行核心逻辑
- `scripts/`：规则处理脚本目录，含脚本模板与推送辅助文件
- `settings/`：全部业务配置（持久化）
- `state/`：运行状态（去重与处理进度）
- `quick_check.py`：快速自检脚本
- `QuantumBot.spec`：Windows 打包配置
- `build_release.ps1`：正式打包脚本

## 配置文件

程序当前以 `settings/*.json` 为主，不再依赖 `.env` 作为运行配置来源。

- `settings/app_config.json`：程序基础配置
  包含窗口参数、下载目录、状态文件、默认 webhook、处理程序超时等
- `settings/webhook_aliases.json`：机器人别名与 webhook 地址
- `settings/mailbox_aliases.json`：IMAP 邮箱别名配置
- `settings/subject_attachment_rules.json`：邮箱检测规则
- `settings/folder_monitor_config.json`：文件夹检测配置

运行状态文件：

- `state/mail_state.json`：邮件规则最近处理 UID 去重状态
- `state/file_sent_state.json`：文件夹监测文件去重状态

## 环境要求

- Python 3.11（推荐）

安装依赖：

```bash
python -m venv .venv311
.venv311\Scripts\activate
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

PySide6 桌面版可使用：

```bat
start_gui_qt.bat
```

也可直接运行：

```bash
python gui_qt_app.py
```

## 使用流程

1. 先在“机器人别名”中维护 webhook 地址。
2. 再在“邮箱配置”中维护 IMAP 邮箱别名，并先做连接测试。
3. 在“邮箱检测规则”中为每条规则选择所属邮箱、推送机器人、检测方式与附件条件。
4. 如需脚本处理，可为规则指定 `.py` 或 `.exe`，并设置输出目录。
5. 如需本地文件推送，可在“文件夹检测”中配置监测目录和目标机器人。
6. 到“邮件检测”或“文件夹检测”页面独立启动、测试并查看日志。

## 稳定性说明

- 邮件规则与文件夹监测均支持运行中热刷新。
- 单条规则保存后不会影响其他槽位，空规则槽位会保留。
- 脚本执行默认超时 `300` 秒，可在“界面设置”中调整“处理程序超时(s)”。
- 脚本超时、邮箱登录失败、文件夹无效、Webhook 发送失败等情况会在日志中给出更明确的排查建议。

## 快速自检

```bash
python quick_check.py
```

当前自检会检查：

- 核心入口文件是否存在
- 邮件规则槽位、启用状态、触发方式
- 文件夹检测配置
- 运行状态文件是否可读取
- 打包资源引用
- 处理程序超时配置

## 打包（Windows EXE）

推荐在独立虚拟环境中打包，避免引入无关依赖：

```powershell
py -3.11 -m venv .venv-pack
.venv-pack\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt -r requirements_gui.txt
pip install pyinstaller
.\build_release.ps1
```

产物：

- `dist/QuantumBot/QuantumBot.exe`
- `dist/QuantumBot/scripts/`

说明：

- `build_release.ps1` 会先执行 PyInstaller，再在最终成品根目录生成 `scripts/`
- `dist/QuantumBot/scripts/` 内默认带上：
  - `rule_processor_template.py`
  - `script_push_helper.py`
- 如需独立分发脚本处理程序，可单独将业务脚本打包为 EXE 后放入 `scripts/`

## 打包（Windows EXE，PySide6 版）

如需单独验证 Qt 版打包，可使用：

```powershell
py -3.11 -m venv .venv-pack
.venv-pack\Scripts\Activate.ps1
python -m pip install --upgrade pip
pip install -r requirements.txt -r requirements_gui_qt.txt
pip install pyinstaller
.\build_release_qt.ps1
```

产物：

- `dist/QuantumBotQt/QuantumBotQt.exe`
- `dist/QuantumBotQt/scripts/`

## Git 上传建议

默认不提交本机敏感配置与运行状态：

- 忽略：`settings/*.json`
- 忽略：`state/*.json`
- 保留目录：`settings/.gitkeep`、`state/.gitkeep`

建议提交前检查：

```bash
git status
```

## PySide6 迁移状态

当前仓库已新增一套不影响现有 Tk 版的 PySide6 并行版本：

- `gui_qt_app.py`
- `qt_pages/`
- `qt_components/`
- `app_services/`
- `requirements_gui_qt.txt`
- `QuantumBotQt.spec`
- `build_release_qt.ps1`
- `PYSIDE6_MIGRATION_PLAN.md`

说明：

- 当前 Qt 版已可独立运行，已具备：
  - 邮件检测页：规则卡片、独立启动/停止/测试、独立日志、日志导出/清空
  - 文件夹检测页：独立启动/停止、独立日志、日志导出/清空
  - 机器人测试页：文字/文件测试、日志导出/清空
  - 设置页：机器人别名、邮箱配置、邮箱检测规则、文件夹检测、路径设置、界面设置
  - 设置保存后可通知其他页面刷新配置
- Qt 入口已支持读取窗口尺寸、启动页、侧栏宽度、页脚文案等现有配置
- 现有正式版本仍以 `gui_app.py` 为主
