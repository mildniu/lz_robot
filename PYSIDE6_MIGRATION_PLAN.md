# PySide6 迁移蓝图（进行中）

## 1. 迁移目标

在不破坏现有 `CustomTkinter` 版本的前提下，并行搭建一套 `PySide6` 桌面 UI，后续逐页替换：

- 保留 `mail_forwarder/` 业务核心
- 保留 `settings/*.json` 与 `state/*.json`
- 先实现可运行预览版，再逐步迁移页面

## 2. 文件级映射

当前 Tk 结构：

- `gui_app.py`
- `desktop_pages/about_page.py`
- `desktop_pages/bot_test_page.py`
- `desktop_pages/execute_page.py`
- `desktop_pages/folder_monitor_page.py`
- `desktop_pages/settings_page.py`
- `desktop_pages/common.py`

建议的 Qt 结构：

- `gui_qt_app.py`
- `qt_pages/about_page.py`
- `qt_pages/bot_test_page.py`
- `qt_pages/execute_page.py`
- `qt_pages/folder_monitor_page.py`
- `qt_pages/settings_page.py`
- `qt_pages/base.py`
- `qt_components/navigation.py`
- `app_services/log_bus.py`
- `app_services/runtime_managers.py`

## 3. 目录结构建议

```text
gui_qt_app.py
qt_pages/
  __init__.py
  base.py
  about_page.py
  bot_test_page.py
  execute_page.py
  folder_monitor_page.py
  settings_page.py
qt_components/
  __init__.py
  navigation.py
app_services/
  __init__.py
  log_bus.py
  runtime_managers.py
```

## 4. 类设计

### 4.1 主窗口

`QuantumMainWindow`

- 继承 `QMainWindow`
- 左侧导航栏
- 中间 `QStackedWidget`
- 页面切换、窗口图标、统一样式

### 4.2 页面基类

`BasePage`

- 统一页面标题、说明、副标题
- 提供 `on_page_activated()`、`on_external_config_updated()` 空实现

### 4.3 日志总线

`LogBus`

- 继承 `QObject`
- 提供 `log_emitted` 信号
- 供运行管理器、页面、业务层适配

### 4.4 运行管理器

`MailRuleRuntimeManager`

- 管理邮件规则启动/停止/热刷新
- 后续接 `QThread` / `QThreadPool`

`FolderMonitorRuntimeManager`

- 管理文件夹监测启动/停止/热刷新

当前已不只是接口骨架，邮件检测页和文件夹检测页都已接入真实业务逻辑。

## 5. 页面迁移顺序

### 阶段一：骨架

- 主窗口
- 左侧导航
- 五个占位页面
- 日志总线
- 运行管理器接口

### 阶段二：设置页

- 邮箱配置
- 机器人别名
- 邮件规则
- 文件夹检测配置
- 路径设置
- 界面设置

### 阶段三：运行页

- 邮件检测页
- 文件夹检测页
- 独立日志与状态卡片
- 日志导出/清空/自动滚动

### 阶段四：辅助页

- 机器人测试页
- 关于页

## 6. 线程与信号建议

后续 Qt 版统一使用：

- UI 更新：`Signal / Slot`
- 长任务：`QThread` 或 `QRunnable + QThreadPool`
- 日志分发：通过 `LogBus.log_emitted`

中长期建议仍然逐步收敛到 Runtime Manager / QThreadPool 模式，但当前预览版允许页面先直接驱动已有业务线程，优先保证迁移速度和功能可用性。

## 7. 打包建议

Qt 版先与 Tk 版并存：

- Tk 入口：`gui_app.py`
- Qt 入口：`gui_qt_app.py`

后续成熟后再决定是否切换主入口，并新增独立 spec，例如：

- `QuantumBotQt.spec`

## 8. 当前进度

当前已落地：

- `gui_qt_app.py`
- `qt_pages/about_page.py`
- `qt_pages/bot_test_page.py`
- `qt_pages/execute_page.py`
- `qt_pages/folder_monitor_page.py`
- `qt_pages/settings_page.py`
- `qt_components/navigation.py`
- `app_services/log_bus.py`
- `app_services/webhook_alias_store.py`
- `app_services/folder_monitor_support.py`
- `requirements_gui_qt.txt`
- `start_gui_qt.bat`
- `QuantumBotQt.spec`
- `build_release_qt.ps1`

当前状态：

- Qt 预览版可以独立运行
- Qt 预览版可以独立打包
- Qt 打包已与 Tk 依赖解耦，体积已做过一轮收紧
- 正式主版本仍然以 Tk 版为准

下一步建议：

- 继续补齐 Qt 页面的细节体验
- 逐步把共享逻辑从 `desktop_pages/` 挪到更中立的 `app_services/`
- 评估是否进入“Qt 作为主入口”的切换阶段
