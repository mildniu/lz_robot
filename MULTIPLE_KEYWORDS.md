# 多主题关键字功能说明

## 🎉 新功能

邮件监测现在支持**多个主题关键字**！只要邮件主题包含任一关键字，就会被处理。

## 📝 配置方法

### 1. 环境变量配置

在 `.env` 文件中设置 `SUBJECT_KEYWORDS`：

```bash
# 多个关键字用逗号分隔
SUBJECT_KEYWORDS=衡水装维营销日报,政企序列日报表,销售报表

# 或者单个关键字
SUBJECT_KEYWORDS=衡水装维营销日报
```

### 2. GUI配置

在设置页面的"📨 邮件处理配置"中：
- **字段名**：主题关键字(逗号分隔)
- **输入示例**：`衡水装维营销日报,政企序列日报表,销售报表`
- **说明**：多个关键字用英文逗号`,`分隔

## 🔧 匹配规则

### 匹配逻辑
- 邮件主题**包含**任一关键字即匹配
- 匹配是**部分匹配**，不需要完全相等
- 关键字匹配**不区分大小写**（取决于IMAP服务器）

### 匹配示例

假设配置关键字：`衡水装维营销日报,政企序列,销售`

| 邮件主题 | 是否匹配 | 匹配关键字 |
|---------|---------|-----------|
| 衡水装维营销日报 | ✅ 是 | 衡水装维营销日报 |
| 政企序列日报表 | ✅ 是 | 政企序列 |
| 销售报表2026 | ✅ 是 | 销售 |
| 周报-销售部 | ✅ 是 | 销售 |
| 财务报表 | ❌ 否 | 无 |
| 人事通知 | ❌ 否 | 无 |

## 🔄 兼容性

### 向后兼容
- **保留** `SUBJECT_EXACT` 支持
- 如果只设置了 `SUBJECT_EXACT`，会自动转换为单关键字
- 如果同时设置，`SUBJECT_KEYWORDS` 优先

### 迁移指南

**旧配置**：
```bash
SUBJECT_EXACT=衡水装维营销日报
```

**新配置**：
```bash
SUBJECT_KEYWORDS=衡水装维营销日报
```

**多关键字配置**：
```bash
SUBJECT_KEYWORDS=衡水装维营销日报,政企序列日报表,销售报表
```

## 📊 测试结果

```
[OK] Config loaded successfully
     Keywords: ['衡水装维营销日报', '政企序列日报表', '销售报表']
     Count: 3 keywords

[OK] IMAP client imported
[INFO] find_latest_uid now accepts list of keywords

Test keywords: ['衡水装维营销日报', '政企序列', '销售']

  [MATCH] '衡水装维营销日报'
         -> Matched keyword: '衡水装维营销日报'
  [MATCH] '政企序列日报表'
         -> Matched keyword: '政企序列'
  [MATCH] '销售报表2026'
         -> Matched keyword: '销售'
  [MATCH] '周报-销售部'
         -> Matched keyword: '销售'
  [SKIP] '其他邮件'
```

## 🎯 使用场景

### 场景1: 多部门报表汇总
```bash
SUBJECT_KEYWORDS=销售日报,财务周报,人事月报,技术总结
```
收集来自不同部门的各种报表。

### 场景2: 多项目监控
```bash
SUBJECT_KEYWORDS=项目A周报,项目B进度,项目C里程碑
```
监控多个项目的相关邮件。

### 场景3: 灵活的关键字
```bash
SUBJECT_KEYWORDS=日报,周报,月报,季报,年报
```
收集各种时间维度的报告。

### 场景4: 品牌+产品组合
```bash
SUBJECT_KEYWORDS=华为,小米,OPPO,vivo
```
监控特定品牌的产品邮件。

## ⚙️ 技术实现

### 配置文件
- **位置**: `mail_forwarder/config.py`
- **变更**: `subject_exact: str` → `subject_keywords: list[str]`
- **解析**: 支持逗号分隔，自动去除空格

### IMAP客户端
- **位置**: `mail_forwarder/imap_client.py`
- **方法**: `find_latest_uid(subject_keywords: list[str])`
- **逻辑**: 遍历所有关键字，任一匹配即返回

### Worker
- **位置**: `mail_forwarder/worker.py`
- **更新**: 使用 `subject_keywords` 而非 `subject_exact`

### GUI
- **位置**: `gui_app_v3.py`
- **字段**: "主题关键字(逗号分隔)"
- **保存**: 自动保存为 `SUBJECT_KEYWORDS`

## 📝 注意事项

1. **关键字长度**
   - 建议关键字不要太短（如"报表"）
   - 过短的关键字可能匹配过多邮件

2. **特殊字符**
   - 支持中文、英文、数字
   - 不需要转义特殊字符

3. **空格处理**
   - 关键字前后的空格会被自动去除
   - 关键字中间的空格会被保留

4. **大小写**
   - 匹配行为取决于IMAP服务器
   - 大多数IMAP服务器不区分大小写

## 🚀 立即使用

### 1. 更新 .env 文件
```bash
SUBJECT_KEYWORDS=关键字1,关键字2,关键字3
```

### 2. 重启程序
```bash
python gui_app_v3.py
```

### 3. 查看日志
```
[INFO] 主题关键字: ['关键字1', '关键字2', '关键字3']
```

### 4. 测试匹配
点击"🧪 测试"按钮查看是否能正确匹配邮件。

## 🎊 总结

**之前**：只能精确匹配一个主题
```
SUBJECT_EXACT=衡水装维营销日报
```

**现在**：可以匹配多个关键字
```
SUBJECT_KEYWORDS=衡水装维营销日报,政企序列日报表,销售报表
```

**优势**：
- ✅ 更灵活的邮件匹配
- ✅ 支持多个监测主题
- ✅ 向后兼容旧配置
- ✅ 部分匹配，容错性强

享受更强大的邮件监测功能！🎉
