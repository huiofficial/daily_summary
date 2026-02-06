# Daily Summary

用于汇总多个 Git 仓库提交记录并生成 Excel 报告的 Python 工具。

[![License](https://img.shields.io/badge/license-Apache%202.0-blue.svg)](LICENSE)
[![Python](https://img.shields.io/badge/python-3.7+-blue.svg)](https://www.python.org/downloads/)

## 目录

- [项目背景](#项目背景)
- [核心特性](#核心特性)
- [技术栈](#技术栈)
- [快速开始](#快速开始)
- [使用示例](#使用示例)
- [输出说明](#输出说明)
- [Roadmap](#roadmap)
- [文档说明](#文档说明)
- [贡献方式](#贡献方式)
- [License](#license)

## 项目背景

在日常开发工作中，开发者经常需要同时维护多个 Git 仓库。为了追踪和汇报工作进展，需要从不同仓库中提取指定时间范围内的提交记录，并进行汇总分析。

本工具解决以下问题：
- 手动遍历多个仓库提取提交信息效率低下
- 需要统一的报告格式用于工作汇报或进度追踪
- 跨仓库提交统计需要重复执行 git 命令

适用于需要定期汇总多个项目工作成果的场景，如日报、周报生成，项目进度追踪等。

**work_summary**：按用户汇总参与情况，输出每个用户在各项目的提交数、代码行贡献（新增/删除/净增）及功能点数量（`feat` 类型提交数），便于工作汇报与按用户统计。

## 核心特性

- 递归扫描目录下的所有 Git 仓库
- 支持多种时间范围查询：今日、昨日、最近一周、自定义日期区间
- 按作者过滤提交记录
- 生成包含提交详情的 Excel 报告
- 统计文件变更、代码增减行数
- 可选使用 Ollama 生成仓库级别的工作总结（基于提交信息）
- **work_summary**：按用户统计参与项目、每项目代码贡献行数及功能点（feat 提交数）

## 技术栈

**后端**
- Python 3.7+
- pandas：数据处理与报告生成（支持 Excel、CSV、JSON）
- openpyxl：Excel 文件操作
- toml：配置文件解析
- ollama（可选）：本地 AI 模型调用

**工具**
- Git：提交记录查询
- Shell：命令行包装脚本

**依赖管理**
- requirements.txt

## 快速开始

### 环境要求

- Python 3.7 或更高版本
- Git 命令行工具
- （可选）Ollama 服务，用于生成 AI 总结

### 安装依赖

```bash
pip install -r requirements.txt
```

### （可选）配置 Ollama

如需使用 AI 总结功能，请安装并启动 Ollama：

```bash
# 安装 Ollama（如未安装）
# 访问 https://ollama.com/ 获取安装说明

# 拉取模型
ollama pull qwen3:0.6b

# 启动服务
ollama serve
```

### 最小示例

```bash
# 汇总今日提交（日报）
python daily_summary.py /path/to/work/directory --today

# 或使用 Shell 包装脚本
./daily_summary.sh /path/to/work/directory --today

# 工作汇总：按项目统计参与项目、代码行数、功能点
python work_summary.py /path/to/work/directory --today
```

### 配置文件（可选）

工具支持通过配置文件设置默认选项。创建 `daily_summary_config.toml` 文件（参考 `daily_summary_config.toml.example`）：

```toml
default_time_range = "today"
output_dir = "."
output_format = "excel"
author = null
ollama_model = "qwen3:0.6b"
enable_ai_summary = true
retry_attempts = 3
retry_delay = 1.0
commit_message_template = "{repository} - {author} - {message}"
incremental_state_file = ".daily_summary_state.json"
```

配置文件搜索顺序：
1. 当前工作目录
2. 脚本所在目录
3. 用户主目录

## 使用示例

### 时间范围查询

**今日提交：**
```bash
python daily_summary.py /path/to/work/dir --today
```

**昨日提交：**
```bash
python daily_summary.py /path/to/work/dir --yesterday
```

**最近一周（7天）：**
```bash
python daily_summary.py /path/to/work/dir --lastweek
```

**自定义日期范围：**
```bash
python daily_summary.py /path/to/work/dir --start 2024-01-01 --end 2024-01-07
```

### 过滤与输出

**按作者过滤：**
```bash
python daily_summary.py /path/to/work/dir --today --author "Your Name"
```

**指定输出文件：**
```bash
python daily_summary.py /path/to/work/dir --today -o custom_report.xlsx
```

### 导出格式

**Excel 格式（默认）：**
```bash
python daily_summary.py /path/to/work/dir --today -o report.xlsx
# 或
python daily_summary.py /path/to/work/dir --today --format excel
```

**CSV 格式：**
```bash
python daily_summary.py /path/to/work/dir --today -o report.csv
# 或
python daily_summary.py /path/to/work/dir --today --format csv
```

**JSON 格式：**
```bash
python daily_summary.py /path/to/work/dir --today -o report.json
# 或
python daily_summary.py /path/to/work/dir --today --format json
```

### 配置文件

**使用配置文件：**
```bash
# 创建配置文件 daily_summary_config.toml（参考 daily_summary_config.toml.example）
python daily_summary.py /path/to/work/dir --config daily_summary_config.toml
```

配置文件支持设置默认时间范围、输出目录、输出格式、作者过滤等选项。

### 增量更新

**启用增量更新模式（只处理新提交）：**
```bash
python daily_summary.py /path/to/work/dir --today --incremental
```

增量模式会记录上次处理的最后一个提交，下次运行时只处理新提交，提高处理效率。

### 提交信息模板化

**使用自定义模板格式化提交信息：**
```bash
python daily_summary.py /path/to/work/dir --today --template "{repository}[{author}]: {message}"
```

可用占位符：`{repository}`, `{author}`, `{email}`, `{date}`, `{commit_hash}`, `{message}`

### 命令行参数

- `work_dir`：包含 Git 仓库的目录路径（可通过配置文件设置）
- `--today`：汇总今日提交
- `--yesterday`：汇总昨日提交
- `--lastweek`：汇总最近一周提交（7天）
- `--start YYYY-MM-DD`：自定义范围的起始日期
- `--end YYYY-MM-DD`：自定义范围的结束日期
- `--output, -o`：输出文件路径（默认自动生成）
- `--format, -f`：输出格式（excel、csv、json、auto），默认 auto（根据文件扩展名自动检测）
- `--author`：按作者姓名或邮箱过滤提交
- `--config, -c`：配置文件路径
- `--incremental`：启用增量更新模式（只处理新提交）
- `--template`：提交信息模板（使用 {field} 占位符）
- `--no-ai-summary`：禁用 AI 总结生成

## 输出说明

### Excel 格式（默认）

生成包含以下工作表的 Excel 文件：

1. **Commits**：提交详情
   - 仓库名称
   - 提交日期时间
   - 作者姓名和邮箱
   - 提交哈希（短哈希）
   - 提交信息
   - 文件变更数
   - 新增行数
   - 删除行数

2. **Summary**：汇总统计
   - 查询时间段
   - 提交总数
   - 仓库总数
   - 文件变更总数
   - 新增代码总行数
   - 删除代码总行数
   - 唯一作者数

3. **Repository Summaries**（可选，需 Ollama 可用）
   - 各仓库的 AI 生成工作总结
   - 基于提交信息自动生成

### CSV 格式

生成以下 CSV 文件：
- `report.csv`：提交详情
- `report.summary.csv`：汇总统计
- `report.summaries.csv`：仓库总结（如果启用 AI 总结）

### JSON 格式

生成包含以下结构的 JSON 文件：
- `period`：查询时间段
- `summary`：汇总统计数据对象
- `commits`：提交详情数组
- `repository_summaries`：仓库总结对象（如果启用 AI 总结）

**注意事项：**
- 工具会跳过隐藏目录（以 `.` 开头）和常见非仓库目录（如 `node_modules`、`venv`、`__pycache__`）
- 如未指定时间选项，默认查询今日提交
- 如未指定输出文件，将按日期自动生成文件名
- 若 Ollama 不可用，工具将跳过仓库总结，仍会生成提交报告
- AI 总结生成时间取决于仓库数量和提交数量

### Work Summary（工作汇总）

`work_summary.py` 在相同时间范围内，**按用户（以邮箱区分）**统计参与情况：每个邮箱参与了哪些项目、每个项目贡献了多少行代码、增加了多少功能点。

**功能点**：统计 commit message 符合 Conventional Commits 中 **feat** 类型的提交数量（如 `feat: 新功能`、`feat(api): 接口`）。可选 `--fix-as-feature` 将 `fix` 类型也计入功能点。

**用法示例：**
```bash
# 今日
python work_summary.py /path/to/work/dir --today

# 最近一周
python work_summary.py /path/to/work/dir --lastweek

# 自定义日期并指定输出
python work_summary.py /path/to/work/dir --start 2024-01-01 --end 2024-01-07 -o work_summary.xlsx
```

**输出（Excel）：**
- **By User & Project**：邮箱、项目、提交数、新增行数、删除行数、净增行数、功能点数量、**commits**（该用户在该项目下的每条 commit：短 hash、日期、message，多行展示）
- **By User**：按邮箱汇总：邮箱、参与项目数、总提交数、总新增/删除/净增行、总功能点数
- **Summary**：统计周期、用户数（按邮箱）、项目数、总提交数、总新增/删除/净增行、总功能点数

支持与 daily_summary 相同的配置（`work_dir`、`author`、`output_dir`、`output_format`、时间范围默认值等），输出格式支持 excel、csv、json，默认文件名为 `work_summary_YYYYMMDD.xlsx`。**默认仅统计项目名以 `zgzl` 开头的仓库**；可在配置中设置 `work_summary_project_prefix` 或使用 `--project-prefix` 修改前缀，使用 `--project-prefix ""` 可统计全部项目。

## Roadmap

- [x] 多仓库递归扫描
- [x] 时间范围查询支持
- [x] Excel 报告生成
- [x] 提交统计信息
- [x] Ollama AI 总结集成
- [x] 配置文件支持（默认时间范围、输出目录等）
- [x] 导出格式扩展（CSV、JSON）
- [x] 增量更新支持
- [x] 提交信息模板化处理
- [x] 错误处理与重试机制优化

## 文档说明

本文档（README.md）提供项目概览和快速上手指南。详细使用说明、API 文档和高级配置请参考：

- `docs/` 目录（如有）
- 代码注释
- 命令行帮助：`python daily_summary.py --help`

**多语言支持：**
- README.md（中文）
- README.en.md（英文，如有需要可补充）

## 贡献方式

欢迎提交 Issue 和 Pull Request。

**Issue 规范：**
- 使用清晰的问题描述
- 包含复现步骤和环境信息
- 对于功能请求，说明使用场景

**Pull Request 规范：**
- 保持代码风格一致
- 添加必要的注释和文档
- 确保现有功能不受影响
- 如适用，更新相关文档

## License

Apache License 2.0

详见 [LICENSE](LICENSE) 文件。
