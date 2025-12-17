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

## 核心特性

- 递归扫描目录下的所有 Git 仓库
- 支持多种时间范围查询：今日、昨日、最近一周、自定义日期区间
- 按作者过滤提交记录
- 生成包含提交详情的 Excel 报告
- 统计文件变更、代码增减行数
- 可选使用 Ollama 生成仓库级别的工作总结（基于提交信息）

## 技术栈

**后端**
- Python 3.7+
- pandas：数据处理与 Excel 生成
- openpyxl：Excel 文件操作
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
# 汇总今日提交
python daily_summary.py /path/to/work/directory --today

# 或使用 Shell 包装脚本
./daily_summary.sh /path/to/work/directory --today
```

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

### 命令行参数

- `work_dir`（必需）：包含 Git 仓库的目录路径
- `--today`：汇总今日提交
- `--yesterday`：汇总昨日提交
- `--lastweek`：汇总最近一周提交（7天）
- `--start YYYY-MM-DD`：自定义范围的起始日期
- `--end YYYY-MM-DD`：自定义范围的结束日期
- `--output, -o`：输出 Excel 文件路径（默认自动生成）
- `--author`：按作者姓名或邮箱过滤提交

## 输出说明

工具生成包含以下工作表的 Excel 文件：

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

**注意事项：**
- 工具会跳过隐藏目录（以 `.` 开头）和常见非仓库目录（如 `node_modules`、`venv`、`__pycache__`）
- 如未指定时间选项，默认查询今日提交
- 如未指定输出文件，将按日期自动生成文件名
- 若 Ollama 不可用，工具将跳过仓库总结，仍会生成提交报告
- AI 总结生成时间取决于仓库数量和提交数量

## Roadmap

- [x] 多仓库递归扫描
- [x] 时间范围查询支持
- [x] Excel 报告生成
- [x] 提交统计信息
- [x] Ollama AI 总结集成
- [ ] 配置文件支持（默认时间范围、输出目录等）
- [ ] 导出格式扩展（CSV、JSON）
- [ ] 增量更新支持
- [ ] 提交信息模板化处理
- [ ] 错误处理与重试机制优化

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
