# 实训报告自动生成工具

这个项目用于把每日实训笔记整理成标准 Word 实训报告。脚本会读取 `daily_notes/YYYY-MM-DD.txt`，调用大模型生成课程目标、课程内容、课程详情和代码任务，再在本地沙盒中执行代码，把运行结果渲染成图片并写入 Word 报告。

## 功能

- 从 TXT 笔记生成结构化报告内容
- 支持 DeepSeek / OpenAI-compatible API
- 自动生成并执行 Python 代码任务
- 捕获标准输出和 matplotlib 图形输出
- 将源代码作为 Word 文本插入，将运行结果作为图片插入
- 支持代码执行失败后的自动修复
- 支持中文结果图片渲染和 Word 模板填充

## 目录结构

```text
.
├── scripts/
│   ├── generate_report.py   # 主入口：生成报告
│   ├── code_sandbox.py      # 代码沙盒执行与自动修复
│   └── render_output.py     # 运行结果图片渲染
├── config.example.yaml      # 示例配置，不包含真实 API Key
├── requirements.txt         # Python 依赖
├── 问题清单.md              # 已发现问题与处理记录
└── .uploads/                # 开发文档
```

本地运行时还需要准备这些目录或文件，但通常不提交到 GitHub：

```text
config.yaml                  # 本地真实配置
daily_notes/                 # 每日 TXT 笔记
template/实训报告模板.docx   # Word 报告模板
outputs/                     # 生成的报告
logs/                        # 日志
temp/                        # 临时文件
```

## 环境准备

建议使用 Python 3.10 或更高版本。

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

如果不激活虚拟环境，也可以直接使用：

```powershell
.venv\Scripts\python.exe scripts\generate_report.py --help
```

## 配置

复制示例配置：

```powershell
copy config.example.yaml config.yaml
```

然后修改 `config.yaml` 中的个人信息和模型配置。

推荐不要把 API Key 写进配置文件，而是设置环境变量：

```powershell
$env:DEEPSEEK_API_KEY="你的 DeepSeek API Key"
```

如果希望长期保存到 Windows 用户变量，可以在系统环境变量中新增：

```text
DEEPSEEK_API_KEY
```

设置后需要重新打开 PowerShell 才会生效。

## 运行

准备输入文件：

```text
daily_notes/2026-07-05.txt
```

执行：

```powershell
.venv\Scripts\python.exe scripts\generate_report.py --date 2026-07-05 --provider deepseek --model deepseek-chat
```

输出文件默认生成到：

```text
outputs/
```

## 常用参数

```powershell
--date       指定笔记日期，格式为 YYYY-MM-DD
--config     指定配置文件，默认 config.yaml
--provider   指定服务商配置，例如 deepseek 或 openai
--model      临时覆盖模型名称
--api-key    临时传入 API Key，不推荐长期使用
--base-url   临时覆盖 API 地址
--review     生成前打印摘要确认
```

示例：

```powershell
.venv\Scripts\python.exe scripts\generate_report.py --date 2026-07-05 --provider deepseek --model deepseek-chat
```

## Git 提交建议

建议 `.gitignore` 忽略本地配置、笔记、模板、输出和虚拟环境：

```gitignore
.venv/
temp/
logs/
outputs/
daily_notes/
template/
config.yaml
__pycache__/
*.pyc
.env
```

日常提交可以使用：

```powershell
git status
git add .
git commit -m "update project"
git push
```

如果某个本地文件已经被 Git 跟踪，但之后不想继续提交，例如 Word 模板，可以使用：

```powershell
git rm --cached "template/实训报告模板.docx"
git commit -m "remove local template from repository"
git push
```

这个命令不会删除本地文件，只会取消 Git 跟踪。

## 注意事项

- 不要提交真实 API Key。
- 不建议提交 `config.yaml`，请提交 `config.example.yaml`。
- 不建议提交 `daily_notes/`、`outputs/`、`logs/`、`temp/`。
- 如果 PowerShell 中中文文件名显示为转义字符，通常不影响 Git 使用。
- 如果模型返回非标准 JSON，脚本会尝试重试、修复解析、纯代码提取和本地模板兜底。
