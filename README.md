<div align="center">

# AI Office MCP Suite

**The most complete open-source toolkit for AI-driven Microsoft Office automation.**

PowerPoint · Excel · Word — three first-class MCP integrations bundled in one repo.

[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)]()
[![MCP](https://img.shields.io/badge/Model%20Context%20Protocol-compatible-purple.svg)](https://modelcontextprotocol.io/)
[![Tools](https://img.shields.io/badge/MCP%20tools-358%20registered-orange.svg)]()
[![Tests](https://img.shields.io/badge/tests-287%2F287%20passed-brightgreen.svg)]()

[English](#english) | [中文](#中文)

</div>

---

## 为什么是这个仓库?

> 把 **PPT 生成、Excel 自动化、Word 实时编辑** 三件最高频的办公任务,一次性交给 AI——而且全部跑在你自己机器上。

市面上的 AI 办公工具要么只覆盖单一产品,要么把你的文档丢到别人服务器上,要么生成的根本不是真正可编辑的 Office 文件。本仓库把三个经过生产验证的 MCP / Skill 项目整合在一起:

| 模块 | 角色 | 已注册工具 | 关键能力 |
|---|---|---:|---|
| 🎨 **ppt-master-main** | PPT 生成 Skill | 工作流式 | PDF/DOCX/URL → **原生可编辑 PPTX**(真 DrawingML 形状,不是图片) |
| 📊 **excel** | Excel MCP Server | **152** (Win) / 67 (跨平台) | COM 直驱 Excel,涵盖透视表、切片器、VBA、图表、命名区域、快照撤销 |
| 📝 **word-mcp-live** | Word MCP Server | **206** | **打开状态下实时编辑**,原生修订标记 + 单步 Ctrl+Z 撤销 |

**Windows 全栈 358 个 MCP 工具实测注册成功** — 在公开可见的 MCP 生态中,这是 Office 三件套覆盖最完整的开源组合之一。

---

## 📈 测试数据 & 对比分析

> **以下所有数字均为实测**(2026-05-08,Windows 11 + Office 16.0 + Python 3.13.12)。
> 复现命令见每个表格下方,benchmark 脚本在 [`benchmarks/`](./benchmarks/)。

### 1. 工具注册数(实测,非 grep 估算)

```bash
$ python benchmarks/bench_startup.py
```

| 项目 | 实测注册工具数 | Windows 模式 | 跨平台模式 |
|---|---:|---:|---:|
| **excel** (本仓库) | **152** | 151 (COM) + 1 (excel_diff) | 66 (openpyxl) + 1 |
| **word-mcp-live** (本仓库) | **206** | 全功能 | 核心功能 |
| Office-Word-MCP-Server (社区主流) | ~50 | — | — |
| office-powerpoint-mcp-server (社区主流) | ~30 | — | — |
| haris-musa/excel-mcp-server (社区主流) | ~35 | — | — |

> 📌 数字含义:Excel 模块比社区主流 Excel MCP 多 **~4 倍** 操作面;Word 模块多 **~4 倍**。

### 2. 端到端测试通过率(本机实测)

```bash
$ cd excel && python test_excel_ops.py && python test_accounting.py && python test_incremental.py
$ cd ../word-mcp-live && python test_formatting.py && pytest tests/
```

| 测试套件 | 通过率 | 耗时 | 验证内容 |
|---|---:|---:|---|
| `excel/test_excel_ops.py` | **32 / 32** | 5.8s | 文件/范围/工作表/Excel 表格/透视表/图表/撤销 全流程 |
| `excel/test_accounting.py` | **197 / 197** | 22.1s | 会计场景:数字格式、公式、合计行、舍入 |
| `excel/test_incremental.py` | **56 / 56** | 13.4s | 增量写入与快照回滚 |
| `word-mcp-live/test_formatting.py` | ✅ 通过 | 0.4s | docx 创建 + 5 段不同字体/字号/粗斜体 |
| `word-mcp-live/tests/test_convert_to_pdf.py` | **1 / 1** | 10.8s | 真实 .docx → .pdf 转换 |
| **总计** | **287 / 287 (100%)** | **52.5s** | |

### 3. MCP 服务冷启动 & 工具枚举性能

```bash
$ python benchmarks/bench_startup.py
```

| Server | 冷启动(子进程,3 次中位数) | 工具枚举(进程内) |
|---|---:|---:|
| `excel.excel_mcp.main` | **1898 ms** | 1859 ms (152 tools) |
| `word_document_server.main` | **1536 ms** | 263 ms (206 tools) |

> Word 工具枚举更快是因为 Excel 模块预加载了 22 个工具子模块的代码。Cold-start 一次,之后 stdio 服务持续响应。

### 4. 与主流 AI 办公方案对比

| 维度 | **AI Office MCP Suite (本仓库)** | Microsoft 365 Copilot | Gamma / Tome | ChatGPT + Code Interpreter |
|---|---|---|---|---|
| 输出真实 .pptx | ✅ 原生 DrawingML 形状 | ✅ | ❌ HTML / 图片型 | ⚠️ 模板填充 |
| 输出真实 .xlsx 公式 | ✅ 152 个 COM 工具 | ✅ | — | ⚠️ 需上传文件 |
| 实时编辑打开中的文档 | ✅ Word 模块独有 | 部分 | ❌ | ❌ |
| 数据本地化 | ✅ 全本地(除调用 LLM) | ❌ 云端 | ❌ 云端 | ❌ 云端 |
| 模型自由选择 | ✅ Claude/GPT/Gemini/Kimi | ❌ 仅 OpenAI | 部分 | ❌ |
| 价格 | **免费 / 开源 (MIT)** | $30/用户/月 | $20/月 起 | $20/月 |
| 单步撤销 (Ctrl+Z) | ✅ Word 原生 + Excel 快照 | 部分 | ❌ | ❌ |
| 跨平台 | Win/macOS/Linux* | Win/Mac | Web | Web |
| 平台锁定 | ✅ 无 | ❌ M365 生态 | ❌ SaaS | ❌ OpenAI |

<sup>* Excel 模块的高级特性(透视表、VBA、切片器)需 Windows + Excel;基础读写在 macOS/Linux 通过 openpyxl 工作。</sup>

### 5. 上游项目质量证据

| 模块 | 上游来源 | 公开口碑 |
|---|---|---|
| `ppt-master-main` | [hugohe3/ppt-master](https://github.com/hugohe3/ppt-master) | 22 个示例项目 / 309 页生产案例 / Star History 持续上升 |
| `word-mcp-live` | [ykarapazar/word-mcp-live](https://github.com/ykarapazar/word-mcp-live) v1.6.2 | PyPI 已发布、Cursor 一键安装、官方宣传"the only MCP server that edits Word documents while they're open" |
| `excel` | 本仓库原生整合 | 152 个工具是当前公开 MCP 生态中已知 Excel 操作面最大者 |

---

## ✨ 三大模块亮点

## 🔬 自己复现测试

```bash
# 1. 安装依赖
pip install -r excel/requirements.txt -r word-mcp-live/requirements.txt

# 2. 跑基准(冷启动 + 工具枚举,无需 Office 弹窗)
python benchmarks/bench_startup.py

# 3. 跑 Excel 端到端(需 Windows + Excel/WPS;会短暂打开 Excel)
cd excel && python test_excel_ops.py && python test_accounting.py && python test_incremental.py

# 4. 跑 Word(需要 UTF-8 控制台):
cd ../word-mcp-live
$env:PYTHONIOENCODING="utf-8"; $env:PYTHONUTF8="1"
python test_formatting.py
pytest tests/
```

**测试硬件**:Windows 11 / Office 16.0 / Python 3.13.12 / 普通笔记本。
所有数字源自单次干净环境运行;欢迎 PR 你自己机器上的结果。

---

### 🎨 ppt-master-main — 让 AI 生成真正可编辑的 PPT
- **Real PowerPoint** — 输出的不是图片,每个形状、文本框、图表都能在 PPT 里点开编辑
- **Source → PPTX** — 支持 PDF / DOCX / URL / Markdown 一键转换
- **原生动画 + 转场** — 真实 OOXML 动画,PowerPoint / Keynote 直接播放
- **TTS 旁白 + 视频导出** — `edge-tts` 默认免费,支持声音克隆 (ElevenLabs/MiniMax/Qwen/CosyVoice)
- **22 个示例,309 页** — 全部开箱即看

### 📊 excel — 152 个工具的 Excel 自动化引擎(实测全通过)
- **COM 直驱** — Windows 下直接控制运行中的 Excel/WPS,所见即所得
- **跨平台降级** — macOS/Linux 自动切换 openpyxl,核心读写不打折
- **完整覆盖** — 透视表 / 切片器 / 图表 / 公式 / 命名区域 / 数据验证 / VBA / 截图 / 窗口管理 / PDF 导出
- **快照撤销** — 每次写操作自动快照,`excel_undo` 一键回滚
- **excel_diff** — 不打开文件就能 diff 两个 .xlsx
- **会话管理** — 多文件并发操作,通过 `workbook_name` 路由

### 📝 word-mcp-live — 唯一支持"打开状态实时编辑"的 Word MCP
- **Live editing** — 文档开着不用关,AI 直接改,实时刷新
- **原生修订标记** — 真正的 Word Track Changes,带你的署名
- **线程化批注** — 添加 / 回复 / 解决 / 删除,像人类审稿人一样
- **每个动作单步 Ctrl+Z** — 撤销不是 XML hack,是真撤销
- **布局诊断** — 打印灾难发生前先报警
- **公式 + 交叉引用** — 自动更新

---

## 🚀 快速开始

### 安装

```bash
git clone https://github.com/lingfan36/ai-office-mcp.git
cd ai-office-mcp
```

按需安装各模块依赖:

```bash
# PPT
pip install -r ppt-master-main/requirements.txt

# Excel (Windows 全功能;macOS/Linux 自动跳过 pywin32)
pip install -r excel/requirements.txt

# Word
pip install -r word-mcp-live/requirements.txt
```

### Claude Desktop 配置示例

```jsonc
{
  "mcpServers": {
    "excel": {
      "command": "python",
      "args": ["-m", "excel_mcp.main"],
      "cwd": "D:/path/to/ai-office-mcp/excel"
    },
    "word": {
      "command": "uvx",
      "args": ["word-mcp-live"],
      "env": {
        "MCP_AUTHOR": "Your Name",
        "MCP_AUTHOR_INITIALS": "YN"
      }
    }
  }
}
```

PPT 模块是 **Skill / 工作流**,不需要注册为 MCP server,放在 AI IDE(Claude Code / Cursor)里直接对话使用,详见 `ppt-master-main/README.md`。

---

## 🗂️ 仓库结构

```
ai-office-mcp/
├── ppt-master-main/    # PPT 生成 Skill (Hugo He, MIT)
│   ├── skills/ppt-master/    # SKILL.md + 工作流 + 模板
│   ├── examples/             # 22 个示例项目,309 页
│   └── docs/                 # 文档、FAQ、技术设计
├── excel/              # Excel MCP Server (152 tools on Windows / 67 cross-platform)
│   ├── excel_mcp/
│   │   ├── tools/            # 22 个工具模块
│   │   ├── core/             # COM 包装 + 快照引擎
│   │   └── main.py           # FastMCP 入口
│   └── test_*.py             # 端到端测试
└── word-mcp-live/      # Word MCP Server (206 tools, ykarapazar, MIT)
    ├── word_document_server/
    └── tests/
```

---

## ⚖️ 许可与归属

本仓库为 **整合发行版**,各子项目原作者及许可声明如下:

| 子项目 | 上游作者 | 许可 |
|---|---|---|
| `ppt-master-main` | Hugo He ([hugohe3](https://github.com/hugohe3)) | MIT |
| `word-mcp-live` | [ykarapazar](https://github.com/ykarapazar) | MIT |
| `excel` | 本仓库整合 | MIT |

整合工作: [@lingfan36](https://github.com/lingfan36)。如果你觉得有用,请给上游项目都点个 ⭐。

---

<div align="center">

## English

**The most complete open-source MCP toolkit for Microsoft Office automation.** Three production-grade projects bundled together — generate native-editable PowerPoint from any source document, drive Excel through 152 registered tools on Windows, and edit Word documents *while they're open* with native tracked changes.

| Module | Role | MCP Tools (measured) | Highlight |
|---|---|---:|---|
| `ppt-master-main` | PPT generation skill | workflow | Native DrawingML output — every shape clickable in PowerPoint |
| `excel` | Excel MCP server | **152** (Win) / 67 (cross-platform) | COM-driven; pivot tables, slicers, VBA, snapshot undo |
| `word-mcp-live` | Word MCP server | **206** | Live editing of open documents, native track changes |

**Why bundled here**: most published MCP Office servers cover ~30–50 tools per app. This suite ships **358 measured tools** across all three on Windows. Everything runs locally; no document leaves your machine except the LLM call itself. Compatible with Claude Desktop / Claude Code / Cursor / Cline / VS Code Copilot / any MCP-aware client.

**Verified test results** (Windows 11 + Office 16.0 + Python 3.13.12, 2026-05-08):
- Excel: **285 / 285 assertions passed in 41.3s** (3 end-to-end suites)
- Word: **2 / 2 test suites passed in 11.2s** (formatting + docx→pdf)
- Cold-start: Excel 1898 ms, Word 1536 ms (3-run median, fresh subprocess)

See the Chinese section above for a detailed feature comparison vs. **Microsoft 365 Copilot**, **Gamma / Tome**, and **ChatGPT + Code Interpreter**.

### One-line install

```bash
git clone https://github.com/lingfan36/ai-office-mcp.git && cd ai-office-mcp
```

Then install per-module requirements as needed (see Quick Start above).

</div>

---

<div align="center">

⭐ If this saves you a few hours, please star the upstream projects too:
[ppt-master](https://github.com/hugohe3/ppt-master) · [word-mcp-live](https://github.com/ykarapazar/word-mcp-live)

Made with ❤️ — MIT licensed.

</div>
