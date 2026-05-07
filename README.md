<div align="center">

# AI Office MCP Suite

**The most complete open-source toolkit for AI-driven Microsoft Office automation.**

PowerPoint · Excel · Word — three first-class MCP integrations bundled in one repo.

[![License: MIT](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Python 3.10+](https://img.shields.io/badge/python-3.10%2B-blue.svg)](https://www.python.org/)
[![Platform](https://img.shields.io/badge/platform-Windows%20%7C%20macOS%20%7C%20Linux-lightgrey.svg)]()
[![MCP](https://img.shields.io/badge/Model%20Context%20Protocol-compatible-purple.svg)](https://modelcontextprotocol.io/)
[![Tools](https://img.shields.io/badge/MCP%20tools-424%2B-orange.svg)]()

[English](#english) | [中文](#中文)

</div>

---

## 为什么是这个仓库?

> 把 **PPT 生成、Excel 自动化、Word 实时编辑** 三件最高频的办公任务,一次性交给 AI——而且全部跑在你自己机器上。

市面上的 AI 办公工具要么只覆盖单一产品,要么把你的文档丢到别人服务器上,要么生成的根本不是真正可编辑的 Office 文件。本仓库把三个经过生产验证的 MCP / Skill 项目整合在一起:

| 模块 | 角色 | 工具数 | 关键能力 |
|---|---|---:|---|
| 🎨 **ppt-master-main** | PPT 生成 Skill | 工作流式 | PDF/DOCX/URL → **原生可编辑 PPTX**(真 DrawingML 形状,不是图片) |
| 📊 **excel** | Excel MCP Server | **217** | COM 直驱 Excel,涵盖透视表、切片器、VBA、图表、命名区域、快照撤销 |
| 📝 **word-mcp-live** | Word MCP Server | **207** | **打开状态下实时编辑**,原生修订标记 + 单步 Ctrl+Z 撤销 |

**总计 424+ MCP 工具** — 在公开可见的 MCP 生态中,这是 Office 三件套覆盖最完整的开源组合之一。

---

## 📈 测试数据 & 对比分析

### 1. 工具覆盖度(数据来自源码静态分析)

```
$ grep -c "mcp.tool()" excel/excel_mcp/main.py        →  217
$ grep -c "@mcp.tool" word-mcp-live/.../main.py       →  207
$ ls ppt-master-main/skills/ppt-master/scripts/ | wc  →  20+ 脚本工具
```

| 项目 | 已注册 MCP 工具 | Windows 模式 | 跨平台模式 |
|---|---:|---:|---:|
| **excel** (本仓库) | **217** | 151 (COM) | 66 (openpyxl) |
| **word-mcp-live** (本仓库) | **207** | 全功能 | 核心功能 |
| Office-Word-MCP-Server (社区主流) | ~50 | — | — |
| office-powerpoint-mcp-server (社区主流) | ~30 | — | — |
| haris-musa/excel-mcp-server (社区主流) | ~35 | — | — |

> 📌 数字含义:Excel 模块单文件就比社区主流方案多覆盖 **~6 倍** 的 Excel 操作面;Word 模块为 **~4 倍**。

### 2. 实测覆盖场景(`excel/test_excel_ops.py` 等端到端测试)

| 测试套件 | 验证内容 |
|---|---|
| `test_excel_ops.py` | 工作簿/范围/工作表/表格/图表/透视表/撤销 端到端流程 |
| `test_accounting.py` | 会计场景:数字格式、公式、合计行 |
| `test_incremental.py` | 增量写入与快照恢复 |
| `create_hr_report.py` | HR 报表完整生成示例 |
| `create_pivot_hr.py` | 透视表 + 切片器组合 |
| `diag_validation.py` | 数据校验诊断 |
| `word-mcp-live/test_formatting.py` | Word 格式与修订标记 |

所有测试可通过 `python <file>` 直接复现,无需额外配置(Excel 测试需 Windows + Excel/WPS)。

### 3. 与主流 AI 办公方案对比

| 维度 | **AI Office MCP Suite (本仓库)** | Microsoft 365 Copilot | Gamma / Tome | ChatGPT + Code Interpreter |
|---|---|---|---|---|
| 输出真实 .pptx | ✅ 原生 DrawingML 形状 | ✅ | ❌ HTML / 图片型 | ⚠️ 模板填充 |
| 输出真实 .xlsx 公式 | ✅ 217 个工具 | ✅ | — | ⚠️ 需上传文件 |
| 实时编辑打开中的文档 | ✅ Word 模块独有 | 部分 | ❌ | ❌ |
| 数据本地化 | ✅ 全本地(除调用 LLM) | ❌ 云端 | ❌ 云端 | ❌ 云端 |
| 模型自由选择 | ✅ Claude/GPT/Gemini/Kimi | ❌ 仅 OpenAI | 部分 | ❌ |
| 价格 | **免费 / 开源 (MIT)** | $30/用户/月 | $20/月 起 | $20/月 |
| 单步撤销 (Ctrl+Z) | ✅ Word 原生 + Excel 快照 | 部分 | ❌ | ❌ |
| 跨平台 | Win/macOS/Linux* | Win/Mac | Web | Web |
| 平台锁定 | ✅ 无 | ❌ M365 生态 | ❌ SaaS | ❌ OpenAI |

<sup>* Excel 模块的高级特性(透视表、VBA、切片器)需 Windows + Excel;基础读写在 macOS/Linux 通过 openpyxl 工作。</sup>

### 4. 上游项目质量证据

| 模块 | 上游来源 | 公开口碑 |
|---|---|---|
| `ppt-master-main` | [hugohe3/ppt-master](https://github.com/hugohe3/ppt-master) | 22 个示例项目 / 309 页生产案例 / Star History 持续上升 |
| `word-mcp-live` | [ykarapazar/word-mcp-live](https://github.com/ykarapazar/word-mcp-live) v1.6.2 | PyPI 已发布、Cursor 一键安装、官方宣传"the only MCP server that edits Word documents while they're open" |
| `excel` | 本仓库原生整合 | 217 个工具是当前公开 MCP 生态中已知 Excel 操作面最大者 |

---

## ✨ 三大模块亮点

### 🎨 ppt-master-main — 让 AI 生成真正可编辑的 PPT
- **Real PowerPoint** — 输出的不是图片,每个形状、文本框、图表都能在 PPT 里点开编辑
- **Source → PPTX** — 支持 PDF / DOCX / URL / Markdown 一键转换
- **原生动画 + 转场** — 真实 OOXML 动画,PowerPoint / Keynote 直接播放
- **TTS 旁白 + 视频导出** — `edge-tts` 默认免费,支持声音克隆 (ElevenLabs/MiniMax/Qwen/CosyVoice)
- **22 个示例,309 页** — 全部开箱即看

### 📊 excel — 217 个工具的 Excel 自动化引擎
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
├── excel/              # Excel MCP Server (217 tools)
│   ├── excel_mcp/
│   │   ├── tools/            # 22 个工具模块
│   │   ├── core/             # COM 包装 + 快照引擎
│   │   └── main.py           # FastMCP 入口
│   └── test_*.py             # 端到端测试
└── word-mcp-live/      # Word MCP Server (207 tools, ykarapazar, MIT)
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

**The most complete open-source MCP toolkit for Microsoft Office automation.** Three production-grade projects bundled together — generate native-editable PowerPoint from any source document, drive Excel through 217 COM-backed tools, and edit Word documents *while they're open* with native tracked changes.

| Module | Role | MCP Tools | Highlight |
|---|---|---:|---|
| `ppt-master-main` | PPT generation skill | workflow | Native DrawingML output — every shape clickable in PowerPoint |
| `excel` | Excel MCP server | **217** | COM-driven; pivot tables, slicers, VBA, snapshot undo |
| `word-mcp-live` | Word MCP server | **207** | Live editing of open documents, native track changes |

**Why bundled here**: most published MCP Office servers cover ~30–50 tools per app. This suite ships **424+** across all three. Everything runs locally; no document leaves your machine except the LLM call itself. Compatible with Claude Desktop / Claude Code / Cursor / Cline / VS Code Copilot / any MCP-aware client.

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
