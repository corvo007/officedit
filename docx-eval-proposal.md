# office-eval：轻量级 OpenXML 脚本运行器

给 AI agent 一个 VM + API 文档，而非封装命令。支持 Word (.docx)、Excel (.xlsx)、PowerPoint (.pptx) 全家桶。

## 动机

### 当前痛点

- **python-docx / openpyxl / docx.js 等三方库**：对 OpenXML 规范实现残缺，编辑复杂文档（多级列表、域代码、交叉引用、复杂样式继承）100% 会坏。Claude 官方的 docx skill 也是基于这些库，只能创建简单文档，不能可靠编辑
- **OpenXML SDK（C#）**：微软自己的格式自己的 SDK，覆盖完整，但需要 ~500 MB .NET SDK 才能 `dotnet run`
- **officecli 等封装工具**：设计者要预判所有需求，漏了就没法用（选择器不灵活、搜索漏数据、不暴露底层属性）
- **信息差**：大部分人不知道 OpenXML SDK 的存在，搜到的全是 python-docx → 体验差 → 结论"AI 不能可靠编辑 Office 文档"。实际上是工具选错了

### 核心洞察

"AI 不能可靠编辑 Word/Excel/PPT" 不是能力问题，是工具链问题。OpenXML SDK 完全能胜任，但被 .NET SDK 的安装门槛挡在了外面。

## 方案

构建一个自包含的 CLI 工具 `office-eval`，内嵌 Roslyn 编译器 + OpenXML SDK，能在运行时接受 C# 脚本并编译执行。下载即用，无需安装 .NET SDK。

### 架构

```
office-eval.exe（自包含，零外部依赖）
├── Roslyn 编译器（Microsoft.CodeAnalysis.CSharp.Scripting）
├── OpenXML SDK（DocumentFormat.OpenXml）
│   ├── WordprocessingDocument → .docx
│   ├── SpreadsheetDocument   → .xlsx
│   └── PresentationDocument  → .pptx
└── 脚本加载器（读取 .csx 文件 → 编译 → 执行）
```

### 依赖包

```xml
<PackageReference Include="Microsoft.CodeAnalysis.CSharp.Scripting" Version="4.*" />
<PackageReference Include="DocumentFormat.OpenXml" Version="3.*" />
```

### 使用方式

```bash
# 运行脚本文件
office-eval fix-subscript.csx
```

```csharp
// Word 示例：批量设置下标
var doc = WordprocessingDocument.Open(@"thesis.docx", true);
var body = doc.MainDocumentPart.Document.Body;
foreach (var run in body.Descendants<Run>().Where(r => ...)) {
    // ...
}
doc.MainDocumentPart.Document.Save();
doc.Dispose();
```

```csharp
// Excel 示例：读取数据
var doc = SpreadsheetDocument.Open(@"data.xlsx", false);
var sheet = doc.WorkbookPart.WorksheetParts.First().Worksheet;
var rows = sheet.Descendants<Row>().ToList();
// LINQ 查询、汇总、生成报告...
```

```csharp
// PowerPoint 示例：批量替换文本
var doc = PresentationDocument.Open(@"slides.pptx", true);
var slides = doc.PresentationPart.SlideParts;
foreach (var slide in slides) {
    foreach (var text in slide.Slide.Descendants<Drawing.Text>()) {
        text.Text = text.Text.Replace("旧标题", "新标题");
    }
}
doc.PresentationPart.Presentation.Save();
```

### 预导入命名空间

脚本中自动可用，无需 using：

```csharp
// 基础
using System;
using System.Linq;
using System.Text.RegularExpressions;
using System.Collections.Generic;
// OpenXML 通用
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
// Word
using DocumentFormat.OpenXml.Wordprocessing;
// Excel
using DocumentFormat.OpenXml.Spreadsheet;
// PowerPoint
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
```

### 发布方式

| 方式 | 命令 | 产物大小 | 运行依赖 |
|------|------|---------|---------|
| self-contained | `dotnet publish -r win-x64 --self-contained` | ~80 MB | 无 |
| framework-dependent | `dotnet publish` | ~15 MB | .NET Runtime (~80 MB) |

## 安全性

- **无沙箱**：Roslyn eval 等价于执行任意 C# 代码，.NET 没有进程内沙箱机制（反射可绕过任何 AST 级别的限制）
- **安全靠操作规范**：只跑自己写的脚本、编辑副本不动原文件、编辑后验证
- **与现状相比无额外风险**：和 `dotnet run Program.cs` 完全等价

## 为什么不用裸 XML 替代 OpenXML SDK

讨论过直接用 Python + lxml 编辑 XML（绕开 SDK），结论：简单操作可以，但以下场景裸 XML 风险高：

| 操作 | 裸 XML 风险 | 原因 |
|------|------------|------|
| 调格式、批量替换 | 低 | 改属性值而已 |
| 设置/应用样式 | 中 | 要同步改 `styles.xml`，继承链搞错会全文走样 |
| **多级列表** | **高** | `numbering.xml` 里 AbstractNum/NumId/Level 绑定关系极复杂，曾导致章标题损坏、计数器跨章累加 |
| **交叉引用/域代码** | **高** | `fldChar` + `instrText` + `separate` + `fldChar end` 顺序和嵌套要求严格，少一个元素 Word 就不识别 |

SDK 的类型系统能保证生成的结构合法，裸 XML 没有这层保障。

## 为什么不做命令式工具（officecli 路线）

讨论过另一种方案：预定义命令，参数化调用（如 `docx-tool replace --paraId xxx --old "a" --new "b"`）。
放弃原因：本质上是重写 officecli，每遇到新需求就要加命令，不灵活。eval 方案一劳永逸。

## 更深层的设计理念：给 agent VM + API 文档，而非封装命令

officecli 和大量 MCP 工具的思路是"给 agent 封装好的命令"，但这个方向从根本上就是错的：

```
officecli 路线:    人类预设命令 → agent 选命令 → 执行
VM + API 路线:     agent 读 API 文档 → 自己写代码 → VM 执行
```

**封装是给人用的，agent 不需要。** 人类需要 CLI 封装是因为人记不住 API、写不了即时代码。但 agent 最擅长的就是"读文档、写代码"，最不擅长的反而是"猜一个 CLI 工具的选择器语法到底怎么写"。

officecli 的问题正是如此 — 设计者要预判所有需求，漏了就没法用（我们碰到的 `w:hint` 属性不暴露、嵌套段落搜索遗漏，都是预判不到的场景）。给 agent 一个 eval 环境 + OpenXML SDK 的类型信息，它能自己组合出任意操作，比任何预设命令集都灵活。

这也是 MCP 生态的一个普遍问题：大量工具在做"给 agent 封装好的高级命令"，但 agent 真正需要的是**原语 + 文档**，不是封装。office-eval 的定位正是如此 — 不封装任何操作，只提供执行能力和类型信息，让 agent 自己组合。

## 本质

相当于把 Roslyn 编译器打包进自己的 exe，做了一个专用的 C# VM。和 `dotnet run` 的区别只是编译器从哪来——SDK 里的独立工具 vs 打包进 exe 的库。运行时都是同一个 CLR。

```
传统:   源码 → SDK(Roslyn) → IL → Runtime(JIT) → 机器码
本工具: 源码 → exe(内嵌Roslyn) → IL → Runtime(JIT) → 机器码
```

## 适用场景

- **Word**：多级列表、域代码、交叉引用、样式继承、批量格式修改 — python-docx 搞不定的全能做
- **Excel**：公式、条件格式、数据透视表、图表 — openpyxl 丢格式的场景
- **PowerPoint**：母版、动画、SmartArt — python-pptx 覆盖不到的结构
- 在没装 .NET SDK 的机器上运行 Office 编辑脚本（self-contained 发布）
- 替代 officecli 和各种封装工具（给 agent VM + API 文档，而非预设命令）
- 替代 python-docx / openpyxl / python-pptx（同一个工具覆盖全家桶，且不丢格式）

## 现状评估

当前 SDK 已安装且工作流通畅（Python 查询 + C# `dotnet run` 编辑），此工具暂无迫切需求。记录于此，后续讨论是否值得实现。

## 待讨论

- [ ] 是否需要预定义辅助函数（如按 paraId 查找段落、批量替换等）
- [ ] 是否支持管道输入（`echo "code" | office-eval`）
- [ ] 是否需要 REPL 交互模式
- [ ] 是否也内嵌 lxml 等价的 XML 查询能力（用于替代 Python 查询步骤）
- [ ] self-contained 发布后实际体积测试（Roslyn + OpenXml 打包可能超 80 MB）
