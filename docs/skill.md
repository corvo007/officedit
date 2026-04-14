# office-eval：OpenXML 脚本运行器

轻量级 CLI，内嵌 Roslyn 编译器 + OpenXML SDK。下载即用，无需安装 .NET SDK。

## 设计理念

**给 agent VM + API 文档，而非封装命令。**

- officecli 等工具的思路是"预设命令 → agent 选命令 → 执行"，但设计者不可能预判所有需求
- office-eval 的思路是"agent 读 API 文档 → 自己写 C# 脚本 → VM 执行"
- agent 最擅长的就是"读文档、写代码"，最不擅长的是"猜 CLI 工具的选择器语法"

## 为什么不用其他工具

| 工具 | 问题 |
|------|------|
| python-docx / openpyxl | OpenXML 规范实现残缺，编辑复杂文档（多级列表、域代码、交叉引用）100% 会坏 |
| officecli 等封装工具 | 设计者要预判所有需求，漏了就没法用（属性不暴露、嵌套搜索遗漏） |
| OpenXML SDK + `dotnet run` | SDK 本身完美，但需要安装 ~500 MB .NET SDK |
| dotnet-script | 需要 .NET SDK 做 NuGet 包恢复 |

office-eval 把 Roslyn 编译器和 OpenXML SDK 打包进单个 exe，消除安装门槛。

## 什么时候用

- 编辑 .docx 文档：批量替换文本、修改格式、调整样式、处理多级列表、域代码
- 编辑 .xlsx 表格：读取/写入单元格、公式、条件格式
- 编辑 .pptx 演示文稿：替换文本、修改形状、批量操作幻灯片
- 需要保证文档结构不被破坏的场景（SDK 的类型系统保证生成的 XML 合法）

## 什么时候不用

- 简单的文本提取 → 用 `markitdown` 或 `python-docx` 就够了
- 创建全新的简单文档 → python-docx 能胜任
- 不涉及 Office 文档的任务

---

## 铁律：三条不可违反的操作纪律

### 1. 必须编辑副本，永远不动原文件

```csharp
File.Copy(src, dst, overwrite: true);
var doc = WordprocessingDocument.Open(dst, true);  // 编辑副本
```

**Agent 不得以任何方式修改、覆盖、移动或删除原文件。** 编辑产物始终是副本（如 `xxx_edited.docx`）。是否用副本替换原文件，完全由用户自行决定。

### 2. 必须按 paraId 精确定位，禁止全局子串匹配替换

```csharp
// 正确 ✓ — 用 paraId 精确定位，只在该段落内替换
var para = body.Descendants<Paragraph>()
    .FirstOrDefault(p => p.ParagraphId?.Value == "0778B88E");

// 错误 ✗ — 遍历全文档子串替换
// 曾把 "0.067%" 误改成 "0.072%"，险些引入数据造假
```

**标准工作流：** 查询脚本列出 paraId → 确认目标 → 编辑脚本只在该段落内替换 → 替换前打印 run 文本确认

### 3. 编辑完必须读取验证，确认修改结果

```csharp
var doc = WordprocessingDocument.Open(dst, false);  // 只读打开验证
```

**完整工作流：查询 → 编辑副本 → 验证 → 交给用户。** Agent 的工作到验证为止。

---

## CLI 用法

```bash
office-eval script.csx -- file.docx arg2 arg3    # 文件执行
office-eval -e "Console.WriteLine(Args[0]);" -- file.docx  # 内联执行
```

脚本内通过 `Args`（`IList<string>`）访问 `--` 后的参数。

## 脚本模板

自动可用：`System`、`System.IO`、`System.Linq`、`System.Collections.Generic`、`System.Text.RegularExpressions`。

OpenXML 命名空间**不自动导入**——脚本开头明确声明用什么，消除类名冲突。程序集已预加载，写 `using` 即可。

```csharp
// Word
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// Excel
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

// PPT
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;  // 别名避免冲突
```

## 文档层级结构速查

### Word (.docx)
```
WordprocessingDocument → MainDocumentPart → Document → Body
  ├── Paragraph (p) → ParagraphProperties (pPr) + Run (r) → RunProperties (rPr) + Text (t)
  ├── Table → TableRow → TableCell → Paragraph → ...
  └── SectionProperties — 页面设置
```

### Excel (.xlsx)
```
SpreadsheetDocument → WorkbookPart → Workbook → Sheets → Sheet
  └── WorksheetPart → Worksheet → SheetData → Row → Cell (DataType: Number|SharedString|Boolean)
```

### PowerPoint (.pptx)
```
PresentationDocument → PresentationPart → Presentation → SlideIdList → SlideId
  └── SlidePart → Slide → ShapeTree → Shape → TextBody → Drawing.Paragraph → Drawing.Text
```

## 关键概念

- **Run** 是最小格式单元。同一 Run 内文本共享格式
- **ParagraphId** 是段落唯一标识符，精确定位用
- **SharedString** 是 Excel 字符串去重机制，读 Cell 时检查 DataType
- **FontSize 单位是半磅**。12pt = `Val = "24"`
- **`using var` 不能用** — Roslyn 脚本模式会解析为命名空间导入，用 `var doc = ...; doc.Dispose();`
- **`SpaceProcessingModeValues.Preserve`** — 含空格的 Text 元素必须设置，否则 Word 吞空格

## 已验证可靠的操作

以下操作已在真实论文编辑中反复验证：批量文本替换、表格边框调整（54 张）、标题加粗/去粗（86 个）、下标格式化（63 处）、± 字体统一（240 处含 Run 拆分）、数字单位间距规范化（250 处）、段落对齐、移除高亮（103 处）、节距/页边距调整、关键词字体修正。

---

## 文档导航

| 需要什么 | 查哪里 |
|---------|--------|
| 基础操作示例（替换、查询、删除、读表格、读批注、Excel、PPT） | [`examples-basic.md`](examples-basic.md) |
| 高级操作示例（表格边框、下标、跨Run替换、±字体、SEQ域、多级列表） | [`examples-advanced.md`](examples-advanced.md) |
| 踩坑记录 + 工作流检查清单 + 常见编译错误 | [`pitfalls.md`](pitfalls.md) |
| OpenXML SDK how-to 指南（样式、图片、页眉、图表等） | `open-xml-docs/docs/` |
| API 签名查表（类名、属性名、方法） | `api-doc/DocumentFormat.OpenXml.xml`（grep 搜索） |
