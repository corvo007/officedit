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

## 铁律：三条不可违反的操作纪律

### 1. 必须编辑副本，永远不动原文件

```csharp
// 正确 ✓
File.Copy(src, dst, overwrite: true);
var doc = WordprocessingDocument.Open(dst, true);

// 错误 ✗ — 直接编辑原文件，一旦脚本有 bug 就无法恢复
var doc = WordprocessingDocument.Open(src, true);
```

原文件是唯一的真相来源。脚本可能有 bug、可能误改、可能中途崩溃。编辑副本意味着随时可以从原文件重来。

**Agent 不得以任何方式修改、覆盖、移动或删除原文件。** 编辑产物始终是副本（如 `xxx_edited.docx`）。是否用副本替换原文件，完全由用户自行决定。Agent 不得代为执行，也不得建议执行破坏性的文件操作。

### 2. 必须按 paraId 精确定位，禁止全局子串匹配替换

```csharp
// 正确 ✓ — 用 paraId 精确定位到目标段落，只在该段落内替换
var para = body.Descendants<Paragraph>()
    .FirstOrDefault(p => p.ParagraphId?.Value == "0778B88E");
if (para != null)
{
    foreach (var te in para.Descendants<Text>())
        if (te.Text.Contains("旧文本")) te.Text = te.Text.Replace("旧文本", "新文本");
}

// 错误 ✗ — 遍历全文档做子串替换，必然误改
foreach (var run in body.Descendants<Run>())  // 遍历全部 Run！
{
    var t = run.GetFirstChild<Text>();
    if (t != null && t.Text.Contains("67%"))
        t.Text = t.Text.Replace("67%", "72%");  // "0.067%" 会被误改为 "0.072%"！
}
```

> **真实事故**：曾用 `"67%"→"72%"` 全局替换，把实验数据中的 `"0.067%"` 误改成 `"0.072%"`，险些引入数据造假。百分比、小数等数字的子串匹配尤其危险。

字符串匹配定位会导致：
- **数据污染** — 数字子串被误匹配（如上述事故）
- **误改其他位置** — 相同文本在脚注、页眉、批注中重复出现
- **不可复现** — 文本被前一次修改改变后，下次运行匹配到别处

paraId 是 Word 为每个段落分配的唯一标识符，不会重复，不受内容变化影响。

**标准工作流：**
1. 先用查询脚本列出目标区域的段落及其 paraId
2. 确认 paraId 后，编辑脚本中只在该段落内做替换
3. 替换前打印完整 run 文本，确认不会误匹配
4. 涉及百分比/数字的替换，必须用更长的上下文字符串匹配

### 3. 编辑完必须读取验证，确认修改结果

```csharp
// 编辑脚本的末尾，或单独写一个验证脚本：
var doc = WordprocessingDocument.Open(dst, false);  // 只读打开
var body = doc.MainDocumentPart!.Document.Body!;

var para = body.Descendants<Paragraph>()
    .FirstOrDefault(p => p.ParagraphId?.Value == "0778B88E");

if (para != null)
{
    string text = string.Concat(para.Descendants<Text>().Select(t => t.Text));
    Console.WriteLine($"验证 [0778B88E]: {text}");
}
```

不要假设脚本执行成功就意味着修改正确。必须验证：
- 目标段落的文本是否变成了预期内容
- 格式是否保留（没有丢失粗体/字体/字号）
- 没有产生副作用（周围段落未被意外修改）

**完整工作流：查询 → 编辑副本 → 验证 → 交给用户。** Agent 的工作到验证为止，绝不触碰原文件。由用户自行决定是否用副本覆盖原文件。

---

## 什么时候不用

- 简单的文本提取 → 用 `markitdown` 或 `python-docx` 就够了
- 创建全新的简单文档 → python-docx 能胜任
- 不涉及 Office 文档的任务

## CLI 用法

```bash
# 文件执行
office-eval script.csx -- file.docx arg2 arg3

# 内联执行
office-eval -e "Console.WriteLine(Args[0]);" -- file.docx
```

脚本内通过 `Args`（`IList<string>`）访问 `--` 后的参数。

## 脚本模板

自动可用的只有基础命名空间（`System`、`System.IO`、`System.Linq`、`System.Collections.Generic`、`System.Text.RegularExpressions`）。

OpenXML 命名空间**不自动导入**——脚本开头必须明确声明用什么。这样做是为了消除类名冲突（`Paragraph`、`Text`、`Run` 在 Word/Excel/PPT 三个命名空间中都有同名类），也让脚本的依赖一目了然。

程序集已全部预加载，写 `using` 即可，不需要 `#r`。

### Word 脚本模板

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

var doc = WordprocessingDocument.Open(Args[0], true);
var body = doc.MainDocumentPart!.Document.Body!;

// ... 操作 ...

doc.MainDocumentPart.Document.Save();
doc.Dispose();
```

### Excel 脚本模板

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

var doc = SpreadsheetDocument.Open(Args[0], false);
var wbPart = doc.WorkbookPart!;

// ... 操作 ...

doc.Dispose();
```

### PPT 脚本模板

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;  // 别名避免冲突

var doc = PresentationDocument.Open(Args[0], false);
var presPart = doc.PresentationPart!;

// ... 操作 ...

doc.Dispose();
```

## 扩展引用

- `#r "path/to/local.dll"` — 引用本地 DLL（Roslyn 原生支持）
- `using` — 可使用 .NET Runtime 标准库中的任何命名空间
- 不支持运行时 NuGet 包引用

---

## 常见示例

### Word：安全编辑模式（先复制再编辑）

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// office-eval edit.csx -- "源文件.docx" "输出文件.docx"
var src = Args[0];
var dst = Args[1];
File.Copy(src, dst, overwrite: true);

var doc = WordprocessingDocument.Open(dst, true);
var body = doc.MainDocumentPart!.Document.Body!;

// ... 编辑操作 ...

doc.MainDocumentPart.Document.Save();
doc.Dispose();
Console.WriteLine($"已保存: {dst}");
```

### Word：批量文本替换（逐 Run 精确替换）

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// office-eval replace.csx -- "thesis.docx"
var doc = WordprocessingDocument.Open(Args[0], true);
var body = doc.MainDocumentPart!.Document.Body!;

int count = 0;
foreach (var run in body.Descendants<Run>().ToList())
{
    foreach (var te in run.Elements<Text>().ToList())
    {
        if (te.Text.Contains("旧文本"))
        {
            te.Text = te.Text.Replace("旧文本", "新文本");
            te.Space = SpaceProcessingModeValues.Preserve;
            count++;
        }
    }
}

doc.MainDocumentPart.Document.Save();
doc.Dispose();
Console.WriteLine($"替换了 {count} 处");
```

### Word：通过 paraId 定位段落并替换内容

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// office-eval fix-para.csx -- "thesis.docx"
var doc = WordprocessingDocument.Open(Args[0], true);
var body = doc.MainDocumentPart!.Document.Body!;

// 建 paraId 查找表
var paraMap = new Dictionary<string, Paragraph>();
foreach (var p in body.Descendants<Paragraph>())
{
    var id = p.ParagraphId?.Value;
    if (id != null) paraMap[id] = p;
}

// 按 paraId 替换段落中的部分文本
if (paraMap.TryGetValue("0778B88E", out var para))
{
    foreach (var te in para.Descendants<Text>())
    {
        if (te.Text.Contains("旧表述"))
        {
            te.Text = te.Text.Replace("旧表述", "新表述");
            Console.WriteLine("已替换");
        }
    }
}

doc.MainDocumentPart.Document.Save();
doc.Dispose();
```

### Word：替换整段文本但保留格式

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// office-eval rewrite-para.csx -- "thesis.docx" "paraId" "新文本内容"
var doc = WordprocessingDocument.Open(Args[0], true);
var body = doc.MainDocumentPart!.Document.Body!;

var para = body.Descendants<Paragraph>()
    .FirstOrDefault(p => p.ParagraphId?.Value == Args[1]);

if (para != null)
{
    // 保存第一个 Run 的格式
    var firstRun = para.Descendants<Run>().FirstOrDefault();
    RunProperties? rpr = null;
    if (firstRun?.RunProperties != null)
        rpr = (RunProperties)firstRun.RunProperties.CloneNode(true);

    // 删除所有 Run（保留 ParagraphProperties）
    para.ChildElements
        .Where(e => !(e is ParagraphProperties))
        .ToList()
        .ForEach(e => e.Remove());

    // 用保存的格式创建新 Run
    var newRun = new Run();
    if (rpr != null) newRun.AppendChild(rpr);
    newRun.AppendChild(new Text(Args[2]) { Space = SpaceProcessingModeValues.Preserve });
    para.AppendChild(newRun);

    Console.WriteLine("段落已替换");
}

doc.MainDocumentPart.Document.Save();
doc.Dispose();
```

### Word：修改 Run 格式（字体、加粗、字号）

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// office-eval set-font.csx -- "doc.docx"
var doc = WordprocessingDocument.Open(Args[0], true);
var body = doc.MainDocumentPart!.Document.Body!;

// 给第一个 Run 设置 Arial 字体 + 加粗
var run = body.Descendants<Run>().First();
var rPr = new RunProperties(
    new RunFonts() { Ascii = "Arial", HighAnsi = "Arial" },
    new Bold(),
    new FontSize() { Val = "24" }  // 单位是半磅，24 = 12pt
);
run.PrependChild(rPr);

doc.MainDocumentPart.Document.Save();
doc.Dispose();
```

### Word：删除段落

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// office-eval delete-para.csx -- "thesis.docx" "paraId"
var doc = WordprocessingDocument.Open(Args[0], true);
var body = doc.MainDocumentPart!.Document.Body!;

var para = body.Descendants<Paragraph>()
    .FirstOrDefault(p => p.ParagraphId?.Value == Args[1]);

if (para != null)
{
    string text = string.Concat(para.Descendants<Text>().Select(t => t.Text));
    Console.WriteLine($"删除: {(text.Length > 60 ? text[..60] + "..." : text)}");
    para.Remove();
}

doc.MainDocumentPart.Document.Save();
doc.Dispose();
```

### Excel：读取单元格值

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

// office-eval read-cell.csx -- "data.xlsx" "Sheet1" "B2"
var doc = SpreadsheetDocument.Open(Args[0], false);
var wbPart = doc.WorkbookPart!;

// 按名称找 Sheet
var sheet = wbPart.Workbook.Descendants<Sheet>()
    .First(s => s.Name == Args[1]);
var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id!);

// 按地址找 Cell
var cell = wsPart.Worksheet!.Descendants<Cell>()
    .FirstOrDefault(c => c.CellReference == Args[2]);

if (cell != null)
{
    string value = cell.InnerText;

    // 处理共享字符串（Excel 的字符串存储机制）
    if (cell.DataType?.Value == CellValues.SharedString)
    {
        var sst = wbPart.GetPartsOfType<SharedStringTablePart>().First();
        value = sst.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
    }

    Console.WriteLine($"{Args[2]} = {value}");
}

doc.Dispose();
```

### PowerPoint：提取所有幻灯片文本

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

// office-eval ppt-text.csx -- "slides.pptx"
var doc = PresentationDocument.Open(Args[0], false);
var presPart = doc.PresentationPart!;
var slideIds = presPart.Presentation.SlideIdList!.ChildElements;

for (int i = 0; i < slideIds.Count; i++)
{
    var slideId = (SlideId)slideIds[i];
    var slidePart = (SlidePart)presPart.GetPartById(slideId.RelationshipId!);

    Console.WriteLine($"--- Slide {i + 1} ---");
    foreach (var para in slidePart.Slide.Descendants<Drawing.Paragraph>())
    {
        string text = string.Concat(
            para.Descendants<Drawing.Text>().Select(t => t.Text));
        if (!string.IsNullOrWhiteSpace(text))
            Console.WriteLine(text);
    }
}

doc.Dispose();
```

### PowerPoint：批量替换文本

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;

// office-eval ppt-replace.csx -- "slides.pptx" "旧标题" "新标题"
var doc = PresentationDocument.Open(Args[0], true);
int count = 0;

foreach (var slidePart in doc.PresentationPart!.SlideParts)
{
    foreach (var text in slidePart.Slide.Descendants<Drawing.Text>())
    {
        if (text.Text.Contains(Args[1]))
        {
            text.Text = text.Text.Replace(Args[1], Args[2]);
            count++;
        }
    }
}

doc.PresentationPart.Presentation.Save();
doc.Dispose();
Console.WriteLine($"替换了 {count} 处");
```

### Word：查询文档结构（列出段落 paraId + 样式 + 文本）

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// office-eval list-paras.csx -- "thesis.docx"
// 编辑前的第一步：先查再改
var doc = WordprocessingDocument.Open(Args[0], false);
var body = doc.MainDocumentPart!.Document.Body!;

foreach (var para in body.Descendants<Paragraph>())
{
    var id = para.ParagraphId?.Value ?? "(none)";
    var style = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "Normal";
    var text = string.Concat(para.Descendants<Text>().Select(t => t.Text));
    if (string.IsNullOrWhiteSpace(text)) continue;

    Console.WriteLine($"[{id}] ({style}) {(text.Length > 70 ? text[..70] + "..." : text)}");
}

doc.Dispose();
```

### Word：读取表格数据

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// office-eval read-table.csx -- "thesis.docx"
var doc = WordprocessingDocument.Open(Args[0], false);
var body = doc.MainDocumentPart!.Document.Body!;

int tableIndex = 0;
foreach (var table in body.Descendants<Table>())
{
    Console.WriteLine($"\n=== Table {++tableIndex} ===");
    foreach (var row in table.Descendants<TableRow>())
    {
        var cells = row.Descendants<TableCell>()
            .Select(c => string.Concat(c.Descendants<Text>().Select(t => t.Text)).Trim());
        Console.WriteLine(string.Join(" | ", cells));
    }
}

doc.Dispose();
```

### Word：读取批注

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// office-eval read-comments.csx -- "thesis.docx"
var doc = WordprocessingDocument.Open(Args[0], false);
var commentsPart = doc.MainDocumentPart!.WordprocessingCommentsPart;

if (commentsPart != null)
{
    foreach (var comment in commentsPart.Comments.Elements<Comment>())
    {
        var id = comment.Id?.Value;
        var author = comment.Author?.Value;
        var text = string.Concat(comment.Descendants<Text>().Select(t => t.Text));
        Console.WriteLine($"[Comment {id}] ({author}): {text}");
    }
}

doc.Dispose();
```

### Word：批量移除高亮标记

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// office-eval remove-highlight.csx -- "thesis.docx"
var src = Args[0];
var dst = Args[0].Replace(".docx", "_edited.docx");
File.Copy(src, dst, true);

var doc = WordprocessingDocument.Open(dst, true);
var body = doc.MainDocumentPart!.Document.Body!;

int count = 0;
foreach (var run in body.Descendants<Run>())
{
    var highlight = run.RunProperties?.Highlight;
    if (highlight != null)
    {
        highlight.Remove();
        count++;
    }
}

doc.MainDocumentPart.Document.Save();
doc.Dispose();
Console.WriteLine($"移除 {count} 处高亮");
```

### Word：图表引用批量校正

正文中引用的图/表编号与实际标注不匹配时的处理流程：
1. 先扫描所有标注，建立正确编号映射
2. 再逐段修正正文引用（用 paraId 精确定位）

```csharp
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// office-eval fix-refs.csx -- "thesis.docx"
// 第一步：扫描所有表/图标注，建立映射
var doc = WordprocessingDocument.Open(Args[0], false);
var body = doc.MainDocumentPart!.Document.Body!;

var regex = new Regex(@"(表|图)\s*(\d+-\d+)");
foreach (var para in body.Descendants<Paragraph>())
{
    var text = string.Concat(para.Descendants<Text>().Select(t => t.Text));
    var match = regex.Match(text);
    if (match.Success)
    {
        var id = para.ParagraphId?.Value ?? "?";
        Console.WriteLine($"[{id}] {match.Value}: {(text.Length > 60 ? text[..60] + "..." : text)}");
    }
}

doc.Dispose();
// 第二步：确认映射后，写编辑脚本按 paraId 逐个修正
```

---

## 工作流检查清单

每次编辑文档前后对照检查：

- [ ] 编辑输出到副本（如 `xxx_edited.docx`），不动原文件
- [ ] 涉及域代码的脚本先在测试文件上验证
- [ ] 替换前打印完整 run 文本确认上下文
- [ ] 编辑后用只读模式打开副本验证修改生效
- [ ] 检查是否有误替换（特别是数字/百分比）
- [ ] 跨 Run 替换后检查相邻 Run 有无残留文本
- [ ] 全角/半角字符都要匹配
- [ ] 确保 Word 已关闭，避免 IOException
- [ ] 含空格的 Text 元素设置了 `SpaceProcessingModeValues.Preserve`
- [ ] Agent 不触碰原文件，交给用户决定是否覆盖

---

## 常见编译错误

| 错误 | 原因 | 解决 |
|------|------|------|
| `CS1002: 应输入 ;` | `using var doc = ...` 在脚本模式下被解析为命名空间导入 | 改用 `var doc = ...; doc.Dispose();` |
| `CS0104: 不明确的引用` | `Paragraph`/`Text` 等类名在多个命名空间中存在 | 只导入需要的命名空间，或用别名 `using Drawing = ...` |
| `CS0117: 不包含定义` | OpenXML SDK 的 API 和直觉不一样 | 查 `docs/api-doc/` 确认 API 存在 |
| `CS8602: 空引用警告` | OpenXML 属性可能为 null | 用 `!` 断言：`doc.MainDocumentPart!.Document.Body!` |

---

## 已验证可靠的操作（来自实际生产使用）

以下操作已在真实论文编辑中反复验证，可放心使用：

- 批量文本替换（℃→°C、术语统一、占位符替换）
- 表格边框粗细调整（54 张表格，中间线/外框分别设置）
- 标题加粗/去粗（86 个标题，操作 StyleRunProperties）
- 下标/上标格式化（63 处 Q(15s) 等，`vertAlign = "subscript"`）
- ± 符号字体统一（240 处，含 8 个需拆分 Run 的混合内容）
- 数字与单位间距规范化（250 处，SI/GB 3101 标准）
- 段落对齐方式修改（`JustificationValues.Both`）
- 移除高亮标记（103 处，删除 `RunProperties.Highlight`）
- 节距/页边距调整（SectionProperties 操作）
- 关键词字体修正（`rFonts` 的 ascii/hAnsi/eastAsia 分别设置）
- `SpaceProcessingModeValues.Preserve` 保留空格

---

## 文档层级结构速查

### Word (.docx)
```
WordprocessingDocument
└── MainDocumentPart
    └── Document
        └── Body
            ├── Paragraph (p)
            │   ├── ParagraphProperties (pPr) — 段落格式
            │   └── Run (r)
            │       ├── RunProperties (rPr) — 字符格式（字体/加粗/字号/颜色）
            │       └── Text (t) — 文本内容
            ├── Table
            │   └── TableRow → TableCell → Paragraph → ...
            └── SectionProperties — 页面设置
```

### Excel (.xlsx)
```
SpreadsheetDocument
└── WorkbookPart
    ├── Workbook → Sheets → Sheet（名称 + 关系 ID）
    └── WorksheetPart
        └── Worksheet → SheetData → Row → Cell
            Cell.DataType: Number | SharedString | Boolean | ...
```

### PowerPoint (.pptx)
```
PresentationDocument
└── PresentationPart
    ├── Presentation → SlideIdList → SlideId（关系 ID）
    └── SlidePart
        └── Slide → CommonSlideData → ShapeTree → Shape
            └── TextBody → Drawing.Paragraph → Drawing.Run → Drawing.Text
```

---

## 进阶参考

### 需要更复杂的操作？

查阅 `docs/open-xml-docs/docs/` 目录下的 How-to 指南：

- **Word**: `docs/word/` — 样式、表格、页眉页脚、图片插入、修订接受
- **Excel**: `docs/spreadsheet/` — 创建工作表、插入图表、合并单元格、大文件 SAX 读取
- **PPT**: `docs/presentation/` — 插入幻灯片、添加动画/音视频、母版操作

### 需要确认函数签名、属性名、类层级？

查阅 `docs/api-doc/` 目录下的 XML 文档文件（从 NuGet 包提取的 IntelliSense 数据）：

- `DocumentFormat.OpenXml.xml` — 所有 OpenXML 类型的完整 API 参考
- 按类名搜索即可找到所有公开属性和方法

### 关键概念

- **Run** 是最小的格式单元。同一个 Run 内的文本共享相同格式（字体、粗体、颜色等）
- **ParagraphId** 是每个段落的唯一标识符，可用于精确定位
- **SharedString** 是 Excel 的字符串去重机制。读单元格时要检查 `DataType` 是否为 `SharedString`，是则去 SharedStringTable 查真实文本
- **Drawing 命名空间** 在 PPT 中用于文本和图形。和 Word 的 `Text`/`Run` 同名但不同类，需要用别名区分
- **FontSize 单位是半磅**（half-point）。12pt = `Val = "24"`

---

## 踩坑记录（来自真实生产事故）

以下是实际使用中踩过的坑，编写脚本时必须注意。

### 多级列表（numbering.xml）

- **不要往章标题段落里插入任何额外元素**（隐藏域、重置 run 等）。曾因此导致章标题被 Word 合并/损坏，整个文档结构崩溃
- **每章用独立 AbstractNum**，比 `LevelOverride + StartOverrideNumberingValue` 更可靠。后者与 `LevelRestart` 交互复杂，容易计数器跨章累加
- **Level 定义**：用 `new Level(子元素...)` 构造器传入，只定义实际使用的 level 数（不需要补全 9 个），不加 `LevelRestart`、`nsid`、`tmpl`、`isLgl` 等额外属性。加了反而 Word 不识别
- `numbering.xml` 里 AbstractNum/NumId/Level 的绑定关系极复杂，裸 XML 操作风险最高的区域

### 域代码（fldChar + instrText）

- `fldChar(begin)` + `instrText` + `fldChar(separate)` + 缓存值 + `fldChar(end)` 顺序和嵌套要求严格，少一个元素 Word 就不识别
- **SEQ 域跨章重置**：不要用 `\s 1`（需要内置 Heading 1 才生效），不要往标题段落插隐藏 `\r 0 \h` 域。安全做法是每章用独立 SEQ 名称（`Tbl1`/`Tbl2`/`Fig3` 等），彻底隔离计数器
- 涉及域代码的脚本，**必须先在副本上测试，确认无误后再应用到正式文件**

### 跨 Run 文本处理

Word 经常把一段文字拆成多个 Run（如 "表 " 和 "3-13" 在不同 Run 里）。直接按单个 Run 的 `Text` 匹配会找不到完整字符串。
- 需要拼接相邻 Run 的文本做匹配
- 替换后必须清理相邻 Run 中的残留文本片段，否则出现重复字符
- 实际案例：修正 "表 3-13" → "表3-14"，涉及跨 Run 替换 + 清理残留

### 关键 XML 属性（只有直接读 XML 才能看到）

- `w:hint="eastAsia"` — 告诉 Word 用 eastAsia 字体渲染该 Run。会导致 ± 等 Latin 字符被宋体渲染而非 TNR
- `w:rFonts` 的 `ascii`/`hAnsi`/`eastAsia`/`cs` 四个属性分别控制不同 Unicode 范围的字体选择
- `xml:space="preserve"` / `SpaceProcessingModeValues.Preserve` — 不设则 Word 吞掉 Text 元素中的空格
- `w:vanish` — 隐藏 Run（域代码中常见），遍历文本时注意跳过
- `w:vertAlign val="subscript"/"superscript"` — 下标/上标格式

### Run 拆分技巧（混合内容处理）

当一个 Run 包含特殊字符和普通字符时，不能直接改整个 Run 的字体（会影响其他字符）。需要：
1. 找到目标字符在 Text 中的位置
2. 把 Run 拆成 3 个：前文本 Run + 目标字符 Run + 后文本 Run
3. 只给目标字符 Run 设置特定字体/格式

### 全角/半角字符

`Q（15s）`（全角括号 U+FF08/FF09）和 `Q(15s)`（半角括号）是不同字符序列。正则匹配时必须同时覆盖两种形式，否则漏改。

### 文件锁定

Word 打开文件时，C# OpenXML SDK 无法访问同一文件（IOException）。必须先关闭 Word 或操作副本。

### Roslyn 脚本模式语法限制

- **`using var` 不能用**。Roslyn 脚本模式把 `using` 开头的语句解析为命名空间导入指令，`using var doc = ...` 会报 `CS1002: 应输入 ;`。必须用 `var doc = ...; doc.Dispose();` 手动管理生命周期
- **命名空间冲突**：`Wordprocessing.Paragraph` / `Drawing.Paragraph` / `Spreadsheet`  等同名类型会导致 `CS0104` 歧义错误。默认只导入 `Wordprocessing`，Excel/PPT 脚本需在开头自行添加 `using`

### 脚本幂等性

- 脚本应设计为可重复运行。已转换的标注再次用正则匹配时，域代码的缓存值会干扰
- 每次运行前应检查是否已经处理过（判断目标状态是否已存在），避免重复操作导致结构损坏
