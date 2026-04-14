# 基础操作示例

所有示例可直接作为 `.csx` 脚本运行：`office-eval script.csx -- args...`

## Word：安全编辑模式（先复制再编辑）

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

## Word：查询文档结构（列出段落 paraId + 样式 + 文本）

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

## Word：批量文本替换（逐 Run 精确替换）

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

## Word：通过 paraId 定位段落并替换内容

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

## Word：替换整段文本但保留格式

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

## Word：修改 Run 格式（字体、加粗、字号）

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

## Word：删除段落

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

## Word：读取表格数据

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

## Word：读取批注

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

## Word：批量移除高亮标记

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

## Word：图表引用批量校正

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

## Excel：读取单元格值

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

## PowerPoint：提取所有幻灯片文本

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

## PowerPoint：批量替换文本

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
