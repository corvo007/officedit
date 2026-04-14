# 高级操作示例

来自实际论文编辑的复杂操作，附完整可运行 `.csx` 代码。

### 表格边框批量调整

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// office-eval fix-borders.csx -- "thesis.docx"
// 盲审要求：外框 1.5pt，内框 1pt
var src = Args[0];
var dst = Args[0].Replace(".docx", "_edited.docx");
File.Copy(src, dst, true);

var doc = WordprocessingDocument.Open(dst, true);
var body = doc.MainDocumentPart!.Document.Body!;

int count = 0;
foreach (var table in body.Descendants<Table>())
{
    var tblPr = table.GetFirstChild<TableProperties>();
    if (tblPr == null) continue;

    var borders = tblPr.GetFirstChild<TableBorders>();
    if (borders == null) continue;

    // 外框 1.5pt = 12 八分之一磅（Size 单位是 1/8 pt）
    foreach (var border in new[] { borders.TopBorder, borders.BottomBorder })
    {
        if (border != null) border.Size = 12;
    }

    // 内框 1pt = 8
    foreach (var border in new[] { borders.InsideHorizontalBorder, borders.InsideVerticalBorder })
    {
        if (border != null) border.Size = 8;
    }

    count++;
}

doc.MainDocumentPart.Document.Save();
doc.Dispose();
Console.WriteLine($"调整了 {count} 张表格的边框");
```

### 下标/上标格式化（如 Q(15s) 的 "15s"）

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// office-eval subscript.csx -- "thesis.docx"
// 将 Q(15s) 中的 "15s" 设为下标
var src = Args[0];
var dst = Args[0].Replace(".docx", "_edited.docx");
File.Copy(src, dst, true);

var doc = WordprocessingDocument.Open(dst, true);
var body = doc.MainDocumentPart!.Document.Body!;

// 匹配全角和半角括号
var regex = new Regex(@"Q[\(（](\d+\s*(?:s|min))[\)）]");
int count = 0;

foreach (var para in body.Descendants<Paragraph>().ToList())
{
    foreach (var run in para.Descendants<Run>().ToList())
    {
        var te = run.GetFirstChild<Text>();
        if (te == null) continue;

        var match = regex.Match(te.Text);
        if (!match.Success) continue;

        // 获取原 RunProperties
        RunProperties? origRpr = run.RunProperties != null
            ? (RunProperties)run.RunProperties.CloneNode(true) : null;

        string full = te.Text;
        int mStart = match.Index;
        int subStart = match.Groups[1].Index;
        int subEnd = subStart + match.Groups[1].Length;
        int mEnd = match.Index + match.Length;

        // 拆成：前文 + "Q(" + "15s"(下标) + ")" + 后文
        var parts = new List<(string text, bool subscript)>();
        if (mStart > 0)
            parts.Add((full[..mStart], false));
        parts.Add((full[mStart..subStart], false));      // "Q("
        parts.Add((full[subStart..subEnd], true));         // "15s" — 下标
        parts.Add((full[subEnd..mEnd], false));            // ")"
        if (mEnd < full.Length)
            parts.Add((full[mEnd..], false));

        // 替换原 Run 为多个 Run
        var parent = run.Parent!;
        foreach (var (text, isSub) in parts)
        {
            var newRun = new Run();
            var rpr = origRpr != null ? (RunProperties)origRpr.CloneNode(true) : new RunProperties();
            if (isSub)
                rpr.VerticalTextAlignment = new VerticalTextAlignment { Val = VerticalPositionValues.Subscript };
            newRun.AppendChild(rpr);
            newRun.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            parent.InsertBefore(newRun, run);
        }
        run.Remove();
        count++;
    }
}

doc.MainDocumentPart.Document.Save();
doc.Dispose();
Console.WriteLine($"处理了 {count} 处下标");
```

### 跨 Run 文本替换

Word 经常把一段文字拆成多个 Run（如 "表 " 和 "3-13" 分属不同 Run）。需要拼接后匹配、替换、清理残留。

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// office-eval cross-run-replace.csx -- "thesis.docx" "paraId" "表 3-13" "表3-14"
var src = Args[0];
var dst = Args[0].Replace(".docx", "_edited.docx");
File.Copy(src, dst, true);

var doc = WordprocessingDocument.Open(dst, true);
var body = doc.MainDocumentPart!.Document.Body!;

var para = body.Descendants<Paragraph>()
    .FirstOrDefault(p => p.ParagraphId?.Value == Args[1]);

if (para != null)
{
    string oldText = Args[2];
    string newText = Args[3];

    // 拼接所有 Run 的文本，记录每个字符属于哪个 Run
    var runs = para.Descendants<Run>().ToList();
    var fullText = "";
    var charToRun = new List<(int runIdx, Text textNode, int charIdx)>();

    for (int r = 0; r < runs.Count; r++)
    {
        foreach (var te in runs[r].Elements<Text>())
        {
            for (int c = 0; c < te.Text.Length; c++)
            {
                charToRun.Add((r, te, c));
                fullText += te.Text[c];
            }
        }
    }

    int pos = fullText.IndexOf(oldText);
    if (pos >= 0)
    {
        Console.WriteLine($"找到: \"{oldText}\" at position {pos}");

        // 在第一个匹配字符所在的 Text 节点写入新文本
        var first = charToRun[pos];
        var firstText = first.textNode;
        int startInNode = first.charIdx;

        // 收集匹配范围涉及的所有 Text 节点
        var affectedTexts = new HashSet<Text>();
        for (int i = pos; i < pos + oldText.Length; i++)
            affectedTexts.Add(charToRun[i].textNode);

        // 替换第一个节点中的匹配部分
        string before = firstText.Text[..startInNode];
        // 最后一个匹配字符
        var last = charToRun[pos + oldText.Length - 1];
        if (last.textNode == firstText)
        {
            // 匹配在同一个 Text 节点内
            string after = firstText.Text[(last.charIdx + 1)..];
            firstText.Text = before + newText + after;
        }
        else
        {
            // 跨节点：第一个节点保留 before + newText
            firstText.Text = before + newText;
            // 最后一个节点删除匹配部分，保留 after
            string after = last.textNode.Text[(last.charIdx + 1)..];
            last.textNode.Text = after;
            // 中间节点清空
            foreach (var te in affectedTexts)
            {
                if (te != firstText && te != last.textNode)
                    te.Text = "";
            }
        }

        firstText.Space = SpaceProcessingModeValues.Preserve;
        Console.WriteLine($"替换为: \"{newText}\"");
    }
    else
    {
        Console.WriteLine($"未找到: \"{oldText}\"");
    }
}

doc.MainDocumentPart.Document.Save();
doc.Dispose();
```

### ± 符号字体统一（含 Run 拆分）

± (U+00B1) 需要用 Times New Roman 渲染。当 ± 与其他字符混在同一个 Run 里时，需要拆分 Run。

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// office-eval fix-pm.csx -- "thesis.docx"
var src = Args[0];
var dst = Args[0].Replace(".docx", "_edited.docx");
File.Copy(src, dst, true);

var doc = WordprocessingDocument.Open(dst, true);
var body = doc.MainDocumentPart!.Document.Body!;

int pure = 0, split = 0;
foreach (var run in body.Descendants<Run>().ToList())
{
    var te = run.GetFirstChild<Text>();
    if (te == null || !te.Text.Contains("\u00B1")) continue;

    // 确保 RunProperties 存在
    var rpr = run.RunProperties ?? run.PrependChild(new RunProperties());

    if (te.Text == "\u00B1" || te.Text.Trim() == "\u00B1")
    {
        // 纯 ± Run：直接改字体
        var fonts = rpr.GetFirstChild<RunFonts>() ?? rpr.AppendChild(new RunFonts());
        fonts.Ascii = "Times New Roman";
        fonts.HighAnsi = "Times New Roman";
        // 移除 hint 避免 eastAsia 字体覆盖
        var hint = fonts.Hint;
        if (hint != null) fonts.Hint = null;
        pure++;
    }
    else
    {
        // 混合 Run：拆成多个
        var parent = run.Parent!;
        RunProperties origRpr = (RunProperties)rpr.CloneNode(true);
        string text = te.Text;
        int i = 0;
        while (i < text.Length)
        {
            int pmIdx = text.IndexOf('\u00B1', i);
            if (pmIdx < 0)
            {
                // 剩余普通文本
                var tail = new Run();
                tail.AppendChild((RunProperties)origRpr.CloneNode(true));
                tail.AppendChild(new Text(text[i..]) { Space = SpaceProcessingModeValues.Preserve });
                parent.InsertBefore(tail, run);
                break;
            }
            if (pmIdx > i)
            {
                // ± 之前的普通文本
                var before = new Run();
                before.AppendChild((RunProperties)origRpr.CloneNode(true));
                before.AppendChild(new Text(text[i..pmIdx]) { Space = SpaceProcessingModeValues.Preserve });
                parent.InsertBefore(before, run);
            }
            // ± Run（TNR 字体）
            var pmRun = new Run();
            var pmRpr = (RunProperties)origRpr.CloneNode(true);
            var pmFonts = pmRpr.GetFirstChild<RunFonts>() ?? pmRpr.AppendChild(new RunFonts());
            pmFonts.Ascii = "Times New Roman";
            pmFonts.HighAnsi = "Times New Roman";
            if (pmFonts.Hint != null) pmFonts.Hint = null;
            pmRun.AppendChild(pmRpr);
            pmRun.AppendChild(new Text("\u00B1") { Space = SpaceProcessingModeValues.Preserve });
            parent.InsertBefore(pmRun, run);

            i = pmIdx + 1;
        }
        run.Remove();
        split++;
    }
}

doc.MainDocumentPart.Document.Save();
doc.Dispose();
Console.WriteLine($"纯 ± Run: {pure}, 拆分混合 Run: {split}");
```

### 数字与单位间距规范化（SI/GB 3101）

规则：数字+单位加半角空格（75℃→75 °C），±两侧加空格（25±1→25 ± 1），数字+%不加空格。

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// office-eval fix-spacing.csx -- "thesis.docx"
var src = Args[0];
var dst = Args[0].Replace(".docx", "_edited.docx");
File.Copy(src, dst, true);

var doc = WordprocessingDocument.Open(dst, true);
var body = doc.MainDocumentPart!.Document.Body!;

int count = 0;
foreach (var te in body.Descendants<Text>().ToList())
{
    string orig = te.Text;
    string text = orig;

    // 数字 + 单位（非%）加空格：75℃ → 75 °C, 25mg → 25 mg
    text = Regex.Replace(text, @"(\d)(?=\s*(?:°C|℃|mg|mL|μm|rpm|kDa|nm|mm|cm|min|MPa|kN))", "$1 ");
    // 清理多余空格
    text = Regex.Replace(text, @"(\d)\s{2,}(°C|℃|mg|mL|μm|rpm|kDa|nm|mm|cm|min|MPa|kN)", "$1 $2");

    // ± 两侧加空格
    text = Regex.Replace(text, @"(\S)\u00B1(\S)", "$1 \u00B1 $2");

    // 数字+% 不加空格（移除误加的）
    text = Regex.Replace(text, @"(\d)\s+%", "$1%");

    if (text != orig)
    {
        te.Text = text;
        te.Space = SpaceProcessingModeValues.Preserve;
        count++;
    }
}

doc.MainDocumentPart.Document.Save();
doc.Dispose();
Console.WriteLine($"修正了 {count} 处间距");
```

### SEQ 域代码结构（图表自动编号）

Word 的 SEQ 域由 5 个元素严格按顺序组成。手工标注转自动编号时需要构建完整的域结构。

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// office-eval add-seq.csx -- "thesis.docx" "paraId" "Tbl3" "1"
// 在指定段落中，将纯文本"表3-1"替换为带 SEQ 域的自动编号
var src = Args[0];
var dst = Args[0].Replace(".docx", "_edited.docx");
File.Copy(src, dst, true);

var doc = WordprocessingDocument.Open(dst, true);
var body = doc.MainDocumentPart!.Document.Body!;

var para = body.Descendants<Paragraph>()
    .FirstOrDefault(p => p.ParagraphId?.Value == Args[1]);

if (para != null)
{
    string seqName = Args[2];   // 每章独立名称：Tbl1/Tbl2/Fig3 等
    string label = Args[3];     // 显示的缓存值（如 "1"）

    // 获取原 Run 的格式
    var firstRun = para.Descendants<Run>().FirstOrDefault();
    RunProperties? rpr = firstRun?.RunProperties != null
        ? (RunProperties)firstRun.RunProperties.CloneNode(true) : null;

    // 构建 SEQ 域：5 个元素严格按顺序
    // fldChar(begin) → instrText → fldChar(separate) → 缓存值 → fldChar(end)
    var runs = new List<Run>();

    // 1. fldChar begin
    var beginRun = new Run();
    if (rpr != null) beginRun.AppendChild((RunProperties)rpr.CloneNode(true));
    beginRun.AppendChild(new FieldChar { FieldCharType = FieldCharValues.Begin });
    runs.Add(beginRun);

    // 2. instrText
    var instrRun = new Run();
    if (rpr != null) instrRun.AppendChild((RunProperties)rpr.CloneNode(true));
    instrRun.AppendChild(new FieldCode($" SEQ {seqName} \\* ARABIC ") { Space = SpaceProcessingModeValues.Preserve });
    runs.Add(instrRun);

    // 3. fldChar separate
    var sepRun = new Run();
    if (rpr != null) sepRun.AppendChild((RunProperties)rpr.CloneNode(true));
    sepRun.AppendChild(new FieldChar { FieldCharType = FieldCharValues.Separate });
    runs.Add(sepRun);

    // 4. 缓存值（Word 打开后会更新为实际编号）
    var valRun = new Run();
    if (rpr != null) valRun.AppendChild((RunProperties)rpr.CloneNode(true));
    valRun.AppendChild(new Text(label));
    runs.Add(valRun);

    // 5. fldChar end
    var endRun = new Run();
    if (rpr != null) endRun.AppendChild((RunProperties)rpr.CloneNode(true));
    endRun.AppendChild(new FieldChar { FieldCharType = FieldCharValues.End });
    runs.Add(endRun);

    // 找到要替换的位置，插入域代码
    // 这里示例是追加到段落末尾，实际使用时根据需求定位
    foreach (var r in runs)
        para.AppendChild(r);

    Console.WriteLine($"已添加 SEQ {seqName} 域到段落 [{Args[1]}]");
}

doc.MainDocumentPart.Document.Save();
doc.Dispose();
```

### 多级列表编号（章节标题自动编号）

给标题段落绑定多级列表编号。每章使用独立的 AbstractNum 实现从 1 重新编号。

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// office-eval add-numbering.csx -- "thesis.docx"
// 论文样式 ID 映射（从 XML 中读取确认，不要猜）：
//   正文=af5, 章标题=11, 一级节=21, 二级节=31, 三级节=41
var src = Args[0];
var dst = Args[0].Replace(".docx", "_edited.docx");
File.Copy(src, dst, true);

var doc = WordprocessingDocument.Open(dst, true);
var body = doc.MainDocumentPart!.Document.Body!;
var numPart = doc.MainDocumentPart!.NumberingDefinitionsPart
    ?? doc.MainDocumentPart.AddNewPart<NumberingDefinitionsPart>();

numPart.Numbering ??= new Numbering();
var numbering = numPart.Numbering;

// 为每章创建独立的 AbstractNum（避免计数器跨章累加）
// 关键：只定义使用到的 level 数，不补全 9 个
// 关键：不加 LevelRestart/nsid/tmpl/isLgl 等额外属性
int abstractNumId = 1000;  // 起始 ID，避免与现有冲突

// 示例：创建一个 3 级列表（章-节-条）
var absNum = new AbstractNum(
    new Level(
        new StartNumberingValue { Val = 1 },
        new NumberingFormat { Val = NumberFormatValues.Decimal },
        new LevelText { Val = "第%1章" },
        new LevelJustification { Val = LevelJustificationValues.Left }
    ) { LevelIndex = 0 },
    new Level(
        new StartNumberingValue { Val = 1 },
        new NumberingFormat { Val = NumberFormatValues.Decimal },
        new LevelText { Val = "%1.%2" },
        new LevelJustification { Val = LevelJustificationValues.Left }
    ) { LevelIndex = 1 },
    new Level(
        new StartNumberingValue { Val = 1 },
        new NumberingFormat { Val = NumberFormatValues.Decimal },
        new LevelText { Val = "%1.%2.%3" },
        new LevelJustification { Val = LevelJustificationValues.Left }
    ) { LevelIndex = 2 }
) { AbstractNumberId = abstractNumId };

// 插入到 Numbering 的最前面（AbstractNum 必须在 NumberingInstance 之前）
numbering.InsertAt(absNum, 0);

// 创建 NumberingInstance 引用该 AbstractNum
int numId = 2000;
var numInstance = new NumberingInstance(
    new AbstractNumId { Val = abstractNumId }
) { NumberID = numId };
numbering.AppendChild(numInstance);

// 给章标题段落绑定编号
string[] headingStyles = { "11", "21", "31" };  // 章、一级节、二级节
int[] levelMap = { 0, 1, 2 };

for (int i = 0; i < headingStyles.Length; i++)
{
    foreach (var para in body.Descendants<Paragraph>()
        .Where(p => p.ParagraphProperties?.ParagraphStyleId?.Val?.Value == headingStyles[i]))
    {
        var pPr = para.ParagraphProperties ?? para.PrependChild(new ParagraphProperties());

        // 移除已有编号（如果有手打的数字在 Run 里，需要另外处理）
        pPr.NumberingProperties?.Remove();

        // 绑定多级列表
        pPr.NumberingProperties = new NumberingProperties(
            new NumberingLevelReference { Val = levelMap[i] },
            new NumberingId { Val = numId }
        );
    }
}

numbering.Save();
doc.MainDocumentPart.Document.Save();
doc.Dispose();
Console.WriteLine("标题编号已绑定");
// 注意：此操作风险较高，务必先在副本上测试，Word 打开后确认编号正确
```
