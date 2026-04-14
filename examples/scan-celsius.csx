using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// 统计 ℃ 出现次数（只读，不修改）
var doc = WordprocessingDocument.Open(Args[0], false);
var body = doc.MainDocumentPart!.Document.Body!;

var singleChar = "\u2103"; // ℃
int count = 0;

foreach (var run in body.Descendants<Run>().ToList())
{
    foreach (var te in run.Elements<Text>().ToList())
    {
        if (te.Text.Contains(singleChar))
        {
            int n = te.Text.Split(singleChar).Length - 1;
            count += n;
            Console.WriteLine($"  Found: ...{te.Text[..Math.Min(50, te.Text.Length)]}...");
        }
    }
}

Console.WriteLine($"\nTotal ℃ occurrences: {count}");
doc.Dispose();
