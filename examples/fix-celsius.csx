using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// 把 ℃ 替换成 °C（编辑副本 + 自动验证）
var src = Args[0];
var dst = Args[0].Replace(".docx", "_test_edited.docx");
File.Copy(src, dst, true);
Console.WriteLine($"副本: {dst}");

var doc = WordprocessingDocument.Open(dst, true);
var body = doc.MainDocumentPart!.Document.Body!;

var singleChar = "\u2103"; // ℃
var siForm = "\u00B0C";    // °C
int count = 0;

foreach (var run in body.Descendants<Run>().ToList())
{
    foreach (var te in run.Elements<Text>().ToList())
    {
        if (te.Text.Contains(singleChar))
        {
            int n = te.Text.Split(singleChar).Length - 1;
            te.Text = te.Text.Replace(singleChar, siForm);
            te.Space = SpaceProcessingModeValues.Preserve;
            count += n;
        }
    }
}

doc.MainDocumentPart!.Document.Save();
doc.Dispose();
Console.WriteLine($"℃ → °C: {count} 处");

// 验证
var verify = WordprocessingDocument.Open(dst, false);
var vBody = verify.MainDocumentPart!.Document.Body!;
int remaining = 0;
foreach (var te in vBody.Descendants<Text>())
{
    if (te.Text.Contains(singleChar)) remaining++;
}
verify.Dispose();
Console.WriteLine($"验证: 剩余 ℃ = {remaining} (应为 0)");
