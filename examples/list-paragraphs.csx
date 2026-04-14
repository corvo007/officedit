using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// 读取前 10 个段落的 paraId 和文本
var doc = WordprocessingDocument.Open(Args[0], false);
var body = doc.MainDocumentPart!.Document.Body!;

int n = 0;
foreach (var para in body.Descendants<Paragraph>())
{
    var id = para.ParagraphId?.Value ?? "(none)";
    var text = string.Concat(para.Descendants<Text>().Select(t => t.Text));
    if (string.IsNullOrWhiteSpace(text)) continue;

    Console.WriteLine($"[{id}] {(text.Length > 80 ? text[..80] + "..." : text)}");
    if (++n >= 10) break;
}

doc.Dispose();
