using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

class FixCelsius
{
    const string SingleChar = "\u2103"; // ℃
    const string SIForm = "\u00B0C";    // °C

    static void Main()
    {
        var dst = @"D:\Download\毕业论文_edited.docx";
        Console.WriteLine($"Editing: {dst}");

        using var doc = WordprocessingDocument.Open(dst, true);
        var body = doc.MainDocumentPart!.Document.Body!;

        int count = 0;

        foreach (var run in body.Descendants<Run>().ToList())
        {
            foreach (var te in run.Elements<Text>().ToList())
            {
                if (te.Text.Contains(SingleChar))
                {
                    int n = te.Text.Split(SingleChar).Length - 1;
                    te.Text = te.Text.Replace(SingleChar, SIForm);
                    te.Space = SpaceProcessingModeValues.Preserve;
                    count += n;
                }
            }
        }

        doc.MainDocumentPart.Document.Save();
        Console.WriteLine($"℃ → °C: {count} 处");
        Console.WriteLine("Saved.");
    }
}
