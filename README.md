# office-eval

Self-contained C# script runner for editing Office documents. Embeds Roslyn compiler + OpenXML SDK in a single executable — no .NET SDK required.

**130 lines of code. 43 MB single file. Full OpenXML SDK access.**

## Why

"AI can't reliably edit Office documents" is a tooling problem, not a capability problem.

| Tool | Problem |
|------|---------|
| python-docx / openpyxl | Incomplete OpenXML spec coverage. Complex documents (multi-level lists, field codes, cross-references) always break. |
| officecli / CLI wrappers | Designers must anticipate every need. Miss something? You're stuck. (80K lines of C# to wrap what the SDK already exposes.) |
| OpenXML SDK + `dotnet run` | The SDK works perfectly — but requires a ~500 MB .NET SDK install. |

office-eval solves this by packaging the Roslyn compiler + OpenXML SDK into a single zero-dependency binary. Download, run, done.

## Philosophy

**Give agents a VM + API docs, not wrapped commands.**

CLI wrappers try to predict what users need: `--prop fill=FF0000`, `--type shape`, `--prop bold=true`. But agents don't need human-friendly CLI syntax. Agents read API docs and write code — that's what they're best at.

```
CLI wrapper:  human predicts commands → agent picks commands → execute
office-eval:  agent reads API docs → agent writes code → VM executes
```

One approach scales linearly with the designer's imagination. The other scales with the SDK.

## Install

Download `office-eval.exe` from [Releases](https://github.com/corvo007/officedit/releases) (43 MB, Windows x64).

Or build from source:

```bash
cd src/office-eval
dotnet publish -r win-x64 --self-contained \
  -p:PublishSingleFile=true \
  -p:IncludeAllContentForSelfExtract=true \
  -p:EnableCompressionInSingleFile=true \
  -o bin/publish
```

## Usage

```bash
# Run a script
office-eval script.csx -- document.docx arg2 arg3

# Inline execution
office-eval -e "Console.WriteLine(Args[0]);" -- document.docx
```

Scripts access arguments via `Args` (`IList<string>`), separated by `--`.

## Script example

```csharp
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

// List paragraphs with their IDs
var doc = WordprocessingDocument.Open(Args[0], false);
var body = doc.MainDocumentPart!.Document.Body!;

foreach (var para in body.Descendants<Paragraph>())
{
    var id = para.ParagraphId?.Value ?? "(none)";
    var text = string.Concat(para.Descendants<Text>().Select(t => t.Text));
    if (!string.IsNullOrWhiteSpace(text))
        Console.WriteLine($"[{id}] {text}");
}

doc.Dispose();
```

More examples in [`examples/`](examples/).

## What's included

- **System namespaces** auto-imported: `System`, `System.IO`, `System.Linq`, `System.Collections.Generic`, `System.Text.RegularExpressions`
- **OpenXML SDK** assemblies pre-loaded. Add `using` for the namespace you need:
  - Word: `using DocumentFormat.OpenXml.Wordprocessing;`
  - Excel: `using DocumentFormat.OpenXml.Spreadsheet;`
  - PPT: `using DocumentFormat.OpenXml.Presentation;` + `using Drawing = DocumentFormat.OpenXml.Drawing;`
- `#r "path/to/local.dll"` for additional assemblies
- Full .NET Runtime standard library available

## Scripting notes

- **No `using var`** — Roslyn scripting mode parses `using` at line start as a namespace import directive. Use `var doc = ...; doc.Dispose();` instead.
- **No runtime NuGet** — packages can't be installed at runtime. The SDK and standard library cover all Office editing needs.

## Project structure

```
officedit/
├── src/office-eval/         # Source code (Program.cs + .csproj)
├── examples/                # Example .csx scripts
├── docs/
│   ├── skill.md             # Agent skill document (usage guide + examples + pitfalls)
│   ├── open-xml-docs/       # Microsoft's official OpenXML SDK how-to guides
│   ├── api-doc/             # XML API reference extracted from NuGet package
│   └── plans/               # Design documents
└── README.md
```

## For AI agents

The [`docs/skill.md`](docs/skill.md) file is designed as an agent skill entry point:

1. **Operational rules** — edit copies only, locate by paraId, verify after editing
2. **Script templates** — Word / Excel / PPT boilerplate ready to copy
3. **15 real-world examples** — from actual thesis editing production use
4. **Pitfall records** — lessons from real data corruption incidents
5. **Workflow checklist** — pre/post editing verification steps
6. **Reference pointers** — advanced operations → `open-xml-docs/`, API lookup → `api-doc/`

For advanced operations, agents consult `docs/open-xml-docs/` (Microsoft's how-to guides). For API signatures, grep `docs/api-doc/DocumentFormat.OpenXml.xml`.

## License

MIT
