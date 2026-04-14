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

## Install (for humans)

Download `office-eval.exe` from [Releases](https://github.com/corvo007/officedit/releases) (~43 MB, Windows x64).

Or build from source:

```bash
cd src/office-eval
dotnet publish -r win-x64 --self-contained \
  -p:PublishSingleFile=true \
  -p:IncludeAllContentForSelfExtract=true \
  -p:EnableCompressionInSingleFile=true \
  -o bin/publish
```

## Install (for AI agents)

Give your agent these instructions to set up office-eval:

### 1. Clone the repo (includes skill docs and API reference)

```bash
git clone --recursive https://github.com/corvo007/officedit.git
```

`--recursive` pulls the OpenXML SDK docs submodule.

### 2. Get the binary

**Option A — Download pre-built** (no .NET SDK needed):

```bash
# Detect platform and download latest release
cd officedit
OS=$(uname -s | tr '[:upper:]' '[:lower:]')
ARCH=$(uname -m)
case "$ARCH" in aarch64|arm64) ARCH="arm64" ;; x86_64|amd64) ARCH="x64" ;; esac
case "$OS" in darwin) RID="mac-$ARCH" ;; linux) RID="linux-$ARCH" ;; mingw*|msys*|cygwin*) RID="win-$ARCH" ;; esac

gh release download v0.1.0 -p "office-eval-${RID}*" -D .
chmod +x office-eval-* 2>/dev/null
```

**Option B — Build from source** (requires .NET 9 SDK):

```bash
cd officedit/src/office-eval
dotnet publish -r win-x64 --self-contained \
  -p:PublishSingleFile=true \
  -p:IncludeAllContentForSelfExtract=true \
  -p:EnableCompressionInSingleFile=true \
  -o ../../bin
```

### 3. Read the skill document

The agent should read `docs/skill.md` first — it contains:
- **3 inviolable rules** (edit copies only, locate by paraId, verify after edit)
- **Script templates** for Word / Excel / PPT
- **Navigation table** pointing to examples, advanced operations, and pitfalls

```
docs/
├── skill.md                ← Start here (~150 lines)
├── examples-basic.md       ← 14 basic operation examples
├── examples-advanced.md    ← 7 advanced examples (borders, SEQ fields, numbering...)
├── pitfalls.md             ← Checklist + known pitfalls + compile errors
├── open-xml-docs/          ← Microsoft's official how-to guides
└── api-doc/                ← Full API reference (grep for class/property names)
```

### 4. Verify installation

```bash
office-eval -e "Console.WriteLine(\"office-eval is ready\");"
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
├── src/office-eval/              # Source code (Program.cs + .csproj, ~130 lines)
├── examples/                     # Example .csx scripts
├── docs/
│   ├── skill.md                  # Agent entry point (~150 lines): rules, templates, navigation
│   ├── examples-basic.md         # 14 basic operation examples
│   ├── examples-advanced.md      # 7 advanced operation examples (borders, subscript, SEQ fields, numbering)
│   ├── pitfalls.md               # Workflow checklist + pitfall records + common errors
│   ├── open-xml-docs/            # Microsoft's official OpenXML SDK how-to guides (git clone)
│   ├── api-doc/                  # XML API reference extracted from NuGet package
│   └── plans/                    # Design documents
└── README.md
```

## For AI agents

The doc system is layered — agents read only what they need:

| File | Lines | When to read |
|------|-------|-------------|
| [`skill.md`](docs/skill.md) | ~150 | **Always** — rules, templates, key concepts |
| [`examples-basic.md`](docs/examples-basic.md) | ~400 | Basic operations (replace, query, delete, read table/comments) |
| [`examples-advanced.md`](docs/examples-advanced.md) | ~500 | Complex operations (borders, subscript, cross-Run, SEQ fields, numbering) |
| [`pitfalls.md`](docs/pitfalls.md) | ~170 | Before editing — checklist, known pitfalls, compile errors |
| `open-xml-docs/` | — | On demand — Microsoft's how-to guides for operations not in examples |
| `api-doc/` | — | On demand — grep for class/property names |

## License

MIT
