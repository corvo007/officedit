using System.Reflection;
using System.Text;
using Microsoft.CodeAnalysis.CSharp.Scripting;
using Microsoft.CodeAnalysis.Scripting;

Console.OutputEncoding = Encoding.UTF8;
Console.InputEncoding = new UTF8Encoding(false);

// -- Parse CLI args --
// Modes:
//   office-eval script.csx [-- arg1 arg2 ...]
//   office-eval -e "code"  [-- arg1 arg2 ...]

string? scriptFile = null;
string? inlineCode = null;
var scriptArgs = new List<string>();

int i = 0;
bool afterSeparator = false;

while (i < args.Length)
{
    if (afterSeparator)
    {
        scriptArgs.Add(args[i]);
        i++;
        continue;
    }

    if (args[i] == "--")
    {
        afterSeparator = true;
        i++;
        continue;
    }

    if (args[i] == "-e" && i + 1 < args.Length)
    {
        inlineCode = args[i + 1];
        i += 2;
        continue;
    }

    if (args[i] == "--version")
    {
        Console.WriteLine("office-eval 0.1.0");
        return 0;
    }

    if (args[i] == "--help" || args[i] == "-h")
    {
        PrintUsage();
        return 0;
    }

    // First positional arg is the script file
    if (scriptFile == null && !args[i].StartsWith("-"))
    {
        scriptFile = args[i];
        i++;
        continue;
    }

    Console.Error.WriteLine($"Unknown option: {args[i]}");
    PrintUsage();
    return 1;
}

// Determine code to run
string code;
if (inlineCode != null)
{
    code = inlineCode;
}
else if (scriptFile != null)
{
    if (!File.Exists(scriptFile))
    {
        Console.Error.WriteLine($"File not found: {scriptFile}");
        return 1;
    }
    code = File.ReadAllText(scriptFile);
}
else
{
    PrintUsage();
    return 1;
}

// -- Configure script options --
var openXmlAssembly = typeof(DocumentFormat.OpenXml.Packaging.WordprocessingDocument).Assembly;
var openXmlFrameworkAssembly = typeof(DocumentFormat.OpenXml.OpenXmlElement).Assembly;

var options = ScriptOptions.Default
    .WithReferences(
        openXmlAssembly,
        openXmlFrameworkAssembly,
        typeof(System.Text.RegularExpressions.Regex).Assembly,
        typeof(System.Linq.Enumerable).Assembly,
        typeof(Console).Assembly,
        typeof(File).Assembly,
        Assembly.GetAssembly(typeof(object))!
    )
    .WithImports(
        // Basics only — no OpenXML namespaces auto-imported.
        // Scripts explicitly declare what they use, avoiding type conflicts.
        "System",
        "System.IO",
        "System.Linq",
        "System.Text.RegularExpressions",
        "System.Collections.Generic"
    );

// -- Run --
var globals = new ScriptGlobals { Args = scriptArgs };

try
{
    await CSharpScript.EvaluateAsync(code, options, globals);
    return 0;
}
catch (CompilationErrorException ex)
{
    Console.Error.WriteLine("Compilation error:");
    foreach (var diag in ex.Diagnostics)
        Console.Error.WriteLine($"  {diag}");
    return 2;
}
catch (Exception ex)
{
    Console.Error.WriteLine($"Runtime error: {ex.GetType().Name}: {ex.Message}");
    Console.Error.WriteLine(ex.StackTrace);
    return 3;
}

void PrintUsage()
{
    Console.Error.WriteLine("Usage:");
    Console.Error.WriteLine("  office-eval <script.csx> [-- args...]");
    Console.Error.WriteLine("  office-eval -e \"code\"    [-- args...]");
    Console.Error.WriteLine();
    Console.Error.WriteLine("Options:");
    Console.Error.WriteLine("  -e <code>    Execute inline C# code");
    Console.Error.WriteLine("  --version    Show version");
    Console.Error.WriteLine("  -h, --help   Show this help");
}

/// <summary>
/// Globals injected into the script. Public members become top-level variables.
/// </summary>
public class ScriptGlobals
{
    public IList<string> Args { get; set; } = new List<string>();
}
