# office-eval 设计文档

> 日期：2026-04-14
> 状态：设计完成，待实施
> 前置依赖：.NET 9 SDK（开发阶段）

## 1. 问题定义

AI agent 无法可靠编辑 Office 文档，不是能力问题，是工具链问题。python-docx / openpyxl 等三方库对 OpenXML 规范实现残缺，编辑复杂文档必坏；微软自己的 OpenXML SDK（C#）覆盖完整但需要 ~500 MB .NET SDK 才能运行。office-eval 的目标是**消除这个安装门槛**：构建一个自包含 CLI，内嵌 Roslyn 编译器 + OpenXML SDK，下载即用，让 agent 和用户能直接运行 C# 脚本编辑 .docx / .xlsx / .pptx。

成功标准：在没有安装 .NET SDK / Runtime 的 Windows 机器上，下载 office-eval.exe 后能直接执行 OpenXML 编辑脚本。

## 2. 调研发现

### 2.1 现有 C# 脚本运行器

调查了 dotnet-script、cs-script、scriptcs、.NET 10 `dotnet run app.cs` 四个工具。**结论：全部不满足"无 SDK 即用"需求。**

| 工具 | NuGet 运行时解析 | 无 SDK 运行 | 状态 |
|------|-----------------|------------|------|
| dotnet-script | `#r "nuget:..."` | 不行，NuGet restore 需要 SDK | 活跃 |
| cs-script | 支持，靠 `dotnet restore` | 不行 | 活跃 |
| scriptcs | 旧版支持 | 不行 | 已废弃 |
| .NET 10 file-based | `#:package` 指令 | 不行，就是 SDK 功能 | 新特性 |

核心瓶颈：NuGet 包恢复依赖 MSBuild / `dotnet restore` 基础设施，这些是 SDK 的组成部分，无法仅靠 Runtime 运行。

自建工具通过**预打包 OpenXML SDK DLL**（而非运行时 NuGet 解析）绕过了这个限制。

**来源：**
- [dotnet-script GitHub](https://github.com/dotnet-script/dotnet-script)
- [cs-script GitHub](https://github.com/oleg-shilo/cs-script)
- [.NET Blog: dotnet run app.cs](https://devblogs.microsoft.com/dotnet/announcing-dotnet-run-app/)

### 2.2 Roslyn 编译路径选择

调研了 Roslyn Scripting API（`CSharpScript`）和全编译路径（`CSharpCompilation.Emit()`）两种方案。

| 方面 | Scripting API | CSharpCompilation.Emit() |
|------|-------------|--------------------------|
| 脚本风格 | 顶层语句，无需 class/Main | 完整 C# 程序 |
| 预导入命名空间 | `WithImports()` 原生支持 | 需自行注入 GlobalUsings |
| 内存泄漏 | 每次 EvaluateAsync 生成不可卸载的 assembly | 可用 collectible AssemblyLoadContext |
| NuGet `#r` | 不原生支持，需自行实现 | 同 |
| NativeAOT | 不兼容（需要 JIT） | 同 |

关键发现：
- `#r "nuget:..."` 不是 Roslyn 原生功能，是 dotnet-script 封装的。原生只支持 `#r "path/to/dll"`
- Scripting API 的内存泄漏问题（roslyn #41722）对一次性执行的 CLI 无影响——进程退出后内存自然释放
- Scripting API 的顶层语句风格更适合 agent 生成的短脚本

**来源：**
- [Roslyn Scripting API Samples](https://github.com/dotnet/roslyn/blob/main/docs/wiki/Scripting-API-Samples.md)
- [roslyn #41722: EvaluateAsync memory leak](https://github.com/dotnet/roslyn/issues/41722)
- [roslyn #5654: #r "nuget:" not native](https://github.com/dotnet/roslyn/issues/5654)

### 2.3 Self-contained 发布体积

| 策略 | 体积 | 可行性 |
|------|------|--------|
| Self-contained（无优化） | ~120-160 MB | 可行 |
| + ReadyToRun | +10-20% 但启动更快 | 可行 |
| + 压缩单文件 | ~60-80 MB，首次解压 | 可行但首次启动慢 |
| Framework-dependent | ~30-60 MB（需 .NET Runtime） | 备选 |
| IL Trimming | — | **不可行**，Roslyn 大量反射 |
| NativeAOT | — | **不可行**，Roslyn 需要 JIT |

关键限制：
- Roslyn 不支持 IL Trimming（roslyn #48873），不支持 NativeAOT
- `PublishSingleFile` 会导致 Roslyn 找不到 assembly（`Assembly.Location` 返回空），需 `IncludeAllContentForSelfExtract=true` 变通
- 压缩单文件可将分发体积压到 ~60-80 MB，但首次运行需解压 1-3 秒

**来源：**
- [roslyn #48873: Trimming not supported](https://github.com/dotnet/roslyn/issues/48873)
- [roslyn #50719: PublishSingleFile breaks scripting](https://github.com/dotnet/roslyn/issues/50719)
- [Native AOT overview](https://learn.microsoft.com/en-us/dotnet/core/deploying/native-aot/)

### 2.4 OpenXML SDK 文档体系

| 来源 | 内容 | 用途 |
|------|------|------|
| [OfficeDev/open-xml-docs](https://github.com/OfficeDev/open-xml-docs) | How-to 指南 + 代码示例（markdown） | 进阶操作参考 |
| NuGet XML doc 文件 | 每个公开类型/成员的签名 + summary | API 字典查表 |
| ECMA-376 规范 | OpenXML 底层 XML schema，5000+ 页 | 太重，不实用 |

## 3. 设计讨论过程

### 3.1 编译路径：Scripting API vs CSharpCompilation.Emit()

调研建议用全编译路径（`CSharpCompilation.Emit()`）以避免内存泄漏和脚本模式限制。但讨论后确认 **Scripting API 更适合 office-eval**，理由：

1. 用途是一次性脚本执行——进程跑完就退出，内存泄漏不是问题
2. 顶层语句风格是 agent 生成代码最自然的形态——不需要 `class`/`Main` 样板
3. `WithImports()` + `WithReferences()` 天然支持预导入，用户脚本开箱即用

用户确认："改成顶层语句更简单，而且不影响功能。"

### 3.2 发布策略：Self-contained vs Framework-dependent

讨论了两种路线。用户指出 **.NET Runtime 覆盖率不高**，选择 **self-contained 为默认发布方式**，优先零依赖体验。framework-dependent 作为备选，开发阶段使用。

### 3.3 脚本全局变量：最小方案 vs 便利方案

两个选项：(A) 只暴露 `Args`，用户自己管理文档生命周期；(B) 预打开文档 + 自动保存。

选择 **(A) 只暴露 `Args`**。用户洞察："反正会给 Agent 一个 API 文档和一个使用手册，告诉他们要干什么就行了。" 这完全契合"不封装操作"的设计理念——agent 自己写 Open/Save/Dispose，比隐式封装更灵活。

### 3.4 v1 输入模式范围

四种模式（文件、管道、内联、REPL）中选择 **v1 支持文件执行 + 内联（`-e`）**。

- 管道：agent 生成代码后写入 `.csx` 文件再执行完全可行，省一个临时文件不重要
- REPL：对 agent 无用，对人类有用但不急

### 3.5 额外包引用

讨论了脚本中能否引用新 NuGet 包。结论：**不支持运行时 NuGet 解析**（需要 SDK 基础设施），但支持 `#r "path/to/local.dll"` 引用本地 DLL。.NET Runtime 标准库 + OpenXML SDK 已覆盖 Office 文档编辑的所有场景。若后续发现某个库高频需要，直接预打包进 office-eval。

用户确认："我们这个应用场景很窄，这些应该够了。"

### 3.6 文档架构：三层 skill 设计

用户提出关键设计——**不是把文档打包进工具，而是封装成 agent skill**：

```
skill.md（入口，agent 首先读这个）
├── 理念：为什么用 office-eval，不用 python-docx / officecli
├── 使用时机：什么场景该用，什么场景不该用
├── 常见示例：Word/Excel/PPT 各高频操作，直接能抄
├── 指路牌：
│   ├── 高级操作 → 查 open-xml-docs
│   └── 确认签名 → 查 API doc
```

这是"给 agent VM + API 文档"理念的完整实现：skill.md 是使用手册（覆盖 80% 场景），open-xml-docs 是进阶指南，API doc 是字典。

## 4. 设计方案

### 4.1 架构

```
office-eval.exe（self-contained，零外部依赖）
├── Roslyn Scripting（Microsoft.CodeAnalysis.CSharp.Scripting）
├── OpenXML SDK（DocumentFormat.OpenXml）
│   ├── WordprocessingDocument → .docx
│   ├── SpreadsheetDocument   → .xlsx
│   └── PresentationDocument  → .pptx
└── 脚本加载器
    ├── 文件模式：读 .csx → ScriptOptions → EvaluateAsync
    └── 内联模式：-e "code" → 同上
```

### 4.2 CLI 接口

```bash
# 文件执行
office-eval script.csx -- thesis.docx arg2

# 内联执行
office-eval -e "var doc = WordprocessingDocument.Open(Args[0], true); ..." -- thesis.docx
```

脚本内通过 `Args`（`IList<string>`）访问 `--` 后的参数。

### 4.3 预导入命名空间

脚本中自动可用，无需 `using`：

```csharp
// 基础
System, System.IO, System.Linq,
System.Text.RegularExpressions, System.Collections.Generic

// OpenXML 通用
DocumentFormat.OpenXml, DocumentFormat.OpenXml.Packaging

// Word
DocumentFormat.OpenXml.Wordprocessing

// Excel
DocumentFormat.OpenXml.Spreadsheet

// PowerPoint
DocumentFormat.OpenXml.Presentation, DocumentFormat.OpenXml.Drawing
```

注意：`Drawing` 和 `Wordprocessing` 存在类名冲突（如 `Text`、`Run`）。混合操作时用户自行在脚本中添加 `using Drawing = DocumentFormat.OpenXml.Drawing;` 别名。

### 4.4 全局变量

仅注入 `Args`：

```csharp
public class ScriptGlobals
{
    public IList<string> Args { get; set; }
}
```

不预打开文档、不自动保存。脚本完全控制文档生命周期。

### 4.5 扩展性

- `#r "path/to/local.dll"` — 引用本地 DLL，Roslyn 原生支持
- `using` — 可使用 .NET Runtime 标准库中的任何命名空间
- 不支持运行时 NuGet 包解析

### 4.6 发布

```bash
dotnet publish -r win-x64 --self-contained -p:PublishSingleFile=true \
  -p:IncludeAllContentForSelfExtract=true \
  -p:EnableCompressionInSingleFile=true
```

预计产物：~60-80 MB 单文件（压缩），首次运行解压到临时目录。

### 4.7 文档体系（与工具并列，非打包进工具）

```
officedit/
├── src/                    # office-eval 源码
├── docs/
│   ├── skill.md            # agent skill 入口文档
│   ├── open-xml-docs/      # 克隆 OfficeDev/open-xml-docs
│   └── api-doc/            # 从 NuGet 提取的 XML doc 文件
└── examples/               # 示例脚本
```

## 5. 实施路线图

### Phase 1：核心 CLI（最小可用）

- 创建 .NET 9 控制台项目，引用 Roslyn Scripting + OpenXML SDK
- 实现文件执行模式：读取 `.csx` → 配置 ScriptOptions（imports + references）→ `CSharpScript.EvaluateAsync()`
- 实现 `Args` 全局变量注入
- 验证：用现有 ThesisEditor 的逻辑改写为 `.csx` 脚本并成功执行
- 预计变更：~200 行 C#

### Phase 2：内联模式 + 错误处理

- 实现 `-e` 内联执行
- 编译错误友好输出（行号、错误信息）
- 运行时异常捕获和堆栈输出
- 预计变更：~100 行 C#

### Phase 3：Self-contained 发布

- 配置 `.csproj` 发布参数
- 测试压缩单文件发布
- 在无 .NET SDK/Runtime 的机器上验证
- 测量实际体积和首次启动时间

### Phase 4：文档体系

- 克隆 `OfficeDev/open-xml-docs` 到 `docs/open-xml-docs/`
- 从 NuGet 提取 XML doc 文件到 `docs/api-doc/`
- 编写 `docs/skill.md`（理念、使用时机、常见示例、指路牌）
- 编写示例脚本集（Word/Excel/PPT 各 2-3 个典型操作）

## 6. 开放问题

1. 项目名称：`office-eval` 还是其他？（proposal 中用的是 `office-eval`）
2. 是否支持 Linux/macOS 的 self-contained 发布？（当前只规划了 win-x64）
3. 压缩单文件首次解压的临时目录清理策略
4. 是否需要版本号 / `--version` 输出

## 7. 信息来源

**竞品/工具：**
- [dotnet-script](https://github.com/dotnet-script/dotnet-script)
- [cs-script](https://github.com/oleg-shilo/cs-script)
- [.NET 10 dotnet run app.cs](https://devblogs.microsoft.com/dotnet/announcing-dotnet-run-app/)

**技术文档：**
- [Roslyn Scripting API Samples](https://github.com/dotnet/roslyn/blob/main/docs/wiki/Scripting-API-Samples.md)
- [Native AOT overview](https://learn.microsoft.com/en-us/dotnet/core/deploying/native-aot/)
- [.NET trimming docs](https://learn.microsoft.com/en-us/dotnet/core/deploying/trimming/trim-self-contained)
- [ReadyToRun deployment](https://learn.microsoft.com/en-us/dotnet/core/deploying/ready-to-run)

**Roslyn Issues（关键限制）：**
- [#41722: EvaluateAsync memory leak](https://github.com/dotnet/roslyn/issues/41722)
- [#48873: Trimming not supported](https://github.com/dotnet/roslyn/issues/48873)
- [#50719: PublishSingleFile breaks scripting](https://github.com/dotnet/roslyn/issues/50719)
- [#5654: #r "nuget:" not native](https://github.com/dotnet/roslyn/issues/5654)

**OpenXML SDK 文档：**
- [OfficeDev/open-xml-docs](https://github.com/OfficeDev/open-xml-docs)
- [dotnet/Open-XML-SDK](https://github.com/dotnet/Open-XML-SDK)
- [DocumentFormat.OpenXml on NuGet](https://www.nuget.org/packages/documentformat.openxml)
- [ECMA-376 标准](https://ecma-international.org/publications-and-standards/standards/ecma-376/)

**用户领域知识：**
- .NET Runtime 覆盖率不高，self-contained 优先
- 应用场景窄，标准库 + OpenXML SDK 足够覆盖
- "给 agent VM + API 文档，而非封装命令"的设计理念
- 三层文档架构：skill.md → open-xml-docs → API doc
