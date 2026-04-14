# 踩坑记录、检查清单与常见错误

## 工作流检查清单

每次编辑文档前后对照检查：

- [ ] 编辑输出到副本（如 `xxx_edited.docx`），不动原文件
- [ ] 涉及域代码的脚本先在测试文件上验证
- [ ] 替换前打印完整 run 文本确认上下文
- [ ] 编辑后用只读模式打开副本验证修改生效
- [ ] 检查是否有误替换（特别是数字/百分比）
- [ ] 跨 Run 替换后检查相邻 Run 有无残留文本
- [ ] 全角/半角字符都要匹配
- [ ] 确保 Word 已关闭，避免 IOException
- [ ] 含空格的 Text 元素设置了 `SpaceProcessingModeValues.Preserve`
- [ ] Agent 不触碰原文件，交给用户决定是否覆盖

---

## 常见编译错误

| 错误 | 原因 | 解决 |
|------|------|------|
| `CS1002: 应输入 ;` | `using var doc = ...` 在脚本模式下被解析为命名空间导入 | 改用 `var doc = ...; doc.Dispose();` |
| `CS0104: 不明确的引用` | `Paragraph`/`Text` 等类名在多个命名空间中存在 | 只导入需要的命名空间，或用别名 `using Drawing = ...` |
| `CS0117: 不包含定义` | OpenXML SDK 的 API 和直觉不一样 | 查 `docs/api-doc/` 确认 API 存在 |
| `CS8602: 空引用警告` | OpenXML 属性可能为 null | 用 `!` 断言：`doc.MainDocumentPart!.Document.Body!` |

---

## 已验证可靠的操作（来自实际生产使用）

以下操作已在真实论文编辑中反复验证，可放心使用：

- 批量文本替换（℃→°C、术语统一、占位符替换）
- 表格边框粗细调整（54 张表格，中间线/外框分别设置）
- 标题加粗/去粗（86 个标题，操作 StyleRunProperties）
- 下标/上标格式化（63 处 Q(15s) 等，`vertAlign = "subscript"`）
- ± 符号字体统一（240 处，含 8 个需拆分 Run 的混合内容）
- 数字与单位间距规范化（250 处，SI/GB 3101 标准）
- 段落对齐方式修改（`JustificationValues.Both`）
- 移除高亮标记（103 处，删除 `RunProperties.Highlight`）
- 节距/页边距调整（SectionProperties 操作）
- 关键词字体修正（`rFonts` 的 ascii/hAnsi/eastAsia 分别设置）
- `SpaceProcessingModeValues.Preserve` 保留空格

---

## 文档层级结构速查

### Word (.docx)
```
WordprocessingDocument
└── MainDocumentPart
    └── Document
        └── Body
            ├── Paragraph (p)
            │   ├── ParagraphProperties (pPr) — 段落格式
            │   └── Run (r)
            │       ├── RunProperties (rPr) — 字符格式（字体/加粗/字号/颜色）
            │       └── Text (t) — 文本内容
            ├── Table
            │   └── TableRow → TableCell → Paragraph → ...
            └── SectionProperties — 页面设置
```

### Excel (.xlsx)
```
SpreadsheetDocument
└── WorkbookPart
    ├── Workbook → Sheets → Sheet（名称 + 关系 ID）
    └── WorksheetPart
        └── Worksheet → SheetData → Row → Cell
            Cell.DataType: Number | SharedString | Boolean | ...
```

### PowerPoint (.pptx)
```
PresentationDocument
└── PresentationPart
    ├── Presentation → SlideIdList → SlideId（关系 ID）
    └── SlidePart
        └── Slide → CommonSlideData → ShapeTree → Shape
            └── TextBody → Drawing.Paragraph → Drawing.Run → Drawing.Text
```

---

## 进阶参考

### 需要更复杂的操作？

查阅 `docs/open-xml-docs/docs/` 目录下的 How-to 指南：

- **Word**: `docs/word/` — 样式、表格、页眉页脚、图片插入、修订接受
- **Excel**: `docs/spreadsheet/` — 创建工作表、插入图表、合并单元格、大文件 SAX 读取
- **PPT**: `docs/presentation/` — 插入幻灯片、添加动画/音视频、母版操作

### 需要确认函数签名、属性名、类层级？

查阅 `docs/api-doc/` 目录下的 XML 文档文件（从 NuGet 包提取的 IntelliSense 数据）：

- `DocumentFormat.OpenXml.xml` — 所有 OpenXML 类型的完整 API 参考
- 按类名搜索即可找到所有公开属性和方法

### 关键概念

- **Run** 是最小的格式单元。同一个 Run 内的文本共享相同格式（字体、粗体、颜色等）
- **ParagraphId** 是每个段落的唯一标识符，可用于精确定位
- **SharedString** 是 Excel 的字符串去重机制。读单元格时要检查 `DataType` 是否为 `SharedString`，是则去 SharedStringTable 查真实文本
- **Drawing 命名空间** 在 PPT 中用于文本和图形。和 Word 的 `Text`/`Run` 同名但不同类，需要用别名区分
- **FontSize 单位是半磅**（half-point）。12pt = `Val = "24"`

---

## 踩坑记录（来自真实生产事故）

以下是实际使用中踩过的坑，编写脚本时必须注意。

### 多级列表（numbering.xml）

- **不要往章标题段落里插入任何额外元素**（隐藏域、重置 run 等）。曾因此导致章标题被 Word 合并/损坏，整个文档结构崩溃
- **每章用独立 AbstractNum**，比 `LevelOverride + StartOverrideNumberingValue` 更可靠。后者与 `LevelRestart` 交互复杂，容易计数器跨章累加
- **Level 定义**：用 `new Level(子元素...)` 构造器传入，只定义实际使用的 level 数（不需要补全 9 个），不加 `LevelRestart`、`nsid`、`tmpl`、`isLgl` 等额外属性。加了反而 Word 不识别
- `numbering.xml` 里 AbstractNum/NumId/Level 的绑定关系极复杂，裸 XML 操作风险最高的区域

### 域代码（fldChar + instrText）

- `fldChar(begin)` + `instrText` + `fldChar(separate)` + 缓存值 + `fldChar(end)` 顺序和嵌套要求严格，少一个元素 Word 就不识别
- **SEQ 域跨章重置**：不要用 `\s 1`（需要内置 Heading 1 才生效），不要往标题段落插隐藏 `\r 0 \h` 域。安全做法是每章用独立 SEQ 名称（`Tbl1`/`Tbl2`/`Fig3` 等），彻底隔离计数器
- 涉及域代码的脚本，**必须先在副本上测试，确认无误后再应用到正式文件**

### 跨 Run 文本处理

Word 经常把一段文字拆成多个 Run（如 "表 " 和 "3-13" 在不同 Run 里）。直接按单个 Run 的 `Text` 匹配会找不到完整字符串。
- 需要拼接相邻 Run 的文本做匹配
- 替换后必须清理相邻 Run 中的残留文本片段，否则出现重复字符
- 实际案例：修正 "表 3-13" → "表3-14"，涉及跨 Run 替换 + 清理残留

### 关键 XML 属性（只有直接读 XML 才能看到）

- `w:hint="eastAsia"` — 告诉 Word 用 eastAsia 字体渲染该 Run。会导致 ± 等 Latin 字符被宋体渲染而非 TNR
- `w:rFonts` 的 `ascii`/`hAnsi`/`eastAsia`/`cs` 四个属性分别控制不同 Unicode 范围的字体选择
- `xml:space="preserve"` / `SpaceProcessingModeValues.Preserve` — 不设则 Word 吞掉 Text 元素中的空格
- `w:vanish` — 隐藏 Run（域代码中常见），遍历文本时注意跳过
- `w:vertAlign val="subscript"/"superscript"` — 下标/上标格式

### Run 拆分技巧（混合内容处理）

当一个 Run 包含特殊字符和普通字符时，不能直接改整个 Run 的字体（会影响其他字符）。需要：
1. 找到目标字符在 Text 中的位置
2. 把 Run 拆成 3 个：前文本 Run + 目标字符 Run + 后文本 Run
3. 只给目标字符 Run 设置特定字体/格式

### 全角/半角字符

`Q（15s）`（全角括号 U+FF08/FF09）和 `Q(15s)`（半角括号）是不同字符序列。正则匹配时必须同时覆盖两种形式，否则漏改。

### 文件锁定

Word 打开文件时，C# OpenXML SDK 无法访问同一文件（IOException）。必须先关闭 Word 或操作副本。

### Roslyn 脚本模式语法限制

- **`using var` 不能用**。Roslyn 脚本模式把 `using` 开头的语句解析为命名空间导入指令，`using var doc = ...` 会报 `CS1002: 应输入 ;`。必须用 `var doc = ...; doc.Dispose();` 手动管理生命周期
- **命名空间冲突**：`Wordprocessing.Paragraph` / `Drawing.Paragraph` / `Spreadsheet`  等同名类型会导致 `CS0104` 歧义错误。默认只导入 `Wordprocessing`，Excel/PPT 脚本需在开头自行添加 `using`

### 脚本幂等性

- 脚本应设计为可重复运行。已转换的标注再次用正则匹配时，域代码的缓存值会干扰
- 每次运行前应检查是否已经处理过（判断目标状态是否已存在），避免重复操作导致结构损坏
