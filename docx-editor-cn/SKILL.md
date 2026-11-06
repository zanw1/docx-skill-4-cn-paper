---
name: docx-editor-cn
description: "Use this skill whenever the user wants to create, read, edit, or manipulate Word documents (.docx files). Triggers include: any mention of \"Word doc\", \"word document\", \".docx\", or requests to produce professional documents with formatting like tables of contents, headings, page numbers, or letterheads. Also use when extracting or reorganizing content from .docx files, inserting or replacing images in documents, performing find-and-replace in Word files, working with tracked changes or comments, or converting content into a polished Word document. If the user asks for a \"report\", \"memo\", \"letter\", \"template\", or similar deliverable as a Word or .docx file, use this skill. Do NOT use for PDFs, spreadsheets, Google Docs, or general coding tasks unrelated to document generation."
license: Proprietary. LICENSE.txt has complete terms
---

# DOCX creation, editing, and analysis

## Overview

A .docx file is a ZIP archive containing XML files.

## Quick Reference

| Task | Approach |
|------|----------|
| Read/analyze content | `pandoc` or unpack for raw XML |
| Create new document | `node scripts/new_doc.js` (edit CONTENT section first) |
| Edit existing document | Unpack → edit XML → repack - see Editing Existing Documents below |
| Insert 三线表 (XML editing) | `python scripts/table.py unpacked/ "1-1" "标题" --headers … --rows …` |
| Insert block formula (XML editing) | `python scripts/formula.py unpacked/ "LaTeX" 1 --anchor "锚文本"` |
| 双栏论文中的通栏图片＋图题 | 连续分节实现 2 → 1 → 2 栏；见下方“通栏图片”章节与 `scripts/new_doc.js` 的 `columnSection()` |

### Converting .doc to .docx

Legacy `.doc` files must be converted before editing:

```bash
python scripts/office/soffice.py --headless --convert-to docx document.doc
```

### Reading Content

```bash
# Text extraction with tracked changes
pandoc --track-changes=all document.docx -o output.md

# Raw XML access
python scripts/office/unpack.py document.docx unpacked/
```

### Converting to Images

```bash
python scripts/office/soffice.py --headless --convert-to pdf document.docx
pdftoppm -jpeg -r 150 document.pdf page
```

### Accepting Tracked Changes

To produce a clean document with all tracked changes accepted (requires LibreOffice):

```bash
python scripts/accept_changes.py input.docx output.docx
```

---

## Creating New Documents

Generate .docx files with JavaScript, then validate. Install: `npm install -g docx`

### Setup
```javascript
const { Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell, ImageRun,
        Header, Footer, AlignmentType, PageOrientation, LevelFormat, ExternalHyperlink,
        TableOfContents, HeadingLevel, BorderStyle, WidthType, ShadingType,
        VerticalAlign, PageNumber, PageBreak } = require('docx');

const doc = new Document({ sections: [{ children: [/* content */] }] });
Packer.toBuffer(doc).then(buffer => fs.writeFileSync("doc.docx", buffer));
```

### Validation
After creating the file, validate it. If validation fails, unpack, fix the XML, and repack.
```bash
python scripts/office/validate.py doc.docx
```

### Page Size

```javascript
// Chinese academic papers use A4 with 2.5cm margins on all sides
sections: [{
  properties: {
    page: {
      size: {
        width: 11906,   // A4 width in DXA
        height: 16838   // A4 height in DXA
      },
      margin: { top: 1418, right: 1418, bottom: 1418, left: 1418 } // 2.5cm = 1418 DXA
    }
  },
  children: [/* content */]
}]
```

**Common page sizes (DXA units, 1440 DXA = 1 inch, 567 DXA = 1 cm):**

| Paper | Width | Height | Content Width (2.5cm margins) |
|-------|-------|--------|-------------------------------|
| A4 (Chinese standard) | 11,906 | 16,838 | 9,070 |
| US Letter | 12,240 | 15,840 | 9,404 |

**Landscape orientation:** docx-js swaps width/height internally, so pass portrait dimensions and let it handle the swap:
```javascript
size: {
  width: 12240,   // Pass SHORT edge as width
  height: 15840,  // Pass LONG edge as height
  orientation: PageOrientation.LANDSCAPE  // docx-js swaps them in the XML
},
// Content width = 15840 - left margin - right margin (uses the long edge)
```

### Styles (Academic Thesis Formatting)

**CRITICAL**: When generating academic papers, course designs, or mathematical documents, strictly adhere to the following Chinese academic formatting standards.

**1. Global Page & Normal Text (【全局页面与正文规范】)**
- **Page**: A4 (11906 × 16838 DXA), margins all 2.5cm (`1418` DXA).
- **Fonts**: Chinese = SimSun (宋体), English/Math/Code = Cambria Math (default); Times New Roman is an acceptable alternative for English text.
- **Size**: 12pt (小四) -> `size: 24` in docx-js (half-points).
- **Paragraph**: First-line indent 2 characters (`firstLine: 480` DXA), Line spacing **single** (`line: 240, lineRule: LineRuleType.AUTO`) by default; alternatives: fixed 20pt (`line: 400, lineRule: LineRuleType.EXACT`) or 1.5x (`line: 360, lineRule: LineRuleType.AUTO`). Before/After spacing 0pt.

**2. Headings (【标题规范】)**
- **Heading 1 (一级标题)**: SimHei (黑体), 16pt (三号, `size: 32`), Bold, Centered, 1.2x line spacing (`line: 288`). Auto-numbering: Arabic `1`, `2`, `3` (displayed in heading text or prefixed manually as 一、二、). See Heading & Reference Auto-Numbering below.
- **Heading 2 (二级标题)**: SimHei (黑体), 14pt (四号, `size: 28`), Bold, Left-aligned, 1.5x line spacing (`line: 360`). Auto-numbering: `1.1`, `1.2`, `2.1` (chapter-tracking Arabic decimal).
- **Heading 3 (三级标题)**: SimHei (黑体), 12pt (小四, `size: 24`), Bold, Left-aligned, 1.1x line spacing (`line: 264`). Auto-numbering: `1.1.1`, `2.4.1`.

**⚠️ Heading Numbering Rule**: Use a **single all-decimal multi-level numbering config** for all three heading levels. This is the ONLY way to get correct cross-chapter tracking (1.1, 2.1, 2.4 etc.) while keeping H2/H3 numbers purely Arabic. If CHINESE_COUNTING is used for H1 level, H2 will render as `二.4` instead of `2.4` — this is a known OOXML rendering issue.

**3. Figures, Tables & Math (【图表与公式规范】)**
- **Figure Captions (图标题)**: SimSun (宋体), 11pt (`size: 22`), Centered, Bold. Spacing: Single spacing, 0.5 lines before (`before: 120`), 3pt after (`after: 60`). Position: Below the figure. Format: "图 章-图序" e.g. "图 1-1".
- **Table Captions (表标题)**: SimSun (宋体), 11pt (`size: 22`), Centered, **no indent**, Bold. Spacing: Single spacing, 0.5 lines before (`before: 120`), 3pt after (`after: 60`). Position: Above the table. Format: "表 章-表序" e.g. "表 1-1". Tables MUST use **三线表 (three-line table)** style (thick top/bottom, thin after header, no other borders).
- **Block Math Formulas**: Use the formula table layout — a 3-column borderless table: [1cm spacer | formula centered | equation number right-aligned 1cm]. Use `scripts/formula.py` to generate the XML. Inline math: write as plain text with italic formatting.

**Creating a new document — use the template script:**
```bash
# Copy the template, edit the CONTENT section, then run:
node scripts/new_doc.js
# Outputs output.docx with all styles, numbering, and page setup pre-configured.
# Change OUTPUT_PATH inside the script if needed.
```

`scripts/new_doc.js` provides ready-to-use helper functions: `h1/h2/h3(text)`, `body(text)`, `tableCaption/figCaption(label)`, `threeLineTable(headers, rows, colWidths)`, `ref(text)`, `blank()`, and optional `columnSection(children, columnCount, options)` for continuous column layouts. The STYLES / NUMBERING / page constants at the top of the file are the canonical implementation of all specs below; refer to that file for exact docx-js values.

### Lists (NEVER use unicode bullets)

```javascript
// ❌ WRONG - never manually insert bullet characters
new Paragraph({ children: [new TextRun("• Item")] })  // BAD
new Paragraph({ children: [new TextRun("\u2022 Item")] })  // BAD

// ✅ CORRECT - use numbering config with LevelFormat.BULLET
const doc = new Document({
  numbering: {
    config: [
      { reference: "bullets",
        levels: [{ level: 0, format: LevelFormat.BULLET, text: "•", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      { reference: "numbers",
        levels: [{ level: 0, format: LevelFormat.DECIMAL, text: "%1.", alignment: AlignmentType.LEFT,
          style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
    ]
  },
  sections: [{
    children: [
      new Paragraph({ numbering: { reference: "bullets", level: 0 },
        children: [new TextRun("Bullet item")] }),
      new Paragraph({ numbering: { reference: "numbers", level: 0 },
        children: [new TextRun("Numbered item")] }),
    ]
  }]
});

// ⚠️ Each reference creates INDEPENDENT numbering
// Same reference = continues (1,2,3 then 4,5,6)
// Different reference = restarts (1,2,3 then 1,2,3)
```

### Tables (三线表 / Three-Line Table)

**Chinese academic papers MUST use 三线表**: thick top/bottom (1.5pt), thin after header (0.75pt), no other borders. Caption goes **above** the table.

**For new documents — use `threeLineTable()` helper in `scripts/new_doc.js`:**
```javascript
// Inside the CONTENT array:
tableCaption('符号说明'),
threeLineTable(
  ['符号', '说明'],                       // headers
  [['S', '状态空间'], ['A', '动作空间']], // data rows
  [1800, 7270]                            // DXA column widths; must sum to 9070
),
```

**For editing existing documents (XML) — use `scripts/table.py`:**
```bash
# Insert after paragraph containing anchor text:
python scripts/table.py unpacked/ "1-1" "符号说明" \\
    --headers "符号,说明" \\
    --rows '[["S","状态空间"],["A","动作空间"]]' \\
    --anchor "以下是符号说明"

# Custom column widths (must sum to 9070):
python scripts/table.py unpacked/ "2-1" "参数" \\
    --headers "参数,取值,说明" \\
    --rows '[["lr","0.001","学习率"]]' \\
    --widths "1800,1500,5770"

# Print XML snippet only (no file modification):
python scripts/table.py --caption "示例" \\
    --headers "列1,列2" --rows '[["a","b"]]'
```

**Key rules:**
- Content width = 9070 DXA (A4 2.5cm margins); `columnWidths` must sum exactly to 9070
- Always use `WidthType.DXA` — never `WidthType.PERCENTAGE`
- Table width = sum of `columnWidths`; set matching `width` on each cell too

### Images

```javascript
// CRITICAL: type parameter is REQUIRED
new Paragraph({
  children: [new ImageRun({
    type: "png", // Required: png, jpg, jpeg, gif, bmp, svg
    data: fs.readFileSync("image.png"),
    transformation: { width: 200, height: 150 },
    altText: { title: "Title", description: "Desc", name: "Name" } // All three required
  })]
})
```

### Page Breaks

```javascript
// CRITICAL: PageBreak must be inside a Paragraph
new Paragraph({ children: [new PageBreak()] })

// Or use pageBreakBefore
new Paragraph({ pageBreakBefore: true, children: [new TextRun("New page")] })
```

### 双栏论文中的通栏图片＋图题（Word / WPS）

适用于“正文双栏，局部图片与图题横跨两栏，之后恢复双栏”的论文。使用连续分节：**双栏正文 → 单栏图片＋图题 → 双栏正文**。分栏属于节属性；只增大图片宽度不能完成这种布局。这里的“通栏”指正文可用宽度，不包括页边距。

#### Word / WPS 手动操作

1. 显示段落标记，将光标放到准备插图的正文位置，选择【布局／页面布局】→【分隔符】→【分节符（连续）】。
2. 将光标放在新节内，选择【分栏】→【一栏】。在“更多分栏／应用于”中确保仅作用于目标节或选定区域，避免误改整篇文档；各版本菜单名称可能不同。
3. 插入图片，优先使用【嵌入型／与文字在一行】并居中。图片宽度不得超过页面宽度减去左右页边距，保持原始宽高比。
4. 紧接图片添加独立的图题段落，居中、无首行缩进。按论文模板选择“图 1-1”或“Fig. 1”；需要自动编号时使用题注／SEQ 字段。
5. 图片所在段落设置【与下段同页】；图题设置【段中不分页】，图题本身不要继续绑定后面的正文。图片段落使用单倍或至少值行距，避免固定的小行高裁切图片。
6. 在图题之后插入第二个【分节符（连续）】，将其后的新节设回【两栏】，恢复原来的栏宽和栏间距。

连续分节不主动要求换页，但不能保证图片一定留在当前页：图片＋图题放不下时会正常移动到后页。图片过高、段落的“段前分页”、脚注或不同的页面尺寸／方向也可能影响结果。双栏节在结束处可能重新平衡两栏文字；这是需要实际预览的排版行为。

#### 使用 `docx` / `scripts/new_doc.js` 生成

模板已提供 `columnSection(children, columnCount, options)`：栏数为 1 或 2，默认栏间距为 360 DXA（18pt），每节沿用模板的 A4 与 2.5cm 页边距，分节类型为 `SectionType.CONTINUOUS`。它是可选布局函数，默认单栏模板不会自动改成双栏。

在复制的 `new_doc.js` 中，用下例替换示例 CONTENT 和文档构建部分；保留模板前面的依赖、样式、编号和辅助函数。`architecture.png` 及原始尺寸必须替换为真实图片。没有单独的导入模块需要安装。

```javascript
// 正文前后分成不同的数组；章节与布局节互相独立。
const beforeFigure = [
  h1Chinese('一、引言'),
  body('这里是图片之前的双栏正文。'),
];
const afterFigure = [body('这里继续双栏正文。')];

// docx 图片尺寸使用像素；DXA / 15 = 96 DPI 下的像素数。
const sourceWidth = 1600;
const sourceHeight = 600;
// 模板导入的 docx.Math 会遮蔽原生 Math，因此显式使用 globalThis.Math。
const figureWidth = globalThis.Math.min(560, globalThis.Math.floor(CONTENT_W / 15));
const imageParagraph = new Paragraph({
  alignment: AlignmentType.CENTER,
  indent: { firstLine: 0 },
  spacing: { line: 240, lineRule: LineRuleType.AUTO },
  keepNext: true,
  keepLines: true,
  children: [new ImageRun({
    type: 'png',
    data: fs.readFileSync('architecture.png'),
    transformation: {
      width: figureWidth,
      height: figureWidth * sourceHeight / sourceWidth,
    },
    altText: { title: '系统架构', description: '系统组件及其连接关系', name: 'architecture' },
  })],
});
const caption = figCaption('Architecture overview', {
  format: 'continuous', label: 'Fig.',
});
const pageFooter = new Footer({ children: [new Paragraph({
  alignment: AlignmentType.CENTER,
  indent: { firstLine: 0 },
  children: [new TextRun({ children: [PageNumber.CURRENT] })],
})] });

const doc = new Document({
  styles: STYLES,
  numbering: NUMBERING,
  features: { updateFields: true },
  sections: [
    columnSection(beforeFigure, 2, { footers: { default: pageFooter } }),
    columnSection([imageParagraph, caption], 1),
    columnSection(afterFigure, 2),
  ],
});
// 保留模板原有 Packer.toBuffer(doc) 保存代码。
```

`figCaption('描述')` 仍生成“图 章-序”；`figCaption('描述', { format: 'continuous', label: 'Fig.' })` 使用共享 `SEQ figure \* ARABIC` 字段生成“Fig. 1、Fig. 2…”（标签与序号间保留空格）。新连续图题通过 `SimpleField` 写入初始编号的缓存显示值，使忽略字段更新的渲染器也能显示数字；正文顺序应与函数调用顺序一致。不要手工写入数字，也不要为每个布局节新建序列或重置序号。在 Word/WPS 中新增、删除或移动图题后更新字段并检查编号；`updateFields` 是更新请求，渲染器是否执行需要验证。

多张通栏图重复单栏图题节与双栏正文节即可，不要重复插入章节标题来切换布局。首页标题／摘要若需单栏，也可用独立的一栏节，但应依据目标期刊模板设置。表格和公式在双栏节内必须按实际栏宽重新计算，不能直接使用本技能其他示例中固定的 9070 DXA 全页宽度；两栏等宽时，单栏宽度为 `(CONTENT_W - columnSpace) / 2`。

#### 编辑现有 DOCX 的 OOXML 规则

按本技能的“解包 → 编辑 → 打包”流程，先定位目标位置所属的原始节及其完整 `w:sectPr`，将该节拆成前正文、图片＋图题、后正文三节：

- 前正文节的 `w:cols` 保留两栏。
- 图片＋图题节使用 `<w:cols w:num="1" w:space="360" w:equalWidth="1"/>`。
- 后正文节恢复原始两栏设置；不要强行把不等宽栏转换成等宽栏。
- 图片＋图题节和后正文节的节属性都包含 `<w:type w:val="continuous"/>`。前正文节保留原始的起节类型，避免改变它如何衔接更前面的内容。
- **位置与归属**：除文档最后一节外，`w:sectPr` 位于该节最后段落的 `w:pPr` 内；最后一节位于 `w:body` 的末尾。属性描述的是其所在节，而 `w:type` 决定该节如何接续前一节。不要把中间单栏节的属性错放进后正文节，或把每个节都写成 `w:body` 的直接子节点。
- 三节复制并保持原有纸张尺寸、方向、页边距、段落样式和相关关系。核查原节的页眉页脚引用、首页不同设置、脚注等：保持预期继承关系；已有 `w:pgNumType/@w:start` 只保留在应重启编号的位置，新插入的布局节通常不复制这个重启值。
- 其他既有节的属性不应被替换。复制带关系的元素必须保留或正确映射 `r:id`；不要重建 `styles.xml`。

图片段落的 `w:pPr` 中加入 `<w:keepNext/>` 和 `<w:keepLines/>`；图题段落加入 `<w:keepLines/>`。保留合法的元素顺序：这些分页控制在 `w:spacing`／`w:ind`／`w:jc` 之前，`w:sectPr` 在 `w:rPr` 之后、`w:pPrChange` 之前（存在时）。用嵌入式 `wp:inline` 引用图片，不用浮动锚点充当分栏实现。

#### 验证与故障定位

1. 校验打包文档的 XML、图片关系与内容类型；检查三个目标节的栏数为 2／1／2，后两节的类型为 continuous，并保留页面设置。
2. 用 Word/WPS 打开或通过 LibreOffice 转 PDF 检查：图片和图题跨正文全宽，上下正文仍为双栏，页码连续，图题没有单独掉到后页。LibreOffice 渲染只能证明该渲染器的结果，不能代替目标 Word/WPS 版本的验收。
3. 出现多余换页时，先查是否误用 `PageBreak`／`NEXT_PAGE`，再查 `pageBreakBefore`、图片高度、图题后的 `keepNext` 链及节间的页面设置差异。分栏符只进入下一栏；分页符只换页，均不能代替连续分节切换栏数。

可按下列场景核查修改后的模板；预期不代表所有 Word/WPS 版本都已实测：

| 输入或场景 | 预期结果 |
|---|---|
| 运行原来的单栏示例 | 仍为一个单栏节 |
| 双栏正文中加入一张通栏图及 Fig. 图题 | 栏数 2／1／2；空间充足时同页显示 Fig. 1 |
| 两张通栏图，中间仍有双栏正文 | 栏数 2／1／2／1／2；图题为 Fig. 1、Fig. 2 |
| 默认 `figCaption('描述')` | 保留“图 章-序”及章内 SEQ 字段 |
| `columnSpace: 720` | 两栏间距为 720 DXA |
| `children` 传入字符串 | 明确抛出 TypeError |
| 栏数传入 0、3 或非整数 | 明确抛出 RangeError |
| 栏间距为负数、非整数或不小于正文宽度 | 明确抛出 RangeError |
| 首节指定页脚，内容跨多页、多节 | 页面尺寸、页边距保留，页码按 1、2…连续 |
| 页末剩余空间不足，图片带多行图题 | 可容纳于单页的图片＋图题一起移到后页，之后恢复双栏 |

研究依据（2026-10-10 核查）：

- [Microsoft：用分节符修改局部布局，连续分节可改变同页栏数](https://support.microsoft.com/en-us/word/use-section-breaks-to-change-the-layout-or-formatting-in-one-section-of-your-word-document)
- [Microsoft：插入分节符，含 Windows / Mac 操作](https://support.microsoft.com/en-us/word/insert-a-section-break)
- [WPS Academy：分栏及栏宽、间距设置](https://www.wps.com/academy/layout-skill-how-can-we-split-columns/1861406/)
- [Microsoft：段中不分页、与下段同页](https://support.microsoft.com/en-us/word/keep-text-together-in-word)
- [Microsoft Open XML：SectionProperties 的位置](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.sectionproperties?view=openxml-3.0.1) 与 [SectionType 的含义](https://learn.microsoft.com/en-us/dotnet/api/documentformat.openxml.wordprocessing.sectiontype?view=openxml-3.0.1)
- [docx 9.6.1：节属性接口](https://github.com/dolanmiu/docx/blob/9.6.1/src/file/document/body/section-properties/section-properties.ts)、[栏属性实现](https://github.com/dolanmiu/docx/blob/9.6.1/src/file/document/body/section-properties/properties/columns.ts)、[节尾属性生成](https://github.com/dolanmiu/docx/blob/9.6.1/src/file/document/body/body.ts)
- [docx 9.6.1：SimpleField 的字段代码与缓存显示值](https://github.com/dolanmiu/docx/blob/9.6.1/src/file/paragraph/run/simple-field.ts)

### Table of Contents

```javascript
// The overridden Heading1/2/3 styles (defined in paragraphStyles above) include outlineLevel,
// which TableOfContents uses to build the TOC. Using heading: HeadingLevel.HEADING_X on
// each paragraph applies that overridden style automatically — giving both TOC support AND
// the custom SimHei/size formatting. Do NOT skip paragraphStyles and rely on built-in styles.
new TableOfContents("目录", { hyperlink: true, headingStyleRange: "1-3" })

// Apply to heading paragraphs — heading property applies the overridden style:
new Paragraph({
  heading: HeadingLevel.HEADING_1, // → w:pStyle "Heading1" (SimHei 16pt, outlineLevel:0)
  children: [new TextRun("一、引言")]
})
new Paragraph({
  heading: HeadingLevel.HEADING_2, // → w:pStyle "Heading2" (SimHei 14pt, outlineLevel:1)
  children: [new TextRun("1.1 研究背景")]
})
```

### Headers/Footers

```javascript
sections: [{
  properties: {
    page: { margin: { top: 1418, right: 1418, bottom: 1418, left: 1418 } } // 2.5cm = 1418 DXA
  },
  headers: {
    default: new Header({ children: [new Paragraph({ children: [new TextRun("Header")] })] })
  },
  footers: {
    default: new Footer({ children: [new Paragraph({
      children: [new TextRun("Page "), new TextRun({ children: [PageNumber.CURRENT] })]
    })] })
  },
  children: [/* content */]
}]
```

### Heading Auto-Numbering (Word per-chapter reset)

H1 uses Chinese numerals written literally in text (一、二、三). H2/H3 use **Word auto-numbering** via per-chapter numbering references (`sections_c1`, `sections_c2`, ...). Each chapter boundary resets H2/H3 counters automatically — adding/removing sections within a chapter triggers full auto-update.

**CRITICAL — numbering format rule:**

NEVER use `LevelFormat.CHINESE_COUNTING` for H1 in a multi-level config — `%1` in H2 text expands to "二", giving `二.4`. Always use `LevelFormat.DECIMAL` for every level.

Numbering references are generated dynamically by `buildNumberingConfig(chapterCount)`. Each chapter gets its own `sections_c{N}` reference with 2 levels (H2, H3). This ensures H2/H3 reset per chapter while Word handles all numbering.

```javascript
// H1 — Chinese numerals in text (no Word numbering)
h1Chinese('一、引言')   // → "一、引言"  (Heading1 style)

// H2 — Word auto-numbered, chapter.section format (1.1, 2.4)
h2('研究背景')          // → "1.1  研究背景"  (Word: sections_c1, level 0)
h2('文献综述')          // → "1.2  文献综述"

// H3 — Word auto-numbered, chapter.section.subsection format (1.1.1)
h3('研究现状')          // → "1.1.1  研究现状"  (Word: sections_c1, level 1)
h3('存在不足')          // → "1.1.2  存在不足"

// Next chapter — H2/H3 auto-reset with new chapter prefix
h1Chinese('二、方法')   // _chapter increments, sections_c2 used
h2('算法设计')          // → "2.1  算法设计"  (reset + "2." prefix)
h3('数据结构')          // → "2.1.1  数据结构"
```

The `_chapter` counter is incremented by `h1Chinese()` and used by `h2()`, `h3()`, `figCaption()`, and `tableCaption()` to select the correct numbering reference / SEQ identifier.

### Caption Auto-Numbering (SEQ Fields)

Captions use Word **SEQ fields** for per-chapter auto-numbering. The chapter number is a literal prefix, and the sequence counter is chapter-specific (e.g., `figure_c3`, `table_c1`).

```javascript
// Figure captions — auto: "图 章-序  描述"
figCaption('系统架构图')   // → "图 1-1  系统架构图"  (SEQ figure_c1)
figCaption('数据流图')     // → "图 1-2  数据流图"

// Table captions — auto: "表 章-序  描述"
tableCaption('符号说明')   // → "表 1-1  符号说明"  (SEQ table_c1)
```

Pass ONLY the description text. Never include manual "图 3-1" or "表 1-1" prefixes — the SEQ field handles numbering automatically.

### Markdown Number Stripping

When converting markdown, use these helpers to strip manual numbers from input text:

```javascript
stripH1Number('一、引言')          // → "引言"
stripH2Number('1.1 研究背景')     // → "研究背景"
stripCaptionNumber('图 3-1 系统') // → "系统"
```

### Block Formula Layout

Block formulas use a **3-column borderless table** so the equation number can be right-aligned while the formula is centered. Use `scripts/formula.py` to generate and insert formula blocks automatically.

```bash
# Generate a formula block and insert after paragraph containing anchor text:
python scripts/formula.py unpacked/ "Q_n(x,a) = r + \\gamma \\max_{a'} Q_{n-1}" 1 --anchor "由此可得"

# Output XML snippet only (no insertion):
python scripts/formula.py --latex "E=mc^2" --number 2
```

**Layout (A4, 9070 DXA content width):**
```
┌──────────┬──────────────────────────┬──────────┐
│  1cm gap │   formula   (centered)   │   (n)    │
│ 567 DXA  │        7936 DXA          │ 567 DXA  │
└──────────┴──────────────────────────┴──────────┘
  no border throughout; single line spacing
```

If direct XML editing is preferred, the skeleton is:
```xml
<w:tbl>
  <w:tblPr>
    <w:tblW w:w="9070" w:type="dxa"/>
    <w:tblBorders>
      <w:top w:val="none"/><w:left w:val="none"/>
      <w:bottom w:val="none"/><w:right w:val="none"/>
      <w:insideH w:val="none"/><w:insideV w:val="none"/>
    </w:tblBorders>
  </w:tblPr>
  <w:tblGrid>
    <w:gridCol w:w="567"/><w:gridCol w:w="7936"/><w:gridCol w:w="567"/>
  </w:tblGrid>
  <w:tr>
    <w:tc><w:tcPr><w:tcW w:w="567" w:type="dxa"/></w:tcPr>
      <w:p/></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="7936" w:type="dxa"/></w:tcPr>
      <w:p><w:pPr><w:jc w:val="center"/>
        <w:spacing w:line="240" w:lineRule="auto"/></w:pPr>
        <m:oMath xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">
          <!-- OMML formula here -->
        </m:oMath>
      </w:p></w:tc>
    <w:tc><w:tcPr><w:tcW w:w="567" w:type="dxa"/></w:tcPr>
      <w:p><w:pPr><w:jc w:val="right"/>
        <w:spacing w:line="240" w:lineRule="auto"/></w:pPr>
        <w:r><w:t>(%NUMBER%)</w:t></w:r>
      </w:p></w:tc>
  </w:tr>
</w:tbl>
```

### References (GB/T 7714-2015)

Use the `"references"` numbering reference defined above. Format text manually per GB/T 7714-2015 entry type.

```javascript
// Reference list style (add to paragraphStyles)
{ id: "Reference", name: "Reference", basedOn: "Normal",
  run: { font: { ascii: "Cambria Math", eastAsia: "SimSun" }, size: 24 },
  paragraph: { spacing: { line: 240, lineRule: LineRuleType.AUTO },
               indent: { left: 480, hanging: 480 } } }

// Usage — one Paragraph per entry:
new Paragraph({
  style: "Reference",
  numbering: { reference: "references", level: 0 },
  children: [new TextRun("Author. "),
             new TextRun({ text: "Title", italics: true }),
             new TextRun("[J]. Journal, Year, Vol(Issue): Pages.")]
})
```

**GB/T 7714-2015 common format templates (fill in fields):**

| Type | Tag | Format |
|------|-----|--------|
| 期刊论文 | [J] | 作者. 题名[J]. 刊名, 年, 卷(期): 起止页. |
| 专著 | [M] | 作者. 书名[M]. 出版地: 出版者, 出版年: 起止页. |
| 学位论文 | [D] | 作者. 题名[D]. 保存地: 保存单位, 年份. |
| 会议论文 | [C] | 作者. 题名[C]//会议名. 出版地: 出版者, 年: 起止页. |
| 网络资源 | [EB/OL] | 作者. 题名[EB/OL]. (发布日期)[引用日期]. URL. |

### Footer (Page Numbers)

Center-aligned page number in the footer. Uses A4 2.5cm margins.

```javascript
const { Footer, Paragraph, TextRun, PageNumber, AlignmentType } = require('docx');

sections: [{
  properties: {
    page: {
      size: { width: 11906, height: 16838 },           // A4
      margin: { top: 1418, right: 1418, bottom: 1418, left: 1418 } // 2.5cm
    }
  },
  footers: {
    default: new Footer({
      children: [
        new Paragraph({
          alignment: AlignmentType.CENTER,
          children: [new TextRun({ children: [PageNumber.CURRENT] })]
        })
      ]
    })
  },
  children: [/* content */]
}]

// For "第 X 页 共 Y 页" style:
new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [
    new TextRun("第 "),
    new TextRun({ children: [PageNumber.CURRENT] }),
    new TextRun(" 页  共 "),
    new TextRun({ children: [PageNumber.TOTAL_PAGES] }),
    new TextRun(" 页")
  ]
})
```

---

### Critical Rules for docx-js

- **Set page size explicitly** - always use A4 (11906 × 16838 DXA) with 2.5cm margins (1418 DXA) for Chinese academic documents
- **Landscape: pass portrait dimensions** - docx-js swaps width/height internally; pass short edge as `width`, long edge as `height`, and set `orientation: PageOrientation.LANDSCAPE`
- **Never use `\n`** - use separate Paragraph elements
- **Never use unicode bullets** - use `LevelFormat.BULLET` with numbering config
- **PageBreak must be in Paragraph** - standalone creates invalid XML
- **双栏正文中的通栏图题** - 用连续分节实现 2 → 1 → 2 栏；图片与图题同处单栏节，图片段落 keepNext，图题 keepLines；见“通栏图片”章节
- **ImageRun requires `type`** - always specify png/jpg/etc
- **Heading numbering: per-chapter Word auto-numbering** — H1 uses Chinese numerals in text (一、二、三) via `h1Chinese()`. H2/H3 use Word auto-numbering via per-chapter `sections_c{N}` references generated by `buildNumberingConfig(chapterCount)`. NEVER use `LevelFormat.CHINESE_COUNTING` in a multi-level config. Pass heading text WITHOUT manual number prefixes — Word handles all numbering.
- **Figure/Table captions: no manual numbers** — use `figCaption('描述')` without "图 3-1 " prefix. SEQ fields generate "图 章-序 描述" automatically. Same for `tableCaption()`. Use `stripCaptionNumber()` to clean markdown input.
- **Tables MUST be 三线表** - use thick top/bottom borders (size:12, 1.5pt) and thin header-bottom border (size:6, 0.75pt); set all other borders to `BorderStyle.NONE`; no fill color on header cells
- **Table caption: no first-line indent** - override Normal style's `firstLine` with `indent: { firstLine: 0 }` on the caption paragraph
- **Block formulas use 3-column table** - [567 DXA spacer | 7936 DXA formula centered | 567 DXA number right-aligned]; use `scripts/formula.py` to generate; never use plain `$$` LaTeX in the final docx
- **Always set table `width` with DXA** - never use `WidthType.PERCENTAGE` (breaks in Google Docs)
- **Tables need dual widths** - `columnWidths` array AND cell `width`, both must match
- **Table width = sum of columnWidths** - for DXA, ensure they add up exactly
- **Always add cell margins** - use `margins: { top: 80, bottom: 80, left: 120, right: 120 }` for readable padding
- **Use `ShadingType.CLEAR`** - never SOLID for table shading
- **TOC: use `heading: HeadingLevel.HEADING_X`** - applies the overridden Heading style (with `outlineLevel`), giving TOC recognition AND custom formatting simultaneously
- **Override built-in styles** - use exact IDs: "Heading1", "Heading2", etc.; always use `font: { ascii, eastAsia, hAnsi }` object form — `font: "SimHei"` string shorthand only sets the ASCII slot and leaves Chinese characters falling back to Normal font
- **Include `outlineLevel`** - required for TOC (0 for H1, 1 for H2, etc.)

---

## Editing Existing Documents

**Follow all 3 steps in order.**

### Step 1: Unpack
```bash
python scripts/office/unpack.py document.docx unpacked/
```
Extracts XML, pretty-prints, merges adjacent runs, and converts smart quotes to XML entities (`&#x201C;` etc.) so they survive editing. Use `--merge-runs false` to skip run merging.

### Step 2: Edit XML

Edit files in `unpacked/word/`. See XML Reference below for patterns.

**Use "Claude" as the author** for tracked changes and comments, unless the user explicitly requests use of a different name.

**Use the Edit tool directly for string replacement. Do not write Python scripts.** Scripts introduce unnecessary complexity. The Edit tool shows exactly what is being replaced.

**CRITICAL: Use smart quotes for new content.** When adding text with apostrophes or quotes, use XML entities to produce smart quotes:
```xml
<!-- Use these entities for professional typography -->
<w:t>Here&#x2019;s a quote: &#x201C;Hello&#x201D;</w:t>
```
| Entity | Character |
|--------|-----------|
| `&#x2018;` | ‘ (left single) |
| `&#x2019;` | ’ (right single / apostrophe) |
| `&#x201C;` | “ (left double) |
| `&#x201D;` | ” (right double) |

**Adding comments:** Use `comment.py` to handle boilerplate across multiple XML files (text must be pre-escaped XML):
```bash
python scripts/comment.py unpacked/ 0 "Comment text with &amp; and &#x2019;"
python scripts/comment.py unpacked/ 1 "Reply text" --parent 0  # reply to comment 0
python scripts/comment.py unpacked/ 0 "Text" --author "Custom Author"  # custom author name
```
Then add markers to document.xml (see Comments in XML Reference).

### Step 3: Pack
```bash
python scripts/office/pack.py unpacked/ output.docx --original document.docx
```
Validates with auto-repair, condenses XML, and creates DOCX. Use `--validate false` to skip.

**Auto-repair will fix:**
- `durableId` >= 0x7FFFFFFF (regenerates valid ID)
- Missing `xml:space="preserve"` on `<w:t>` with whitespace

**Auto-repair won't fix:**
- Malformed XML, invalid element nesting, missing relationships, schema violations

### Common Pitfalls

- **Replace entire `<w:r>` elements**: When adding tracked changes, replace the whole `<w:r>...</w:r>` block with `<w:del>...<w:ins>...` as siblings. Don't inject tracked change tags inside a run.
- **Preserve `<w:rPr>` formatting**: Copy the original run's `<w:rPr>` block into your tracked change runs to maintain bold, font size, etc.
- **Never regenerate `styles.xml`**: When editing an existing document, all template styles live in `word/styles.xml`. Edit only `document.xml` content; do not overwrite or recreate `styles.xml` unless explicitly asked — doing so erases the user's template.
- **Preserve `<w:pStyle>` references**: When inserting new paragraphs, copy the `<w:pStyle w:val="..."/>` from an adjacent paragraph of the same type. Omitting `<w:pStyle>` silently falls back to the document default style, losing all heading/body formatting from the template.
- **Smart quotes are XML-encoded after unpack**: The unpack step converts `"` / `"` to `&#x201C;` / `&#x201D;` entities. When searching `document.xml` for heading text that contains Chinese quotation marks, search for the entity form (`&#x201C;`), not the raw Unicode character.

---

## XML Reference

### Schema Compliance

- **Element order in `<w:pPr>`**: `<w:pStyle>`, `<w:numPr>`, `<w:spacing>`, `<w:ind>`, `<w:jc>`, `<w:rPr>` last
- **Whitespace**: Add `xml:space="preserve"` to `<w:t>` with leading/trailing spaces
- **RSIDs**: Must be 8-digit hex (e.g., `00AB1234`)

### Tracked Changes

**Insertion:**
```xml
<w:ins w:id="1" w:author="Claude" w:date="2025-01-01T00:00:00Z">
  <w:r><w:t>inserted text</w:t></w:r>
</w:ins>
```

**Deletion:**
```xml
<w:del w:id="2" w:author="Claude" w:date="2025-01-01T00:00:00Z">
  <w:r><w:delText>deleted text</w:delText></w:r>
</w:del>
```

**Inside `<w:del>`**: Use `<w:delText>` instead of `<w:t>`, and `<w:delInstrText>` instead of `<w:instrText>`.

**Minimal edits** - only mark what changes:
```xml
<!-- Change "30 days" to "60 days" -->
<w:r><w:t>The term is </w:t></w:r>
<w:del w:id="1" w:author="Claude" w:date="...">
  <w:r><w:delText>30</w:delText></w:r>
</w:del>
<w:ins w:id="2" w:author="Claude" w:date="...">
  <w:r><w:t>60</w:t></w:r>
</w:ins>
<w:r><w:t> days.</w:t></w:r>
```

**Deleting entire paragraphs/list items** - when removing ALL content from a paragraph, also mark the paragraph mark as deleted so it merges with the next paragraph. Add `<w:del/>` inside `<w:pPr><w:rPr>`:
```xml
<w:p>
  <w:pPr>
    <w:numPr>...</w:numPr>  <!-- list numbering if present -->
    <w:rPr>
      <w:del w:id="1" w:author="Claude" w:date="2025-01-01T00:00:00Z"/>
    </w:rPr>
  </w:pPr>
  <w:del w:id="2" w:author="Claude" w:date="2025-01-01T00:00:00Z">
    <w:r><w:delText>Entire paragraph content being deleted...</w:delText></w:r>
  </w:del>
</w:p>
```
Without the `<w:del/>` in `<w:pPr><w:rPr>`, accepting changes leaves an empty paragraph/list item.

**Rejecting another author's insertion** - nest deletion inside their insertion:
```xml
<w:ins w:author="Jane" w:id="5">
  <w:del w:author="Claude" w:id="10">
    <w:r><w:delText>their inserted text</w:delText></w:r>
  </w:del>
</w:ins>
```

**Restoring another author's deletion** - add insertion after (don't modify their deletion):
```xml
<w:del w:author="Jane" w:id="5">
  <w:r><w:delText>deleted text</w:delText></w:r>
</w:del>
<w:ins w:author="Claude" w:id="10">
  <w:r><w:t>deleted text</w:t></w:r>
</w:ins>
```

### Comments

After running `comment.py` (see Step 2), add markers to document.xml. For replies, use `--parent` flag and nest markers inside the parent's.

**CRITICAL: `<w:commentRangeStart>` and `<w:commentRangeEnd>` are siblings of `<w:r>`, never inside `<w:r>`.**

```xml
<!-- Comment markers are direct children of w:p, never inside w:r -->
<w:commentRangeStart w:id="0"/>
<w:del w:id="1" w:author="Claude" w:date="2025-01-01T00:00:00Z">
  <w:r><w:delText>deleted</w:delText></w:r>
</w:del>
<w:r><w:t> more text</w:t></w:r>
<w:commentRangeEnd w:id="0"/>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>

<!-- Comment 0 with reply 1 nested inside -->
<w:commentRangeStart w:id="0"/>
  <w:commentRangeStart w:id="1"/>
  <w:r><w:t>text</w:t></w:r>
  <w:commentRangeEnd w:id="1"/>
<w:commentRangeEnd w:id="0"/>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="0"/></w:r>
<w:r><w:rPr><w:rStyle w:val="CommentReference"/></w:rPr><w:commentReference w:id="1"/></w:r>
```

### Images

1. Add image file to `word/media/`
2. Add relationship to `word/_rels/document.xml.rels`:
```xml
<Relationship Id="rId5" Type=".../image" Target="media/image1.png"/>
```
3. Add content type to `[Content_Types].xml`:
```xml
<Default Extension="png" ContentType="image/png"/>
```
4. Reference in document.xml:
```xml
<w:drawing>
  <wp:inline>
    <wp:extent cx="914400" cy="914400"/>  <!-- EMUs: 914400 = 1 inch -->
    <a:graphic>
      <a:graphicData uri=".../picture">
        <pic:pic>
          <pic:blipFill><a:blip r:embed="rId5"/></pic:blipFill>
        </pic:pic>
      </a:graphicData>
    </a:graphic>
  </wp:inline>
</w:drawing>
```

---

## Dependencies

- **pandoc**: Text extraction AND LaTeX→OMML formula conversion (`scripts/formula.py` requires pandoc ≥ 2.0)
- **docx**: `npm install docx` (new documents)
- **temml**: `npm install temml` (LaTeX → MathML conversion for Word native math)
- **fast-xml-parser**: `npm install fast-xml-parser` (MathML parsing for docx conversion)
- **LibreOffice**: PDF conversion (auto-configured for sandboxed environments via `scripts/office/soffice.py`)
- **Poppler**: `pdftoppm` for images

---

## Markdown to Word Conversion (Chinese Academic Papers)

This section documents comprehensive solutions for converting Markdown papers with LaTeX formulas, tables, and citations to properly formatted Word documents following Chinese academic standards.

### Quick Start

```bash
# Install dependencies
npm install docx temml fast-xml-parser

# Run conversion
node scripts/new_doc.js
```

### Critical Issues & Solutions

The following 8 issues were identified and solved during production usage. All solutions are implemented in `scripts/new_doc.js` and `scripts/mathml-to-docx.js`.

#### Issue 1: Three-Line Table Middle Borders Visible

**Problem**: Body row borders in 三线表 appeared visible instead of invisible.

**Solution**: Set ALL body row borders to `NONE`, only keep:
- Header top: `THICK` (1.5pt)
- Header bottom: `THIN` (0.75pt)
- Last row bottom: `THICK` (1.5pt)

```javascript
const THICK = { style: BorderStyle.SINGLE, size: 12, color: '000000' };
const THIN  = { style: BorderStyle.SINGLE, size: 6,  color: '000000' };
const NONE  = { style: BorderStyle.NONE,   size: 0,  color: 'FFFFFF' };

// Header row
cellOf(h, colWidths[i], { top: THICK, bottom: THIN, left: NONE, right: NONE }, true)

// Body rows - ALL borders NONE except last row bottom
cellOf(cell, colWidths[i], {
  top: NONE,
  bottom: isLastRow ? THICK : NONE,
  left: NONE,
  right: NONE,
})
```

#### Issue 2: Formula Table Borders Visible

**Problem**: Block formula tables (3-column layout) showed visible borders.

**Solution**: Set ALL borders including `insideHorizontal` and `insideVertical` to `NONE`:

```javascript
return new Table({
  width: { size: CONTENT_W, type: WidthType.DXA },
  columnWidths: [567, 7936, 567],
  borders: {
    top: NONE,
    bottom: NONE,
    left: NONE,
    right: NONE,
    insideHorizontal: NONE,  // CRITICAL: Must include these
    insideVertical: NONE,    // CRITICAL: Must include these
  },
  rows: [new TableRow({ children: [leftCell, formulaCell, numberCell] })],
});
```

#### Issue 3: Heading Numbering — Word Auto-Numbering with Per-Chapter Reset

**Problem**: Originally, heading numbering was baked into paragraph text via manual JS counters. When content was added/removed in Word, numbers did not update.

**Solution**: H2/H3 use Word's native auto-numbering via per-chapter numbering references. H1 keeps Chinese numerals in text (一、二、三). The `buildNumberingConfig(chapterCount)` function generates a `sections_c{N}` numbering reference for each chapter with 2 levels. At each chapter boundary, `_chapter` increments and H2/H3 switch to a fresh numbering reference, achieving automatic reset.

```javascript
let _chapter = 0;

function h1Chinese(text) {  // Chinese numeral in text
  _chapter++;
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    indent: { firstLine: 0 },
    children: [new TextRun(text)],
  });
}

function h2(text) {  // Word auto-numbered (level 0 in per-chapter ref)
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    numbering: { reference: `sections_c${_chapter}`, level: 0 },
    indent: { firstLine: 0 },
    children: [new TextRun(text)],
  });
}

function h3(text) {  // Word auto-numbered (level 1 in per-chapter ref)
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    numbering: { reference: `sections_c${_chapter}`, level: 1 },
    indent: { firstLine: 0 },
    children: [new TextRun(text)],
  });
}
```

#### Issue 4: Heading English/Numbers in Wrong Font (SimSun instead of Cambria Math)

**Problem**: English text and numbers in headings, figure captions, and table captions displayed in SimSun (宋体) instead of Cambria Math.

**Solution**: Use mixed font configuration with `ascii`, `eastAsia`, and `hAnsi` properties:

```javascript
// Style definition for headings
{
  id: 'Heading1', name: 'Heading 1', basedOn: 'Normal',
  run: {
    font: {
      ascii: 'Cambria Math',    // English letters, numbers
      eastAsia: 'SimHei',        // Chinese characters (黑体)
      hAnsi: 'Cambria Math',     // Western European characters
    },
    size: 32, bold: true,
  },
  // ... paragraph settings
}

// For captions (SimSun body font for Chinese)
font: { ascii: 'Cambria Math', eastAsia: 'SimSun', hAnsi: 'Cambria Math' }
```

#### Issue 5: Block Formulas Not Using Word Equation Editor

**Problem**: LaTeX formulas were rendered as plain text or images instead of native Word equations.

**Solution**: Use `temml` (LaTeX→MathML) + `mathml-to-docx.js` (MathML→OMML) pipeline:

```javascript
const temml = require('temml');
const { mathmlToDocxChildren } = require('./mathml-to-docx');

function latexToMath(latex) {
  const mathml = temml.renderToString(latex, { displayMode: true, throwOnError: false });
  const children = mathmlToDocxChildren(mathml);
  if (children && children.length) {
    return new Math({ children });
  }
  // Fallback
  return new Math({ children: [new MathRun(latex)] });
}

// Usage in formula table
const mathObj = latexToMath('Q_n(x, a) = r + \\gamma V_{n-1}(y)');
```

#### Issue 6: Main Title English in Wrong Font

**Problem**: Document title's English text showed in SimSun instead of Cambria Math.

**Solution**: Apply same mixed font to title paragraph:

```javascript
new Paragraph({
  alignment: AlignmentType.CENTER,
  children: [new TextRun({
    text: '论文标题 Paper Title',
    bold: true,
    size: 36,
    font: { ascii: 'Cambria Math', eastAsia: 'SimHei', hAnsi: 'Cambria Math' },
  })],
})
```

#### Issue 7: Inline Math Detection Too Aggressive

**Problem**: Inline math regex matched plain numbers (like "1992") and English words (like "Agent", "Watkins"), incorrectly converting them to formula objects.

**Solution**: Use strict regex that ONLY matches actual mathematical content:

```javascript
// Detection function - returns true only for real math content
function containsMath(text) {
  // Greek letters
  if (/[αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ]/.test(text)) return true;
  // Unicode subscript/superscript characters
  if (/[₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]/.test(text)) return true;
  // Math operators and special symbols
  if (/[∞∑∏∫≤≥≠≈→←↔∈∉⊂⊃∀∃∧∨×÷±∓·…⋯′″⟨⟩]/.test(text)) return true;
  // Starred symbols like π*, Q*
  if (/[A-Z]\*/.test(text)) return true;
  // Explicit $...$ LaTeX
  if (/\$[^$]+\$/.test(text)) return true;
  return false;
}

// Parsing regex - strict matching for inline formulas
const mathPattern = /\$([^$]+)\$|([A-Z][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]+\*?\s*\([^)]+\))|([A-Z]\s*\([^)]*[αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ][^)]*\))|([αβγδεζηθικλμνξπρστυφχψωΓΔΘΛΞΠΣΦΨΩ][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]*\*?)|([A-Za-z][₀₁₂₃₄₅₆₇₈₉ₙₓᵢₜₛ⁰¹²³⁴⁵⁶⁷⁸⁹ⁿⁱ]+\*?)|([A-Z]\*)/g;

// What it matches:
// - $...$ explicit LaTeX
// - Qₙ(x,a) - function with subscripts
// - V(αₙ) - function with Greek params
// - αₙ - Greek with subscript
// - xₙ - variable with subscript
// - Q* - starred variable

// What it does NOT match:
// - Plain numbers: 1992, 500
// - English words: Agent, Watkins, Dayan
// - Plain parentheses: (1992), (optional)
```

#### Issue 8: Citations Not in Superscript

**Problem**: Reference citations like [1], [2] appeared as normal text instead of superscript.

**Solution**: Detect citation pattern and apply `superScript: true`:

```javascript
function containsCitation(text) {
  return /\[\d+\]/.test(text);
}

// In parseInlineContentWithCitations():
const combinedPattern = /(\[\d+\])|...; // Citation first in alternation

if (match[1]) {
  // Citation [n] - convert to superscript
  children.push(new TextRun({
    text: match[1],
    superScript: true,
  }));
}
```

#### Issue 10: Page Break After Abstract and Before References

**Problem**: Abstract page should end after keywords, and references should start on a new page.

**Solution**: Add `pageBreak()` helper function:

```javascript
// ISSUE 10 FIX: Page break helper
function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

// Usage in CONTENT:
// After keywords (end of abstract)
body('关键词：强化学习；Q-learning'),
pageBreak(),  // <- Start new page after abstract

// Before references
pageBreak(),  // <- Start new page for references
new Paragraph({
  heading: HeadingLevel.HEADING_1,
  children: [new TextRun('参考文献')],
}),
```

**Important**: `PageBreak` MUST be wrapped in a `Paragraph` - it cannot be used standalone.

#### Issue 11: Formula Table Cell Vertical Alignment

**Problem**: In the 3-column borderless formula table, the formula and equation number `(n)` are not vertically centered, causing misalignment.

**Solution**: Add `verticalAlign: VerticalAlign.CENTER` to all three TableCell definitions:

```javascript
// First, add VerticalAlign to imports:
const {
  // ... other imports
  VerticalAlign,  // <-- Add this
} = require('docx');

// Then apply to each cell in the formula() function:
const leftCell = new TableCell({
  width: { size: 567, type: WidthType.DXA },
  borders: noBorders,
  shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
  verticalAlign: VerticalAlign.CENTER,  // <-- Add this
  children: [new Paragraph({ indent: { firstLine: 0 }, children: [] })],
});

const formulaCell = new TableCell({
  width: { size: 7936, type: WidthType.DXA },
  borders: noBorders,
  shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
  verticalAlign: VerticalAlign.CENTER,  // <-- Add this
  children: [new Paragraph({
    alignment: AlignmentType.CENTER,
    indent: { firstLine: 0 },
    children: [mathObj],
  })],
});

const numberCell = new TableCell({
  width: { size: 567, type: WidthType.DXA },
  borders: noBorders,
  shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
  verticalAlign: VerticalAlign.CENTER,  // <-- Add this
  children: [new Paragraph({
    alignment: AlignmentType.RIGHT,
    indent: { firstLine: 0 },
    children: [new TextRun(`(${number})`)],
  })],
});
```

**Result**: Formula and equation number now align horizontally on the same baseline.

### File Structure

```
scripts/
├── new_doc.js          # Main template with Issues 1-8, 10-11 fixes (~820 lines)
├── mathml-to-docx.js   # MathML→docx Math converter (~250 lines)
├── formula.py          # Block formula XML insertion (legacy)
├── table.py            # Table XML insertion (legacy)
└── office/
    ├── unpack.py       # DOCX→XML extraction
    ├── pack.py         # XML→DOCX assembly
    └── validate.py     # Document validation
```

### Unicode to LaTeX Mapping

The script includes comprehensive Unicode math symbol conversion:

```javascript
const UNICODE_TO_LATEX = {
  // Greek letters
  'α': '\\alpha', 'β': '\\beta', 'γ': '\\gamma', // ... etc
  // Subscripts
  '₀': '_0', '₁': '_1', '₂': '_2', // ... etc
  // Superscripts
  '²': '^2', '³': '^3', // ... etc
  // Special symbols
  '∞': '\\infty', '∑': '\\sum', '∫': '\\int', // ... etc
};
```

### MathML to DOCX Conversion

`scripts/mathml-to-docx.js` converts MathML (from temml) to docx Math components:

- Fractions: `<mfrac>` → `MathFraction`
- Subscripts: `<msub>` → `MathSubScript`
- Superscripts: `<msup>` → `MathSuperScript`
- Combined: `<msubsup>` → `MathSubSuperScript`
- Radicals: `<msqrt>`, `<mroot>` → `MathRadical`
- Summation: `<munderover>` with ∑ → `MathSum`
- Integrals: `<munderover>` with ∫ → `MathIntegral`
- Matrices: `<mtable>` → `MathMatrix`

### Workflow for New Conversions

1. **Copy `new_doc.js` template** to your project
2. **Install dependencies**: `npm install docx temml fast-xml-parser`
3. **Edit CONTENT section** with your document structure
4. **Use helper functions**:
   - `h1/h2/h3(text)` - Headings with auto-numbering
   - `body(text)` - Body paragraph (auto-detects math/citations)
   - `formula(latex, number)` - Block formula
   - `threeLineTable(headers, rows, colWidths)` - Three-line table
   - `tableCaption/figCaption(label)` - Captions
   - `pageBreak()` - Page break (Issue 10)
   - `ref(text)` - Reference entry
5. **Run**: `node new_doc.js`
