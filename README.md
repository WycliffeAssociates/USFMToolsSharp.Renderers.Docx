![GitHub](https://img.shields.io/github/license/WycliffeAssociates/USFMToolsSharp.Renderers.Docx?color=blue)
![Travis (.com) branch](https://img.shields.io/travis/com/WycliffeAssociates/USFMToolsSharp.Renderers.Docx/master)
![Nuget](https://img.shields.io/nuget/v/USFMToolsSharp.Renderers.Docx?color=blue)
![Nuget](https://img.shields.io/nuget/dt/USFMToolsSharp.Renderers.Docx?color=blue)

# USFMToolsSharp.Renderers.Docx
A .net Docx rendering tool for USFM.

# Description
USFMToolsSharp.Renderers.Docx is a Docx renderer for USFM. 

# Installation

You can install this package from nuget https://www.nuget.org/packages/USFMToolsSharp.Renderers.Docx/

# Requirements

We targeted .net standard 2.0 so .net core 2.0, .net framework 4.6.1, and mono 5.4 and
higher are the bare minimum.

# Building

With Visual Studio just build the solution. With the .net core tooling use `dotnet build`

# Dependencies

[WycliffeAssociates.NPOI](https://www.nuget.org/packages/WycliffeAssociates.NPOI/)

# Contributing

Yes please! A couple things would be very helpful

- Testing: Because I can't test every single possible USFM document in existance. If you find something that doesn't look right in the parsing or rendering please submit an issue.
- Adding support for other markers to the parser. There are still plenty of things in the USFM spec that aren't implemented.
- Adding support for other markers to the DOCX renderer

# Usage

There are two main renderer classes that you can use:

## OOXMLDocxRenderer (Recommended)

This is the **preferred** renderer class. It transforms a USFMDocument into a Stream using OpenXML and is actively maintained.

### Basic Example:
```csharp
using USFMToolsSharp;
using USFMToolsSharp.Renderers.Docx;

var parser = new USFMParser();
var contents = File.ReadAllText("01-GEN.usfm");
USFMDocument document = parser.ParseFromString(contents);

OOXMLDocxRenderer renderer = new OOXMLDocxRenderer();
Stream docxStream = renderer.Render(document);

using (var fs = new FileStream("output.docx", FileMode.Create, FileAccess.Write))
{
    docxStream.Position = 0;
    docxStream.CopyTo(fs);
}
```

### Using Configuration:
```csharp
var config = new DocxConfig
{
    fontSize = 14,
    separateChapters = false,
    showPageNumbers = true,
    renderTableOfContents = true
};

OOXMLDocxRenderer renderer = new OOXMLDocxRenderer(config);
renderer.FrontMatter = frontMatterDoc;  // Optional
Stream docxStream = renderer.Render(document);
```

### Adding Front Matter:
```csharp
var frontMatterParser = new USFMParser();
var frontMatterContents = File.ReadAllText("front-matter.usfm");
USFMDocument frontMatterDoc = frontMatterParser.ParseFromString(frontMatterContents);

var config = new DocxConfig
{
    fontSize = 14,
    renderTableOfContents = true
};

OOXMLDocxRenderer renderer = new OOXMLDocxRenderer(config);
renderer.FrontMatter = frontMatterDoc;
Stream docxStream = renderer.Render(document);

using (var fs = new FileStream("output.docx", FileMode.Create, FileAccess.Write))
{
    docxStream.Position = 0;
    docxStream.CopyTo(fs);
}
```

## DocxRenderer (Obsolete)

**Note:** This renderer is obsolete. Please use `OOXMLDocxRenderer` for new projects.

This class transforms a USFMDocument into a XWPFDocument (using NPOI).

### Basic Example:
```csharp
using USFMToolsSharp;
using USFMToolsSharp.Renderers.Docx;

var parser = new USFMParser();
var contents = File.ReadAllText("01-GEN.usfm");
USFMDocument document = parser.ParseFromString(contents);

DocxRenderer docxRenderer = new DocxRenderer();
XWPFDocument docxOutput = docxRenderer.Render(document);

using (var fs = new FileStream("output.docx", FileMode.Create, FileAccess.Write))
{
    docxOutput.Write(fs);
}
```

### Using DocxConfig:
```csharp
var config = new DocxConfig
{
    fontSize = 12,
    separateChapters = true,
    separateVerses = false,
    showPageNumbers = true,
    renderTableOfContents = false,
    columnCount = 2,
    lineSpacing = 1.5,
    textAlign = TextAlignment.LEFT,
    rightToLeft = false,
    marginLeft = 2,  // in CM
    marginRight = 2  // in CM
};

DocxRenderer docxRenderer = new DocxRenderer(config);
XWPFDocument docxOutput = docxRenderer.Render(document);
```

### Adding Front Matter:
```csharp
var frontMatterParser = new USFMParser();
var frontMatterContents = File.ReadAllText("front-matter.usfm");
USFMDocument frontMatterDoc = frontMatterParser.ParseFromString(frontMatterContents);

DocxRenderer docxRenderer = new DocxRenderer(config);
docxRenderer.FrontMatter = frontMatterDoc;
XWPFDocument docxOutput = docxRenderer.Render(document);
```

## Configuration Options

### DocxConfig Properties

| Property | Type | Default | Description |
|----------|------|---------|-------------|
| `fontSize` | int | 12 | Base font size for the document |
| `textAlign` | TextAlignment | LEFT | Text alignment (see TextAlignment enum below) |
| `rightToLeft` | bool | false | Enable right-to-left text direction |
| `rightToLeftLangCode` | string | null | Language code for RTL text |
| `marginLeft` | int | 0 | Left margin in centimeters |
| `marginRight` | int | 0 | Right margin in centimeters |
| `columnCount` | int | 1 | Number of columns for text layout |
| `lineSpacing` | double | 1 | Line spacing multiplier (1.5 = 1.5x spacing) |
| `separateChapters` | bool | false | Add page breaks between chapters |
| `separateVerses` | bool | false | Add line breaks between verses |
| `showPageNumbers` | bool | true | Display page numbers in headers |
| `renderTableOfContents` | bool | false | Generate a table of contents |

### TextAlignment Enum Values

```csharp
public enum TextAlignment
{
    LEFT = 1,
    CENTER = 2,
    RIGHT = 3,
    BOTH = 4,              // Justified
    MEDIUM_KASHIDA = 5,    // Arabic justification
    DISTRIBUTE = 6,
    NUM_TAB = 7,
    HIGH_KASHIDA = 8,      // Arabic justification (high)
    LOW_KASHIDA = 9,       // Arabic justification (low)
    THAI_DISTRIBUTE = 10   // Thai justification
}
```

### StyleConfig

StyleConfig is used internally for text styling but can be referenced when understanding the rendering output:

```csharp
public class StyleConfig
{
    public int fontSize = 14;
    public bool isBold = false;
    public bool isItalics = false;
    public bool isAlignRight = false;
    public bool isSmallCaps = false;
}
```

## Advanced Examples

### Multi-Column Layout with Table of Contents:
```csharp
var config = new DocxConfig
{
    fontSize = 11,
    columnCount = 2,
    renderTableOfContents = true,
    showPageNumbers = true,
    lineSpacing = 1.2
};

DocxRenderer renderer = new DocxRenderer(config);
XWPFDocument docxOutput = renderer.Render(document);
```

### Right-to-Left Languages (e.g., Arabic, Hebrew):
```csharp
var config = new DocxConfig
{
    rightToLeft = true,
    textAlign = TextAlignment.RIGHT,
    rightToLeftLangCode = "ar",
    fontSize = 14
};

DocxRenderer renderer = new DocxRenderer(config);
XWPFDocument docxOutput = renderer.Render(document);
```

### Print-Ready Format with Margins and Chapters:
```csharp
var config = new DocxConfig
{
    fontSize = 12,
    marginLeft = 3,
    marginRight = 3,
    separateChapters = true,
    showPageNumbers = true,
    columnCount = 1,
    lineSpacing = 1.5
};

DocxRenderer renderer = new DocxRenderer(config);
XWPFDocument docxOutput = renderer.Render(document);
```
