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

There a couple useful classes that you'll want to use

## DocxRenderer

This class transforms a USFMDocument into a XWPFDocument

Example:
```csharp
var contents = File.ReadAllText("01-GEN.usfm");
USFMDocument document = parser.ParseFromString(contents);
DocxRenderer docxRenderer = new DocxRenderer();
XWPFDocument docxOutput = docxRenderer.Render(document);
using (var fs = new FileStream(FilePath, FileMode.Create, FileAccess.Write))
{
    docxOutput.Write(fs);
}
```
