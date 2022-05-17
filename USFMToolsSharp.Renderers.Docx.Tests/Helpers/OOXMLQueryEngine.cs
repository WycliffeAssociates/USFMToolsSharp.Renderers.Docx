using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;

namespace USFMToolsSharp.Renderers.Docx.Tests.Helpers
{
    internal class OOXMLQueryEngine
    {
        public WordprocessingDocument WPDocument { get; }
        public OOXMLQueryEngine(Stream contentStream)
        {
            contentStream.Position = 0;
            this.WPDocument = WordprocessingDocument.Open(contentStream, false);
        }
        public string QueryValue(string path)
        {
            var mainPart = WPDocument.MainDocumentPart;
            var document = mainPart.Document;
            var parsedPath = ParseQuery(path);
            OpenXmlElement outputElement = null;
            OpenXmlElementList elements;
            if (parsedPath[0].element == "document")
            {
                elements = document.ChildElements;
            }
            else if (parsedPath[0].element == "footnotes")
            {
                elements = mainPart.FootnotesPart.Footnotes.ChildElements;
            }
            else if (parsedPath[0].element == "settings")
            {
                elements = mainPart.DocumentSettingsPart.Settings.ChildElements;
            }
            else if (parsedPath[0].element == "styles")
            {
                elements = mainPart.StyleDefinitionsPart.Styles.ChildElements;
            }
            else
            {
                return null;
            }

            foreach (var pathItem in parsedPath.Skip(1))
            {
                var count = 0;
                bool found = false;
                foreach(var element in elements) 
                {
                    if (element.LocalName == pathItem.element)
                    {
                        if (count == pathItem.count)
                        {
                            elements = element.ChildElements;
                            outputElement = element;
                            found = true;
                            break;
                        }
                        count++;
                    }
                }
                if (!found)
                {
                    Console.WriteLine("Didn't find element");
                    return null;
                }
            }

            var property = parsedPath[^1].property;

            switch (outputElement)
            {
                case Text text:
                    return text.InnerText;
                case FontSize size:
                    return size.Val;
                case VerticalTextAlignment alignment:
                    if (alignment.Val == VerticalPositionValues.Superscript)
                    {
                        return "superscript";
                    }
                    if (alignment.Val == VerticalPositionValues.Subscript)
                    {
                        return "subscript";
                    }
                    if (alignment.Val == VerticalPositionValues.Baseline)
                    {
                        return "baseline";
                    }
                    return alignment.Val;
                case BiDi bidi:
                    return bidi.Val;
                case Columns columns:
                    return columns.ColumnCount;
                case PageNumberType pageNumberType:
                    if (property == "format")
                    {
                        return pageNumberType.Format;
                    }
                    if (property == "chapterSeperator") 
                    {
                        return pageNumberType.ChapterSeparator;
                    }
                    return pageNumberType.Format;
                case SpacingBetweenLines spacing:
                    if (property == "before")
                    {
                        return spacing.Before;
                    }
                    else if (property == "after")
                    {
                        return spacing.After;
                    }
                    else if (property == "line")
                    {
                        return spacing.After;
                    }
                    return $"{spacing.Line}:{spacing.After}";
                case Indentation indentation:
                    if (property == "left")
                    {
                        return indentation.Left;
                    }
                    else if (property == "right")
                    {
                        return indentation.Right;
                    }
                    return $"{indentation.Left}:{indentation.Right}";
                case UpdateFieldsOnOpen updateFieldsOnOpen:
                    return updateFieldsOnOpen.Val;
                case Footnote footnote:
                    if (property == "id")
                    {
                        return footnote.Id;
                    }
                    if (property == "type")
                    {
                        return footnote.Type;
                    }
                    return footnote.Id;
                case FootnoteReference footnoteReference:
                    return footnoteReference.Id;

            }
            return null;

        }
        public List<(string element, int count, string property)> ParseQuery(string input)
        {
            // Capture groups from the following regex: operator[level].property
            var regex = new Regex(@"(?<operator>[^\[\.]+)(?<level>\[\d+\])?\.?(?<sub>[^\.]+)?");
            var splitInput = input.Split("/");
            var output = new List<(string element, int count, string property)>(splitInput.Length);
            foreach (var item in splitInput)
            {
                var match = regex.Match(item);
                var @operator = match.Groups.GetValueOrDefault("operator").Value;
                var levelAsString = match.Groups.GetValueOrDefault("level").Value;
                var property = match.Groups.GetValueOrDefault("sub").Value;
                if (levelAsString == null)
                {
                    levelAsString = "0";
                }
                levelAsString = levelAsString.TrimStart('[').TrimEnd(']');
                if (!int.TryParse(levelAsString, out int level))
                {
                    level = 0;
                }
                output.Add((@operator, level, property));
            }
            return output;
        }
    }
}
