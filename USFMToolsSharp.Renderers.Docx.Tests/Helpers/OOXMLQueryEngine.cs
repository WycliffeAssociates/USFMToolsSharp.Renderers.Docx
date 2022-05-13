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
            if (parsedPath[0].element == "document")
            {
                var elements = document.ChildElements;
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
                        return null;
                    }
                }
            }

            switch (outputElement)
            {
                case Text text:
                    return text.InnerText;
                case FontSize size:
                    return size.Val;
                case VerticalTextAlignment alignment:
                    return alignment.Val;
                case BiDi bidi:
                    return bidi.Val;
                case Columns columns:
                    return columns.ColumnCount;
                case PageNumberType pageNumberType:
                    return pageNumberType.Format;
                case SpacingBetweenLines spacing:
                    return $"{spacing.Line}:{spacing.After}";
                case Indentation indentation:
                    return $"{indentation.Left}:{indentation.Right}";

            }
            return null;

        }
        public List<(string element, int count, string property)> ParseQuery(string input)
        {
            // Capture groups from the following regex: operator[level].property
            var regex = new Regex(@"(?<operator>[^\[]+)(?<level>\[\d+\])?.?(?<property>[^\.]+)?");
            var splitInput = input.Split("/");
            var output = new List<(string element, int count, string property)>(splitInput.Length);
            foreach (var item in splitInput)
            {
                var match = regex.Match(item);
                var @operator = match.Groups.GetValueOrDefault("operator").Value;
                var levelAsString = match.Groups.GetValueOrDefault("level").Value;
                if (levelAsString == null)
                {
                    levelAsString = "0";
                }
                levelAsString = levelAsString.TrimStart('[').TrimEnd(']');
                var level = 0;
                if (!int.TryParse(levelAsString, out level))
                {
                    level = 0;
                }
                output.Add((@operator, level, null));
            }
            return output;
        }
    }
}
