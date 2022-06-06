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
            if (outputElement == null)
            {
                return null;
            }

            var property = parsedPath[^1].property;
            foreach(var element in outputElement.GetAttributes())
            {
                if (property == string.Empty && element.LocalName == "val")
                {
                    return element.Value;
                }
                if (element.LocalName == property)
                {
                    return element.Value;
                }
            }

            return outputElement.InnerText;
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
