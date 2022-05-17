using Microsoft.VisualStudio.TestTools.UnitTesting;
using USFMToolsSharp.Renderers.Docx;
using System;
using System.Collections.Generic;
using System.Text;
using CsvHelper;
using USFMToolsSharp.Renderers.Docx.Tests.Helpers;
using System.IO;
using System.Threading.Tasks;
using System.Linq;
using CsvHelper.Configuration;

namespace USFMToolsSharp.Renderers.Docx.Tests
{
    [TestClass]
    public class OOXmlDocxRendererTests
    {
        [TestMethod]
        public void InitialTest()
        {
            var tests = LoadTestsFromFileAsync("tests.csv");
            foreach (var test in tests)
            {
                Stream stream = new MemoryStream();
                try
                {
                    var parser = new USFMParser();
                    var config = CreateConfig(test);
                    var renderer = new OOXMLDocxRenderer(config);
                    stream = renderer.Render(parser.ParseFromString(test.USFM));
                    var queryEngine = new OOXMLQueryEngine(stream);
                    var tmp = queryEngine.QueryValue(test.Query);
                    if (test.Value != tmp)
                    {
                        var file = File.Create(Path.Join("debugfiles", test.Label + ".docx"));
                        stream.Position = 0;
                        stream.CopyTo(file);
                        file.Close();
                        Assert.Fail($"{test.Label} failed expected {test.Value} got {tmp} writing file for debugging");
                    }
                }
                catch (Exception ex)
                {
                    var file = File.Create(Path.Join("debugfiles", test.Label + ".docx"));
                    stream.Position = 0;
                    stream.CopyTo(file);
                    file.Close();
                    Assert.Fail($"{test.Label} failed {ex.Message} writing file for debugging");
                }
            }
        }
        private List<OOXMLDocxRendererTestRow> LoadTestsFromFileAsync(string fileName)
        {
            using var stream = new StreamReader(fileName);
            using var reader = new CsvReader(stream, new CsvConfiguration(System.Globalization.CultureInfo.InvariantCulture) { HeaderValidated = null, MissingFieldFound = null });
            return reader.GetRecords<OOXMLDocxRendererTestRow>().ToList();
        }
        private DocxConfig CreateConfig(OOXMLDocxRendererTestRow input)
        {
            var config = new DocxConfig();
            if (input.FontSize != null)
            {
                config.fontSize = input.FontSize.Value;
            }
            if (input.RightToLeft != null)
            {
                input.RightToLeft = input.RightToLeft.Value;
            }
            if (input.ColumnCount != null)
            {
                config.columnCount = input.ColumnCount.Value;
            }
            if (input.MarginLeft != null)
            {
                config.marginLeft = input.MarginLeft.Value;
            }
            if (input.MarginRight != null)
            {
                config.marginRight = input.MarginRight.Value;
            }
            if (input.RenderTableOfContents != null)
            {
                config.renderTableOfContents = input.RenderTableOfContents.Value;
            }
            return config;
        }
    }

    public class OOXMLDocxRendererTestRow
    {
        public string USFM { get; set; }
        public string Query { get; set; }
        public string Value { get; set; }
        public string Label { get; set; }
        public string? TextAlign { get; set; }
        public bool? RightToLeft { get; set; }
        public int? ColumnCount { get; set; }
        public int? MarginLeft { get; set; }
        public int? MarginRight { get; set; }
        public double? LineSpacing { get; set; }
        public bool? SeparateChapters { get; set; }
        public bool? SeperateVerses { get; set; }
        public bool? ShowPageNumbers { get; set; }
        public bool? RenderTableOfContents { get; set; }
        public int? FontSize { get; set; }
    }
}
