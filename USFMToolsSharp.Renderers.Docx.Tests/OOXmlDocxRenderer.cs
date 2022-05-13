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
                    var renderer = new OOXMLDocxRenderer();
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
            using var reader = new CsvReader(stream, System.Globalization.CultureInfo.InvariantCulture);
            return reader.GetRecords<OOXMLDocxRendererTestRow>().ToList();
        }
    }

    public class OOXMLDocxRendererTestRow
    {
        public string USFM { get; set; }
        public string Query { get; set; }
        public string Value { get; set; }
        public string Label { get; set; }
    }
}
