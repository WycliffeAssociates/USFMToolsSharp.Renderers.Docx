using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.XWPF.UserModel;
using System;
using System.Collections.Generic;
using System.Text;
using USFMToolsSharp.Models.Markers;

namespace USFMToolsSharp.Renderers.Docx.Tests
{
    [TestClass]
    public class DocxRendererTests
    {
        private USFMParser parser;
        private DocxRenderer render;
        private DocxConfig configDocx;

        [TestInitialize]
        public void SetUpTestCase()
        {
            configDocx = new DocxConfig();
            parser = new USFMParser();
            render = new DocxRenderer();
        }
        [TestMethod]
        public void TestHeaderRender()
        {
            Assert.AreEqual("Genesis",renderDoc("\\h Genesis").Paragraphs[0].Text);
            Assert.AreEqual("1 John", renderDoc("\\h 1 John").Paragraphs[0].Text);
            Assert.AreEqual("", renderDoc("\\h      ").Paragraphs[0].Text);

        }
        [TestMethod]
        public void TestChapterRender()
        {
            Assert.AreEqual("5", renderDoc("\\c 5").Paragraphs[0].Text);
            Assert.AreEqual("1", renderDoc("\\c 1").Paragraphs[0].Text);
            Assert.AreEqual("-1", renderDoc("\\c -1").Paragraphs[0].Text);
            Assert.AreEqual("0", renderDoc("\\c 0").Paragraphs[0].Text);
        }
        [TestMethod]
        public void TestVerseRender()
        {
            Assert.AreEqual("1This is a simple verse.", renderDoc("\\c 1 \\v 1 This is a simple verse.").Paragraphs[1].ParagraphText);
            Assert.AreEqual("1This is a simple verse.2verse 2", renderDoc("\\c 1 \\v 1 This is a simple verse. \\v 2 verse 2").Paragraphs[1].ParagraphText);
            Assert.AreEqual("2verse 2", renderDoc("\\c 1 \\v 1 This is a simple verse. \\c 2 \\v 2 verse 2").Paragraphs[3].ParagraphText);
        }
        [TestMethod]
        public void TestFootnoteRender()
        {
            Assert.AreEqual("Footnotes", renderDoc("\\c 1 \\v 1 This is a simple verse. \\f + \\ft Hello Friend \\f*").Paragraphs[2].ParagraphText);
            Assert.AreEqual("1Hello Friend", renderDoc("\\c 1 \\v 1 This is a simple verse. \\f + \\ft Hello Friend \\f*").Paragraphs[3].ParagraphText);
            Assert.AreEqual("1Hello Fried Friend", renderDoc("\\c 1 \\v 1 This is a simple verse. \\f + \\ft \\fqa Hello Fried Friend \\f*").Paragraphs[3].ParagraphText);
        }

        public XWPFDocument renderDoc(string usfm)
        {
            render.clearDocumentElements();
            USFMDocument markerTree = parser.ParseFromString(usfm);
            XWPFDocument testDoc = render.Render(markerTree);
            return testDoc;
        }

    }
}
