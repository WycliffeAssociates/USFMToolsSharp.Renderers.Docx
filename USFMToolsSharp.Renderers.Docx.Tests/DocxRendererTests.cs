﻿using Microsoft.VisualStudio.TestTools.UnitTesting;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;
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
        }
        
        [TestMethod]
        public void TestHeaderRenderTwoWords()
        {
            Assert.AreEqual("1 John", renderDoc("\\h 1 John").Paragraphs[0].Text);
        }
        
        [TestMethod]
        public void TestHeaderRenderBlank()
        {
            Assert.AreEqual("", renderDoc("\\h      ").Paragraphs[0].Text);
        }

        [TestMethod]
        public void TestHeadersCreateSections()
        {
            XWPFDocument doc = renderDoc("\\h 1 John \\c 1 \\v 1 Text \\h 2 John \\c 1 \\v 1 Text");

            // 7 paragraphs: H C V (pagebreak) H C V
            Assert.AreEqual(7, doc.Paragraphs.Count);
            // 9 body items: same as above plus two section headers
            Assert.AreEqual(9, doc.Document.body.Items.Count);

            // Header
            Assert.AreEqual("1 John", doc.Paragraphs[0].Text);
            // Chapter
            Assert.AreEqual("1", doc.Paragraphs[1].Text);
            // Verse
            Assert.AreEqual("1Text", doc.Paragraphs[2].Text);
            // Line break
            Assert.AreEqual("\n", doc.Paragraphs[3].Text);
            // New book: Section break exists at end and has a header
            Assert.IsNotNull(((CT_P)doc.Document.body.Items[3]).pPr.sectPr.headerReference);

            // Header
            Assert.AreEqual("2 John", doc.Paragraphs[4].Text);
            // Chapter
            Assert.AreEqual("1", doc.Paragraphs[5].Text);
            // Verse
            Assert.AreEqual("1Text", doc.Paragraphs[6].Text);
            // Final book: Section break exists at end and has a header
            Assert.IsNotNull(((CT_P)doc.Document.body.Items[8]).pPr.sectPr.headerReference);

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
            Assert.AreEqual("1This is a simple verse.2Another one.", renderDoc("\\c 1 \\v 1 This is a simple verse. \\v 2 Another one.").Paragraphs[1].ParagraphText);
            Assert.AreEqual("2Another one.", renderDoc("\\c 1 \\v 1 This is a simple verse. \\c 2 \\v 2 Another one.").Paragraphs[3].ParagraphText);
        }

        [TestMethod]
        public void TestFootnoteRender()
        {
            Assert.AreEqual("1Hello Friend", renderDoc("\\c 1 \\v 1 This is a simple verse. \\f + \\ft Hello Friend \\f*").Paragraphs[3].ParagraphText);
            Assert.AreEqual("1Hello Fried Friend", renderDoc("\\c 1 \\v 1 This is a simple verse. \\f + \\ft \\fqa Hello Fried Friend \\f*").Paragraphs[3].ParagraphText);
        }

        public XWPFDocument renderDoc(string usfm)
        {
            USFMDocument markerTree = parser.ParseFromString(usfm);
            XWPFDocument testDoc = render.Render(markerTree);
            return testDoc;
        }

    }
}