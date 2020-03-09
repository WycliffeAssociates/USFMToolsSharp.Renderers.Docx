using Microsoft.VisualStudio.TestTools.UnitTesting;
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
            Assert.AreEqual("Chapter 1", doc.Paragraphs[1].Text);
            // Verse
            Assert.AreEqual("1 Text ", doc.Paragraphs[2].Text);
            // Line break
            Assert.AreEqual("\n", doc.Paragraphs[3].Text);
            // New book: Section break exists at end and has a header
            Assert.IsNotNull(((CT_P)doc.Document.body.Items[3]).pPr.sectPr.headerReference);

            // Header
            Assert.AreEqual("2 John", doc.Paragraphs[4].Text);
            // Chapter
            Assert.AreEqual("Chapter 1", doc.Paragraphs[5].Text);
            // Verse
            Assert.AreEqual("1 Text", doc.Paragraphs[6].Text);
            // Final book: Section break exists at end and has a header
            Assert.IsNotNull(((CT_P)doc.Document.body.Items[8]).pPr.sectPr.headerReference);

        }

        [TestMethod]
        public void TestHeadersCreateSectionsOneBook()
        {
            XWPFDocument doc = renderDoc("\\h 1 John \\c 1 \\v 1 Text");

            // 3 paragraphs: H C V
            Assert.AreEqual(3, doc.Paragraphs.Count);
            // 4 body items: same as above plus one section header
            Assert.AreEqual(4, doc.Document.body.Items.Count);

            // Header
            Assert.AreEqual("1 John", doc.Paragraphs[0].Text);
            // Chapter
            Assert.AreEqual("Chapter 1", doc.Paragraphs[1].Text);
            // Verse
            Assert.AreEqual("1 Text", doc.Paragraphs[2].Text);
            // New book: Section break exists at end and has a header
            Assert.IsNotNull(((CT_P)doc.Document.body.Items[3]).pPr.sectPr.headerReference);

        }

        [TestMethod]
        public void TestHeadersCreateSectionsNoBooks()
        {
            XWPFDocument doc = renderDoc("\\c 1 \\v 1 Text");

            // 2 paragraphs: C V
            Assert.AreEqual(2, doc.Paragraphs.Count);
            // 2 body items: same as above (no section headers)
            Assert.AreEqual(2, doc.Document.body.Items.Count);

            // Chapter
            Assert.AreEqual("Chapter 1", doc.Paragraphs[0].Text);
            // Verse
            Assert.AreEqual("1 Text", doc.Paragraphs[1].Text);

        }

        [TestMethod]
        public void TestChapterRender()
        {
            Assert.AreEqual("Chapter 5", renderDoc("\\c 5").Paragraphs[0].Text);
            Assert.AreEqual("Chapter 1", renderDoc("\\c 1").Paragraphs[0].Text);
            Assert.AreEqual("Chapter -1", renderDoc("\\c -1").Paragraphs[0].Text);
            Assert.AreEqual("Chapter 0", renderDoc("\\c 0").Paragraphs[0].Text);
        }

        [TestMethod]
        public void TestVerseRender()
        {
            Assert.AreEqual("1 This is a simple verse.", renderDoc("\\c 1 \\v 1 This is a simple verse.").Paragraphs[1].ParagraphText);
            Assert.AreEqual("1 This is a simple verse. 2 Another one.", renderDoc("\\c 1 \\v 1 This is a simple verse. \\v 2 Another one.").Paragraphs[1].ParagraphText);
            Assert.AreEqual("2 Another one.", renderDoc("\\c 1 \\v 1 This is a simple verse. \\c 2 \\v 2 Another one.").Paragraphs[3].ParagraphText);
        }

        [TestMethod]
        public void TestSpaceBetweenVerses()
        {
            XWPFDocument doc = renderDoc("\\c 1 \\v 1 First Verse. \\v 2 Second verse.");
            Assert.AreEqual("1 First Verse. 2 Second verse.", doc.Paragraphs[1].ParagraphText);
        }

        [TestMethod]
        public void TestSpaceBetweenVersesInParagraph()
        {
            XWPFDocument doc = renderDoc("\\c 1 \\p \\v 1 First Verse. \\v 2 Second verse.");
            Assert.AreEqual("1 First Verse. 2 Second verse.", doc.Paragraphs[2].ParagraphText);
        }

        [TestMethod]
        public void TestFootnoteRender()
        {
            Assert.AreEqual("1Hello Friend ", renderDoc("\\c 1 \\v 1 This is a simple verse. \\f + \\ft Hello Friend \\f*").Paragraphs[3].ParagraphText);
            Assert.AreEqual("1Hello Fried Friend ", renderDoc("\\c 1 \\v 1 This is a simple verse. \\f + \\ft \\fqa Hello Fried Friend \\f*").Paragraphs[3].ParagraphText);
        }

        [TestMethod]
        public void TestChapterLabelNone()
        {
            string usfm = "\\c 1 \\v 1 First verse. \\v 2 Second verse.";
            XWPFDocument doc = renderDoc(usfm);
            Assert.AreEqual("Chapter 1",doc.Paragraphs[0].Text);
        }

        [TestMethod]
        public void TestChapterLabelDoc()
        {
            string usfm = "\\cl Psalm \\c 1 \\v 1 First verse. \\c 2 \\v 1 First verse.";
            XWPFDocument doc = renderDoc(usfm);
            Assert.AreEqual("Psalm 1",doc.Paragraphs[0].Text);
            Assert.AreEqual("Psalm 2",doc.Paragraphs[2].Text);
        }

        [TestMethod]
        public void TestChapterLabelIndividual()
        {
            string usfm = "\\c 1 \\cl Psalm One \\v 1 First verse. \\c 2 \\v 1 First verse.";
            XWPFDocument doc = renderDoc(usfm);
            Assert.AreEqual("Psalm One",doc.Paragraphs[0].Text);
            Assert.AreEqual("Chapter 2",doc.Paragraphs[2].Text);
        }

        public XWPFDocument renderDoc(string usfm)
        {
            USFMDocument markerTree = parser.ParseFromString(usfm);
            XWPFDocument testDoc = render.Render(markerTree);
            return testDoc;
        }

    }
}
