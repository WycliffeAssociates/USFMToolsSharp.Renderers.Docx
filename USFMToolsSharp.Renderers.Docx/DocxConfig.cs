using NPOI.XWPF.UserModel;

namespace USFMToolsSharp.Renderers.Docx
{
    public class DocxConfig
    {

        public ParagraphAlignment textAlign;
        public bool rightToLeft;
        public int marginLeft; // cm unit
        public int marginRight; // cm unit
        public string rightToLeftLangCode;
        public int columnCount;
        public double lineSpacing;

        public bool separateChapters;
        public bool separateVerses;
        public bool showPageNumbers;
        public bool renderTableOfContents;

        public int fontSize;

        public DocxConfig()
        {
            textAlign = ParagraphAlignment.LEFT;
            rightToLeft = false;
            marginLeft = 0;
            marginRight = 0;
            columnCount = 1;
            lineSpacing = 1;
            showPageNumbers = true;
            fontSize = 12;
            renderTableOfContents = false;
        }
        public DocxConfig(
            int fontSize = 12,
            bool separateChapters = false, 
            bool separateVerses = false, 
            bool showPageNumbers = true
        ) : this()
        {
            this.fontSize = fontSize;
            this.separateChapters = separateChapters;
            this.separateVerses = separateVerses;
            this.showPageNumbers = showPageNumbers;
        }
        
    }
}
