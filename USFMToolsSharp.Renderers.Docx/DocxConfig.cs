using NPOI.XWPF.UserModel;

namespace USFMToolsSharp.Renderers.Docx
{
    public class DocxConfig
    {

        public TextAlignment textAlign;
        public bool rightToLeft;
        /// <summary>
        /// Left margin in CM
        /// </summary>
        public int marginLeft;
        /// <summary>
        /// Right margin in CM
        /// </summary>
        public int marginRight;
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
            textAlign = TextAlignment.LEFT;
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
    public enum TextAlignment
    {
        LEFT = 1,
        CENTER = 2,
        RIGHT = 3,
        BOTH = 4,
        MEDIUM_KASHIDA = 5,
        DISTRIBUTE = 6,
        NUM_TAB = 7,
        HIGH_KASHIDA = 8,
        LOW_KASHIDA = 9,
        THAI_DISTRIBUTE = 10
    }
}
