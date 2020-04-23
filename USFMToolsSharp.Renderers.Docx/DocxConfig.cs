using System;
using System.Collections.Generic;
using System.Text;
using NPOI.XWPF.UserModel;
using NPOI.OpenXmlFormats.Wordprocessing;

namespace USFMToolsSharp.Renderers.Docx
{
    public class DocxConfig
    {

        public ParagraphAlignment textAlign;
        public ST_TextDirection textDirection;
        public int columnCount;
        public double lineSpacing;

        public bool separateChapters;
        public bool separateVerses;
        public bool showPageNumbers;

        public int fontSize;

        public DocxConfig()
        {
            textAlign = ParagraphAlignment.LEFT;
            textDirection = ST_TextDirection.lrTb;
            columnCount = 1;
            lineSpacing = 1;
            showPageNumbers = true;
            fontSize = 12;
        }
        public DocxConfig(int fontSize = 12, bool separateChapters = false, bool separateVerses = false, bool showPageNumbers = true) : this()
        {
            this.fontSize = fontSize;
            this.separateChapters = separateChapters;
            this.separateVerses = separateVerses;
            this.showPageNumbers = showPageNumbers;
        }
        
    }
}
