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

        public int fontSize;

        public DocxConfig()
        {
            textAlign = ParagraphAlignment.LEFT;
            textDirection = ST_TextDirection.lrTb;
            columnCount = 1;
            lineSpacing = 1;
            
            fontSize = 12;
        }
        public DocxConfig(int fontSize, Tuple<ParagraphAlignment, ST_TextDirection, int, double> styles, bool separateChapters = false, bool separateVerses = false)
        {
            textAlign = styles.Item1;
            textDirection = styles.Item2;
            columnCount = styles.Item3;
            lineSpacing = styles.Item4;

            this.fontSize = fontSize;
            this.separateChapters = separateChapters;
            this.separateVerses = separateVerses;
        }
        
    }
}
