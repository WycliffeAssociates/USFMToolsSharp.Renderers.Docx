using System;
using System.Collections.Generic;
using System.Text;

namespace USFMToolsSharp.Renderers.Docx
{
    public class DocxConfig
    {
        public bool separateChapters;
        public bool separateVerses;
        public int fontSize;

        public DocxConfig()
        {
            separateChapters = false;
            separateVerses = true;
            fontSize = 12;
        }
        public DocxConfig(int fontSize, bool separateChapters = false, bool separateVerses = false)
        {
            this.fontSize = fontSize;
            this.separateChapters = separateChapters;
            this.separateVerses = separateVerses;
        }
    }
}
