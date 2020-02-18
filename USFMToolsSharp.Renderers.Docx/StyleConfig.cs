using System;
using System.Collections.Generic;
using System.Text;
using NPOI.XWPF.UserModel;

namespace USFMToolsSharp.Renderers.Docx
{
    public class StyleConfig : ICloneable
    {
        public int fontSize = 14;
        public bool isBold = false;
        public bool isItalics = false;
        public ParagraphAlignment Alignment = ParagraphAlignment.LEFT;
        public bool isSmallCaps = false;
        public int indentationLevel = 0;

        public object Clone()
        {
            return this.MemberwiseClone();
        }
    }
}
