using System;
using System.Collections.Generic;
using System.Text;

namespace USFMToolsSharp.Renderers.Docx
{
    public class StyleConfig : ICloneable
    {
        public int fontSize = 14;
        public bool isBold = false;
        public bool isItalics = false;
        public bool isAlignRight = false;
        public bool isSmallCaps = false;

        public object Clone()
        {
            return this.MemberwiseClone();
        }
    }
}
